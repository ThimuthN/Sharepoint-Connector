"""Profile setup and target-binding commands for the CLI."""
import argparse
import sys
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import unquote, urlparse


def resolve_user_email(user_info: dict) -> str:
    """Resolve best available user email identifier from Graph /me payload."""
    return (
        user_info.get("mail")
        or user_info.get("userPrincipalName")
        or "unknown"
    )


def save_profile(
    profile_name: str,
    store_dir: str,
    tokens: dict,
    user_info: dict,
    client_id: str,
    tenant_id: str,
    token_store_cls,
) -> str:
    """Persist token/user data in the standard profile format."""
    store = token_store_cls(store_dir=store_dir)
    user_email = resolve_user_email(user_info)
    profile_data = {
        "access_token": tokens["access_token"],
        "refresh_token": tokens.get("refresh_token", ""),
        "expires_at": (
            datetime.utcnow() + timedelta(seconds=tokens.get("expires_in", 3600))
        ).isoformat(),
        "site_id": "",
        "site_name": "",
        "drive_id": "",
        "drive_name": "",
        "folder_id": "",
        "folder_path": "",
        "user_id": user_info.get("id", ""),
        "user_email": user_email,
        "client_id": client_id,
        "tenant_id": tenant_id,
    }
    store.save_profile(profile_name, profile_data)
    return user_email


def normalize_name(value: str) -> str:
    """Normalize text for case/spacing-insensitive matching."""
    return "".join(ch.lower() for ch in value if ch.isalnum())


def parse_sharepoint_url(sharepoint_url: str) -> dict:
    """Parse a SharePoint browser URL into host/site/library/folder hints."""
    parsed = urlparse(sharepoint_url)
    if not parsed.scheme.startswith("http") or not parsed.hostname:
        raise ValueError("Invalid SharePoint URL.")

    decoded_path = unquote(parsed.path or "")
    for marker in ("/:f:/r", "/:u:/r", "/:x:/r", "/:w:/r", "/:b:/r"):
        if marker in decoded_path:
            decoded_path = decoded_path.split(marker, 1)[1]
            break

    if not decoded_path.startswith("/"):
        decoded_path = "/" + decoded_path

    segments = [seg for seg in decoded_path.split("/") if seg]
    if len(segments) < 2 or segments[0] not in ("sites", "teams"):
        raise ValueError(
            "Could not parse site path from URL. Expected /sites/<name> or /teams/<name>."
        )

    site_path = f"/{segments[0]}/{segments[1]}"
    library_name = segments[2] if len(segments) >= 3 else "Documents"
    folder_path = "/".join(segments[3:]) if len(segments) >= 4 else ""

    return {
        "hostname": parsed.hostname,
        "site_path": site_path,
        "library_name": library_name,
        "folder_path": folder_path,
    }


def select_drive(drives: list, requested_library: str) -> dict:
    """Select best matching document library drive."""
    requested = normalize_name(requested_library)
    aliases = {requested}
    if requested == "shareddocuments":
        aliases.add("documents")
    if requested == "documents":
        aliases.add("shareddocuments")

    for drive in drives:
        candidates = {normalize_name(drive.get("name", ""))}
        web_url = drive.get("webUrl", "")
        if web_url:
            tail = unquote(urlparse(web_url).path.split("/")[-1])
            candidates.add(normalize_name(tail))
        if aliases.intersection(candidates):
            return drive

    raise ValueError(
        f"Library '{requested_library}' not found. Available libraries: "
        f"{', '.join(d.get('name', '<unknown>') for d in drives)}"
    )


def ensure_profile_token(
    profile_name: str,
    profile_data: dict,
    store,
    auth_cls,
) -> dict:
    """Refresh profile access token if required and persist updates."""
    expires_at = datetime.fromisoformat(profile_data["expires_at"])
    auth = auth_cls(
        client_id=profile_data.get("client_id"),
        tenant_id=profile_data.get("tenant_id"),
    )
    if auth.is_token_expired(expires_at):
        token_response = auth.refresh_token(profile_data["refresh_token"])
        profile_data["access_token"] = token_response["access_token"]
        profile_data["refresh_token"] = token_response.get(
            "refresh_token", profile_data["refresh_token"]
        )
        profile_data["expires_at"] = (
            datetime.utcnow() + timedelta(seconds=token_response.get("expires_in", 3600))
        ).isoformat()
        store.save_profile(profile_name, profile_data)
    return profile_data


def cmd_set_target(
    args,
    token_store_cls,
    graph_client_cls,
    auth_cls,
    parse_sharepoint_url_fn,
    select_drive_fn,
    ensure_profile_token_fn,
    print_actionable_error,
) -> None:
    """Bind configured profile to a SharePoint site/library/folder target."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir
    sharepoint_url = args.sharepoint_url
    my_drive = bool(getattr(args, "my_drive", False))
    library_override = args.library
    folder_override = args.folder
    graph = None

    try:
        store = token_store_cls(store_dir=store_dir)
        profile_data = store.load_profile(profile_name)
        if not profile_data:
            print(
                f"Profile '{profile_name}' not found. "
                f"Run: python -m rpa_sharepoint_connector configure --profile {profile_name}"
            )
            sys.exit(1)

        profile_data = ensure_profile_token_fn(profile_name, profile_data, store, auth_cls)

        graph = graph_client_cls(profile_data["access_token"])
        if my_drive:
            if sharepoint_url:
                raise ValueError("Use either --my-drive or --sharepoint-url, not both.")
            drive = graph._get("/me/drive")
            site_id = "me"
            site_name = "My Drive"
            drive_name = drive.get("name") or "OneDrive"
            folder_path = folder_override or ""
        else:
            if not sharepoint_url:
                raise ValueError("Provide --sharepoint-url or use --my-drive.")
            parsed = parse_sharepoint_url_fn(sharepoint_url)
            hostname = parsed["hostname"]
            site_path = parsed["site_path"]
            library_name = library_override or parsed["library_name"] or "Documents"
            folder_path = (
                folder_override if folder_override is not None else parsed["folder_path"]
            )

            site = graph._get(f"/sites/{hostname}:{site_path}")
            drives = graph.list_drives(site["id"])
            drive = select_drive_fn(drives, library_name)
            site_id = site["id"]
            site_name = site.get("displayName", site_path)
            drive_name = drive.get("name", library_name)

        folder_id = "root"
        if folder_path:
            folder_id = graph._ensure_folder_path(drive["id"], folder_path)

        profile_data.update(
            {
                "site_id": site_id,
                "site_name": site_name,
                "drive_id": drive["id"],
                "drive_name": drive_name,
                "folder_id": folder_id,
                "folder_path": folder_path,
            }
        )
        store.save_profile(profile_name, profile_data)

        print(f"OK Target configured for profile '{profile_name}'")
        print(f"Site: {profile_data['site_name']}")
        print(f"Library: {profile_data['drive_name']}")
        print(f"Folder: {profile_data['folder_path'] or '(root)'}")
        print()
    except Exception as error:
        print_actionable_error(error, command="set-target", profile=profile_name)
    finally:
        close = getattr(graph, "close", None)
        if callable(close):
            close()


def cmd_configure(
    args,
    token_store_cls,
    browser_auth_cls,
    save_profile_fn,
    print_actionable_error,
) -> None:
    """Configure a new profile using browser OAuth (Authorization Code + PKCE)."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir
    force = bool(getattr(args, "force", False))
    redirect_uri = getattr(args, "redirect_uri", None)
    client_id = getattr(args, "client_id", None)
    tenant_id = getattr(args, "tenant_id", None)

    print(f"\nConfiguring profile: {profile_name}")
    print("=" * 60)

    try:
        store = token_store_cls(store_dir=store_dir)
        existing_profile = store.load_profile(profile_name)
        if existing_profile and not force:
            print(
                f"ERROR: Profile '{profile_name}' already exists. "
                f"Run configure again with --force to replace it."
            )
            sys.exit(1)

        auth = browser_auth_cls(
            client_id=client_id,
            tenant_id=tenant_id,
            redirect_uri=redirect_uri,
        )
        request = auth.build_authorization_request()

        print("\nMicrosoft Browser Login Required")
        print("-" * 60)
        print(f"Client ID: {auth.client_id}")
        print(f"Tenant: {auth.tenant_id}")
        print(f"Go to: {request['authorization_url']}")
        print(f"Redirect URI: {request['redirect_uri']}")
        print("\nOpening browser for Microsoft sign-in...")
        print("Waiting for callback on localhost...")
        print("-" * 60)

        result = auth.authenticate(
            open_browser=True,
            authorization_request=request,
        )
        tokens = result["tokens"]
        user_info = result["user_info"]

        print("OK Authorization successful!")
        user_email = resolve_user_email(user_info)
        print(f"OK Logged in as: {user_email}")

        save_profile_fn(
            profile_name=profile_name,
            store_dir=store_dir,
            tokens=tokens,
            user_info=user_info,
            client_id=auth.client_id,
            tenant_id=auth.tenant_id,
            token_store_cls=token_store_cls,
        )
        print(f"OK Profile '{profile_name}' saved")
        print()

    except KeyboardInterrupt:
        print("\n\nCancelled")
        sys.exit(0)
    except Exception as error:
        print_actionable_error(error, command="configure", profile=profile_name)


def cmd_setup(
    args,
    configure_cmd,
    set_target_cmd,
    test_upload_cmd,
    unlink_if_exists,
) -> None:
    """Run one-command onboarding: configure, bind target, and smoke test."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir
    force = bool(getattr(args, "force", False))
    redirect_uri = getattr(args, "redirect_uri", None)
    client_id = getattr(args, "client_id", None)
    tenant_id = getattr(args, "tenant_id", None)
    sharepoint_url = getattr(args, "sharepoint_url", None)
    my_drive = bool(getattr(args, "my_drive", False))
    library = getattr(args, "library", None)
    folder = getattr(args, "folder", None)
    skip_smoke_test = bool(getattr(args, "skip_smoke_test", False))

    if not sharepoint_url and not my_drive:
        my_drive = True
        print("No target provided. Defaulting to --my-drive.")

    print(f"\nSetup profile: {profile_name}")
    print("=" * 60)
    print("Step 1/3: Authenticate and save profile")
    configure_cmd(
        argparse.Namespace(
            profile=profile_name,
            store_dir=store_dir,
            force=force,
            redirect_uri=redirect_uri,
            client_id=client_id,
            tenant_id=tenant_id,
        )
    )

    print("Step 2/3: Bind target")
    set_target_cmd(
        argparse.Namespace(
            profile=profile_name,
            store_dir=store_dir,
            sharepoint_url=sharepoint_url,
            my_drive=my_drive,
            library=library,
            folder=folder,
        )
    )

    if not skip_smoke_test:
        print("Step 3/3: Smoke test upload")
        temp_file = Path(tempfile.gettempdir()) / (
            f"rpa_setup_smoke_{int(datetime.utcnow().timestamp())}.txt"
        )
        temp_file.write_text(
            f"rpa-sharepoint-connector setup smoke test {datetime.utcnow().isoformat()}\n",
            encoding="utf-8",
        )
        try:
            test_upload_cmd(
                argparse.Namespace(
                    profile=profile_name,
                    store_dir=store_dir,
                    file=str(temp_file),
                )
            )
        finally:
            try:
                unlink_if_exists(temp_file)
            except Exception:
                pass
    else:
        print("Step 3/3: Smoke test skipped")

    print("OK Setup complete")
    print()
