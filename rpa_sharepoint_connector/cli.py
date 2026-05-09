"""Command-line interface for SharePoint connector."""
import argparse
import json
import logging
import sys
import tempfile
from pathlib import Path
from datetime import datetime, timedelta
from urllib.parse import unquote, urlparse

from .browser_auth import MicrosoftBrowserAuth
from .auth import MicrosoftAuth
from .graph_client import GraphClient
from .token_store import TokenStore
from .sdk import SharePointClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


def _resolve_user_email(user_info: dict) -> str:
    """Resolve best available user email identifier from Graph /me payload."""
    return (
        user_info.get("mail")
        or user_info.get("userPrincipalName")
        or "unknown"
    )


def _save_profile(
    profile_name: str,
    store_dir: str,
    tokens: dict,
    user_info: dict,
    client_id: str,
    tenant_id: str,
) -> str:
    """Persist token/user data in the standard profile format."""
    store = TokenStore(store_dir=store_dir)
    user_email = _resolve_user_email(user_info)
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


def _normalize_name(value: str) -> str:
    """Normalize text for case/spacing-insensitive matching."""
    return "".join(ch.lower() for ch in value if ch.isalnum())


def _parse_sharepoint_url(sharepoint_url: str) -> dict:
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


def _select_drive(drives: list, requested_library: str) -> dict:
    """Select best matching document library drive."""
    requested = _normalize_name(requested_library)
    aliases = {requested}
    if requested == "shareddocuments":
        aliases.add("documents")
    if requested == "documents":
        aliases.add("shareddocuments")

    for drive in drives:
        candidates = {_normalize_name(drive.get("name", ""))}
        web_url = drive.get("webUrl", "")
        if web_url:
            tail = unquote(urlparse(web_url).path.split("/")[-1])
            candidates.add(_normalize_name(tail))
        if aliases.intersection(candidates):
            return drive

    raise ValueError(
        f"Library '{requested_library}' not found. Available libraries: "
        f"{', '.join(d.get('name', '<unknown>') for d in drives)}"
    )


def _ensure_profile_token(
    profile_name: str,
    profile_data: dict,
    store: TokenStore,
) -> dict:
    """Refresh profile access token if required and persist updates."""
    expires_at = datetime.fromisoformat(profile_data["expires_at"])
    auth = MicrosoftAuth(
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


def cmd_set_target(args):
    """Bind configured profile to a SharePoint site/library/folder target."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir
    sharepoint_url = args.sharepoint_url
    my_drive = bool(getattr(args, "my_drive", False))
    library_override = args.library
    folder_override = args.folder

    try:
        store = TokenStore(store_dir=store_dir)
        profile_data = store.load_profile(profile_name)
        if not profile_data:
            print(
                f"Profile '{profile_name}' not found. "
                f"Run: python -m rpa_sharepoint_connector configure --profile {profile_name}"
            )
            sys.exit(1)

        profile_data = _ensure_profile_token(profile_name, profile_data, store)

        graph = GraphClient(profile_data["access_token"])
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
            parsed = _parse_sharepoint_url(sharepoint_url)
            hostname = parsed["hostname"]
            site_path = parsed["site_path"]
            library_name = library_override or parsed["library_name"] or "Documents"
            folder_path = (
                folder_override if folder_override is not None else parsed["folder_path"]
            )

            site = graph._get(f"/sites/{hostname}:{site_path}")
            drives = graph.list_drives(site["id"])
            drive = _select_drive(drives, library_name)
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
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


def cmd_configure(args):
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
        store = TokenStore(store_dir=store_dir)
        existing_profile = store.load_profile(profile_name)
        if existing_profile and not force:
            print(
                f"ERROR: Profile '{profile_name}' already exists. "
                f"Run configure again with --force to replace it."
            )
            sys.exit(1)

        auth = MicrosoftBrowserAuth(
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

        user_email = _resolve_user_email(user_info)
        print(f"OK Logged in as: {user_email}")

        _save_profile(
            profile_name=profile_name,
            store_dir=store_dir,
            tokens=tokens,
            user_info=user_info,
            client_id=auth.client_id,
            tenant_id=auth.tenant_id,
        )
        print(f"OK Profile '{profile_name}' saved")
        print()

    except KeyboardInterrupt:
        print("\n\nCancelled")
        sys.exit(0)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


def cmd_status(args):
    """Show profile connection status."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir

    try:
        store = TokenStore(store_dir=store_dir)
        profile_data = store.load_profile(profile_name)

        if not profile_data:
            print(f"Profile '{profile_name}' not found")
            sys.exit(1)

        print(f"\nProfile: {profile_name}")
        print("=" * 60)
        print(f"User: {profile_data.get('user_email', 'unknown')}")
        print(f"Site: {profile_data.get('site_name', 'unknown')}")
        print(f"Library: {profile_data.get('drive_name', 'unknown')}")
        print(f"Folder: {profile_data.get('folder_path', '(root)')}")

        expires_at = datetime.fromisoformat(profile_data["expires_at"])
        if datetime.utcnow() < expires_at:
            mins = int((expires_at - datetime.utcnow()).total_seconds() / 60)
            print(f"Token expires in: {mins} minutes")
        else:
            print("Token expired (will auto-refresh)")

        print()
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


def cmd_test_upload(args):
    """Test profile with a file upload."""
    profile_name = args.profile or "default"
    file_path = args.file
    store_dir = args.store_dir

    if not file_path or not Path(file_path).exists():
        print(f"Error: File not found: {file_path}")
        sys.exit(1)

    try:
        sp = SharePointClient(profile=profile_name, store_dir=store_dir)
        filename = Path(file_path).name
        remote_path = f".test/{filename}"

        print(f"\nTesting upload...")
        sp.upload(file_path, remote_path)

        # Try to delete it
        sp.delete(remote_path)
        print(f"OK Test successful!")
        print()

    except Exception as e:
        print(f"ERROR: Test failed: {e}")
        sys.exit(1)


def cmd_list_profiles(args):
    """List all saved profiles."""
    store_dir = args.store_dir

    try:
        store = TokenStore(store_dir=store_dir)
        profiles = store.list_profiles()

        if not profiles:
            print("No profiles saved")
            return

        print("Saved profiles:")
        print("=" * 40)
        for profile_name in profiles:
            try:
                profile_data = store.load_profile(profile_name)
                email = profile_data.get("user_email", "unknown")
                print(f"  {profile_name:20} - {email}")
            except:
                print(f"  {profile_name:20} - (error reading)")

        print()
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


def cmd_disconnect(args):
    """Disconnect and delete a profile."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir

    try:
        store = TokenStore(store_dir=store_dir)
        confirm = input(f"Delete profile '{profile_name}'? (y/n): ").strip().lower()

        if confirm == "y":
            store.delete_profile(profile_name)
            print(f"OK Profile '{profile_name}' deleted")
        else:
            print("Cancelled")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


def cmd_setup(args):
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
    cmd_configure(
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
    cmd_set_target(
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
            cmd_test_upload(
                argparse.Namespace(
                    profile=profile_name,
                    store_dir=store_dir,
                    file=str(temp_file),
                )
            )
        finally:
            try:
                if temp_file.exists():
                    temp_file.unlink()
            except Exception:
                pass
    else:
        print("Step 3/3: Smoke test skipped")

    print("OK Setup complete")
    print()


def cmd_run(args):
    """Run one SharePoint/OneDrive operation for automation bots."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir
    op = args.op
    as_json = bool(getattr(args, "json", False))

    try:
        sp = SharePointClient(profile=profile_name, store_dir=store_dir)

        if op == "upload":
            if not args.local_path or not args.remote_path:
                raise ValueError("upload requires --local-path and --remote-path")
            item_id = sp.upload(args.local_path, args.remote_path, conflict=args.conflict)
            result = {
                "operation": "upload",
                "success": True,
                "item_id": item_id,
                "remote_path": args.remote_path,
            }
        elif op == "download":
            if not args.remote_path or not args.local_path:
                raise ValueError("download requires --remote-path and --local-path")
            sp.download(args.remote_path, args.local_path)
            result = {"operation": "download", "success": True, "local_path": args.local_path}
        elif op == "list":
            folder_path = args.folder_path or ""
            items = sp.list(folder_path)
            result = {
                "operation": "list",
                "success": True,
                "folder_path": folder_path,
                "count": len(items),
                "items": items,
            }
        elif op == "delete":
            if not args.remote_path:
                raise ValueError("delete requires --remote-path")
            sp.delete(args.remote_path)
            result = {"operation": "delete", "success": True, "remote_path": args.remote_path}
        elif op == "move":
            if not args.source_path or not args.target_path:
                raise ValueError("move requires --source-path and --target-path")
            sp.move(args.source_path, args.target_path, new_name=args.new_name)
            result = {
                "operation": "move",
                "success": True,
                "source_path": args.source_path,
                "target_path": args.target_path,
                "new_name": args.new_name or "",
            }
        elif op == "mkdir":
            if not args.folder_path:
                raise ValueError("mkdir requires --folder-path")
            item_id = sp.mkdir(args.folder_path)
            result = {
                "operation": "mkdir",
                "success": True,
                "folder_path": args.folder_path,
                "item_id": item_id,
            }
        elif op == "exists":
            if not args.remote_path:
                raise ValueError("exists requires --remote-path")
            exists = sp.exists(args.remote_path)
            result = {
                "operation": "exists",
                "success": True,
                "remote_path": args.remote_path,
                "exists": exists,
            }
        else:
            raise ValueError(f"Unsupported operation: {op}")

        if as_json:
            print(json.dumps(result, indent=2))
        else:
            print(f"OK {op} successful")
            if op == "list":
                for item in result["items"]:
                    kind = "folder" if item.get("is_folder") else "file"
                    print(f"- {item.get('name')} ({kind})")
            elif op == "exists":
                print(f"Exists: {result['exists']}")

    except Exception as e:
        if as_json:
            print(
                json.dumps(
                    {"operation": op, "success": False, "error": str(e)},
                    indent=2,
                )
            )
        else:
            print(f"ERROR: {e}")
        sys.exit(1)


def main():
    """Parse arguments and dispatch to commands."""
    parser = argparse.ArgumentParser(
        description="Microsoft SharePoint Connector for RPA bots",
        prog="python -m rpa_sharepoint_connector",
    )

    parser.add_argument(
        "--store-dir",
        help="Token store directory (default: ~/.rpa_sharepoint_connector)",
    )

    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # configure
    configure_parser = subparsers.add_parser("configure", help="Configure a new profile")
    configure_parser.add_argument(
        "--profile", "-p", help="Profile name (default: default)"
    )
    configure_parser.add_argument(
        "--force",
        action="store_true",
        help="Replace existing profile configuration",
    )
    configure_parser.add_argument(
        "--redirect-uri",
        help="Override browser callback redirect URI (default: http://localhost/callback)",
    )
    configure_parser.add_argument(
        "--client-id",
        help="Override Microsoft public-client app ID for this profile",
    )
    configure_parser.add_argument(
        "--tenant-id",
        help="Override tenant endpoint (e.g., organizations, common, or tenant GUID)",
    )
    configure_parser.set_defaults(func=cmd_configure)

    # setup
    setup_parser = subparsers.add_parser(
        "setup",
        help="One-command setup: configure + target bind + smoke upload",
    )
    setup_parser.add_argument("--profile", "-p", help="Profile name (default: default)")
    setup_parser.add_argument(
        "--force",
        action="store_true",
        help="Replace existing profile configuration",
    )
    setup_parser.add_argument(
        "--redirect-uri",
        help="Override browser callback redirect URI (default: http://localhost/callback)",
    )
    setup_parser.add_argument(
        "--client-id",
        help="Override Microsoft public-client app ID for this profile",
    )
    setup_parser.add_argument(
        "--tenant-id",
        help="Override tenant endpoint (e.g., organizations, common, or tenant GUID)",
    )
    setup_parser.add_argument(
        "--sharepoint-url",
        help="SharePoint site/library/folder URL from browser",
    )
    setup_parser.add_argument(
        "--my-drive",
        action="store_true",
        help="Use signed-in user's OneDrive as target",
    )
    setup_parser.add_argument(
        "--library",
        help="Override library name (defaults to URL-derived value)",
    )
    setup_parser.add_argument(
        "--folder",
        help="Override folder path inside library (defaults to URL-derived value)",
    )
    setup_parser.add_argument(
        "--skip-smoke-test",
        action="store_true",
        help="Skip upload/delete smoke test step",
    )
    setup_parser.set_defaults(func=cmd_setup)

    # run
    run_parser = subparsers.add_parser(
        "run",
        help="Run a single operation (upload/download/list/delete/move/mkdir/exists)",
    )
    run_parser.add_argument("--profile", "-p", help="Profile name (default: default)")
    run_parser.add_argument(
        "--op",
        required=True,
        choices=["upload", "download", "list", "delete", "move", "mkdir", "exists"],
        help="Operation to run",
    )
    run_parser.add_argument("--local-path", help="Local file path for upload/download")
    run_parser.add_argument("--remote-path", help="Remote SharePoint/OneDrive path")
    run_parser.add_argument("--folder-path", help="Folder path for list/mkdir")
    run_parser.add_argument("--source-path", help="Source path for move")
    run_parser.add_argument("--target-path", help="Target folder path for move")
    run_parser.add_argument("--new-name", help="Optional new name for move")
    run_parser.add_argument(
        "--conflict",
        choices=["overwrite", "fail_if_exists", "rename"],
        default="overwrite",
        help="Upload conflict behavior (default: overwrite)",
    )
    run_parser.add_argument(
        "--json",
        action="store_true",
        help="Print result as JSON for bot parsing",
    )
    run_parser.set_defaults(func=cmd_run)

    # status
    status_parser = subparsers.add_parser("status", help="Show profile status")
    status_parser.add_argument(
        "--profile", "-p", help="Profile name (default: default)"
    )
    status_parser.set_defaults(func=cmd_status)

    # test-upload
    test_parser = subparsers.add_parser(
        "test-upload", help="Test profile with file upload"
    )
    test_parser.add_argument("file", help="File to upload")
    test_parser.add_argument("--profile", "-p", help="Profile name (default: default)")
    test_parser.set_defaults(func=cmd_test_upload)

    # set-target
    target_parser = subparsers.add_parser(
        "set-target",
        help="Bind profile to SharePoint site/library/folder from a URL",
    )
    target_parser.add_argument(
        "--profile", "-p", help="Profile name (default: default)"
    )
    target_parser.add_argument(
        "--sharepoint-url",
        help="SharePoint site/library/folder URL from browser",
    )
    target_parser.add_argument(
        "--my-drive",
        action="store_true",
        help="Use signed-in user's OneDrive as target (easy personal-account smoke test)",
    )
    target_parser.add_argument(
        "--library",
        help="Override library name (defaults to URL-derived value)",
    )
    target_parser.add_argument(
        "--folder",
        help="Override folder path inside library (defaults to URL-derived value)",
    )
    target_parser.set_defaults(func=cmd_set_target)

    # list
    list_parser = subparsers.add_parser("list", help="List saved profiles")
    list_parser.set_defaults(func=cmd_list_profiles)

    # disconnect
    disconnect_parser = subparsers.add_parser("disconnect", help="Delete a profile")
    disconnect_parser.add_argument(
        "--profile", "-p", help="Profile name (default: default)"
    )
    disconnect_parser.set_defaults(func=cmd_disconnect)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    try:
        args.func(args)
    except KeyboardInterrupt:
        print("\n\nCancelled")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
