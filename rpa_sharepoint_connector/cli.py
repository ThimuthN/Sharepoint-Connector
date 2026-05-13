"""Command-line interface for SharePoint connector."""
import argparse
import logging
import sys
from pathlib import Path
from datetime import datetime

from .browser_auth import MicrosoftBrowserAuth
from .auth import MicrosoftAuth
from .cli_doctor import cmd_doctor as _cmd_doctor
from .cli_output import (
    classify_error_for_json,
    print_actionable_error,
    print_run_result,
)
from .cli_run import cmd_run as _cmd_run
from .cli_setup import (
    cmd_configure as _cmd_configure,
    cmd_set_target as _cmd_set_target,
    cmd_setup as _cmd_setup,
    ensure_profile_token,
    normalize_name,
    parse_sharepoint_url,
    resolve_user_email,
    save_profile,
    select_drive,
)
from .graph_client import GraphClient
from .token_store import TokenStore
from .sdk import SharePointClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


def _unlink_if_exists(path: Path) -> None:
    """Remove a file if present, without relying on Python 3.8+ APIs."""
    try:
        if path.exists():
            path.unlink()
    except FileNotFoundError:
        return


def _classify_error_for_json(error: Exception) -> dict:
    """Backwards-compatible wrapper for JSON error classification."""
    return classify_error_for_json(error)


def _print_actionable_error(error: Exception, command: str, profile: str = "default") -> None:
    """Backwards-compatible wrapper for CLI error printing."""
    print_actionable_error(error, command=command, profile=profile)


def _print_run_result(result: dict) -> None:
    """Backwards-compatible wrapper for run result formatting."""
    print_run_result(result)


def _resolve_user_email(user_info: dict) -> str:
    """Resolve best available user email identifier from Graph /me payload."""
    return resolve_user_email(user_info)


def _save_profile(
    profile_name: str,
    store_dir: str,
    tokens: dict,
    user_info: dict,
    client_id: str,
    tenant_id: str,
) -> str:
    """Persist token/user data in the standard profile format."""
    return save_profile(
        profile_name=profile_name,
        store_dir=store_dir,
        tokens=tokens,
        user_info=user_info,
        client_id=client_id,
        tenant_id=tenant_id,
        token_store_cls=TokenStore,
    )


def _normalize_name(value: str) -> str:
    """Normalize text for case/spacing-insensitive matching."""
    return normalize_name(value)


def _parse_sharepoint_url(sharepoint_url: str) -> dict:
    """Parse a SharePoint browser URL into host/site/library/folder hints."""
    return parse_sharepoint_url(sharepoint_url)


def _select_drive(drives: list, requested_library: str) -> dict:
    """Select best matching document library drive."""
    return select_drive(drives, requested_library)


def _ensure_profile_token(
    profile_name: str,
    profile_data: dict,
    store: TokenStore,
) -> dict:
    """Refresh profile access token if required and persist updates."""
    return ensure_profile_token(
        profile_name=profile_name,
        profile_data=profile_data,
        store=store,
        auth_cls=MicrosoftAuth,
    )


def cmd_set_target(args):
    """Bind configured profile to a SharePoint site/library/folder target."""
    _cmd_set_target(
        args,
        token_store_cls=TokenStore,
        graph_client_cls=GraphClient,
        auth_cls=MicrosoftAuth,
        parse_sharepoint_url_fn=_parse_sharepoint_url,
        select_drive_fn=_select_drive,
        ensure_profile_token_fn=_ensure_profile_token,
        print_actionable_error=_print_actionable_error,
    )


def cmd_configure(args):
    """Configure a new profile using browser OAuth (Authorization Code + PKCE)."""
    _cmd_configure(
        args,
        token_store_cls=TokenStore,
        browser_auth_cls=MicrosoftBrowserAuth,
        save_profile_fn=_save_profile,
        print_actionable_error=_print_actionable_error,
    )


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
        _print_actionable_error(e, command="status", profile=profile_name)


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
        _print_actionable_error(e, command="test-upload", profile=profile_name)


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
            except Exception:
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
        confirm = "y" if getattr(args, "yes", False) else input(
            f"Delete profile '{profile_name}'? (y/n): "
        ).strip().lower()

        if confirm == "y":
            store.delete_profile(profile_name)
            print(f"OK Profile '{profile_name}' deleted")
        else:
            print("Cancelled")

    except Exception as e:
        _print_actionable_error(e, command="disconnect", profile=profile_name)


def cmd_setup(args):
    """Run one-command onboarding: configure, bind target, and smoke test."""
    _cmd_setup(
        args,
        configure_cmd=cmd_configure,
        set_target_cmd=cmd_set_target,
        test_upload_cmd=cmd_test_upload,
        unlink_if_exists=_unlink_if_exists,
    )


def cmd_run(args):
    """Run one SharePoint/OneDrive operation for automation bots."""
    _cmd_run(
        args,
        sharepoint_client_cls=SharePointClient,
        classify_error=_classify_error_for_json,
        print_run_result=_print_run_result,
        print_actionable_error=_print_actionable_error,
    )


def cmd_doctor(args):
    """Run preflight diagnostics for local/bot execution."""
    _cmd_doctor(
        args,
        unlink_if_exists=_unlink_if_exists,
        print_actionable_error=_print_actionable_error,
    )


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

    # doctor
    doctor_parser = subparsers.add_parser(
        "doctor",
        help="Run local/profile diagnostics for setup and bot runtime",
    )
    doctor_parser.add_argument("--profile", "-p", help="Profile name (default: default)")
    doctor_parser.add_argument(
        "--json",
        action="store_true",
        help="Print doctor report as JSON",
    )
    doctor_parser.set_defaults(func=cmd_doctor)

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
    disconnect_parser.add_argument(
        "--yes",
        action="store_true",
        help="Delete without interactive prompt (bot-safe)",
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
