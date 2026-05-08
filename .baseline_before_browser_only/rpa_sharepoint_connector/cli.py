"""Command-line interface for SharePoint connector."""
import argparse
import logging
import sys
from pathlib import Path
from datetime import datetime, timedelta

from .auth import MicrosoftAuth
from .browser_auth import MicrosoftBrowserAuth
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
    }
    store.save_profile(profile_name, profile_data)
    return user_email


def cmd_configure(args):
    """Configure a new profile using Device Code Flow."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir

    print(f"\nConfiguring profile: {profile_name}")
    print("=" * 60)

    try:
        # Start device code flow
        auth = MicrosoftAuth()
        device_flow = auth.start_device_flow()
        verification_uri = device_flow.get("verification_uri", "")
        user_code = device_flow.get("user_code", "")
        interval = device_flow.get("interval", 5)
        expires_in = device_flow.get("expires_in", 900)

        print("\nMicrosoft Device Login Required")
        print("-" * 60)
        print(f"Tenant: {auth.tenant_id}")
        print(f"Go to: {verification_uri}")
        if device_flow.get("verification_uri_complete"):
            print(f"Direct link: {device_flow['verification_uri_complete']}")
        print(f"Enter code: {user_code}")
        print(f"Code expires in: {expires_in} seconds")
        print(f"Polling interval: {interval} seconds")
        print(f"\nWaiting for you to authorize...")
        print("-" * 60)

        # Poll for completion
        tokens = auth.poll_device_flow(
            device_code=device_flow["device_code"],
            interval=interval,
            expires_in=expires_in,
        )

        print("OK Authorization successful!")

        # Get user info
        user_info = auth.get_user_info(tokens["access_token"])
        user_email = _resolve_user_email(user_info)
        print(f"OK Logged in as: {user_email}")

        # Store profile
        _save_profile(profile_name, store_dir, tokens, user_info)
        print(f"OK Profile '{profile_name}' saved")
        print("\nNext: Set up SharePoint site and folder (optional)")
        print("Tip: You can configure site/drive/folder later or use root for now")
        print()

    except KeyboardInterrupt:
        print("\n\nCancelled")
        sys.exit(0)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


def cmd_configure_browser(args):
    """Configure a new profile using browser OAuth (Authorization Code + PKCE)."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir

    print(f"\nConfiguring profile: {profile_name} (browser)")
    print("=" * 60)

    try:
        auth = MicrosoftBrowserAuth()
        request = auth.build_authorization_request()

        print("\nMicrosoft Browser Login Required")
        print("-" * 60)
        print(f"Tenant: {auth.tenant_id}")
        print(f"Go to: {request['authorization_url']}")
        print(f"Redirect URI: {auth.redirect_uri}")
        print("\nOpening browser for Microsoft sign-in...")
        print("Waiting for callback on localhost...")
        print("-" * 60)

        result = auth.authenticate(
            open_browser=True,
            authorization_request=request,
        )
        tokens = result["tokens"]
        user_info = result["user_info"]
        user_email = _resolve_user_email(user_info)

        print("OK Authorization successful!")
        print(f"OK Logged in as: {user_email}")

        _save_profile(profile_name, store_dir, tokens, user_info)
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
    configure_parser.set_defaults(func=cmd_configure)

    # configure-browser
    configure_browser_parser = subparsers.add_parser(
        "configure-browser",
        help="Configure a new profile via browser OAuth (PKCE)",
    )
    configure_browser_parser.add_argument(
        "--profile", "-p", help="Profile name (default: default)"
    )
    configure_browser_parser.set_defaults(func=cmd_configure_browser)

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
