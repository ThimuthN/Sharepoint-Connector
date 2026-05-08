"""Tests for CLI commands."""
import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta
from pathlib import Path
import json
from rpa_sharepoint_connector.cli import (
    cmd_status, cmd_configure, cmd_configure_browser, cmd_test_upload, cmd_list_profiles, cmd_disconnect
)


class TestStatusCommand:
    """Test status command."""

    def test_status_shows_profile_info(self):
        """Test status displays profile information."""
        with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = {
                "user_email": "user@example.com",
                "site_name": "My Site",
                "drive_name": "Documents",
                "folder_path": "Invoices",
                "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            }
            MockStore.return_value = mock_store

            args = MagicMock()
            args.profile = "test"
            args.store_dir = None

            with patch('builtins.print') as mock_print:
                cmd_status(args)
                output = "\n".join([str(call) for call in mock_print.call_args_list])
                assert "user@example.com" in output

    def test_status_missing_profile(self):
        """Test status with missing profile."""
        with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            args = MagicMock()
            args.profile = "missing"
            args.store_dir = None

            with pytest.raises(SystemExit):
                cmd_status(args)

    def test_status_shows_token_expiry(self):
        """Test status shows token expiry time."""
        future = datetime.utcnow() + timedelta(hours=2)

        with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = {
                "user_email": "user@example.com",
                "site_name": "Site",
                "drive_name": "Docs",
                "folder_path": "Folder",
                "expires_at": future.isoformat(),
            }
            MockStore.return_value = mock_store

            args = MagicMock()
            args.profile = "test"
            args.store_dir = None

            with patch('builtins.print') as mock_print:
                cmd_status(args)
                output = "\n".join([str(call) for call in mock_print.call_args_list])
                assert "expires" in output.lower() or "token" in output.lower()


class TestListProfilesCommand:
    """Test list command."""

    def test_list_profiles_shows_all(self):
        """Test list shows all profiles."""
        with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.list_profiles.return_value = ["profile1", "profile2"]
            mock_store.load_profile.side_effect = [
                {"user_email": "user1@example.com"},
                {"user_email": "user2@example.com"}
            ]
            MockStore.return_value = mock_store

            args = MagicMock()
            args.store_dir = None

            with patch('builtins.print') as mock_print:
                cmd_list_profiles(args)
                output = "\n".join([str(call) for call in mock_print.call_args_list])
                assert "profile1" in output
                assert "profile2" in output

    def test_list_profiles_empty(self):
        """Test list when no profiles."""
        with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.list_profiles.return_value = []
            MockStore.return_value = mock_store

            args = MagicMock()
            args.store_dir = None

            with patch('builtins.print') as mock_print:
                cmd_list_profiles(args)
                output = "\n".join([str(call) for call in mock_print.call_args_list])
                assert "No profiles" in output or "profiles" in output.lower()


class TestDisconnectCommand:
    """Test disconnect command."""

    def test_disconnect_deletes_profile(self):
        """Test disconnect deletes profile after confirmation."""
        with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
            mock_store = MagicMock()
            MockStore.return_value = mock_store

            args = MagicMock()
            args.profile = "test"
            args.store_dir = None

            with patch('builtins.input', return_value="y"):
                with patch('builtins.print') as mock_print:
                    cmd_disconnect(args)
                    mock_store.delete_profile.assert_called_once_with("test")

    def test_disconnect_cancels_on_no(self):
        """Test disconnect cancels if user says no."""
        with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
            mock_store = MagicMock()
            MockStore.return_value = mock_store

            args = MagicMock()
            args.profile = "test"
            args.store_dir = None

            with patch('builtins.input', return_value="n"):
                with patch('builtins.print'):
                    cmd_disconnect(args)
                    mock_store.delete_profile.assert_not_called()


class TestTestUploadCommand:
    """Test upload command."""

    def test_test_upload_success(self):
        """Test upload test succeeds."""
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp_path = tmp.name
            tmp.write(b"test content")

        try:
            with patch('rpa_sharepoint_connector.cli.SharePointClient') as MockClient:
                mock_sp = MagicMock()
                MockClient.return_value = mock_sp

                args = MagicMock()
                args.file = tmp_path
                args.profile = "test"
                args.store_dir = None

                with patch('builtins.print') as mock_print:
                    cmd_test_upload(args)
                    mock_sp.upload.assert_called()
                    mock_sp.delete.assert_called()
                    output = "\n".join([str(call) for call in mock_print.call_args_list])
                    assert "successful" in output.lower() or "test" in output.lower()
        finally:
            import os
            os.unlink(tmp_path)

    def test_test_upload_file_not_found(self):
        """Test upload fails if file not found."""
        args = MagicMock()
        args.file = "/nonexistent/file.pdf"
        args.profile = "test"
        args.store_dir = None

        with pytest.raises(SystemExit):
            cmd_test_upload(args)


class TestConfigureCommand:
    """Test configure command with Device Code Flow."""

    def test_configure_device_code_flow(self):
        """Test configure uses Device Code Flow."""
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None

        with patch('rpa_sharepoint_connector.cli.MicrosoftAuth') as MockAuth:
            mock_auth = MagicMock()
            MockAuth.return_value = mock_auth

            # Mock device flow start
            mock_auth.start_device_flow.return_value = {
                "device_code": "ABC123",
                "user_code": "ABC-DEF",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5
            }

            # Mock polling
            mock_auth.poll_device_flow.return_value = {
                "access_token": "token_123",
                "refresh_token": "refresh_456",
                "expires_in": 3600
            }

            # Mock user info
            mock_auth.get_user_info.return_value = {
                "id": "user_id",
                "mail": "user@example.com"
            }

            with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
                mock_store = MagicMock()
                MockStore.return_value = mock_store

                with patch('builtins.print') as mock_print:
                    cmd_configure(args)

                    # Verify device flow was initiated
                    mock_auth.start_device_flow.assert_called_once()
                    # Verify polling was called
                    mock_auth.poll_device_flow.assert_called_once()
                    # Verify user info was fetched
                    mock_auth.get_user_info.assert_called_once()
                    # Verify profile was saved
                    mock_store.save_profile.assert_called_once()

                    output = "\n".join([str(call) for call in mock_print.call_args_list])
                    assert "Authorization successful" in output

    def test_configure_prints_device_code(self):
        """Test configure prints device code and verification URL."""
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None

        with patch('rpa_sharepoint_connector.cli.MicrosoftAuth') as MockAuth:
            mock_auth = MagicMock()
            MockAuth.return_value = mock_auth

            mock_auth.start_device_flow.return_value = {
                "device_code": "DEV123",
                "user_code": "ABC-XYZ",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5
            }

            mock_auth.poll_device_flow.return_value = {
                "access_token": "token",
                "refresh_token": "refresh",
                "expires_in": 3600
            }

            mock_auth.get_user_info.return_value = {
                "id": "user",
                "mail": "user@example.com"
            }

            with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
                MockStore.return_value = MagicMock()

                with patch('builtins.print') as mock_print:
                    cmd_configure(args)

                    output = "\n".join([str(call) for call in mock_print.call_args_list])
                    assert "ABC-XYZ" in output  # User code
                    assert "microsoft.com/devicelogin" in output  # Verification URI

    def test_configure_no_client_secret_required(self):
        """Test configure does not ask for client secret."""
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None

        with patch('rpa_sharepoint_connector.cli.MicrosoftAuth') as MockAuth:
            mock_auth = MagicMock()
            MockAuth.return_value = mock_auth

            mock_auth.start_device_flow.return_value = {
                "device_code": "DEV",
                "user_code": "ABC",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5
            }

            mock_auth.poll_device_flow.return_value = {
                "access_token": "token",
                "refresh_token": "refresh",
                "expires_in": 3600
            }

            mock_auth.get_user_info.return_value = {
                "id": "user",
                "mail": "user@example.com"
            }

            with patch('rpa_sharepoint_connector.cli.TokenStore') as MockStore:
                MockStore.return_value = MagicMock()

                with patch('builtins.input') as mock_input:
                    with patch('builtins.print'):
                        cmd_configure(args)

                    # Should NOT ask for client_id or secret
                    # Device Code Flow doesn't require user input
                    mock_input.assert_not_called()

    def test_configure_handles_authorization_denied(self):
        """Test configure handles user denying authorization."""
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None

        with patch('rpa_sharepoint_connector.cli.MicrosoftAuth') as MockAuth:
            mock_auth = MagicMock()
            MockAuth.return_value = mock_auth

            mock_auth.start_device_flow.return_value = {
                "device_code": "DEV",
                "user_code": "ABC",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5
            }

            mock_auth.poll_device_flow.side_effect = ValueError("Authorization denied")

            with pytest.raises(SystemExit):
                cmd_configure(args)

    def test_configure_handles_timeout(self):
        """Test configure handles device code timeout."""
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None

        with patch('rpa_sharepoint_connector.cli.MicrosoftAuth') as MockAuth:
            mock_auth = MagicMock()
            MockAuth.return_value = mock_auth

            mock_auth.start_device_flow.return_value = {
                "device_code": "DEV",
                "user_code": "ABC",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5
            }

            mock_auth.poll_device_flow.side_effect = ValueError("timed out")

            with pytest.raises(SystemExit):
                cmd_configure(args)


class TestConfigureBrowserCommand:
    """Test configure-browser command."""

    def test_configure_browser_stores_encrypted_refresh_token(self, tmp_path):
        """Successful browser callback persists encrypted refresh token."""
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = str(tmp_path)

        with patch("rpa_sharepoint_connector.cli.MicrosoftBrowserAuth") as MockAuth:
            mock_auth = MagicMock()
            MockAuth.return_value = mock_auth
            mock_auth.tenant_id = "common"
            mock_auth.redirect_uri = "http://localhost:8765/callback"
            mock_auth.build_authorization_request.return_value = {
                "state": "state_123",
                "code_verifier": "verifier_abc",
                "code_challenge": "challenge_xyz",
                "authorization_url": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=state_123",
            }
            mock_auth.authenticate.return_value = {
                "tokens": {
                    "access_token": "access_token_value",
                    "refresh_token": "refresh_token_value",
                    "expires_in": 3600,
                },
                "user_info": {
                    "id": "user_1",
                    "mail": "user@example.com",
                },
                "authorization_url": "https://example.com",
                "state": "state_123",
                "code_challenge": "challenge_xyz",
                "callback_closed": True,
            }

            with patch("builtins.print"):
                cmd_configure_browser(args)

        profile_file = Path(tmp_path) / "demo.json"
        raw = json.loads(profile_file.read_text())
        assert raw["refresh_token"] != "refresh_token_value"

        from rpa_sharepoint_connector.token_store import TokenStore
        store = TokenStore(store_dir=str(tmp_path))
        loaded = store.load_profile("demo")
        assert loaded["refresh_token"] == "refresh_token_value"
