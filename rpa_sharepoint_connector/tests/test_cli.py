"""Tests for CLI commands."""
import json
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from rpa_sharepoint_connector.cli import (
    cmd_configure,
    cmd_disconnect,
    cmd_list_profiles,
    cmd_status,
    cmd_test_upload,
)


class TestStatusCommand:
    """Test status command."""

    def test_status_shows_profile_info(self):
        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
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

            with patch("builtins.print") as mock_print:
                cmd_status(args)
                output = "\n".join([str(call) for call in mock_print.call_args_list])
                assert "user@example.com" in output

    def test_status_missing_profile(self):
        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            args = MagicMock()
            args.profile = "missing"
            args.store_dir = None

            with pytest.raises(SystemExit):
                cmd_status(args)


class TestListProfilesCommand:
    """Test list command."""

    def test_list_profiles_shows_all(self):
        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.list_profiles.return_value = ["profile1", "profile2"]
            mock_store.load_profile.side_effect = [
                {"user_email": "user1@example.com"},
                {"user_email": "user2@example.com"},
            ]
            MockStore.return_value = mock_store

            args = MagicMock()
            args.store_dir = None

            with patch("builtins.print") as mock_print:
                cmd_list_profiles(args)
                output = "\n".join([str(call) for call in mock_print.call_args_list])
                assert "profile1" in output
                assert "profile2" in output


class TestDisconnectCommand:
    """Test disconnect command."""

    def test_disconnect_deletes_profile(self):
        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            MockStore.return_value = mock_store

            args = MagicMock()
            args.profile = "test"
            args.store_dir = None

            with patch("builtins.input", return_value="y"):
                cmd_disconnect(args)
                mock_store.delete_profile.assert_called_once_with("test")


class TestTestUploadCommand:
    """Test upload command."""

    def test_test_upload_success(self):
        import tempfile

        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp_path = tmp.name
            tmp.write(b"test content")

        try:
            with patch("rpa_sharepoint_connector.cli.SharePointClient") as MockClient:
                mock_sp = MagicMock()
                MockClient.return_value = mock_sp

                args = MagicMock()
                args.file = tmp_path
                args.profile = "test"
                args.store_dir = None

                with patch("builtins.print") as mock_print:
                    cmd_test_upload(args)
                    mock_sp.upload.assert_called_once()
                    mock_sp.delete.assert_called_once()
                    output = "\n".join([str(call) for call in mock_print.call_args_list])
                    assert "successful" in output.lower()
        finally:
            import os

            os.unlink(tmp_path)


class TestConfigureCommand:
    """Test configure command (browser OAuth)."""

    def test_configure_uses_browser_auth(self):
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None
        args.force = False
        args.redirect_uri = None

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

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
                    "redirect_uri": "http://localhost:8765/callback",
                }
                mock_auth.authenticate.return_value = {
                    "tokens": {
                        "access_token": "token_123",
                        "refresh_token": "refresh_456",
                        "expires_in": 3600,
                    },
                    "user_info": {
                        "id": "user_id",
                        "mail": "user@example.com",
                    },
                }

                with patch("builtins.print") as mock_print:
                    cmd_configure(args)
                    mock_auth.build_authorization_request.assert_called_once()
                    mock_auth.authenticate.assert_called_once()
                    assert mock_store.save_profile.called
                    output = "\n".join([str(call) for call in mock_print.call_args_list])
                    assert "Authorization successful" in output

    def test_configure_requires_force_for_existing_profile(self):
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None
        args.force = False
        args.redirect_uri = None

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = {"user_email": "existing@example.com"}
            MockStore.return_value = mock_store

            with pytest.raises(SystemExit):
                cmd_configure(args)

    def test_configure_force_overwrites_existing_profile(self):
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None
        args.force = True
        args.redirect_uri = None

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = {"user_email": "existing@example.com"}
            MockStore.return_value = mock_store

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
                    "redirect_uri": "http://localhost:8765/callback",
                }
                mock_auth.authenticate.return_value = {
                    "tokens": {
                        "access_token": "token_123",
                        "refresh_token": "refresh_456",
                        "expires_in": 3600,
                    },
                    "user_info": {
                        "id": "user_id",
                        "mail": "user@example.com",
                    },
                }

                with patch("builtins.print"):
                    cmd_configure(args)
                assert mock_store.save_profile.called

    def test_configure_saves_encrypted_refresh_token(self, tmp_path):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = str(tmp_path)
        args.force = False
        args.redirect_uri = None

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
                "redirect_uri": "http://localhost:8765/callback",
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
            }

            with patch("builtins.print"):
                cmd_configure(args)

        profile_file = Path(tmp_path) / "demo.json"
        raw = json.loads(profile_file.read_text())
        assert raw["refresh_token"] != "refresh_token_value"

    def test_configure_allows_redirect_override(self):
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None
        args.force = False
        args.redirect_uri = "http://localhost:9999/custom"

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            with patch("rpa_sharepoint_connector.cli.MicrosoftBrowserAuth") as MockAuth:
                mock_auth = MagicMock()
                MockAuth.return_value = mock_auth
                mock_auth.tenant_id = "common"
                mock_auth.build_authorization_request.return_value = {
                    "state": "state_123",
                    "code_verifier": "verifier_abc",
                    "code_challenge": "challenge_xyz",
                    "authorization_url": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=state_123",
                    "redirect_uri": "http://localhost:9999/custom",
                }
                mock_auth.authenticate.return_value = {
                    "tokens": {
                        "access_token": "token_123",
                        "refresh_token": "refresh_456",
                        "expires_in": 3600,
                    },
                    "user_info": {
                        "id": "user_id",
                        "mail": "user@example.com",
                    },
                }

                with patch("builtins.print"):
                    cmd_configure(args)

                MockAuth.assert_called_once_with(
                    redirect_uri="http://localhost:9999/custom"
                )
