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
    cmd_run,
    cmd_setup,
    cmd_set_target,
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
        args.client_id = None
        args.tenant_id = None

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

                    saved_profile = mock_store.save_profile.call_args[0][1]
                    assert saved_profile["client_id"] == mock_auth.client_id
                    assert saved_profile["tenant_id"] == mock_auth.tenant_id

    def test_configure_requires_force_for_existing_profile(self):
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None
        args.force = False
        args.redirect_uri = None
        args.client_id = None
        args.tenant_id = None

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
        args.client_id = None
        args.tenant_id = None

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
        args.client_id = None
        args.tenant_id = None

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
        args.client_id = None
        args.tenant_id = None

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
                    client_id=None,
                    tenant_id=None,
                    redirect_uri="http://localhost:9999/custom"
                )

    def test_configure_allows_client_and_tenant_override(self):
        args = MagicMock()
        args.profile = "test"
        args.store_dir = None
        args.force = False
        args.redirect_uri = None
        args.client_id = "client_abc"
        args.tenant_id = "organizations"

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            with patch("rpa_sharepoint_connector.cli.MicrosoftBrowserAuth") as MockAuth:
                mock_auth = MagicMock()
                MockAuth.return_value = mock_auth
                mock_auth.client_id = "client_abc"
                mock_auth.tenant_id = "organizations"
                mock_auth.build_authorization_request.return_value = {
                    "state": "state_123",
                    "code_verifier": "verifier_abc",
                    "code_challenge": "challenge_xyz",
                    "authorization_url": "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?state=state_123",
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

                MockAuth.assert_called_once_with(
                    client_id="client_abc",
                    tenant_id="organizations",
                    redirect_uri=None,
                )

                saved_profile = mock_store.save_profile.call_args[0][1]
                assert saved_profile["client_id"] == "client_abc"
                assert saved_profile["tenant_id"] == "organizations"


class TestSetTargetCommand:
    """Test target binding command."""

    def test_set_target_updates_profile_from_sharepoint_url(self):
        args = MagicMock()
        args.profile = "sharepointlocation"
        args.store_dir = None
        args.sharepoint_url = (
            "https://mysliit.sharepoint.com/:f:/r/sites/Sharepointlocation/"
            "Shared%20Documents/Folder?csf=1&web=1"
        )
        args.my_drive = False
        args.library = None
        args.folder = None

        profile_data = {
            "access_token": "access_token",
            "refresh_token": "refresh_token",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "site_id": "",
            "site_name": "",
            "drive_id": "",
            "drive_name": "",
            "folder_id": "",
            "folder_path": "",
            "client_id": "client_123",
            "tenant_id": "organizations",
            "user_email": "user@example.com",
        }

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = profile_data
            MockStore.return_value = mock_store

            with patch("rpa_sharepoint_connector.cli.MicrosoftAuth") as MockAuth:
                mock_auth = MagicMock()
                mock_auth.is_token_expired.return_value = False
                MockAuth.return_value = mock_auth

                with patch("rpa_sharepoint_connector.cli.GraphClient") as MockGraph:
                    mock_graph = MagicMock()
                    mock_graph._get.return_value = {
                        "id": "site_123",
                        "displayName": "Sharepointlocation",
                    }
                    mock_graph.list_drives.return_value = [
                        {"id": "drive_1", "name": "Documents"},
                    ]
                    mock_graph._ensure_folder_path.return_value = "folder_456"
                    MockGraph.return_value = mock_graph

                    with patch("builtins.print"):
                        cmd_set_target(args)

                    saved_profile = mock_store.save_profile.call_args[0][1]
                    assert saved_profile["site_id"] == "site_123"
                    assert saved_profile["drive_id"] == "drive_1"
                    assert saved_profile["folder_id"] == "folder_456"
                    assert saved_profile["folder_path"] == "Folder"

    def test_set_target_requires_existing_profile(self):
        args = MagicMock()
        args.profile = "missing"
        args.store_dir = None
        args.sharepoint_url = "https://mysliit.sharepoint.com/sites/Sharepointlocation"
        args.my_drive = False
        args.library = None
        args.folder = None

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            with pytest.raises(SystemExit):
                cmd_set_target(args)

    def test_set_target_my_drive_updates_profile(self):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = None
        args.sharepoint_url = None
        args.my_drive = True
        args.library = None
        args.folder = "ConnectorSmoke"

        profile_data = {
            "access_token": "access_token",
            "refresh_token": "refresh_token",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "site_id": "",
            "site_name": "",
            "drive_id": "",
            "drive_name": "",
            "folder_id": "",
            "folder_path": "",
            "client_id": "client_123",
            "tenant_id": "common",
            "user_email": "demo@outlook.com",
        }

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = profile_data
            MockStore.return_value = mock_store

            with patch("rpa_sharepoint_connector.cli.MicrosoftAuth") as MockAuth:
                mock_auth = MagicMock()
                mock_auth.is_token_expired.return_value = False
                MockAuth.return_value = mock_auth

                with patch("rpa_sharepoint_connector.cli.GraphClient") as MockGraph:
                    mock_graph = MagicMock()
                    mock_graph._get.return_value = {
                        "id": "drive_abc",
                        "name": "OneDrive",
                    }
                    mock_graph._ensure_folder_path.return_value = "folder_xyz"
                    MockGraph.return_value = mock_graph

                    with patch("builtins.print"):
                        cmd_set_target(args)

                    saved_profile = mock_store.save_profile.call_args[0][1]
                    assert saved_profile["site_id"] == "me"
                    assert saved_profile["site_name"] == "My Drive"
                    assert saved_profile["drive_id"] == "drive_abc"
                    assert saved_profile["drive_name"] == "OneDrive"
                    assert saved_profile["folder_id"] == "folder_xyz"
                    assert saved_profile["folder_path"] == "ConnectorSmoke"

    def test_set_target_rejects_conflicting_target_args(self):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = None
        args.sharepoint_url = "https://mysliit.sharepoint.com/sites/Sharepointlocation"
        args.my_drive = True
        args.library = None
        args.folder = None

        profile_data = {
            "access_token": "access_token",
            "refresh_token": "refresh_token",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "site_id": "",
            "site_name": "",
            "drive_id": "",
            "drive_name": "",
            "folder_id": "",
            "folder_path": "",
            "client_id": "client_123",
            "tenant_id": "common",
            "user_email": "demo@outlook.com",
        }

        with patch("rpa_sharepoint_connector.cli.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = profile_data
            MockStore.return_value = mock_store

            with patch("rpa_sharepoint_connector.cli.MicrosoftAuth") as MockAuth:
                mock_auth = MagicMock()
                mock_auth.is_token_expired.return_value = False
                MockAuth.return_value = mock_auth

                with pytest.raises(SystemExit):
                    cmd_set_target(args)


class TestSetupCommand:
    """Test one-command setup flow."""

    def test_setup_defaults_to_my_drive_and_runs_smoke(self):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = None
        args.force = False
        args.redirect_uri = None
        args.client_id = None
        args.tenant_id = None
        args.sharepoint_url = None
        args.my_drive = False
        args.library = None
        args.folder = None
        args.skip_smoke_test = False

        with patch("rpa_sharepoint_connector.cli.cmd_configure") as mock_configure:
            with patch("rpa_sharepoint_connector.cli.cmd_set_target") as mock_target:
                with patch("rpa_sharepoint_connector.cli.cmd_test_upload") as mock_test:
                    with patch("builtins.print"):
                        cmd_setup(args)

                    assert mock_configure.call_count == 1
                    target_args = mock_target.call_args.args[0]
                    assert target_args.my_drive is True
                    assert target_args.sharepoint_url is None
                    assert mock_test.call_count == 1

    def test_setup_can_skip_smoke_test(self):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = None
        args.force = False
        args.redirect_uri = None
        args.client_id = None
        args.tenant_id = None
        args.sharepoint_url = "https://mysliit.sharepoint.com/sites/Sharepointlocation"
        args.my_drive = False
        args.library = None
        args.folder = None
        args.skip_smoke_test = True

        with patch("rpa_sharepoint_connector.cli.cmd_configure") as mock_configure:
            with patch("rpa_sharepoint_connector.cli.cmd_set_target") as mock_target:
                with patch("rpa_sharepoint_connector.cli.cmd_test_upload") as mock_test:
                    with patch("builtins.print"):
                        cmd_setup(args)

                    assert mock_configure.call_count == 1
                    assert mock_target.call_count == 1
                    assert mock_test.call_count == 0


class TestRunCommand:
    """Test run command for bot-friendly operations."""

    def test_run_upload_calls_sdk_with_conflict(self):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = None
        args.op = "upload"
        args.local_path = "local.txt"
        args.remote_path = "Folder/local.txt"
        args.folder_path = None
        args.source_path = None
        args.target_path = None
        args.new_name = None
        args.conflict = "rename"
        args.json = False

        with patch("rpa_sharepoint_connector.cli.SharePointClient") as MockClient:
            mock_sp = MagicMock()
            mock_sp.upload.return_value = "item_123"
            MockClient.return_value = mock_sp

            with patch("builtins.print") as mock_print:
                cmd_run(args)

            mock_sp.upload.assert_called_once_with(
                "local.txt",
                "Folder/local.txt",
                conflict="rename",
            )
            output = "\n".join(str(c) for c in mock_print.call_args_list)
            assert "successful" in output.lower()

    def test_run_list_json_output(self):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = None
        args.op = "list"
        args.local_path = None
        args.remote_path = None
        args.folder_path = "Inbox"
        args.source_path = None
        args.target_path = None
        args.new_name = None
        args.conflict = "overwrite"
        args.json = True

        with patch("rpa_sharepoint_connector.cli.SharePointClient") as MockClient:
            mock_sp = MagicMock()
            mock_sp.list.return_value = [{"name": "a.txt", "is_folder": False}]
            MockClient.return_value = mock_sp

            with patch("builtins.print") as mock_print:
                cmd_run(args)

            output = "".join(call.args[0] for call in mock_print.call_args_list)
            payload = json.loads(output)
            assert payload["operation"] == "list"
            assert payload["success"] is True
            assert payload["count"] == 1

    def test_run_upload_missing_required_args_exits(self):
        args = MagicMock()
        args.profile = "demo"
        args.store_dir = None
        args.op = "upload"
        args.local_path = None
        args.remote_path = "Folder/a.txt"
        args.folder_path = None
        args.source_path = None
        args.target_path = None
        args.new_name = None
        args.conflict = "overwrite"
        args.json = False

        with patch("rpa_sharepoint_connector.cli.SharePointClient"):
            with pytest.raises(SystemExit):
                cmd_run(args)
