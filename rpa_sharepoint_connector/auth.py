"""Runtime Microsoft auth helpers for token refresh and Graph profile lookup."""
import logging
import os
from datetime import datetime, timedelta
from typing import Dict, Optional, Tuple

import httpx

logger = logging.getLogger(__name__)

# Innobot-owned public-client app registration
MICROSOFT_CLIENT_ID = os.getenv(
    "MICROSOFT_CLIENT_ID",
    "4765d1f0-7a2e-4797-b3c8-5ce6e4a8c3a9",
)


class MicrosoftAuth:
    """Handle runtime token refresh for public-client profiles."""

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
    ):
        """Initialize Microsoft auth client.

        Args:
            client_id: Microsoft client ID (defaults to Innobot public app)
            tenant_id: Microsoft tenant identifier
        """
        self.client_id = client_id or MICROSOFT_CLIENT_ID
        self.tenant_id = tenant_id or os.getenv("MICROSOFT_TENANT_ID", "common")
        self.scopes = [
            "offline_access",
            "User.Read",
            "Sites.ReadWrite.All",
            "Files.ReadWrite.All",
        ]
        self.authority = (
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0"
        )

    @staticmethod
    def _extract_oauth_error(response: httpx.Response) -> Tuple[str, str]:
        """Extract OAuth error details from a response body."""
        try:
            payload = response.json()
            return payload.get("error", ""), payload.get("error_description", "")
        except Exception:
            return "", ""

    def refresh_token(self, refresh_token: str) -> Dict:
        """Refresh access token using a refresh token.

        Public-client refresh does not use client secret.
        """
        data = {
            "client_id": self.client_id,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
            "scope": " ".join(self.scopes),
        }

        token_url = f"{self.authority}/token"

        try:
            with httpx.Client() as client:
                response = client.post(token_url, data=data, timeout=15.0)

                if not response.is_success:
                    error, description = self._extract_oauth_error(response)
                    if error == "invalid_grant":
                        raise ValueError(
                            "Refresh token expired or invalid. "
                            "Run: python -m rpa_sharepoint_connector configure"
                        )
                    raise ValueError(
                        f"Failed to refresh token: {error or response.status_code}. "
                        f"{description}".strip()
                    )

                return response.json()

        except httpx.HTTPError as exc:
            logger.error(f"Token refresh failed: {exc}")
            raise ValueError(f"Failed to refresh token: {exc}") from exc

    def get_user_info(self, access_token: str) -> Dict:
        """Get current user profile from Microsoft Graph."""
        try:
            with httpx.Client() as client:
                response = client.get(
                    "https://graph.microsoft.com/v1.0/me",
                    headers={"Authorization": f"Bearer {access_token}"},
                    timeout=15.0,
                )
                response.raise_for_status()
                return response.json()
        except httpx.HTTPError as exc:
            logger.error(f"Failed to get user info: {exc}")
            raise ValueError(f"Failed to get user info: {exc}") from exc

    def is_token_expired(self, expires_at: datetime) -> bool:
        """Check whether token is expired or near expiry (5 minute buffer)."""
        return datetime.utcnow() >= expires_at - timedelta(minutes=5)
