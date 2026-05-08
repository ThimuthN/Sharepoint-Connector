"""Microsoft OAuth authentication and token refresh using Device Code Flow."""
import httpx
import logging
import time
import os
from typing import Dict, Optional, Tuple
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

# Innobot-owned public-client app registration
# No client secret exists
MICROSOFT_CLIENT_ID = os.getenv(
    "MICROSOFT_CLIENT_ID",
    "4765d1f0-7a2e-4797-b3c8-5ce6e4a8c3a9"  # Innobot public client app
)


class MicrosoftAuth:
    """Handle Microsoft OAuth with Device Code Flow (no secrets)."""

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
    ):
        """Initialize auth client for Device Code Flow.

        Args:
            client_id: Microsoft client ID (defaults to Innobot public app)
            tenant_id: Azure tenant identifier.
                Common values:
                - "common": work/school + personal Microsoft accounts (default)
                - "organizations": work/school only
                - "consumers": personal Microsoft accounts only

        Note:
            No client_secret is used. This is a public-client Device Code Flow.
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
        """Extract OAuth error and description from response body."""
        try:
            payload = response.json()
            return payload.get("error", ""), payload.get("error_description", "")
        except Exception:
            return "", ""

    def start_device_flow(self) -> Dict:
        """Initiate device code flow.

        Returns:
            Dict with:
                device_code: Code used to poll for token
                user_code: Code user enters at verification_uri
                verification_uri: URL where user logs in
                verification_uri_complete: Full URL (if available)
                expires_in: Seconds until device code expires
                interval: Seconds to wait between polls
                message: User-friendly instructions from Microsoft

        Raises:
            ValueError: If device flow initiation fails
        """
        data = {
            "client_id": self.client_id,
            "scope": " ".join(self.scopes),
        }
        device_code_url = f"{self.authority}/devicecode"

        try:
            logger.info(
                "Starting device flow: tenant=%s endpoint=%s",
                self.tenant_id,
                device_code_url,
            )
            with httpx.Client() as client:
                response = client.post(
                    device_code_url,
                    data=data,
                    timeout=10.0
                )

                if not response.is_success:
                    error, description = self._extract_oauth_error(response)
                    logger.error(
                        "Device flow start failed: status=%s error=%s",
                        response.status_code,
                        error or "unknown",
                    )
                    raise ValueError(
                        f"Failed to start device flow: {error or response.status_code}. "
                        f"{description}".strip()
                    )

                result = response.json()

                logger.info(
                    "Device flow started: user_code=%s expires_in=%s interval=%s verification_uri=%s",
                    result.get("user_code"),
                    result.get("expires_in"),
                    result.get("interval"),
                    result.get("verification_uri"),
                )
                if result.get("verification_uri_complete"):
                    logger.info(
                        "Device flow verification_uri_complete available for one-click login."
                    )
                return result

        except httpx.HTTPError as e:
            logger.error(f"Failed to start device flow: {e}")
            raise ValueError(f"Failed to start device flow: {str(e)}")

    def poll_device_flow(
        self,
        device_code: str,
        interval: int,
        expires_in: int
    ) -> Dict:
        """Poll for device code completion.

        Args:
            device_code: Device code from start_device_flow
            interval: Seconds to wait between polls
            expires_in: Total seconds to poll before timeout

        Returns:
            Dict with access_token, refresh_token, expires_in

        Raises:
            ValueError: If polling fails or times out
        """
        data = {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": self.client_id,
            "device_code": device_code,
        }

        start_time = time.time()
        current_interval = interval
        token_url = f"{self.authority}/token"
        poll_count = 0

        logger.info(
            "Polling device flow started: tenant=%s endpoint=%s expires_in=%ss interval=%ss",
            self.tenant_id,
            token_url,
            expires_in,
            interval,
        )

        elapsed = 0
        while True:
            elapsed = int(time.time() - start_time)
            if elapsed >= expires_in:
                break

            try:
                poll_count += 1
                logger.info(
                    "Polling device flow: attempt=%s elapsed=%ss interval=%ss",
                    poll_count,
                    elapsed,
                    current_interval,
                )
                with httpx.Client() as client:
                    response = client.post(
                        token_url,
                        data=data,
                        timeout=10.0
                    )

                    # Check for error response
                    if not response.is_success:
                        try:
                            error_data = response.json()
                            error = error_data.get("error", "")
                            logger.info(
                                "Polling response: status=%s error=%s elapsed=%ss",
                                response.status_code,
                                error or "unknown",
                                elapsed,
                            )

                            if error == "authorization_pending":
                                # User hasn't signed in yet, keep polling
                                time.sleep(current_interval)
                                continue

                            elif error == "slow_down":
                                # Microsoft asks us to slow down polling
                                current_interval += 5
                                time.sleep(current_interval)
                                continue

                            elif error == "expired_token":
                                logger.error(
                                    "Device flow expired: elapsed=%ss expires_in=%ss",
                                    elapsed,
                                    expires_in,
                                )
                                raise ValueError(
                                    "Device code expired. "
                                    "Run: python -m rpa_sharepoint_connector configure"
                                )

                            elif error == "access_denied":
                                logger.error(
                                    "Authorization denied by user: elapsed=%ss",
                                    elapsed,
                                )
                                raise ValueError(
                                    "Authorization denied. "
                                    "Run: python -m rpa_sharepoint_connector configure"
                                )

                            else:
                                description = error_data.get("error_description", "")
                                raise ValueError(
                                    f"Authorization failed: {error}. {description}".strip()
                                )

                        except ValueError:
                            raise
                        except:
                            response.raise_for_status()

                    # Success
                    result = response.json()
                    logger.info(
                        "Device flow authentication successful: elapsed=%ss attempts=%s",
                        elapsed,
                        poll_count,
                    )
                    return result

            except httpx.HTTPError as e:
                logger.error(f"Device flow poll error: {e}")
                raise ValueError(f"Device flow error: {str(e)}")

        logger.error(
            "Device flow polling timed out: elapsed=%ss expires_in=%ss",
            elapsed,
            expires_in,
        )
        raise ValueError(
            "Device code polling timed out. "
            "Run: python -m rpa_sharepoint_connector configure"
        )

    def refresh_token(self, refresh_token: str) -> Dict:
        """Refresh access token using refresh token.

        Public client (no secret required).

        Args:
            refresh_token: Stored refresh token

        Returns:
            Dict with new access_token, refresh_token, expires_in

        Raises:
            ValueError: If refresh fails or token is invalid
        """
        data = {
            "client_id": self.client_id,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
            "scope": " ".join(self.scopes),
        }

        try:
            with httpx.Client() as client:
                response = client.post(f"{self.authority}/token", data=data, timeout=10.0)

                if not response.is_success:
                    try:
                        error_data = response.json()
                        if error_data.get("error") == "invalid_grant":
                            raise ValueError(
                                "Refresh token expired or invalid. "
                                "Run: python -m rpa_sharepoint_connector configure"
                            )
                    except ValueError:
                        raise
                    except:
                        pass

                response.raise_for_status()
                return response.json()

        except httpx.HTTPError as e:
            logger.error(f"Token refresh failed: {e}")
            raise ValueError(f"Failed to refresh token: {str(e)}")

    def get_user_info(self, access_token: str) -> Dict:
        """Get user info from Microsoft Graph.

        Args:
            access_token: Valid access token

        Returns:
            User info dict with id, mail, etc.

        Raises:
            ValueError: If request fails
        """
        try:
            with httpx.Client() as client:
                response = client.get(
                    "https://graph.microsoft.com/v1.0/me",
                    headers={"Authorization": f"Bearer {access_token}"},
                )
                response.raise_for_status()
                return response.json()
        except httpx.HTTPError as e:
            logger.error(f"Failed to get user info: {e}")
            raise ValueError(f"Failed to get user info: {str(e)}")

    def is_token_expired(self, expires_at: datetime) -> bool:
        """Check if token is expired or near expiry (5 min buffer).

        Args:
            expires_at: Token expiration datetime

        Returns:
            True if token needs refresh
        """
        return datetime.utcnow() >= expires_at - timedelta(minutes=5)
