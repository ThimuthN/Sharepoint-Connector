"""Configuration management for Microsoft SharePoint connector."""
import os
from typing import Optional
from dotenv import load_dotenv

# Load .env file
load_dotenv()


class Config:
    """Configuration from environment variables."""
    
    # Microsoft OAuth
    MICROSOFT_CLIENT_ID: str = os.getenv("MICROSOFT_CLIENT_ID")
    MICROSOFT_CLIENT_SECRET: str = os.getenv("MICROSOFT_CLIENT_SECRET")
    MICROSOFT_TENANT_ID: str = os.getenv("MICROSOFT_TENANT_ID", "common")
    MICROSOFT_REDIRECT_URI: str = os.getenv("MICROSOFT_REDIRECT_URI")
    
    # Token encryption
    TOKEN_ENCRYPTION_KEY: Optional[str] = os.getenv("TOKEN_ENCRYPTION_KEY")
    
    # App config
    APP_BASE_URL: str = os.getenv("APP_BASE_URL", "http://localhost:3000")
    DATABASE_URL: str = os.getenv("DATABASE_URL", "sqlite:///./sharepoint.db")
    ENVIRONMENT: str = os.getenv("ENVIRONMENT", "development")
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "info")
    
    # Microsoft Graph endpoints
    MICROSOFT_AUTHORITY = f"https://login.microsoftonline.com/{MICROSOFT_TENANT_ID}/oauth2/v2.0"
    MICROSOFT_AUTHORIZE_URL = f"{MICROSOFT_AUTHORITY}/authorize"
    MICROSOFT_TOKEN_URL = f"{MICROSOFT_AUTHORITY}/token"
    MICROSOFT_GRAPH_API = "https://graph.microsoft.com/v1.0"
    
    # OAuth scopes
    OAUTH_SCOPES = [
        "offline_access",
        "User.Read",
        "Sites.Read.All",
        "Files.Read.All",
    ]
    
    @classmethod
    def validate(cls) -> None:
        """Validate required configuration is present."""
        required = [
            "MICROSOFT_CLIENT_ID",
            "MICROSOFT_CLIENT_SECRET",
            "MICROSOFT_REDIRECT_URI",
            "TOKEN_ENCRYPTION_KEY",
        ]
        missing = [
            key for key in required
            if not getattr(cls, key, None)
        ]
        if missing:
            raise ValueError(
                f"Missing required configuration: {', '.join(missing)}\n"
                f"See .env.example for template."
            )


def get_config() -> Config:
    """Get validated configuration."""
    Config.validate()
    return Config
