"""Test configuration initialization."""
import pytest
import os


def test_config_validates_required_fields():
    """Config.validate() raises error for missing fields."""
    # Temporarily remove env vars
    saved = {}
    for key in ["MICROSOFT_CLIENT_ID", "MICROSOFT_CLIENT_SECRET", 
                "MICROSOFT_REDIRECT_URI", "TOKEN_ENCRYPTION_KEY"]:
        saved[key] = os.environ.pop(key, None)
    
    try:
        from config import Config
        Config.MICROSOFT_CLIENT_ID = None
        
        with pytest.raises(ValueError, match="Missing required configuration"):
            Config.validate()
    finally:
        # Restore env vars
        for key, value in saved.items():
            if value:
                os.environ[key] = value
