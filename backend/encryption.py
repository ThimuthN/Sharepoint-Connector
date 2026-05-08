"""Token encryption/decryption utilities."""
import base64
from cryptography.fernet import Fernet
from typing import Tuple


class TokenEncryption:
    """Encrypt and decrypt tokens."""
    
    def __init__(self, key: str):
        """Initialize with encryption key.
        
        Args:
            key: Base64-encoded Fernet key
        """
        try:
            self.cipher = Fernet(key.encode())
        except Exception as e:
            raise ValueError(
                "Invalid TOKEN_ENCRYPTION_KEY. Must be valid base64 Fernet key. "
                "Generate with: python -c \"from cryptography.fernet import Fernet; "
                "print(Fernet.generate_key().decode())\""
            ) from e
    
    def encrypt(self, token: str) -> str:
        """Encrypt a token.
        
        Args:
            token: Plain text token
            
        Returns:
            Base64-encoded encrypted token
        """
        encrypted = self.cipher.encrypt(token.encode())
        return base64.b64encode(encrypted).decode()
    
    def decrypt(self, encrypted_token: str) -> str:
        """Decrypt a token.
        
        Args:
            encrypted_token: Base64-encoded encrypted token
            
        Returns:
            Plain text token
        """
        try:
            encrypted = base64.b64decode(encrypted_token.encode())
            decrypted = self.cipher.decrypt(encrypted)
            return decrypted.decode()
        except Exception as e:
            raise ValueError("Failed to decrypt token") from e
