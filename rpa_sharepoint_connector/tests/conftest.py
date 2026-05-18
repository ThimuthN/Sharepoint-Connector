"""Shared test fixtures."""
import pytest
from unittest.mock import patch


@pytest.fixture
def mock_file_ops():
    """Auto-mock file operations for tests that don't create real files."""
    with patch('os.path.isfile', return_value=True):
        with patch('os.path.islink', return_value=False):
            with patch('os.path.getsize', return_value=1024):
                yield
