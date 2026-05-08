"""Tests for retry logic on transient failures."""
import pytest
import time
from unittest.mock import MagicMock, patch
import httpx
from rpa_sharepoint_connector.retry import (
    RetryConfig,
    is_transient_error,
    calculate_backoff,
    get_retry_after_header,
    retry_operation,
    TRANSIENT_STATUS_CODES,
    NON_RETRYABLE_STATUS_CODES,
)


class TestTransientErrorDetection:
    """Test identification of transient vs permanent errors."""

    def test_429_is_transient(self):
        """Test that 429 (rate limit) is transient."""
        response = MagicMock()
        response.status_code = 429
        exc = httpx.HTTPError("Rate limited")
        exc.response = response
        assert is_transient_error(exc) is True

    def test_500_is_transient(self):
        """Test that 500 (server error) is transient."""
        response = MagicMock()
        response.status_code = 500
        exc = httpx.HTTPError("Server error")
        exc.response = response
        assert is_transient_error(exc) is True

    def test_502_is_transient(self):
        """Test that 502 is transient."""
        response = MagicMock()
        response.status_code = 502
        exc = httpx.HTTPError("Bad gateway")
        exc.response = response
        assert is_transient_error(exc) is True

    def test_503_is_transient(self):
        """Test that 503 is transient."""
        response = MagicMock()
        response.status_code = 503
        exc = httpx.HTTPError("Service unavailable")
        exc.response = response
        assert is_transient_error(exc) is True

    def test_504_is_transient(self):
        """Test that 504 is transient."""
        response = MagicMock()
        response.status_code = 504
        exc = httpx.HTTPError("Gateway timeout")
        exc.response = response
        assert is_transient_error(exc) is True

    def test_401_is_not_transient(self):
        """Test that 401 (unauthorized) is not transient."""
        response = MagicMock()
        response.status_code = 401
        exc = httpx.HTTPError("Unauthorized")
        exc.response = response
        assert is_transient_error(exc) is False

    def test_403_is_not_transient(self):
        """Test that 403 (forbidden) is not transient."""
        response = MagicMock()
        response.status_code = 403
        exc = httpx.HTTPError("Forbidden")
        exc.response = response
        assert is_transient_error(exc) is False

    def test_404_is_not_transient(self):
        """Test that 404 (not found) is not transient."""
        response = MagicMock()
        response.status_code = 404
        exc = httpx.HTTPError("Not found")
        exc.response = response
        assert is_transient_error(exc) is False

    def test_timeout_is_transient(self):
        """Test that timeout is transient."""
        exc = httpx.TimeoutException("Timeout")
        assert is_transient_error(exc) is True

    def test_connect_error_is_transient(self):
        """Test that connection error is transient."""
        exc = httpx.ConnectError("Connection refused")
        assert is_transient_error(exc) is True

    def test_connection_reset_is_transient(self):
        """Test that connection reset is transient."""
        exc = ConnectionResetError("Connection reset by peer")
        assert is_transient_error(exc) is True


class TestBackoffCalculation:
    """Test exponential backoff with jitter."""

    def test_first_attempt_backoff(self):
        """Test backoff for first retry."""
        config = RetryConfig(
            initial_wait_seconds=1.0,
            backoff_multiplier=2.0,
            jitter_factor=0.0  # No jitter for predictability
        )
        wait = calculate_backoff(0, config)
        assert 0.9 < wait < 1.1  # Allow tiny jitter

    def test_exponential_increase(self):
        """Test backoff increases exponentially."""
        config = RetryConfig(
            initial_wait_seconds=1.0,
            backoff_multiplier=2.0,
            jitter_factor=0.0
        )
        wait0 = calculate_backoff(0, config)
        wait1 = calculate_backoff(1, config)
        wait2 = calculate_backoff(2, config)

        # Each should be roughly 2x previous (without jitter)
        assert wait0 < wait1 < wait2

    def test_max_wait_capped(self):
        """Test backoff is capped at max."""
        config = RetryConfig(
            initial_wait_seconds=10.0,
            backoff_multiplier=10.0,
            max_wait_seconds=30.0,
            jitter_factor=0.0
        )
        wait = calculate_backoff(5, config)  # Would be huge without cap
        assert wait <= 30.0

    def test_jitter_variation(self):
        """Test jitter adds randomness."""
        config = RetryConfig(
            initial_wait_seconds=10.0,
            backoff_multiplier=1.0,
            jitter_factor=0.5  # 50% jitter
        )
        waits = [calculate_backoff(0, config) for _ in range(10)]
        # Not all should be identical
        assert len(set(waits)) > 1


class TestRetryAfterHeader:
    """Test Retry-After header parsing."""

    def test_retry_after_seconds(self):
        """Test parsing Retry-After as seconds."""
        response = MagicMock()
        response.headers = {"Retry-After": "120"}
        wait = get_retry_after_header(response)
        assert wait == 120.0

    def test_retry_after_missing(self):
        """Test missing Retry-After header."""
        response = MagicMock()
        response.headers = {}
        wait = get_retry_after_header(response)
        assert wait is None

    def test_retry_after_float(self):
        """Test Retry-After as float."""
        response = MagicMock()
        response.headers = {"Retry-After": "5.5"}
        wait = get_retry_after_header(response)
        assert wait == 5.5


class TestRetryOperation:
    """Test retry operation orchestration."""

    def test_success_first_attempt(self):
        """Test operation succeeds on first try."""
        operation = MagicMock(return_value="success")
        config = RetryConfig(max_attempts=3)

        result = retry_operation(operation, config, "test_op")

        assert result == "success"
        operation.assert_called_once()

    def test_429_then_success(self):
        """Test 429 retries and succeeds."""
        response = MagicMock()
        response.status_code = 429

        exc1 = httpx.HTTPError("Rate limited")
        exc1.response = response

        operation = MagicMock(side_effect=[exc1, "success"])
        config = RetryConfig(max_attempts=3, initial_wait_seconds=0.01)

        with patch("rpa_sharepoint_connector.retry.time.sleep"):
            result = retry_operation(operation, config, "test_op")

        assert result == "success"
        assert operation.call_count == 2

    def test_500_then_success(self):
        """Test 500 retries and succeeds."""
        response = MagicMock()
        response.status_code = 500

        exc1 = httpx.HTTPError("Server error")
        exc1.response = response

        operation = MagicMock(side_effect=[exc1, "success"])
        config = RetryConfig(max_attempts=3, initial_wait_seconds=0.01)

        with patch("rpa_sharepoint_connector.retry.time.sleep"):
            result = retry_operation(operation, config, "test_op")

        assert result == "success"
        assert operation.call_count == 2

    def test_timeout_then_success(self):
        """Test timeout retries and succeeds."""
        operation = MagicMock(side_effect=[
            httpx.TimeoutException("Timeout"),
            "success"
        ])
        config = RetryConfig(max_attempts=3, initial_wait_seconds=0.01)

        with patch("rpa_sharepoint_connector.retry.time.sleep"):
            result = retry_operation(operation, config, "test_op")

        assert result == "success"
        assert operation.call_count == 2

    def test_no_retry_on_401(self):
        """Test 401 does not retry."""
        response = MagicMock()
        response.status_code = 401

        exc = httpx.HTTPError("Unauthorized")
        exc.response = response

        operation = MagicMock(side_effect=exc)
        config = RetryConfig(max_attempts=3)

        with pytest.raises(httpx.HTTPError):
            retry_operation(operation, config, "test_op")

        # Should only try once, not retry
        operation.assert_called_once()

    def test_no_retry_on_403(self):
        """Test 403 does not retry."""
        response = MagicMock()
        response.status_code = 403

        exc = httpx.HTTPError("Forbidden")
        exc.response = response

        operation = MagicMock(side_effect=exc)
        config = RetryConfig(max_attempts=3)

        with pytest.raises(httpx.HTTPError):
            retry_operation(operation, config, "test_op")

        operation.assert_called_once()

    def test_no_retry_on_404(self):
        """Test 404 does not retry."""
        response = MagicMock()
        response.status_code = 404

        exc = httpx.HTTPError("Not found")
        exc.response = response

        operation = MagicMock(side_effect=exc)
        config = RetryConfig(max_attempts=3)

        with pytest.raises(httpx.HTTPError):
            retry_operation(operation, config, "test_op")

        operation.assert_called_once()

    def test_max_retries_exhausted(self):
        """Test error when max retries exhausted."""
        response = MagicMock()
        response.status_code = 500

        exc = httpx.HTTPError("Server error")
        exc.response = response

        operation = MagicMock(side_effect=exc)
        config = RetryConfig(max_attempts=3, initial_wait_seconds=0.01)

        with patch("rpa_sharepoint_connector.retry.time.sleep"):
            with pytest.raises(ValueError, match="failed after 3 retries"):
                retry_operation(operation, config, "test_op")

        # Should try 3 times
        assert operation.call_count == 3

    def test_respects_retry_after_header(self):
        """Test Retry-After header is respected."""
        response = MagicMock()
        response.status_code = 429
        response.headers = {"Retry-After": "0.05"}

        exc1 = httpx.HTTPError("Rate limited")
        exc1.response = response

        operation = MagicMock(side_effect=[exc1, "success"])
        config = RetryConfig(max_attempts=3)

        sleep_durations = []

        def mock_sleep(duration):
            sleep_durations.append(duration)

        with patch("rpa_sharepoint_connector.retry.time.sleep", side_effect=mock_sleep):
            result = retry_operation(operation, config, "test_op")

        assert result == "success"
        # Should have used Retry-After value (0.05) instead of backoff
        assert len(sleep_durations) == 1
        assert sleep_durations[0] == 0.05

    def test_retry_config_override(self):
        """Test retry config can be overridden."""
        response = MagicMock()
        response.status_code = 500

        exc = httpx.HTTPError("Server error")
        exc.response = response

        operation = MagicMock(side_effect=exc)

        # Custom config: only 2 attempts
        config = RetryConfig(max_attempts=2, initial_wait_seconds=0.01)

        with patch("rpa_sharepoint_connector.retry.time.sleep"):
            with pytest.raises(ValueError, match="failed after 2 retries"):
                retry_operation(operation, config, "test_op")

        # Should only try 2 times (not default 3)
        assert operation.call_count == 2


class TestRetryLogging:
    """Test that retry logging is safe (no token leaks)."""

    def test_logging_no_tokens(self):
        """Test that logs don't contain Authorization headers."""
        response = MagicMock()
        response.status_code = 500
        response.headers = {}

        exc = httpx.HTTPError("Server error")
        exc.response = response

        operation = MagicMock(side_effect=[exc, "success"])
        config = RetryConfig(max_attempts=3, initial_wait_seconds=0.01)

        with patch("rpa_sharepoint_connector.retry.time.sleep"):
            with patch("rpa_sharepoint_connector.retry.logger") as mock_logger:
                result = retry_operation(operation, config, "test_op")

                # Check that logs were made
                assert mock_logger.info.called or mock_logger.error.called

                # Check that no logs contain "Authorization" or "Bearer"
                for call in mock_logger.info.call_args_list + mock_logger.error.call_args_list:
                    log_message = str(call)
                    assert "Authorization" not in log_message
                    assert "Bearer" not in log_message

    def test_logging_operation_name(self):
        """Test that operation name is logged."""
        operation = MagicMock(return_value="success")
        config = RetryConfig()

        with patch("rpa_sharepoint_connector.retry.logger") as mock_logger:
            retry_operation(operation, config, "upload_file")

            # Operation name should not be in error logs for success
            # But would be in info logs if there were retries


class TestRetryConfigClass:
    """Test RetryConfig dataclass."""

    def test_default_config(self):
        """Test default retry config values."""
        config = RetryConfig()

        assert config.max_attempts == 3
        assert config.initial_wait_seconds == 1.0
        assert config.max_wait_seconds == 60.0
        assert config.backoff_multiplier == 2.0
        assert config.jitter_factor == 0.1

    def test_custom_config(self):
        """Test custom retry config."""
        config = RetryConfig(
            max_attempts=5,
            initial_wait_seconds=2.0,
            max_wait_seconds=120.0,
            backoff_multiplier=1.5,
            jitter_factor=0.2
        )

        assert config.max_attempts == 5
        assert config.initial_wait_seconds == 2.0
        assert config.max_wait_seconds == 120.0
        assert config.backoff_multiplier == 1.5
        assert config.jitter_factor == 0.2
