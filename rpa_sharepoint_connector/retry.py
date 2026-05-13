"""Retry layer for transient failures in Microsoft Graph operations."""
import time
import random
import logging
from typing import Callable, TypeVar, Optional, Set
from dataclasses import dataclass
import httpx

logger = logging.getLogger(__name__)

T = TypeVar("T")

# Transient error codes that should be retried
TRANSIENT_STATUS_CODES = {429, 500, 502, 503, 504}

# Error codes that should NOT be retried (client errors)
NON_RETRYABLE_STATUS_CODES = {400, 401, 403, 404}


@dataclass
class RetryConfig:
    """Configuration for retry behavior."""

    max_attempts: int = 3
    initial_wait_seconds: float = 1.0
    max_wait_seconds: float = 60.0
    backoff_multiplier: float = 2.0
    jitter_factor: float = 0.1  # 10% jitter


def is_transient_error(exc: Exception, status_code: Optional[int] = None) -> bool:
    """Check if exception is a transient error worth retrying.

    Args:
        exc: Exception that occurred
        status_code: HTTP status code if available

    Returns:
        True if error is transient and should be retried
    """
    # Check HTTP status codes first
    if isinstance(exc, httpx.HTTPError):
        if hasattr(exc, "response") and exc.response is not None:
            return exc.response.status_code in TRANSIENT_STATUS_CODES
        # Network errors on HTTP calls
        if isinstance(exc, (httpx.TimeoutException, httpx.ConnectError)):
            return True
        # Other HTTP errors might not be retryable
        return False

    # Handle timeout errors
    if isinstance(exc, (TimeoutError, httpx.TimeoutException)):
        return True

    # Connection errors
    if isinstance(exc, (ConnectionError, ConnectionResetError, httpx.ConnectError)):
        return True

    # Check explicit status code parameter
    if status_code is not None:
        return status_code in TRANSIENT_STATUS_CODES

    return False


def calculate_backoff(
    attempt: int, config: RetryConfig
) -> float:
    """Calculate wait time with exponential backoff and jitter.

    Args:
        attempt: Current attempt number (0-indexed)
        config: Retry configuration

    Returns:
        Seconds to wait before next attempt
    """
    # Exponential backoff: initial * (multiplier ^ attempt)
    wait = config.initial_wait_seconds * (config.backoff_multiplier ** attempt)

    # Cap at max
    wait = min(wait, config.max_wait_seconds)

    # Add jitter: ±10% of calculated wait
    jitter = wait * config.jitter_factor * (2 * random.random() - 1)
    wait = wait + jitter

    # Ensure positive
    return max(0.1, wait)


def get_retry_after_header(response: httpx.Response) -> Optional[float]:
    """Extract Retry-After value from response headers.

    Returns:
        Seconds to wait, or None if header not present
    """
    retry_after = response.headers.get("Retry-After")
    if not retry_after:
        return None

    try:
        # Try parsing as seconds (integer)
        return float(retry_after)
    except ValueError:
        # Could be HTTP-date format, ignore for now
        return None


def retry_operation(
    operation: Callable[[], T],
    config: RetryConfig,
    operation_name: str = "operation",
) -> T:
    """Execute operation with retry logic.

    Args:
        operation: Callable that performs the operation
        config: Retry configuration
        operation_name: Name for logging

    Returns:
        Result of operation

    Raises:
        ValueError: If all retry attempts exhausted
        Exception: Original exception if not transient
    """
    last_exception = None
    last_status_code = None

    for attempt in range(config.max_attempts):
        try:
            return operation()

        except httpx.HTTPError as e:
            # Extract status code if available
            if hasattr(e, "response") and e.response is not None:
                last_status_code = e.response.status_code

                # Don't retry non-retryable status codes
                if last_status_code in NON_RETRYABLE_STATUS_CODES:
                    logger.debug(
                        f"Not retrying {operation_name}: status {last_status_code} is not retryable"
                    )
                    raise

                # Check if transient
                if not is_transient_error(e, last_status_code):
                    logger.debug(
                        f"Not retrying {operation_name}: status {last_status_code} is not transient"
                    )
                    raise

            elif not is_transient_error(e):
                logger.debug(f"Not retrying {operation_name}: error is not transient")
                raise

            last_exception = e

        except Exception as e:
            # Check if it's a transient network error
            if not is_transient_error(e):
                logger.debug(f"Not retrying {operation_name}: error is not transient")
                raise

            last_exception = e

        # Determine wait time
        wait_seconds = None

        # Check Retry-After header if available
        if (
            isinstance(last_exception, httpx.HTTPError)
            and hasattr(last_exception, "response")
            and last_exception.response is not None
        ):
            wait_seconds = get_retry_after_header(last_exception.response)

        # Fall back to exponential backoff
        if wait_seconds is None:
            wait_seconds = calculate_backoff(attempt, config)

        # Don't sleep after last attempt
        if attempt < config.max_attempts - 1:
            logger.info(
                f"Retrying {operation_name}: attempt {attempt + 1}/{config.max_attempts}, "
                f"status={last_status_code}, wait={wait_seconds:.2f}s"
            )
            time.sleep(wait_seconds)

    # All retries exhausted
    if last_exception:
        logger.error(
            f"Exhausted retries for {operation_name}: "
            f"{config.max_attempts} attempts failed, last status={last_status_code}"
        )
        raise ValueError(
            f"Operation failed after {config.max_attempts} retries: "
            f"{operation_name}. Last error: {last_exception}"
        ) from last_exception

    raise ValueError(f"Unexpected error in retry logic for {operation_name}")
