"""Output and error-formatting helpers for the CLI."""
import sys


def classify_error_for_json(error: Exception) -> dict:
    """Classify runtime errors into stable bot-facing codes."""
    raw = str(error)
    lower = raw.lower()

    if "requires --" in lower or "unsupported operation" in lower:
        return {"error_code": "INVALID_INPUT", "retryable": False}
    if "profile '" in lower and "not found" in lower:
        return {"error_code": "PROFILE_NOT_FOUND", "retryable": False}
    if "refresh token expired or invalid" in lower:
        return {"error_code": "AUTH_EXPIRED", "retryable": False}
    if "unauthorized" in lower or "token may be expired" in lower:
        return {"error_code": "AUTH_FAILED", "retryable": False}
    if "forbidden" in lower or "permission" in lower:
        return {"error_code": "PERMISSION_DENIED", "retryable": False}
    if "not found" in lower:
        return {"error_code": "NOT_FOUND", "retryable": False}
    if "rate limited" in lower or "429" in lower:
        return {"error_code": "RATE_LIMITED", "retryable": True}
    if "timed out" in lower or "timeout" in lower:
        return {"error_code": "TIMEOUT", "retryable": True}
    if "failed to refresh token" in lower:
        return {"error_code": "AUTH_REFRESH_FAILED", "retryable": False}
    if "connection" in lower or "network" in lower:
        return {"error_code": "NETWORK_ERROR", "retryable": True}

    return {"error_code": "OPERATION_FAILED", "retryable": False}


def print_actionable_error(
    error: Exception,
    command: str,
    profile: str = "default",
) -> None:
    """Print actionable error details with targeted recovery hints."""
    raw = str(error)
    lower = raw.lower()
    hints = []

    if "unauthorized_client" in lower and "consumers" in lower:
        hints.append(
            "Azure app is not enabled for personal accounts. In App Registration set "
            "Supported account types to 'Any Entra ID tenant + Personal Microsoft accounts', "
            "then retry with --tenant-id common."
        )
    if "access_denied" in lower and command in ("configure", "setup"):
        hints.append(
            "Sign-in or consent was denied/cancelled. Retry and click Accept, and keep "
            "the terminal open until callback completes."
        )
    if "oauth callback timed out" in lower:
        hints.append(
            "Open the printed Go to URL immediately and finish sign-in before timeout."
        )
    if "failed to start callback server" in lower or "refused to connect" in lower:
        hints.append(
            "Local callback failed. Close old login tabs and retry without VPN/proxy/adblock."
        )
    if "profile '" in lower and "not found" in lower:
        hints.append(
            f"Profile missing. Run setup first: "
            f"python -m rpa_sharepoint_connector setup --profile {profile} --my-drive"
        )
    if "refresh token expired or invalid" in lower:
        hints.append(
            f"Profile needs reconnection. Run: "
            f"python -m rpa_sharepoint_connector setup --profile {profile} --force --my-drive"
        )
    if "forbidden" in lower:
        hints.append(
            "Account lacks permission to the selected site/library/folder."
        )

    print(f"ERROR: {raw}")
    for idx, hint in enumerate(hints, 1):
        print(f"HINT {idx}: {hint}")
    sys.exit(1)


def print_run_result(result: dict) -> None:
    """Print human-readable command result."""
    operation = result["operation"]
    print(f"OK {operation} successful")
    if operation == "list":
        for item in result["items"]:
            kind = "folder" if item.get("is_folder") else "file"
            print(f"- {item.get('name')} ({kind})")
    elif operation == "exists":
        print(f"Exists: {result['exists']}")
