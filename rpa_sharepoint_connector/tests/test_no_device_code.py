"""Guardrail tests to prevent legacy device-code auth from returning."""
from pathlib import Path


def test_no_device_code_references_remain():
    """Connector source should not contain device-code auth references."""
    source_root = Path(__file__).resolve().parents[1]
    disallowed_patterns = [
        "start_device_flow",
        "poll_device_flow",
        "/devicecode",
        "Device Login Required",
        "Enter code:",
        "configure-browser",
        "cmd_configure_browser",
    ]

    violations = []
    for file_path in source_root.glob("*.py"):
        content = file_path.read_text(encoding="utf-8")
        for pattern in disallowed_patterns:
            if pattern in content:
                violations.append(f"{file_path.name}: {pattern}")

    assert not violations, f"Found disallowed legacy auth references: {violations}"
