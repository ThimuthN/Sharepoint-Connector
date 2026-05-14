"""Run-command orchestration for bot-facing CLI operations."""
import json
import sys
from typing import Callable


def _run_upload_operation(sp, args) -> dict:
    """Execute upload operation and build bot-facing result payload."""
    if not args.local_path or not args.remote_path:
        raise ValueError("upload requires --local-path and --remote-path")

    item_id = sp.upload(args.local_path, args.remote_path, conflict=args.conflict)
    return {
        "operation": "upload",
        "success": True,
        "item_id": item_id,
        "remote_path": args.remote_path,
    }


def _run_download_operation(sp, args) -> dict:
    """Execute download operation and build bot-facing result payload."""
    if not args.remote_path or not args.local_path:
        raise ValueError("download requires --remote-path and --local-path")

    sp.download(args.remote_path, args.local_path)
    return {
        "operation": "download",
        "success": True,
        "local_path": args.local_path,
    }


def _run_list_operation(sp, args) -> dict:
    """Execute list operation and build bot-facing result payload."""
    folder_path = args.folder_path or ""
    items = sp.list(folder_path)
    return {
        "operation": "list",
        "success": True,
        "folder_path": folder_path,
        "count": len(items),
        "items": items,
    }


def _run_delete_operation(sp, args) -> dict:
    """Execute delete operation and build bot-facing result payload."""
    if not args.remote_path:
        raise ValueError("delete requires --remote-path")

    sp.delete(args.remote_path)
    return {
        "operation": "delete",
        "success": True,
        "remote_path": args.remote_path,
    }


def _run_move_operation(sp, args) -> dict:
    """Execute move operation and build bot-facing result payload."""
    if not args.source_path or not args.target_path:
        raise ValueError("move requires --source-path and --target-path")

    sp.move(args.source_path, args.target_path, new_name=args.new_name)
    return {
        "operation": "move",
        "success": True,
        "source_path": args.source_path,
        "target_path": args.target_path,
        "new_name": args.new_name or "",
    }


def _run_mkdir_operation(sp, args) -> dict:
    """Execute mkdir operation and build bot-facing result payload."""
    if not args.folder_path:
        raise ValueError("mkdir requires --folder-path")

    item_id = sp.mkdir(args.folder_path)
    return {
        "operation": "mkdir",
        "success": True,
        "folder_path": args.folder_path,
        "item_id": item_id,
    }


def _run_exists_operation(sp, args) -> dict:
    """Execute exists operation and build bot-facing result payload."""
    if not args.remote_path:
        raise ValueError("exists requires --remote-path")

    exists = sp.exists(args.remote_path)
    return {
        "operation": "exists",
        "success": True,
        "remote_path": args.remote_path,
        "exists": exists,
    }


RUN_OPERATION_HANDLERS = {
    "upload": _run_upload_operation,
    "download": _run_download_operation,
    "list": _run_list_operation,
    "delete": _run_delete_operation,
    "move": _run_move_operation,
    "mkdir": _run_mkdir_operation,
    "exists": _run_exists_operation,
}


def cmd_run(
    args,
    sharepoint_client_cls,
    classify_error: Callable[[Exception], dict],
    print_run_result: Callable[[dict], None],
    print_actionable_error: Callable[[Exception, str, str], None],
) -> None:
    """Run one SharePoint/OneDrive operation for automation bots."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir
    op = args.op
    as_json = bool(getattr(args, "json", False))
    sharepoint_url = getattr(args, "sharepoint_url", None)
    folder_url = getattr(args, "folder_url", None)

    # Handle --folder-url by extracting site URL and updating remote-path
    if folder_url:
        from urllib.parse import urlparse, unquote
        parsed = urlparse(folder_url)
        sharepoint_url = f"{parsed.scheme}://{parsed.hostname}"
        # Extract folder path from URL
        path = unquote(parsed.path or "")
        if path and path != "/":
            # Append folder path to remote-path
            folder_path = path.lstrip("/")
            if hasattr(args, "remote_path") and args.remote_path:
                args.remote_path = f"{folder_path}/{args.remote_path}"
            else:
                args.remote_path = folder_path

    sp = None

    try:
        sp = sharepoint_client_cls(profile=profile_name, store_dir=store_dir, sharepoint_url=sharepoint_url)
        handler = RUN_OPERATION_HANDLERS.get(op)
        if handler is None:
            raise ValueError(f"Unsupported operation: {op}")
        result = handler(sp, args)

        if as_json:
            print(json.dumps(result, indent=2))
        else:
            print_run_result(result)

    except Exception as error:
        if as_json:
            error_meta = classify_error(error)
            print(
                json.dumps(
                    {
                        "operation": op,
                        "success": False,
                        "error_code": error_meta["error_code"],
                        "retryable": error_meta["retryable"],
                        "error": str(error),
                    },
                    indent=2,
                )
            )
        else:
            print_actionable_error(error, command="run", profile=profile_name)
        sys.exit(1)
    finally:
        close = getattr(sp, "close", None)
        if callable(close):
            close()
