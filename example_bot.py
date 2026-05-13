"""Example RPA bot using SharePoint connector.

Before running this bot:
1. Configure SharePoint connection:
   python -m rpa_sharepoint_connector configure --profile client_a

2. Run the bot:
   python example_bot.py
"""
from rpa_sharepoint_connector import SharePointClient
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

WORK_DIR = Path("example_bot_artifacts")


def run_step(title: str, action) -> None:
    """Run one bot step and stop immediately if it fails."""
    logger.info(title)
    try:
        action()
    except Exception:
        logger.exception("Step failed")
        raise


def main():
    """Example bot workflow."""
    # Initialize client (uses stored token - no login required)
    sp = SharePointClient(profile="client_a")
    WORK_DIR.mkdir(exist_ok=True)
    upload_path = WORK_DIR / "test_upload.txt"
    download_path = WORK_DIR / "test_download.txt"

    logger.info("Starting SharePoint operations...")

    run_step("\n1. Listing files in default folder:", lambda: _list_files(sp))
    run_step("\n2. Creating test folder:", lambda: _create_folder(sp))
    run_step("\n3. Uploading test file:", lambda: _upload_file(sp, upload_path))
    run_step("\n4. Checking if file exists:", lambda: _check_exists(sp))
    run_step("\n5. Downloading file:", lambda: _download_file(sp, download_path))
    run_step("\n6. Moving file to Processed folder:", lambda: _move_file(sp))
    run_step("\n7. Deleting file:", lambda: _delete_file(sp))

    logger.info("\nBot completed.")


def _list_files(sp: SharePointClient) -> None:
    files = sp.list()
    for file_info in files:
        kind = "folder" if file_info["is_folder"] else "file"
        logger.info("   %s (%s)", file_info["name"], kind)


def _create_folder(sp: SharePointClient) -> None:
    sp.mkdir("RPA_Test/Processed")
    logger.info("   Created RPA_Test/Processed")


def _upload_file(sp: SharePointClient, upload_path: Path) -> None:
    upload_path.write_text(
        "This is a test file uploaded by the bot\n",
        encoding="utf-8",
    )
    sp.upload(str(upload_path), "RPA_Test/test_upload.txt")
    logger.info("   Uploaded %s", upload_path.name)


def _check_exists(sp: SharePointClient) -> None:
    exists = sp.exists("RPA_Test/test_upload.txt")
    logger.info("   File exists: %s", exists)


def _download_file(sp: SharePointClient, download_path: Path) -> None:
    sp.download("RPA_Test/test_upload.txt", str(download_path))
    content = download_path.read_text(encoding="utf-8")
    logger.info("   Downloaded: %s", content.strip())


def _move_file(sp: SharePointClient) -> None:
    sp.move("RPA_Test/test_upload.txt", "RPA_Test/Processed")
    logger.info("   Moved to RPA_Test/Processed")


def _delete_file(sp: SharePointClient) -> None:
    sp.delete("RPA_Test/Processed/test_upload.txt")
    logger.info("   Deleted")


if __name__ == "__main__":
    main()
