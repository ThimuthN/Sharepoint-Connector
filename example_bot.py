"""Example RPA bot using SharePoint connector.

Before running this bot:
1. Configure SharePoint connection:
   python -m rpa_sharepoint_connector configure --profile client_a

2. Run the bot:
   python example_bot.py
"""
from rpa_sharepoint_connector import SharePointClient
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def main():
    """Example bot workflow."""
    # Initialize client (uses stored token - no login required)
    sp = SharePointClient(profile="client_a")

    logger.info("Starting SharePoint operations...")

    # List files in default folder
    logger.info("\n1. Listing files in default folder:")
    try:
        files = sp.list()
        for f in files:
            logger.info(f"   {f['name']} ({'folder' if f['is_folder'] else 'file'})")
    except Exception as e:
        logger.error(f"   Error: {e}")

    # Create a folder
    logger.info("\n2. Creating test folder:")
    try:
        sp.mkdir("RPA_Test/Processed")
        logger.info("   Created RPA_Test/Processed")
    except Exception as e:
        logger.error(f"   Error: {e}")

    # Upload a sample file
    logger.info("\n3. Uploading test file:")
    try:
        # Create a test file
        with open("test_upload.txt", "w") as f:
            f.write("This is a test file uploaded by the bot\n")

        sp.upload("test_upload.txt", "RPA_Test/test_upload.txt")
        logger.info("   Uploaded test_upload.txt")
    except Exception as e:
        logger.error(f"   Error: {e}")

    # Check if file exists
    logger.info("\n4. Checking if file exists:")
    try:
        exists = sp.exists("RPA_Test/test_upload.txt")
        logger.info(f"   File exists: {exists}")
    except Exception as e:
        logger.error(f"   Error: {e}")

    # Download the file
    logger.info("\n5. Downloading file:")
    try:
        sp.download("RPA_Test/test_upload.txt", "test_download.txt")
        with open("test_download.txt", "r") as f:
            content = f.read()
        logger.info(f"   Downloaded: {content}")
    except Exception as e:
        logger.error(f"   Error: {e}")

    # Move file
    logger.info("\n6. Moving file to Processed folder:")
    try:
        sp.move("RPA_Test/test_upload.txt", "RPA_Test/Processed")
        logger.info("   Moved to RPA_Test/Processed")
    except Exception as e:
        logger.error(f"   Error: {e}")

    # Delete file
    logger.info("\n7. Deleting file:")
    try:
        sp.delete("RPA_Test/Processed/test_upload.txt")
        logger.info("   Deleted")
    except Exception as e:
        logger.error(f"   Error: {e}")

    logger.info("\nBot completed.")


if __name__ == "__main__":
    main()
