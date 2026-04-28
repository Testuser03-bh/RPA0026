import os
import pyodbc
import pandas as pd
import shutil
import time
import pyautogui
import pygetwindow as gw
from robot.api import logger
from datetime import datetime
from robot.api.deco import keyword
def open_latest_ica(timeout=60, poll_interval=2, max_age_seconds=180):
    """
    Waits for a NEW ICA file in Downloads.
    Discards files older than max_age_seconds (default 3 minutes).
    Opens the newest valid ICA file immediately.
    """

    logger.console("Waiting for new ICA file in Downloads...")

    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    start_time = time.time()

    while time.time() - start_time < timeout:

        now = time.time()

        ica_files = [
            os.path.join(downloads, f)
            for f in os.listdir(downloads)
            if f.lower().endswith(".ica")
        ]

        valid_files = []

        for file in ica_files:
            file_age = now - os.path.getctime(file)

            # Keep only files newer than 3 minutes
            if file_age <= max_age_seconds:
                valid_files.append(file)

        if valid_files:
            latest_file = max(valid_files, key=os.path.getctime)

            logger.console(
                f"Opening ICA file: {os.path.basename(latest_file)} "
                f"(age: {int(now - os.path.getctime(latest_file))} sec)"
            )

            os.startfile(latest_file)
            return True

        time.sleep(poll_interval)

    logger.error("Timeout: No new ICA file detected within allowed age.")
    return False