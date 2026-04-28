import os
import shutil
import tempfile

from robot.api import logger

class CreateDirectory:
    # Create a directory, optionally deleting it first if it exists
    def create_directory(self, in_strNewFileFolder=None, in_bDeleteFolderIfExist=False):
        """
        Creates a directory. Optionally deletes it first if it exists.
        Args:
            in_strNewFileFolder (str): The path to create. If None or empty, uses a default temp path.
            in_bDeleteFolderIfExist (bool): If True, deletes the folder if it exists before creating.
        Returns:
            out_strCompleteFolderPath (str): The full path of the created directory.
        """
        # Decide on the path
        if not in_strNewFileFolder or in_strNewFileFolder.strip() == "":
            strTempPath = os.path.join(tempfile.gettempdir(), "Voith_RPA")
        else:
            strTempPath = in_strNewFileFolder

        try:
            if in_bDeleteFolderIfExist:
                if os.path.exists(strTempPath):
                    try:
                        shutil.rmtree(strTempPath)
                        logger.console(f"Deleted existing folder: {strTempPath}")
                        logger.info(f"Deleted existing folder: {strTempPath}")
                    except Exception as e:
                        logger.console(f"The folder already exists and could not be cleared: {e}")
                        logger.info(f"The folder already exists and could not be cleared: {e}")
                        return None

            # Create the directory (if it doesn't exist)
            os.makedirs(strTempPath, exist_ok=True)
            logger.console(f"The path: {strTempPath}, has been created successfully")
            logger.info(f"The path: {strTempPath}, has been created successfully")
            out_strCompleteFolderPath = strTempPath
            return out_strCompleteFolderPath
        except Exception as e:
            logger.console(f"{self.__class__.__name__} -> The folder has not been created ({e})")
            logger.error(f"{self.__class__.__name__} -> The folder has not been created ({e})")
            return None