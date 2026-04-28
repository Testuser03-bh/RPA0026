from robot.api import logger
from Resources.Config import Config

class InitAllApplications:
    # Open applications used in the process and do necessary initialization procedures (e.g., login)
    def init_all_applications(self, config: Config):
        logger.console("Opening applications...")
        logger.info("Opening applications...")

        #print(f"Applications initialized for process ID: {config.ProcessIDCode}")