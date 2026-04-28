import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from robot.api import logger
from Resources.Config import Config

class Process:
    def process_data(self, config: Config):
        # Process data logic here"
        logger.console(f"Processing data... {config.TransactionNumber}{config.TransactionItem}")
        logger.info(f"Processing data... {config.TransactionNumber}{config.TransactionItem}")

        # To-Do: set the TransactionStatus based on actual processing results of your Try-Catch block
        config.TransactionStatus = "Success"
        #config.TransactionStatus = "BusinessException"
        #config.TransactionStatus = "SystemException"