import sys
import os
from datetime import datetime

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from Framework.Main import Main
from Resources.Config import Config
from robot.api import logger

class RobotProcessLibrary:
    """Library to call Python framework from Robot"""
    
    def __init__(self):
        self.main = Main()
        self.config = Config()
        self.start_time = None
        self.end_time = None
    
    def initialize_robot_process(self):
        """Call at start of test - creates database log entry with StartDateTime"""
        self.start_time = datetime.now()
        logger.console("=== ROBOT PROCESS INITIALIZED ===")
        logger.console(f"Process Start Time: {self.start_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}")
        
        # Initialize with Log_Transactions = 0
        self.config.Log_Transactions = 0
        self.config.Log_Done = 0
        self.config.Log_Looping = 0
        
        self.main.initialize(self.config)
    
    def get_config(self):
        """Return the config object so TransactionTracker can access it"""
        return self.config
    
    def run_transaction_process(self):
        """Run the main transaction processing - NOT NEEDED for keyword-level tracking"""
        logger.console("=== RUNNING TRANSACTION PROCESS ===")
        while self.config.TransactionItem is not None or self.config.TransactionNumber < self.config.Log_Transactions:
            self.main.get_transaction_data(self.config)
            if self.config.TransactionItem is not None:
                self.main.process_transaction(self.config)
            else:
                break
    
    def end_robot_process(self):
        """Call at end of test - updates database with EndDateTime and saves timing"""
        self.end_time = datetime.now()
        logger.console("=== ROBOT PROCESS ENDED ===")
        logger.console(f"Process End Time: {self.end_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}")
        
        # Calculate total time taken
        if self.start_time and self.end_time:
            total_seconds = (self.end_time - self.start_time).total_seconds()
            logger.console(f"Total Execution Time: {total_seconds} seconds")
        
        # Log all values being sent to database
        logger.console("=== VALUES SENT TO DATABASE ===")
        logger.console(f"Process ID: {self.config.ProcessIDCode}")
        logger.console(f"Environment: {self.config.Log_Environment}")
        logger.console(f"Log Index: {self.config.Log_Index}")
        logger.console(f"Total Transactions: {self.config.Log_Transactions}")
        logger.console(f"Completed: {self.config.Log_Done}")
        logger.console(f"Technically Checked: {self.config.Log_Looping}")
        logger.console("================================")
        
        self.main.end_process(self.config)