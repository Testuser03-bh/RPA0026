import sys
import os
import pyautogui

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from robot.api import logger
from Resources.Config import Config
from Framework.KillAllProcess import KillAllProcess
from Library.InitAllSettingsSQL import InitAllSettingsSQL
from Library.SaveRobotUsageTime import SaveRobotUsageTime
from Library.CreateDirectory import CreateDirectory
from Framework.InitAllApplications import InitAllApplications

class Initialization:
    # Initialization of the process
    def initialize(self, config: Config):
        logger.console("Initializing settings and applications...")
        logger.info("Initializing settings and applications...")

        # Initializing settings and applications       
        try:            
            # Load configuration and open applications            
            config.SystemException = None             
            
            if config:
                # To-Do: Get the Process ID Code from Robot Framework variable
                config.ProcessIDCode = "RPA0024-VTI"    #Define the Process ID Code here

                # Get all settings from the database            
                init_settings = InitAllSettingsSQL()
                config.Config = init_settings.get_all_settings(config.ProcessIDCode)

                # Save robot usage time - Create log entry
                save_robot_time = SaveRobotUsageTime()
                config.Log_Index = save_robot_time.save_usage_time(config.ProcessIDCode, config.Log_Environment, "", "0", "0", "0", config.Config.get("EmailOnlyWithDifferences"))

                logger.console(f"Starting {config.ProcessIDCode}")
                logger.info(f"Starting {config.ProcessIDCode}")

                logger.console(f"Screen Resolution: {self.get_screen_resolution()}")
                logger.info(f"Screen Resolution: {self.get_screen_resolution()}")

                # Create necessary directories
                create_dir = CreateDirectory()
                config.Config["TempFolder"] = create_dir.create_directory(config.Config.get("TempFolder"), in_bDeleteFolderIfExist=True) + "/"

                # DON'T set Log_Transactions here - it will be counted by TransactionTracker
                # Build Data Table if needed for reference
                config.dt_TransactionData = []
                config.dt_TransactionData.append({"Column1": 1, "Column2": "row 1"})
                config.dt_TransactionData.append({"Column1": 2, "Column2": "row 2"})
                
                # Initialize transaction counters to 0
                if not hasattr(config, 'Log_Transactions') or config.Log_Transactions is None:
                    config.Log_Transactions = 0
                if not hasattr(config, 'Log_Done') or config.Log_Done is None:
                    config.Log_Done = 0
                if not hasattr(config, 'Log_Looping') or config.Log_Looping is None:
                    config.Log_Looping = 0

                logger.console(f"Transaction counters initialized (will be incremented by keyword execution)")
                logger.info(f"Transaction counters initialized")
                
                # Kill all processes if needed
                kill_process = KillAllProcess()   
                kill_process.kill_all_process(['sap.exe']) # Example process names to kill 

                # Add log fields
                config.Config["logF_BusinessProcessName"] = config.Config.get("logF_BusinessProcessName")

                # Initialize all applications
                init_apps = InitAllApplications()
                init_apps.init_all_applications(config)

            # Check max consecutive system exceptions
            max_exceptions = int(config.Config.get("MaxConsecutiveSystemExceptions", 3))
            if max_exceptions > 0 and config.ConsecutiveSystemExceptions >= max_exceptions:
                raise Exception(config.Config.get("ExceptionMessage_ConsecutiveErrors", "Max consecutive errors") +
                                f" Consecutive retry number: {config.ConsecutiveSystemExceptions + 1}") 
                
        except Exception as e:
            self.SystemException = e
            logger.console(f"{self.__class__.__name__} -> System exception at initialization: {e}")
            logger.error(f"{self.__class__.__name__} -> System exception at initialization: {e}")
            config.Config["ShouldMarkJobAsFaulted"] = True  

    def get_screen_resolution(self):
        width, height = pyautogui.size()
        return width, height

class GetTransactionData:
    def get_transaction_data(self, config: Config):
        try:
            if config.TransactionNumber == 0:
                config.TransactionItem = {"id": 1}
                config.TransactionNumber = 1
            else:
                config.TransactionItem = None
        except Exception as e:
            logger.console(f"Error: {e}")
            logger.error(f"Error: {e}")

class ProcessTransaction:
    def process_data(self, config: Config):
        try:
            config.Log_Done += 1
            config.Log_Looping += 1
        except Exception as e:
            logger.console(f"Error: {e}")
            logger.error(f"Error: {e}")