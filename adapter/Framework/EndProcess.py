import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from robot.api import logger
from Resources.Config import Config
from Framework.CloseAllApplications import CloseAllApplications
from Framework.KillAllProcess import KillAllProcess
from Library.SaveRobotUsageTime import SaveRobotUsageTime

class EndProcess:
    # End process and close all applications used 
    def end_process(self, config: Config):      
        try:
            logger.console("Closing all applications...")
            logger.info("Closing all applications...")
            close_app = CloseAllApplications()
            close_app.close_all_applications()
            # Simulate writing log file
            #self.write_log_report()
            # Simulate moving files
            #self.move_files()
            # Simulate sending email
            #self.send_email()
        except Exception as e:
            logger.console(f"{self.__class__.__name__} -> Applications failed to close gracefully. {e}")
            logger.warn(f"{self.__class__.__name__} -> Applications failed to close gracefully. {e}")
            kill_process = KillAllProcess()
            kill_process.kill_all_process(['notepad.exe', 'calc.exe']) # Example process names to kill
        finally:
            # Save robot usage time with actual transaction counts
            save_robot_time = SaveRobotUsageTime()
            save_robot_time.save_usage_time(config.ProcessIDCode, config.Log_Environment, config.Log_Index, str(config.Log_Transactions), str(config.Log_Done), str(config.Log_Looping), config.Config.get("EmailOnlyWithDifferences"))