import os
from robot.api import logger

class KillAllProcess:
    # Kill all processes with the given names
    def kill_all_process(self, processes):
        for proc in processes:
            logger.console(f"Killing process: {proc}")
            logger.info(f"Killing process: {proc}")
            os.system(f'taskkill /f /im {proc}')