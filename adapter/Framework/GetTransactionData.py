import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from robot.api import logger
from Resources.Config import Config

class GetTransactionData:
    # Get transaction data
    def get_transaction_data(self, config: Config):        
        logger.console(config.Config["LogMessage_GetTransactionData"])
        logger.info(config.Config["LogMessage_GetTransactionData"])
        try:
            if config.TransactionNumber < config.Log_Transactions:
                config.TransactionItem = config.dt_TransactionData[config.TransactionNumber]
                config.TransactionNumber += 1
                logger.console(f"New transaction retrieved: {config.TransactionNumber}")
                logger.info(f"New transaction retrieved: {config.TransactionNumber}")
            else:
                config.TransactionItem = None
                logger.console("Process finished due to no more transaction data")
                logger.info("Process finished due to no more transaction data")
        except Exception as e:            
            config.SystemException = e
            logger.console(f"{config.Config['LogMessage_GetTransactionDataError']}{config.TransactionNumber}. {e}")
            logger.fatal(f"{config.Config['LogMessage_GetTransactionDataError']}{config.TransactionNumber}. {e}")
            config.TransactionItem = None
        try:
            if config.TransactionNumber == 0:
                config.TransactionItem = {"id": 1}
                config.TransactionNumber = 1
            else:
                config.TransactionItem = None
        except Exception as e:
            logger.console(f"Error: {e}")
            logger.error(f"Error: {e}")