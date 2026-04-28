from robot.api import logger
from Resources.Config import Config
from Framework.CloseAllApplications import CloseAllApplications
from Framework.KillAllProcess import KillAllProcess

class BusinessRuleException(Exception):
    pass # This class has no extra code, just inherits from Exception

# Increment transaction index and reset retries
class SetTransactionStatus:   
    def set_transaction_status(self, config: Config, 
                   success: bool, 
                   business_exception: bool = False, 
                   system_exception: bool = False, 
                   error_message: str = "",
                   in_BusinessException: BusinessRuleException = None,
                   in_TransactionField1: str = "",
                   in_TransactionField2: str = "", 
                   in_TransactionID: str = "",
                   in_SystemException: Exception = None, 
                   in_TransactionItem = {}, 
                   io_RetryNumber: int = 0, 
                   io_TransactionNumber: int = 0, 
                   io_ConsecutiveSystemExceptions: int = 0
                ):
        try:
            if success:
                self.status = "Success"
                logger.console("Transaction processed successfully.")
                logger.info("Transaction processed successfully.")
                io_TransactionNumber += 1
                io_RetryNumber = 0
                io_ConsecutiveSystemExceptions = 0
            elif business_exception:
                self.status = "BusinessException"
                self.error_message = error_message or "Business rule exception occurred."
                logger.console(f"SetTransactionStatus failed: {self.error_message}")
                logger.error(f"SetTransactionStatus failed: {self.error_message}")
                io_TransactionNumber += 1
                io_RetryNumber = 0
                io_ConsecutiveSystemExceptions = 0
            elif system_exception:
                self.status = "SystemException"
                self.error_message = error_message or "System exception occurred."
                logger.console(f"SetTransactionStatus failed: {self.error_message}")
                logger.error(f"SetTransactionStatus failed: {self.error_message}")

                close_app = CloseAllApplications()
                close_app.close_all_applications()

                kill_process = KillAllProcess()
                kill_process.kill_all_process(['notepad.exe', 'calc.exe']) # Example processes to kill
            else:
                # Fallback to Unknown exception
                status = "UnknownException"
                logger.console("Unknown exception type, defaulting to SystemException.")
                logger.error("Unknown exception type, defaulting to SystemException.")
        except Exception as e:
            logger.console(f"SetTransactionStatus failed: {e}")
            logger.error(f"SetTransactionStatus failed: {e}")
        finally:
            # Update config with the new status and other info
            config.TransactionStatus = self.status
            config.TransactionField1 = in_TransactionField1
            config.TransactionField2 = in_TransactionField2
            config.TransactionID = in_TransactionID
            config.BusinessException = in_BusinessException
            config.SystemException = in_SystemException
            config.TransactionItem = in_TransactionItem
            config.RetryNumber = io_RetryNumber
            config.TransactionNumber = io_TransactionNumber
            config.ConsecutiveSystemExceptions = io_ConsecutiveSystemExceptions