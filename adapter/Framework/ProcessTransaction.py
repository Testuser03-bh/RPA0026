import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from robot.api import logger
from Resources.Config import Config
from Framework.Process import Process
from Framework.SetTransactionStatus import SetTransactionStatus

class BusinessRuleException(Exception):
    pass # This class has no extra code, just inherits from Exception

# Process a single transaction
class ProcessTransaction:
    def process_data(self, config: Config):
        set_status = SetTransactionStatus()
        try:
            logger.console("Processing transaction...")
            logger.info("Processing transaction...")
            
            config.BusinessException = None
            config.SystemException = None

            process = Process()
            process.process_data(config)

            # Simulate processing logic
            if config.TransactionItem and config.TransactionStatus == "BusinessException":
                raise BusinessRuleException("Business rule exception occurred.")
            elif config.TransactionItem and config.TransactionStatus == "SystemException":
                raise Exception("System exception occurred.")

            config.Log_Looping += 1
            config.Log_Done += 1                                

            set_status.set_transaction_status(config, 
                                              success=True, 
                                              business_exception=False,
                                              system_exception=False,
                                              error_message="",
                                              in_BusinessException=config.BusinessException,
                                              in_TransactionField1=config.TransactionField1,
                                              in_TransactionField2=config.TransactionField2,
                                              in_TransactionID=config.TransactionID,
                                              in_SystemException=config.SystemException,
                                              in_TransactionItem=config.TransactionItem,
                                              io_RetryNumber=config.RetryNumber,
                                              io_TransactionNumber=config.TransactionNumber,
                                              io_ConsecutiveSystemExceptions=config.ConsecutiveSystemExceptions                                                
                                              )
        except BusinessRuleException as bre:
            config.BusinessException = bre

            config.Log_Looping += 1

            logger.console(f" {self.__class__.__name__} -> Business exception: {bre}")
            logger.error(f"{self.__class__.__name__} -> Business exception: {bre}")

            set_status.set_transaction_status(config, 
                                    success=False, 
                                    business_exception=True,
                                    system_exception=False,
                                    error_message=bre,
                                    in_BusinessException=config.BusinessException,
                                    in_TransactionField1=config.TransactionField1,
                                    in_TransactionField2=config.TransactionField2,
                                    in_TransactionID=config.TransactionID,
                                    in_SystemException=config.SystemException,
                                    in_TransactionItem=config.TransactionItem,
                                    io_RetryNumber=config.RetryNumber,
                                    io_TransactionNumber=config.TransactionNumber,
                                    io_ConsecutiveSystemExceptions=config.ConsecutiveSystemExceptions                                                
                                    )
        except Exception as se:
            config.SystemException = se
            config.ConsecutiveSystemExceptions += 1

            logger.console(f" {self.__class__.__name__} -> System exception: {se}")
            logger.error(f"{self.__class__.__name__} -> System exception: {se}")

            set_status.set_transaction_status(config, 
                        success=False, 
                        business_exception=False,
                        system_exception=True,
                        error_message=se,
                        in_BusinessException=config.BusinessException,
                        in_TransactionField1=config.TransactionField1,
                        in_TransactionField2=config.TransactionField2,
                        in_TransactionID=config.TransactionID,
                        in_SystemException=config.SystemException,
                        in_TransactionItem=config.TransactionItem,
                        io_RetryNumber=config.RetryNumber,
                        io_TransactionNumber=config.TransactionNumber,
                        io_ConsecutiveSystemExceptions=config.ConsecutiveSystemExceptions                                                
                        )