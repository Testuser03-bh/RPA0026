import sys
import os

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from Framework.Initialization import Initialization
from Framework.GetTransactionData import GetTransactionData
from Framework.ProcessTransaction import ProcessTransaction
from Framework.EndProcess import EndProcess
from Resources.Config import Config

class BusinessRuleException(Exception):
    pass

class Main:
    def initialize(self, config: Config):
        Init = Initialization()        
        return Init.initialize(config)
        
    def get_transaction_data(self, config: Config):
        TransactionData = GetTransactionData()
        return TransactionData.get_transaction_data(config)        

    def process_transaction(self, config: Config):
        ProcTransaction = ProcessTransaction()
        return ProcTransaction.process_data(config)        
        
    def end_process(self, config: Config):
        EndProc = EndProcess()
        return EndProc.end_process(config)

    def run(self):       
        process = Main() 
        config = Config()         
        
        process.initialize(config)

        while config.TransactionItem is not None or config.TransactionNumber < config.Log_Transactions:
            process.get_transaction_data(config)
            if config.TransactionItem is not None:
                process.process_transaction(config)
            else:
                break

        process.end_process(config)

if __name__ == "__main__":
    main_process = Main()
    main_process.run()      
