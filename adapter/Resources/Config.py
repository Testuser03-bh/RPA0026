import pyodbc

from robot.api import logger

class Config:

    def __init__(self):
        # Variables
        self.TransactionItem = None
        self.SystemException = None
        self.BusinessException = None
        self.TransactionStatus = ""
        self.TransactionNumber = -1
        self.RetryNumber = 0
        self.TransactionField1 = ""
        self.TransactionField2 = ""
        self.TransactionID = ""
        self.dt_TransactionData = []
        self.ConsecutiveSystemExceptions = 0
        self.Log_Looping = 0
        self.Log_Done = 0
        self.Log_Transactions = 0
        """
        self.Config = {
            "LogMessage_GetTransactionData": "Getting transaction data...",
            "MaxConsecutiveSystemExceptions": 3,
            "ExceptionMessage_ConsecutiveErrors": "Max consecutive system exceptions exceeded.",
            "TempFolder": "./temp/",
            "E-mail_Team": "natalia.abrao-extern@voith.com",
            "Email_Sender": "rpa@voith.com",
            "EmailOnlyWithDifferences": "True",
            "UserID_Test": "SAP_USER_TEST",
            "SID_Test": "SID_TEST",
            "Path_Folder_Files": "J:/Systems/UiPath/PROD/FilesTest/",
            "ShouldMarkJobAsFaulted": False,
            "logF_BusinessProcessName": "Shipment Documentation Upload"
        }
        """
        self.Config = {}
        self.ShouldMarkJobAsFaulted = False
        self.Log_Environment = "0" # 0=Development, 1=Production
        self.Log_Index = 0
        self.ProcessIDCode = "" 
        self.ConnectionString = "Driver={SQL Server};Server=ITLSQLOTHERSCONS\ITLSQL65;Database=UIPath_Param;Integrated Security=True" #UID=your_username;PWD=your_password;"

    # Connect to the database
    def get_db_connection(self):        
        try:
            connection = pyodbc.connect(self.ConnectionString)
            return connection
        except pyodbc.Error as e:
            print("Error connecting to database: ", e)
            return None

    # Retrieve all settings from the database        
    def get_all_settings(self):
        conn = self.get_db_connection()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT Name, Value FROM tb_ProcessConfig WHERE ID_Process = 'RPA' OR ID_Process = ?", self.ProcessIDCode)
                row = cursor.fetchall()                
                if row:
                    return row
            except Exception as e:
                logger.console(f"Starting {e}")
                logger.error(f"Starting {e}")
            finally:
                if conn:
                    conn.close()
        return None