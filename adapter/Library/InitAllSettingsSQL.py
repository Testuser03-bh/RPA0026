import sys
import os
import pyodbc

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from robot.api import logger
from Resources.Config import Config

class InitAllSettingsSQL:
    # Connect to the database
    def get_db_connection(self):        
        try:
            config = Config()
            connection = pyodbc.connect(config.ConnectionString)
            return connection
        except pyodbc.Error as e:
            logger.console("Error connecting to database: ", e)
            logger.error("Error connecting to database: ", e)
            return None    

    # Retrieve all settings from the database 
    def get_all_settings(self, in_ProcessName):        
        out_Config = {}

        conn = self.get_db_connection()
        if conn:
            try:
                cursor = conn.cursor()

                # If process name is not 'RPA', check and insert process if needed
                if in_ProcessName != "RPA":
                    cursor.execute("SELECT * FROM tb_Process WHERE ID_Process = ?", in_ProcessName)
                    rows = cursor.fetchall()
                    if len(rows) == 0:
                        cursor.execute("INSERT INTO [dbo].[tb_Process] ([ID_Process],[Creation_Date]) VALUES (?, GETDATE())", in_ProcessName)
                        conn.commit()

                # Get settings from tb_ProcessConfig
                cursor.execute("SELECT * FROM tb_ProcessConfig WHERE ID_Process = ? OR ID_Process = 'RPA'", in_ProcessName)
                rows = cursor.fetchall()

                for row in rows:
                    # Assuming Name is at index 2 and Value at index 3 (adjust if needed)
                    name = row[2].strip() if row[2] else ""
                    value = row[3].strip() if row[3] else ""
                    if name:
                        out_Config[name] = value
            except Exception as e:
                logger.console(f"{self.__class__.__name__} -> Error retrieving settings: {e}")
                logger.error(f"{self.__class__.__name__} -> Error retrieving settings: {e}")
            finally:
                if conn:
                    conn.close()

        return out_Config