import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

import pyodbc
from datetime import datetime

from robot.api import logger
from Resources.Config import Config
from Library.InitAllSettingsSQL import InitAllSettingsSQL

class SaveRobotUsageTime:
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
    
    def send_success_email(config, process_id, process_desc, dtStart, dtEnd, transactions, completed, technically_checked):
        subject = f"{process_id} ({process_desc}) has been finished"
        body = (
            f"It started at <b>{dtStart}</b> and finished at <b>{dtEnd}</b>.<br/><br/>"
            f"Transactions: {transactions} <br/><br/>"
            f"Completed: {completed} <br/><br/>"
            f"Technically Checked: {technically_checked}"
        )
        """
        send_email(
            config,
            subject=subject,
            body=body,
            to=config.get("E-mail_Team", ""),
            cc=config.get("Admin_Email", ""),
            sender=config.get("Email_Sender", "rpa@voith.com"),
            sender_name="RPA - Robot Usage"
        )
        """        

    def send_error_email(config, process_id, error_message):
        subject = f"{process_id} - Error to save Robot Usage Time"
        body = f"Error to save Robot usage time: {error_message}"
        """
        send_email(
            config,
            subject=subject,
            body=body,
            to=config.get("E-mail_Team", ""),
            cc=config.get("Admin_Email", ""),
            sender=config.get("Email_Sender", "rpa@voith.com"),
            sender_name="RPA - Robot Usage"
        )
        """        

    def send_email(config, subject, body, to, cc, sender, sender_name):
        # This is a stub. Replace with your actual email sending logic.
        print(f"Sending email:\nSubject: {subject}\nTo: {to}\nCC: {cc}\nFrom: {sender} ({sender_name})\nBody:\n{body}\n")

    # Logic to save usage time to a database
    def save_usage_time(self, in_IDProcess, in_Environment, io_IndexNo, in_TotalTransaction, in_TotalDone, in_TechnicallyChecked, in_EmailOnlyForDifferences):
        # Get the global parameters according with the default process
        settings_loader = InitAllSettingsSQL()
        self.io_Config = settings_loader.get_all_settings("RPA")

        try:
            conn = self.get_db_connection()
            if conn:
                cursor = conn.cursor()
                if not io_IndexNo or io_IndexNo.strip() == "":
                    # Insert new usage time log
                    cursor.execute(
                        "INSERT INTO tb_RobotUsageTime (ID_Process, Environment, StartDateTime) VALUES (?, ?, ?)", 
                        in_IDProcess, in_Environment, datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                    )
                    conn.commit()
                    # Get the new index
                    cursor.execute(
                        "SELECT TOP 1 [ID_UseTimeLog] FROM [UIPath_Param].[dbo].[tb_RobotUsageTime] WHERE (ID_PROCESS = ?) AND (EndDateTime IS NULL) ORDER BY ID_UseTimeLog desc", 
                        in_IDProcess
                    )
                    row = cursor.fetchone()
                    if row:
                        io_IndexNo = str(row[0])
                else:
                    # Update usage time log
                    # Convert string values to int if needed
                    total_trans = int(in_TotalTransaction) if in_TotalTransaction else 0
                    total_done = int(in_TotalDone) if in_TotalDone else 0
                    tech_checked = int(in_TechnicallyChecked) if in_TechnicallyChecked else 0
                    
                    cursor.execute(
                        "UPDATE [dbo].[tb_RobotUsageTime] SET EndDateTime = GetDate(), Transactions = ?, Completed = ?, TechnicallyChecked = ? WHERE ID_UseTimeLog = ?", 
                        total_trans, total_done, tech_checked, io_IndexNo
                    )
                    conn.commit()

                    # Get the updated row
                    cursor.execute(
                        """SELECT R.[ID_UseTimeLog],R.[ID_Process],R.[Environment],R.[StartDateTime],
                                ISNULL(R.[EndDateTime], '') EndDateTime,ISNULL(R.[Transactions],0)[Transactions],
                                ISNULL(R.[Completed],0)[Completed], ISNULL(R.[TechnicallyChecked],0)[TechnicallyChecked],
                                P.Process_Description
                        FROM [UIPath_Param].[dbo].[tb_RobotUsageTime] R
                        JOIN [UIPath_Param].[dbo].[tb_Process] P ON R.[ID_Process] = P.[ID_Process]
                        WHERE (R.ID_UseTimeLog = ?)""",
                        io_IndexNo
                    )
                    TableRow = cursor.fetchone()

                    if TableRow:
                        dtStart = TableRow[3].strftime("%Y-%m-%d %H:%M:%S.%f")[:-3] if TableRow[3] else ""
                        dtEnd = TableRow[4].strftime("%Y-%m-%d %H:%M:%S.%f")[:-3] if TableRow[4] else ""
                        
                        # Calculate duration in seconds if both times are available
                        if TableRow[3] and TableRow[4]:
                            duration_seconds = (TableRow[4] - TableRow[3]).total_seconds()
                        
                        # Decide if email should be sent
                        if in_EmailOnlyForDifferences:
                            bSendEmail = (int(in_TotalTransaction) - int(in_TechnicallyChecked)) != 0
                        else:
                            bSendEmail = True

                        if in_Environment == "0":
                            bSendEmail = False

                        if bSendEmail and TableRow:
                            dtStart = TableRow[3].strftime("%d/%m/%y %H:%M:%S") if TableRow[3] else ""
                            dtEnd = TableRow[4].strftime("%d/%m/%y %H:%M:%S") if TableRow[4] else ""
                            """
                            send_success_email(
                                config,
                                in_IDProcess,
                                TableRow[8],  # Process_Description
                                dtStart,
                                dtEnd,
                                TableRow[5],  # Transactions
                                TableRow[6],  # Completed
                                TableRow[7],  # TechnicallyChecked
                            )
                            """                                                 
        except Exception as e:
            logger.console(f"{self.__class__.__name__} -> Error saving robot usage time: {e}")
            logger.error(f"{self.__class__.__name__} -> Error saving robot usage time: {e}")
        finally:
            if conn:
                conn.close()        

        return io_IndexNo