import os
import pyodbc
import pandas as pd
import shutil
import time
import pyautogui
import pygetwindow as gw
from robot.api import logger
from datetime import datetime
from robot.api.deco import keyword

DB_CONFIG = {
    "server": "ITLSQLOTHERSCONS\ITLSQL65",
    "database": "UIPath_Param",
    "trusted": "yes"
}

def open_latest_ica(timeout=60, poll_interval=2, max_age_seconds=180):
    """
    Waits for a NEW ICA file in Downloads.
    Discards files older than max_age_seconds (default 3 minutes).
    Opens the newest valid ICA file immediately.
    """

    logger.console("Waiting for new ICA file in Downloads...")

    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    start_time = time.time()

    while time.time() - start_time < timeout:

        now = time.time()

        ica_files = [
            os.path.join(downloads, f)
            for f in os.listdir(downloads)
            if f.lower().endswith(".ica")
        ]

        valid_files = []

        for file in ica_files:
            file_age = now - os.path.getctime(file)

            # Keep only files newer than 3 minutes
            if file_age <= max_age_seconds:
                valid_files.append(file)

        if valid_files:
            latest_file = max(valid_files, key=os.path.getctime)

            logger.console(
                f"Opening ICA file: {os.path.basename(latest_file)} "
                f"(age: {int(now - os.path.getctime(latest_file))} sec)"
            )

            os.startfile(latest_file)
            return True

        time.sleep(poll_interval)

    logger.error("Timeout: No new ICA file detected within allowed age.")
    return False

def get_db_connection():
    drivers = ["{ODBC Driver 17 for SQL Server}", "{SQL Server Native Client 11.0}", "{SQL Server}"]
    for driver in drivers:
        try:
            conn_str = f"DRIVER={driver};SERVER={DB_CONFIG['server']};DATABASE={DB_CONFIG['database']};Trusted_Connection={DB_CONFIG['trusted']};"
            conn = pyodbc.connect(conn_str, timeout=5)
            return conn
        except pyodbc.Error:
            continue
    raise ConnectionError(f"Could not connect to {DB_CONFIG['server']}.")

@keyword("Process Downloaded Report")
def process_downloaded_report(row, temp_file_path):
    """
    Validates the downloaded file, distributes it to all target directories, 
    deletes the temp file, and updates the row's Log_Message.
    """
    if not os.path.exists(temp_file_path):
        row['Log_Message'] = f"Error to download the report {row.get('Report_Name', '')}"
        return row
    
    try:
        storage_locations = str(row['Storage_Location']).split(';')
        file_conventions = str(row['FileName_Convention']).split(';')
        today_str = datetime.now().strftime("%Y_%m_%d")
        
        for i in range(len(storage_locations)):
            target_dir = storage_locations[i].strip()
            target_filename = file_conventions[i].strip().replace("YYYY_MM_DD", today_str)
            target_path = os.path.join(target_dir, target_filename)
            
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
                
            shutil.copy(temp_file_path, target_path)
        
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
            
        row['Log_Message'] = "Report saved with success"
        
    except Exception as e:
        row['Log_Message'] = f"Error processing file distribution: {str(e)}"
        
    return row

def close_citrix_popup():
    logger.console("Closing Citrix protocol popup using ESC")
    time.sleep(3)
    pyautogui.press("esc")

@keyword("Generate Final Log File")
def generate_final_log_file(process_queue, downloads_folder):
    """
    Dumps the updated internal memory table into a CSV file.
    """
    today_dash = datetime.now().strftime("%Y-%m-%d")
    log_filename = f"report_primavera_{today_dash}.csv"
    log_filepath = os.path.join(downloads_folder, log_filename)
    
    df = pd.DataFrame(process_queue)
    df.to_csv(log_filepath, index=False, sep=';', encoding='utf-8-sig')
    
    print(f"Final log successfully generated at: {log_filepath}")
    return log_filepath

@keyword("Prepare Environment")
def prepare_environment(temp_folder, path_to_save_report_files):
    """Creates the necessary local directories for processing."""
    download_folder = os.path.join(path_to_save_report_files, "Primavera")
    
    for folder in [temp_folder, download_folder]:
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"Created folder: {folder}")
    
    return temp_folder, download_folder

@keyword("Get Aggregated Primavera Data")
def get_aggregated_primavera_data():
        try:
            # We now use your centralized connection method instead of passed arguments
            conn = get_db_connection()
            
            fetch_query = "SELECT * FROM tb_RPA0026 WHERE (Status = 1) AND (SystemName = 'Primavera')"
            df_prod = pd.read_sql(fetch_query, conn)

            if df_prod.empty:
                conn.close()
                return []

            cursor = conn.cursor()

            create_temp_table = """
            CREATE TABLE #RPA0026T (
                [ID_Lines] [int] NOT NULL, [Variant_Name] [varchar](max) NULL,
                [Storage_Location] [varchar](max) NULL, [FileName_Convention] [varchar](max) NULL,
                [Archieve_Location] [varchar](max) NULL, [Creation_Date] [date] NULL,
                [Update_Date] [datetime] NULL, [UserId] [varchar](30) NULL,
                [Status] [bit] NULL, [Company] [varchar](50) NULL,
                [SAP_URLName] [varchar](300) NULL, [SystemName] [varchar](30) NULL,
                [Module] [varchar](300) NULL, [SAPLink] [varchar](max) NULL,
                [Layout] [varchar](200) NULL, [Field_Delimiter] [varchar](10) NULL,
                [Text_Qualifier] [varchar](100) NULL, [Report_Name] [varchar](200) NULL
            )
            """
            cursor.execute(create_temp_table)

            for index, row in df_prod.iterrows():
                clean_row = tuple(None if pd.isna(x) else x for x in row)

                cursor.execute("""
                    INSERT INTO #RPA0026T ([ID_Lines],[Variant_Name],[Storage_Location],[FileName_Convention],
                    [Archieve_Location],[Creation_Date],[Update_Date],[UserId],[Status],[Company],
                    [SAP_URLName],[SystemName],[Module],[SAPLink],[Layout],[Field_Delimiter],
                    [Text_Qualifier],[Report_Name]) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, clean_row)

            agg_query = """
            SELECT Layout, Report_Name,
                STRING_AGG(Storage_Location, ';') as Storage_Location,
                STRING_AGG(FileName_Convention, ';') as FileName_Convention,
                Field_Delimiter, Text_Qualifier
            FROM #RPA0026T
            WHERE (Status = 1) AND (SystemName = 'Primavera')  
            GROUP BY Layout, Report_Name, Field_Delimiter, Text_Qualifier 
            ORDER BY Layout DESC
            """
            df_final = pd.read_sql(agg_query, conn)
            
            conn.close()

            if not df_final.empty:
                df_final['Reports'] = 0
                df_final['Log_Message'] = "Status Message"
                return df_final.to_dict('records')
            
            return []

        except Exception as e:
            print(f"Database Error: {str(e)}")
            raise e




