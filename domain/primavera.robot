*** Settings ***
Library    RPA.Desktop
Library    SeleniumLibrary
Library    OperatingSystem
Library    String
Library    Collections
Library    DateTime
Library    RPA.Email.ImapSmtp
Library    Process
Library    dotenv
Library    ..${/}adapter${/}Library${/}InitAllSettingsSQL.py
Library    ..${/}adapter${/}Resources${/}Config.py
Library    ..${/}adapter${/}Helper.py
Resource    ..${/}data${/}locator.resource
*** Variables ***
${PRIMARY_PROCESS_NAME}    RPA0026
${SECONDARY_PROCESS_NAME}    RPA
${secondary_config}
${primary_config}
*** Keywords ***
Initialize Environment and Data
    ${primary_fetched_config}=    InitAllSettingsSQL.Get All Settings    ${PRIMARY_PROCESS_NAME}
    ${secondary_fetched_config}=    InitAllSettingsSQL.Get All Settings    ${SECONDARY_PROCESS_NAME}
    Set Global Variable    ${primary_config}    ${primary_fetched_config}
    Set Global Variable    ${secondary_config}    ${secondary_fetched_config}
    ${temp}    ${downloads}=    Prepare Environment
    ...    ${primary_config['TempFolder_Primavera']}
    ...    ${primary_config['PathToSave_ReportFiles']}
    Set Global Variable         \${TEMP_DIR}           ${temp}
    Set Global Variable         \${DOWNLOAD_DIR}       ${downloads}
    ${DATA_TABLE}=    Get Aggregated Primavera Data
    Set Global Variable    \${PROCESS_QUEUE}    ${DATA_TABLE}
    ${count}=        Get Length    ${PROCESS_QUEUE}
    IF    ${count} == 0
        Log To Console    No records to process. Terminating.
        Fatal Error       No data found in database.
    END
Launch Citrix
    Open Browser    ${secondary_config['URL_Citrix_Prod']}     ${BROWSER}
    Maximize Browser Window
    Wait Until Element Is Visible    id:protocolhandler-welcome-installButton    30s
    Click Element                    id:protocolhandler-welcome-installButton
    Close Citrix Popup
    Wait Until Element Is Visible    id:legalstatement-checkbox2                 30s
    Click Element                    id:legalstatement-checkbox2
    Click Element                    id:protocolhandler-detect-alreadyInstalledLink
    Wait Until Element Is Visible    id:username                                 30s
    Input Text                       id:username    ${secondary_config['RPAUser']}
    Evaluate                __import__('dotenv').load_dotenv(".env")
    ${citrix_password}=     Get Environment Variable    Citrix_Password
    Input Password          id:password    ${citrix_password}
    Click Element           id:loginBtn
    Log To Console          Logged into Citrix Workspace
    Wait Until Element Is Visible    xpath=//p[normalize-space()="Oracle Primavera P6 R23 UiPath"]    60s
    Click Element                    xpath=//p[normalize-space()="Oracle Primavera P6 R23 UiPath"]
    Wait Until Element Is Visible    xpath=//div[normalize-space()="Open"]                            60s
    Click Element                    xpath=//div[normalize-space()="Open"]
Open Latest ICA File
    Open Latest Ica
Launching Primavera Application
    Wait For Element        image:${primavera}           timeout=600
    Wait For Element        image:${login_primavera}     timeout=10
    Click                   image:${login_primavera}
    RPA.Desktop.Type Text   ${primary_config['PrimaveraP6_User_Test'].split("_")[0].strip()}
    Wait For Element        image:${password_primavera}  timeout=10
    Click                   image:${password_primavera}
    Evaluate                __import__('dotenv').load_dotenv(".env")
    ${primavera_password}=  Get Environment Variable     Primavera_Password
    RPA.Desktop.Type Text   ${primavera_password}
    RPA.Desktop.Press Keys  ENTER
    ${invalid_creds}=       Run Keyword And Return Status    Wait For Element    image:${invalid_credentials}    timeout=5
    IF    ${invalid_creds}
        ${ok_exists}=       Run Keyword And Return Status    Wait For Element    image:${ok}    timeout=2
        IF    ${ok_exists}
            Click           image:${ok}
        END
        Log To Console      Invalid credentials dialog detected. Terminating process.
        ${login_error}=     Set Variable    True
        Set Global Variable    ${login_error}
        Fatal Error         Process terminated due to Invalid Credentials.
    END
    ${login_success}=       Run Keyword And Return Status    Wait For Element    image:${launched}    timeout=200
    ${login_error}=         Evaluate    not ${login_success}
    Set Global Variable     ${login_error}
    IF    ${login_success}
        Log To Console      Primavera application is launched
    ELSE
        Log To Console      Primavera login failed - Error flag is set for SMTP
    END
Process Records
    FOR    ${row}    IN    @{PROCESS_QUEUE}
        Log To Console    Processing record: Layout=${row['Layout']}, Report=${row['Report_Name']}
        IF    not $row['Layout']
            Log To Console       Skipping row due to missing Layout
            Set To Dictionary    ${row}    Log_Message    Field Layout is empty
            Set To Dictionary    ${row}    Reports    0
            CONTINUE
        END
        IF    not $row['Report_Name']
            Log To Console       Skipping row due to missing Report_Name
            Set To Dictionary    ${row}    Log_Message    Field Report_Name is empty
            Set To Dictionary    ${row}    Reports    0
            CONTINUE
        END
        IF    not $row['Storage_Location']
            Log To Console       Skipping row due to missing Storage_Location
            Set To Dictionary    ${row}    Log_Message    Field Storage_Location is empty
            Set To Dictionary    ${row}    Reports    0
            CONTINUE
        END
        Wait For Element    image:${enterprise}    timeout=10
        Click               image:${enterprise}
        Wait For Element    image:${projects}      timeout=10
        Click               image:${projects}
        Wait For Element    image:${view}          timeout=10
        Click               image:${view}
        Wait For Element    image:${layout}        timeout=100
        Click               image:${layout}
        Wait For Element    image:${openlayout}    timeout=10
        Click               image:${openlayout}
        Log To Console      message=Open Project button is clicked
        ${no_button_exists}=      Run Keyword And Return Status    Wait For Element    image:${no}    timeout=10
        IF    ${no_button_exists}
            Click                 image:${no}
            Log To Console        'No' button appeared and was clicked.
        ELSE
            Log To Console        'No' button did not appear. Continuing flow...
        END
        RPA.Desktop.Press Keys    CTRL    f
        FOR    ${i}    IN RANGE    15
            ${image_found}=       Run Keyword And Return Status    Wait For Element    image:${suchen}    timeout=1s
            IF    ${image_found}    BREAK
            RPA.Desktop.Press Keys    backspace
        END
        Click               image:${suchen}
        RPA.Desktop.Type Text     ${row}[Layout]
        Wait For Element    image:${weitersuchen}    timeout=10
        Click               image:${weitersuchen}
        ${layout_missing}=        Run Keyword And Return Status    Wait For Element    image:${nomatches}    timeout=5
        IF    ${layout_missing}
            RPA.Desktop.Press Keys    ENTER
            Wait For Element          image:${abbrechen}     timeout=10
            Click                     image:${abbrechen}
            Set To Dictionary         ${row}    Log_Message    Layout + ${row}[Layout] + not found
            Log To Console            Layout + ${row}[Layout] + not found. Moving to next row.
            Continue For Loop
        END
        Wait For Element    image:${abbrechen}     timeout=10
        Click               image:${abbrechen}
        Wait For Element    image:${open}          timeout=10
        Click               image:${open}
        Wait For Element    image:${panel_layout}  timeout=80
        Click               image:${panel_layout}
        RPA.Desktop.Press Keys    ctrl    a
        Wait For Element    image:${panel_layout}  timeout=10
        Click               image:${panel_layout}  right_click
        Wait For Element    image:${open_project}  timeout=10
        Click               image:${open_project}
        ${activities_found}=      Run Keyword And Return Status    Wait For Element    image:${activities}    timeout=100
        IF    not ${activities_found}
            # Step 12.10.1 – Activities tab not found; log and skip to next row
            Set To Dictionary    ${row}    Log_Message    Error opening project selection
            Set To Dictionary    ${row}    Reports    0
            Log To Console       Error opening project selection for Layout: ${row}[Layout]
            Continue For Loop
        END
        Click               image:${activities}
        Wait For Element    image:${reports}       timeout=20
        Click               image:${reports}
        Wait For Element    image:${reports_tab}   timeout=50
        RPA.Desktop.Press Keys    ctrl    f
        FOR    ${counter}    IN RANGE    50
            RPA.Desktop.Press Keys    BACKSPACE
            ${cleared}=    Run Keyword And Return Status    Wait For Element    image:${reports_suchen}    timeout=0.5
            IF    ${cleared}    BREAK
        END
        Wait For Element    image:${reports_suchen}    timeout=10
        Click               image:${reports_suchen}
        RPA.Desktop.Type Text     ${row}[Report_Name]
        Wait For Element    image:${checkbox_1}    timeout=10
        Click               image:${checkbox_1}
        Wait For Element    image:${checkbox_2}    timeout=10
        Click               image:${checkbox_2}
        Wait For Element    image:${weitersuchen}  timeout=10
        Click               image:${weitersuchen}
        ${found_exact}=     Set Variable    False
        FOR    ${match_idx}    IN RANGE    20
            ${report_missing}=    Run Keyword And Return Status    Wait For Element    image:${nomatches}    timeout=2
            IF    ${report_missing}
                RPA.Desktop.Press Keys    ENTER
                Wait For Element          image:${abbrechen}     timeout=10
                Click                     image:${abbrechen}
                Set To Dictionary         ${row}    Log_Message    Report + ${row}[Report_Name] + not found
                Log To Console            Report + ${row}[Report_Name] + not found
                BREAK
            END
            RPA.Desktop.Press Keys    ctrl    c
            ${copied_text}=    Evaluate    pyperclip.paste()    modules=pyperclip
            ${match}=    Evaluate    $copied_text.strip() == $row['Report_Name'].strip()
            IF    ${match}
                ${found_exact}=    Set Variable    True
                BREAK
            ELSE
                Click    image:${weitersuchen}
            END
        END
        IF    not ${found_exact}
            Continue For Loop
        END
        Wait For Element    image:${abbrechen}     timeout=10
        Click               image:${abbrechen}
        Wait For Element    image:${related_report}    timeout=10
        Click               image:${related_report}    right_click
        Wait For Element    image:${run}           timeout=10
        Click               image:${run}
        Wait For Element    image:${report}        timeout=10
        Click               image:${report}
        ${error_appeared}=    Run Keyword And Return Status    Wait For Element    image:${error}    timeout=5
        IF    ${error_appeared}
            Wait For Element    image:${error_ok}    timeout=10
            Click               image:${error_ok}
            Set To Dictionary   ${row}    Log_Message    Report + ${row}[Report_Name] + not found
            Log To Console      Error popup closed, skipped to next row.
            Continue For Loop
        END
        Wait For Element        image:${delimited_text}           timeout=10
        FOR    ${attempt}    IN RANGE    10
            Click                   image:${delimited_text}
            ${clicked_state}=       Run Keyword And Return Status    Wait For Element    image:${delimited_text_clicked}    timeout=1
            IF    ${clicked_state}    BREAK
        END
        Wait For Element        image:${field_delimiter_field}    timeout=10
        Click                   image:${field_delimiter_field}
        RPA.Desktop.Type Text   ${row}[Field_Delimiter]
        RPA.Desktop.Press Keys    ENTER
        Wait For Element        image:${text_qualifier_field}     timeout=10
        Click                   image:${text_qualifier_field}
        RPA.Desktop.Type Text   ${row}[Text_Qualifier]
        RPA.Desktop.Press Keys    ENTER
        ${output_1_exists}=    Run Keyword And Return Status    Wait For Element    image:${output_file}    timeout=5
        IF    ${output_1_exists}
            ${target_field}=    Set Variable    image:${output_file}
        ELSE
            Wait For Element    image:${output_file_2}    timeout=10
            ${target_field}=    Set Variable    image:${output_file_2}
        END
        Click                   ${target_field}
        RPA.Desktop.Press Keys  ctrl    a
        RPA.Desktop.Press Keys  delete
        ${temp_file_path}=      Set Variable    ${TEMP_DIR}\\Primavera.csv
        RPA.Desktop.Type Text   text=${temp_file_path}
        Wait For Element        image:${ok}    timeout=10
        Click                   image:${ok}
        ${file_exists}=         Run Keyword And Return Status    Wait For Element    image:${yes}    timeout=10
        IF    ${file_exists}
            Click               image:${yes}
            Log To Console      'Yes' button appeared and was clicked to overwrite file
        ELSE
            Log To Console      'File already exists' button did not appear. Continuing flow...
        END
        ${excel_appeared}=      Run Keyword And Return Status    Wait For Element    image:${excel_open}    timeout=20
        IF    ${excel_appeared}
            RPA.Desktop.Press Keys    alt    f4
        ELSE
            Log To Console      message=Excel report did not open within timeout, proceeding to file check...
        END
        ${file_downloaded}=     Run Keyword And Return Status    File Should Exist    ${temp_file_path}
        IF    not ${file_downloaded}
            Set To Dictionary    ${row}    Log_Message    Error to download the report ${row}[Report_Name]
            Log To Console       Status for ${row}[Report_Name]: ${row}[Log_Message]
            Continue For Loop
        ELSE
            @{storage_locations}=    Split String    ${row}[Storage_Location]      ;
            @{file_names}=           Split String    ${row}[FileName_Convention]   ;
            ${current_date}=         Get Current Date    result_format=%Y_%m_%d
            ${array_length}=         Get Length      ${storage_locations}
            FOR    ${index}    IN RANGE    ${array_length}
                ${target_folder}=    Set Variable    ${storage_locations}[${index}]
                ${target_file}=      Set Variable    ${file_names}[${index}]
                ${target_file}=      Replace String  ${target_file}    YYYY_MM_DD    ${current_date}
                Create Directory     ${target_folder}
                Copy File            ${temp_file_path}    ${target_folder}\\${target_file}
            END
            Remove File          ${temp_file_path}
            Set To Dictionary    ${row}    Log_Message    Report saved with success
            Log To Console       Status for ${row}[Report_Name]: ${row}[Log_Message]
        END
    END
    ${log_path}=    Generate Final Log File    ${PROCESS_QUEUE}    ${primary_config['PathToSave_ReportFiles']}
    Log To Console  Processing completed. Master log saved to: ${log_path}
Send Email Final Report
    Authorize SMTP
    ...    account=${primary_config['E-mail_Sender']}
    ...    password=${EMPTY}
    ...    smtp_server=${secondary_config['SMTP_Server']}
    ...    smtp_port=${secondary_config['SMTP_Port']}
    Log To Console    Preparing to send final report...
    ${current_date_email}=    Get Current Date    result_format=%Y-%m-%d
    ${email_subject}=         Set Variable        Primavera Report Data Extraction - ${current_date_email}
    IF    ${login_error}
        ${body_msg}=    Set Variable    Please find attached the Data Export Primavera report.<br><br>Attention - Error to open Primavera with message: "Login Error"<br><br>RPA Voith
    ELSE
        ${body_msg}=    Set Variable    Please find attached the Data Export Primavera report.<br><br>This is an automatic e-mail. Please do not reply to this message as it is not verified.<br><br>RPA Voith
    END
    Send Message
    ...    sender=${primary_config['E-mail_Sender']}
    ...    recipients=${primary_config['E-mail_Address']}
    ...    subject=${email_subject}
    ...    body=${body_msg}
    ...    html=True
    ...    attachments=${log_path}
    Log To Console    Report successfully sent!
