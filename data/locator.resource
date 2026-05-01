*** Variables ***

${BROWSER}              edge

${BASE_DOWNLOAD_PATH}   ${EXECDIR}${/}TEMP${/}RPA0026${/}Downloads
${DOWNLOAD_DIR}         ${BASE_DOWNLOAD_PATH}/Primavera
${TEMP_DIR}             ${EXECDIR}${/}TEMP${/}TempFolder_Primavera

@{PROCESS_QUEUE}

${login_error}      ${FALSE}
${log_path}         ${EMPTY}

${primavera}            ${EXECDIR}${/}data${/}images${/}primavera.png
${login_primavera}      ${EXECDIR}${/}data${/}images${/}login.png
${password_primavera}    ${EXECDIR}${/}data${/}images${/}password.png
${invalid_credentials}    ${EXECDIR}${/}data${/}images${/}invalid_credentials.png
${launched}    ${EXECDIR}${/}data${/}images${/}launched.png
${enterprise}    ${EXECDIR}${/}data${/}images${/}enterprise.png
${openfile}             ${EXECDIR}${/}data${/}images${/}openfile.png
${projects}             ${EXECDIR}${/}data${/}images${/}projects.png
${view}                 ${EXECDIR}${/}data${/}images${/}view.png    
${layout}               ${EXECDIR}${/}data${/}images${/}layout.png
${openlayout}           ${EXECDIR}${/}data${/}images${/}openlayout.png
${no}                   ${EXECDIR}${/}data${/}images${/}no.png
${suchen}               ${EXECDIR}${/}data${/}images${/}suchen.png
${weitersuchen}         ${EXECDIR}${/}data${/}images${/}weitersuchen.png
${nomatches}            ${EXECDIR}${/}data${/}images${/}nomatches.png
${abbrechen}            ${EXECDIR}${/}data${/}images${/}abbrechen.png
${open}                 ${EXECDIR}${/}data${/}images${/}open.png
${panel_layout}         ${EXECDIR}${/}data${/}images${/}panellayout.png
${open_project}         ${EXECDIR}${/}data${/}images${/}openproject.png
${activities}           ${EXECDIR}${/}data${/}images${/}activities.png
${reports}              ${EXECDIR}${/}data${/}images${/}reports.png
${reports_tab}          ${EXECDIR}${/}data${/}images${/}reportstab.png
${reports_suchen}       ${EXECDIR}${/}data${/}images${/}reports_suchen.png
${checkbox_1}           ${EXECDIR}${/}data${/}images${/}checkbox_1.png
${checkbox_2}           ${EXECDIR}${/}data${/}images${/}checkbox_2.png
${related_report}       ${EXECDIR}${/}data${/}images${/}related_report.png
${run}                  ${EXECDIR}${/}data${/}images${/}run.png
${report}               ${EXECDIR}${/}data${/}images${/}report.png
${report_blue}          ${EXECDIR}${/}data${/}images${/}report_blue.png
${run_report}           ${EXECDIR}${/}data${/}images${/}run_report.png
${output_file}          ${EXECDIR}${/}data${/}images${/}output_file.png
${ok}                   ${EXECDIR}${/}data${/}images${/}ok.png
${excel_open}           ${EXECDIR}${/}data${/}images${/}excel_open.png
${yes}                  ${EXECDIR}${/}data${/}images${/}yes.png
${output_file_2}        ${EXECDIR}${/}data${/}images${/}output_file_2.png
${error}    ${EXECDIR}${/}data${/}images${/}error.png
${error_ok}    ${EXECDIR}${/}data${/}images${/}error_ok.png
${delimited_text}    ${EXECDIR}${/}data${/}images${/}delimited_text.png
${delimited_text_clicked}    ${EXECDIR}${/}data${/}images${/}delimited_text_clicked.png
${field_delimiter_field}    ${EXECDIR}${/}data${/}images${/}field_delimiter_field.png
${text_qualifier_field}    ${EXECDIR}${/}data${/}images${/}text_qualifier_field.png
