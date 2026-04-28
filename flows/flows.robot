*** Settings ***
Documentation       Entry point for RPA0026 - Primavera Data Extraction.
Resource            ..${/}domain${/}primavera.robot

*** Keywords ***
Run Primavera Extraction Flow
    [Documentation]    Executes the full end-to-end process.
    Initialize Environment and Data
    Launch Citrix
    Open Latest ICA File
    Launching Primavera Application
    Process Records
    Send Email Final Report