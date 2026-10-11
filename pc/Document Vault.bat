@echo off
REM Document Vault - opens the app in its own window (like a desktop app).
REM Your documents stay in OneDrive\Document Vault; this only opens the app.
set "URL=https://jacovanstraten-glitch.github.io/doc-vault/"

set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe"

if exist "%CHROME%" (
  start "" "%CHROME%" --app="%URL%" --window-size=1200,900
) else (
  start "" msedge --app="%URL%" --window-size=1200,900
)
