@ECHO OFF
REM #################################################################################
REM # 処理名　｜FileCopybackupTool（起動用バッチ）
REM # 機能　　｜PowerShell起動用のバッチ
REM #--------------------------------------------------------------------------------
REM # 　　　　｜-
REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  FileCopybackupTool
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.
powershell -NoProfile -ExecutionPolicy Unrestricted -File .\source\powershell\Main.ps1 "Copy"
SET RETURNCODE=%ERRORLEVEL%

IF %ERRORLEVEL%==0 (
    powershell -NoProfile -ExecutionPolicy Unrestricted -file .\source\powershell\Main.ps1 "Rotation"
	SET RETURNCODE=%ERRORLEVEL%
)

ECHO.
ECHO 処理が終了しました。
ECHO いずれかのキーを押すとウィンドウが閉じます。
PAUSE > NUL
EXIT %RETURNCODE%
