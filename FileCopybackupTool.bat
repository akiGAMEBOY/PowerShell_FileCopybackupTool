@ECHO OFF
REM #################################################################################
REM # �������@�bFileCopybackupTool�i�N���p�o�b�`�j
REM # �@�\�@�@�bPowerShell�N���p�̃o�b�`
REM #--------------------------------------------------------------------------------
REM # �@�@�@�@�b-
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
ECHO �������I�����܂����B
ECHO �����ꂩ�̃L�[�������ƃE�B���h�E�����܂��B
PAUSE > NUL
EXIT %RETURNCODE%
