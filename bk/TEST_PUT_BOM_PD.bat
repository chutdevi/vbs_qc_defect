@echo off
	SET PATHS=G:\vbs_develop
	REM SET NEWLN=^& echo. 
	SET "TASK=%PATHS% RUN TASK RECEIVEIN OF %DATE% %TIME% %NEWLN% %NEWLN%

	cscript %PATHS%\MASTER_BOM_PD_INSERT.vbs >nul
	timeout /t 2 >nul	

 
 
 pause