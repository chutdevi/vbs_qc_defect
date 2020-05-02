@echo off
	SET PATHS=G:\vbs_qc_defect
	SET NEWLN=^& echo. 
	SET "TASK=%PATHS% RUN TASK RECEIVEIN OF %DATE% %TIME% %NEWLN% %NEWLN%

	cscript %PATHS%\DEFECT_DATA_INSERT.vbs >nul
	timeout /t 1 >nul	
	
	ECHO.
	
	cscript %PATHS%\DEFECT_PRODUCTION_DATA_INSERT.vbs >nul
	timeout /t 1 >nul
	
	ECHO.
	
	cscript %PATHS%\DEFECT_RECEIVE_DATA_INSERT.vbs >nul
	timeout /t 1 >nul 
 
 pause