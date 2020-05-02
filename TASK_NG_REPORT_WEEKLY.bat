@echo off
	SET PATHS=G:\vbs_qc_defect

	cscript %PATHS%\NG_EXPL_INSERT.vbs >nul
	rem timeout /t 1 >nul	
	
	cscript %PATHS%\NG_MSQL_INSERT.vbs >nul
	rem timeout /t 1 >nul

	ECHO  EXPORT...
	cscript %PATHS%\EXPORT.vbs >nul
	timeout /t 1 >nul


	ECHO  SENDMAIL...
	cscript %PATHS%\SENDMAIL_WEEKLY.vbs >nul
	timeout /t 1 >nul
	REM 
	REM cscript %PATHS%\DEFECT_RECEIVE_DATA_INSERT.vbs >nul
	REM timeout /t 1 >nul 
 timeout /t 10 >nul
 rem pause