@echo off
	SET PATHS=G:\vbs_qc_defect

	cscript %PATHS%\NG_MONTH_EXPL_INSERT.vbs >nul
	rem timeout /t 1 >nul
	
	cscript %PATHS%\EXPORT_QUERY_NG_DAILY_SUM.vbs >nul

	cscript %PATHS%\NG_MONTH_EXPL_DAY_INSERT.vbs >nul
    timeout /t 1 >nul	
	
	cscript %PATHS%\NG_MONTH_MSQL_INSERT.vbs >nul
	rem timeout /t 1 >nul
	
	
	ECHO  EXPORT...
	cscript %PATHS%\EXPORT_REPORT_MONTH.vbs >nul
	cscript %PATHS%\EXPORT_REPORT_DAY.vbs >nul
     timeout /t 1 >nul


	REM ECHO  SENDMAIL...
	REM cscript %PATHS%\SENDMAIL.vbs >nul
	REM timeout /t 1 >nul
	REM 
	REM cscript %PATHS%\DEFECT_RECEIVE_DATA_INSERT.vbs >nul
	REM timeout /t 1 >nul 
 timeout /t 10 >nul
 rem pause