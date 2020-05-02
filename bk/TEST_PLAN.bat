@echo off
	SET PATHS=G:\vbs_develop

	cscript %PATHS%\GENERATE_PLAN_MONTH.vbs >nul
	timeout /t 20 

 
 
 REM pause