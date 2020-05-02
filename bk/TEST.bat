@echo off
	SET PATHS=G:\vbs_develop
	SET NEWLN=^& echo. 
	SET "TASK=%PATHS% RUN TASK RECEIVEIN OF %DATE% %TIME% %NEWLN% %NEWLN%

	cscript %PATHS%\INSERT.vbs >nul
	timeout /t 2 >nul	

 
 
 pause