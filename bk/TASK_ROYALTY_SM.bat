@echo off
	SET PATHS=G:\task_royalty_sale
	SET NEWLN=^& echo. 
	SET "TASK=%PATHS% RUN TASK RECEIVEIN OF %DATE% %TIME% %NEWLN% %NEWLN%

	cscript %PATHS%\ROYALTY_SM_INSERT.vbs >nul
	timeout /t 2 >nul	

 
 
 pause