@echo off
	SET PATHS=G:\vbs_develop
	cscript %PATHS%\GENERATE_COMP.vbs >nul
	timeout /t 30 	
 
 pause