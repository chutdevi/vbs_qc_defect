@echo off

	SET PATHS=G:\vbs_develop

	cscript %PATHS%\VBS_RECURSIVE.vbs >nul
	timeout /t 30

 pause