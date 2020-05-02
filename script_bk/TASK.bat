@echo off
	SET PATHS=G:\prod_task
	SET NEWLN=^& echo. 
	SET "TASK=%PATHS% RUN TASK PRODUCTION REPORT OF %DATE% %TIME% %NEWLN% %NEWLN%
	DEL /Q %PATHS%\temp\*<nul
	
	ECHO. %TASK%
	timeout /t 2 >nul
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[==                                                ] 00%%^<
			cscript %PATHS%\INSERT0.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT0_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[==                                                ] 00%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[==                                                ] 05%%^<
			cscript %PATHS%\INSERT.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[==                                                ] 05%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[=======                                           ] 15%%^<
			cscript %PATHS%\INSERT1.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[=======                                           ] 15%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[==============                                    ] 30%%^<
			cscript %PATHS%\INSERT2.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[==============                                    ] 30%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[=====================                             ] 45%%^<
			cscript %PATHS%\INSERT3.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[=====================                             ] 45%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[=========================                         ] 50%%^<
			cscript %PATHS%\INSERT4.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[=========================                         ] 50%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
			
	CLS
		ECHO. %TASK%INSERT    TASK ^>[=====================================             ] 65%%^<
			cscript %PATHS%\INSERT5.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[=====================================             ] 65%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[============================================      ] 80%%^<
			cscript %PATHS%\INSERT6.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[============================================      ] 80%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
	
	CLS
		ECHO. %TASK%INSERT    TASK ^>[================================================= ] 95%%^<
			cscript %PATHS%\INSERT7.vbs >nul
			timeout /t 2 >nul
		If EXIST %PATHS%\temp\INSERT_Error.log ( 
				SET "TASK=%TASK%INSER     TASK ^>[================================================= ] 95%%^<   FAILED!!!!!! %NEWLN%" 				
				GOTO :ERROR 	
		)
	
	REM =================================================================================================  INSERT =========================================================================================
		
	CLS			
		  ECHO. %TASK%INSER     TASK ^>[==================================================] 100%%^<  COMPLETE!!!!		  
	  SET "TASK=%TASK%INSER     TASK ^>[==================================================] 100%%^<  COMPLETE!!!! %NEWLN%"
	timeout /t 2 >nul
	

	
	
	REM =================================================================================================  EXPORT =========================================================================================
	 CLS	
	 	  ECHO. %TASK%EXPORT    TASK ^>[=========================                         ] 50%%^<
	 cscript %PATHS%\EXPORT.vbs >nul
	 timeout /t 2 >nul
	 If EXIST %PATHS%\temp\EXPORT_Error.log ( 
	 		SET "TASK=%TASK%EXPORT    TASK ^>[=========================                         ] 50%%^<   FAILED!!!!!! %NEWLN%" 				
	 		GOTO :ERROR 	
	 )	
	 CLS	
	 	  ECHO. %TASK%EXPORT    TASK ^>[==================================================] 100%%^<  COMPLETE!!!!
	 	  
	   SET "TASK=%TASK%EXPORT    TASK ^>[==================================================] 100%%^<  COMPLETE!!!! %NEWLN%"
	 timeout /t 2 >nul
	 
     
	 CLS	
	 	  ECHO. %TASK%SEND MAIL TASK ^>[=========================                         ] 50%%^<
	 cscript %PATHS%\SENDMAIL.vbs >nul
	 timeout /t 2 >nul
	 If EXIST %PATHS%\temp\SENDMAIL_Error.log ( 
	 		SET "TASK=%TASK%EXPORT    TASK ^>[=========================                         ] 50%%^<   FAILED!!!!!! %NEWLN%" 				
	 		GOTO :ERROR 	
	 )		 
	 CLS			
	 	  ECHO. %TASK%SEND MAIL TASK ^>[==================================================] 100%%^<  COMPLETE!!!!
	   SET "TASK=%TASK%SEND MAIL TASK ^>[==================================================] 100%%^<  COMPLETE!!!! %NEWLN%"
	 timeout /t 2 >nul		  
     
	timeout /t 4 
	
	EXIT
	
	REM ============================================== ERROR ==========================================
	
	:ERROR
	CLS
		ECHO. %TASK%
	cscript %PATHS%\ALERT.vbs >nul
	timeout /t 2 >nul			
	timeout /t 30
		EXIT
		REM start http://192.168.82.31/report_service/Report_sendmail/only_send_mail

		REM timeout /t 120 >nul

		REM start http://192.168.82.31/report_service/Report_sendmail/pro_send_mail