@echo off
 SET  "TA=INSER TASK ^>\[==========================================\] 100%%^<  COMPLETE!!!!"
 set newline=^& echo. 
 
 echo. %TA%
 SET  "TA=%TA%%newline%INSER TASK ^>\[==========================================\] 100%%^<  COMPLETE!!!!"
 echo. %TA%
 pause