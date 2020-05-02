On Error Resume Next

dim ej, cn, path, log, filename, filedata, fileins, namelog, tab
 
    path = "G:\prod_task"
	filename = "INSERT3"
	filedata = "PROD_PD03_DAT.sql"
	fileins  = "PROD_PD03_INS.sql"
	namelog = "prod_pd03_datLog"
	tab = "PROD_PD03"
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile( path & "\work\" & filedata, 1)
content = file.ReadAll
'MsgBox content

'Wscript.Quit
set ej = CreateObject("ADODB.Connection")
set cn = CreateObject("ADODB.Connection")

If Not fso.FileExists( path & "\log\Log.log" ) Then	
	Set objFSO=CreateObject("Scripting.FileSystemObject")					
		outFile= path & "\log\Log.log"
	Set objFile= fso.CreateTextFile(outFile,True)
		objFile.WriteLine "[ LOG FOR TASK AUTOMATIC PRODUCTION REPORT ]"
		objFile.Close
END IF




cn.connectionstring = "Provider=OraOLEDB.Oracle;Data Source=EXPK;User ID=EXPK;Password=EXPK"
cn.open
ej.connectionstring = "Driver={MySQL ODBC 8.0 Driver}; Data Source=DBEJ; User=monty; Password=some_pass;"
ej.open
ej.Execute("TRUNCATE " & tab)
Set sql_sv = ej.Execute(content)

Set file = fso.OpenTextFile( path & "\work\" & fileins, 1)
content = file.ReadAll
options = ""
	   Do Until sql_sv.eof
				options = content
			For Each fld In sql_sv.Fields
				options = options & Put_Sq(sql_sv(fld.Name).value) & ","
			Next				
				options = Mid(options,1,Len(options)-1)
				options = options & " );"
		ej.Execute(options)
		sql_sv.MoveNext
	   LOOP
	   call Write_File(options, path & "\log\" & namelog & ".sql")
cn.close
ej.close
    Set cn = Nothing
	set ej = Nothing
	Myfile = path & "\log\" & "Log.log"
	If Err.Number <> 0 Then
				
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set file = fso.OpenTextFile(Myfile, 1)
				conten = file.ReadAll
				file.Close	
			log = Now & " [ " & WeekdayName(DatePart("w", Date())) &" ]" & "[ " & Err.Description & " ] " & "[ " & filename & " ] "
			'MsgBox log
			call Write_File(conten & log, Myfile)
			Myfile = path & "\temp\" & filename & "_Error.log"
			call Write_File(conten & log, Myfile)
			Wscript.Quit
	Else
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set file = fso.OpenTextFile(Myfile, 1)
				conten = file.ReadAll
				file.Close	
			log = Now & " [ " & WeekdayName(DatePart("w", Date())) &" ]" & "[ " & " Complete! " & " ] " & "[ " & filename & " ] "
			'MsgBox log
			call Write_File(conten & log, Myfile)
			Wscript.Quit
	End If



'======================================== Function =================================================

Function myDateFormat(myDate,opr)
    d = WhatEver(Day(myDate))
    m = WhatEver(Month(myDate))    
    y = Year(myDate)
    myDateFormat= y & opr & m & opr & d
End Function

Function myTimeFormat(myTime,opr)
    h = WhatEver(Hour(myTime))
    m = WhatEver(Minute(myTime))    
    s = WhatEver(Second(myTime))
    myTimeFormat= h & opr & m & opr & s
End Function

Function WhatEver(num)
    If(Len(num)=1) Then
        WhatEver="0"&num
    Else
        WhatEver=num
    End If
End Function

Function Put_Sq(mydata)	
	Put_Sq = "'" & mydata & "'"
End Function


Sub Write_File(myStr, myFname)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
        outFile = myFname
	Set	objFile = objFSO.CreateTextFile(outFile,True)
		objFile.WriteLine myStr
		objFile.Close
End Sub