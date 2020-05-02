On Error Resume Next

dim ej, cn, path, log, filename, namelog
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2) 
    path = "G:\vbs_develop"
	filename = "MASTER_BOM_PD_INSERT"
	namelog = filename & "_Log"
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile( path & "\work\BOM_PD_MASTER.sql", 1)
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

'MsgBox Now

'Wscript.Quit


cn.connectionstring = "Provider=OraOLEDB.Oracle;Data Source=EXPK;User ID=EXPK;Password=EXPK"
cn.open
ej.connectionstring = "Driver={MySQL ODBC 8.0 Driver}; Data Source=DBEJ; User=monty; Password=some_pass;"
ej.open
ej.Execute("TRUNCATE BOM_PD_MASTER")
	Set sql_sv = cn.Execute(content)

	Set count_sql = cn.Execute( " SELECT COUNT(KEY_UPDATE) CC FROM ( " &  content  & ")  " )

	Set file = fso.OpenTextFile( path & "\work\BOM_PD_MASTER_INS.sql", 1 )
	content = file.ReadAll

	count_pro = count_sql(0).value
	'call Echo(count_pro)
	'Wscript.'Quit
	
	GH = CInt(count_pro)
	STEP_PRO = 2
	PG = ( GH * STEP_PRO ) / 100
	PROGRESS = 0
	itt = 0
	Stdout.WriteLine " INPUT DATA ON ORACLE TO MYSQL BY OOR CHICKEN " 
    Stdout.WriteLine " START TASK" & Chr(32) & Now	
    Stdout.Write " INPUT DATA" & Chr(32)
	'MsgBox CInt(PG)
	   Do Until sql_sv.eof
		options = content
		
			For Each fld In sql_sv.Fields
				options = options & Put_Sq(sql_sv(fld.Name).value) & ","
			Next				
				options = Mid(options,1,Len(options)-1)
				options = options & " );"
				call Write_File(options, path & "\log\" & namelog & ".sql")
	
		ej.Execute(options)
		IF CInt(itt) >= CInt(PG) THEN
				Stdout.Write Chr(254)
				
				STEP_PRO = STEP_PRO + 2
				PG = (GH * STEP_PRO) / 100
				PROGRESS = STEP_PRO
				'WScript.Sleep 100
		End IF			
		'MsgBox CInt(PG)
		
		sql_sv.MoveNext
		itt = itt + 1			
	   LOOP
		Stdout.Write Chr(32) & " Complete " & Chr(33) & Chr(32) & PROGRESS & Chr(32) &  Chr(37) & Chr(32) & "Record data " & FormatNumber(itt,0) & " record " & vbLf 
		Stdout.WriteLine " END   TASK" & Chr(32) & Now
		'MsgBox itt
	   
	   
	   
	   
	   
	   
	   
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

Sub Echo(myStr)
		MsgBox myStr
		Wscript.Quit
End Sub

Sub Write_File(myStr, myFname)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
        outFile = myFname
	Set	objFile = objFSO.CreateTextFile(outFile,True)
		objFile.WriteLine myStr
		objFile.Close
End Sub