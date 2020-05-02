'On Error Resume Next

dim ej, cn, path, log, filename, namelog
Dim Chr_str(2)
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2) 
    path = "G:\vbs_qc_defect"
	filename = "NG_MSQL_INSERT"
	namelog = filename & "_Log"
	table = "DEFECT_REPORT"
	data_file =   "NG_DATA_MSQL"
	insert_file = "NG_INST_MSQL" 
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile( path & "\work\" & data_file & ".sql", 1)
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
ej.Execute("TRUNCATE " & table)

	Set sql_sv = ej.Execute(content)

	Set count_sql = ej.Execute( " SELECT COUNT(FG.ITEM_CD) CC FROM ( " &  content  & " ) FG " )
	'call Write_File(" SELECT COUNT(FG.ITEM_CD) CC FROM ( " &  content  & " ) FG ", path & "\log\FGH.sql")
	Set file = fso.OpenTextFile( path & "\work\" & insert_file & ".sql", 1 )
	content = file.ReadAll

	count_pro = count_sql(0).value
	'call Echo(CInt(count_pro))
	'Wscript.'Quit
	
	GH = CInt(count_pro)
	Step_p = 1
	Ind = 0
	
	Digit = 50 \ GH
	
	
	If GH > 50 Then
	Step_p = Cint(GH \ 50)	
	Digit = 1				   
	End If				
	
	Chr_str(0) =   CInt( Digit )
	Chr_str(1) =  ( (GH *  Chr_str(0)) MOD 50 )
	'MsgBox Chr_str(0) & ", " & Chr_str(1) & ", " & Step_p
	If Chr_str(1) < 0 Then
		Chr_str(1) = 0				
	End If	
	itt = 0
	'Stdout.WriteLine " INPUT DATA ON ORACLE TO MYSQL BY OOR CHICKEN " 
    'Stdout.WriteLine " START TASK" & Chr(32) & Now	
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

		itt = itt + 1						
						IF (itt MOD Step_p) = 0 AND Ind < 50 Then
							Ind = Ind + 1	
						   Stdout.Write String(Chr_str(0),Chr(254))
						END if
		
	
		sql_sv.MoveNext
		
	   LOOP
	   
	    'Stdout.Write String(Chr_str(1),Chr(254))
		PROGRESS = 100
		Stdout.Write Chr(32) & " Complete " & Chr(33) & Chr(32) & PROGRESS & Chr(32) &  Chr(37) & Chr(32) & "Record data " & FormatNumber(itt,0) & " record " & vbLf 
		Stdout.WriteLine " END   TASK" & Chr(32) & Now
		'MsgBox itt
	   
	   'MsgBox Ind
	   
	   
	   
	   
	   
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