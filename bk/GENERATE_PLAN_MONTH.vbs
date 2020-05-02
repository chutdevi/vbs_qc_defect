'On Error Resume Next
	
dim ej, cn, path, log, filename, namelog, StdOut, Ng, Ref
 
    path = "G:\vbs_develop"
	filename = "GENNERATE"
	namelog = filename & "_Log"
	Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.OpenTextFile( path & "\work\" & "BOM_PLAN_MONTH.sql", 1)
	content = file.ReadAll 
	
	Str_comp = content '& "'" & "10297584" & "'"
	Str_par = "SELECT ITEM_MASTER FROM BOM_MASTER" ' WHERE ITEM_MASTER IN ( '898320-8638-SP', '898232-6241-EP' ) "
	Rcs_comp = content
	
	

	set ej = CreateObject("ADODB.Connection")
	set cn = CreateObject("ADODB.Connection")

	If Not fso.FileExists( path & "\log\Log.log" ) Then	
	
		Set objFSO=CreateObject("Scripting.FileSystemObject")					
			outFile= path & "\log\Log.log"
		Set objFile= fso.CreateTextFile(outFile,True)
			objFile.WriteLine "[ LOG FOR TASK AUTOMATIC GENNERATE BOM ]"
			objFile.Close
	END IF

'MsgBox Now

'Wscript.Quit


cn.connectionstring = "Provider=OraOLEDB.Oracle;Data Source=EXPK;User ID=EXPK;Password=EXPK"
cn.open
ej.connectionstring = "Driver={MySQL ODBC 8.0 Driver}; Data Source=DBEJ; User=monty; Password=some_pass;"
ej.open


ej.Execute("TRUNCATE BOM_PLANMONTHS")
	Set sql_master = cn.Execute( " SELECT COUNT( BC.ITEM_CD ) MC FROM ( " & Rcs_comp & ") BC" )
	Set sql_v      = cn.Execute( Rcs_comp )	
		Cnt_master = CInt( sql_master(0).value )
		'call Echo(Cnt_master)
	Set file = fso.OpenTextFile( path & "\work\BOM_PLAN_MONTH_INS.sql", 1)
	content = file.ReadAll
	options = content
		GH = Cnt_master
		STEP_PRO = 2
		PG = ( GH * STEP_PRO ) / 100
		PROGRESS = 0
		itt = 0
		Stdout.WriteLine " INPUT DATA ON ORACLE TO MYSQL BY OOR CHICKEN " 
		Stdout.WriteLine ""	
		Stdout.Write " INPUT DATA" & Chr(32)
	'MsgBox CInt(PG)
	
		            Do Until sql_v.eof
					
							options = content
							
								'For Each fld In sql_v.Fields
								options = options & Put_Sq( sql_v(0).value ) & "," & Put_Sq( sql_v(1).value ) & " );"
								'call Echo(options)
								'Next				
									'options = Mid(options,1,Len(options)-1)
									'options = options & "," & Put_Sq(Ref) & ",2 );"
								call Write_File(options, path & "\log\" & namelog & ".sql")
									
								ej.Execute(options)				
							IF CInt(itt) >= CInt(PG) THEN
								Stdout.Write Chr(254)
								
								STEP_PRO = STEP_PRO + 2
								PG = (GH * STEP_PRO) / 100
								PROGRESS = STEP_PRO
								'WScript.Sleep 100
							End IF	
				    itt = itt + 1
					sql_v.MoveNext
					LOOP		

		
		'call Echo(Lvl)		
		
		Stdout.Write Chr(32) & " Complete " & Chr(33) & Chr(32) & PROGRESS & Chr(32) &  Chr(37) & Chr(32) & "Item data " & FormatNumber(itt,0) & " item " & vbLf 
		Stdout.WriteLine ""	
		call Echo(PROGRESS)

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
			call Write_File( conten & log, Myfile )
			Wscript.Quit
	End If



'======================================== Function =================================================






	Function Recursive( FullRecord, ByRef Ind, ByRef PG, ByRef STEP_PRO )
				'  Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2)			
						IF Ind >= PG Then
							Stdout.Write Chr(219)
							WScript.Sleep 100
							'MsgBox PG
							Recursive = Recursive(FullRecord ,(Ind + 1), ( (FullRecord * STEP_PRO) / 100 ), (STEP_PRO + 2) )
						ElseIf  Ind >= FullRecord Then
							'MsgBox PG
							Recursive = 100
						Else
							Recursive = Recursive( FullRecord ,(Ind + 1), PG, STEP_PRO )
						End IF					
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