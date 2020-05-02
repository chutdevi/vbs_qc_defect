'On Error Resume Next
	
dim ej, cn, path, log, filename, namelog, StdOut, Ng, Ref, Ind, GH, Pd, Line, Upd
 
    path = "G:\vbs_develop"
	filename = "GENNERATE_PD_COMP"
	namelog = filename & "_Log"
	Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2)
	Set fso = CreateObject("Scripting.FileSystemObject")

	content = Read_File(path & "\work\" , "BOM_COMP_DEP.sql") 
	
	
	
	
	'Str_comp = content '& "'" & "10297584" & "'"
	Str_par 	= "SELECT ITEM_MASTER, PD, LINE, KEY_UPD FROM BOM_PD_MASTER"' WHERE ITEM_MASTER IN ( '898351-8350', '821011-4741' ) "
	Rcs_comp    = Read_File( path & "\work\" , "BOM_COMP_DEP.sql" )


	Rcs_comp_hd = Read_File( path & "\work\" , "BOM_PD_COMP_HEAD.sql" )
	Rcs_comp_ft = Read_File( path & "\work\" , "BOM_PD_COMP_FOOT.sql" )
	
	

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


ej.Execute("TRUNCATE BOM_PD_COMP")
'ej.Execute("TRUNCATE BOM_COMP")
Set sql_master = ej.Execute( " SELECT COUNT( BC.ITEM_MASTER ) MC FROM ( " & Str_par & ") BC" )
Set sql_sv     = ej.Execute( Str_par )	
	Cnt_master = CInt( sql_master(0).value )
	'call Echo(Cnt_master)
	Ref = sql_sv("ITEM_MASTER").value
	Pd   = sql_sv(1).value
	Line = sql_sv(2).value
	Upd  = sql_sv(3).value
	content = Read_File( path & "\work\" , "BOM_PD_COMP_INS.sql" )
	options = content
	GH = Cnt_master
	'STEP_PRO = 2
	'PG = ( GH * STEP_PRO ) / 100
	'PROGRESS = 0
	itt = 1
	Stdout.WriteLine " INPUT DATA ON ORACLE TO MYSQL BY OOR CHICKEN " 
    Stdout.WriteLine " START TASK" & Chr(32) & Now	
	'MsgBox CInt(PG)
 Dim Chr_str(2)

	   Do Until sql_sv.eof
				'call Echo( Put_Sq(sql_sv(0).value ) )
				'Rcs_comp & Put_Sq(part)
				Str_comp = Rcs_comp & Put_Sq(sql_sv(0).value)
				set sql_m  = cn.Execute( " SELECT COUNT( BC.PARENT_ITEM_CD ) MC FROM ( " & Str_comp & ") BC" )				
				set sql_v  = cn.Execute( Str_comp )						
				Ref = sql_sv("ITEM_MASTER").value
				Pd   = sql_sv(1).value
				Line = sql_sv(2).value
				Upd  = sql_sv(3).value				
				Ind = 0
				GH = Cint(sql_m(0).value)
				
				'call Echo (PD & vbNewLine & LINE)
				Stdout.Write " INPUT DATA ITEM " & WhatEver(itt) & Chr(32)
		            Do Until sql_v.eof
					
							options = content

							
								For Each fld In sql_v.Fields
									options = options & Put_Sq(sql_v(fld.Name).value) & ","
								Next				
									options = Mid(options,1,Len(options)-1)
									options = options &  "," & Put_Sq(Pd) & "," & Put_Sq(Line) & "," & Put_Sq(Upd) & "," & Put_Sq(Ref) & ",2 );"
									call Write_File(options, path & "\log\" & namelog & ".sql")
						'call Echo(options)
							ej.Execute(options)
							Lvl = 3
							
							Cnt_ck = Cnt_Sq( sql_v("COMP_ITEM_CD").value )
							
							
							Ng = 1
							IF	Cnt_ck > 0 Then
								'call Echo( sql_sv("COMP_ITEM_CD").value )
								Ng = Put_bom( sql_v("COMP_ITEM_CD").value, content, Lvl )
									
							Else
								call Write_File( "LVL " & Ng, path & "\temp\" & namelog & ".sql")
							End If
							
							Step_p = 1
							
							
							Digit = 50 \ GH
						
						
							If GH > 50 Then
							Step_p = Cint(GH \ 50)	
							Digit = 1				   
							End If				
							
							Chr_str(0) =  CInt( Digit )
							Chr_str(1) =   50 MOD ( GH * CInt( Chr_str(0) ) )
							'MsgBox Chr_str(0) & ", " & Chr_str(1)
							If Chr_str(1) < 0 Then
								Chr_str(1) = 0				
							End If						
						'IF CDbl(Ind) < GH Then
							Stdout.Write String(Chr_str(0),Chr(254))
							
					
													
						'End IF						
					
					Ind = Ind + Step_p	
					sql_v.MoveNext
					LOOP
		Stdout.Write String(Chr_str(1),Chr(254))
		PROGRESS = 100			
		Stdout.Write Chr(32) & " Complete " & Chr(33) & Chr(32) & PROGRESS & Chr(32) &  Chr(37) & Chr(32) & "Item data " & FormatNumber( Ind ,0) & " item " & vbLf 
			
		
		'call Echo(Lvl)		
		'IF CInt(itt) >= CInt(PG) THEN
		'		Stdout.Write Chr(254)
		'		
		'		STEP_PRO = STEP_PRO + 2
		'		PG = (GH * STEP_PRO) / 100
		'		PROGRESS = STEP_PRO
		'		'WScript.Sleep 100
		'End IF			
		'MsgBox CInt(PG)
		itt = itt+1
		sql_sv.MoveNext
				
	   LOOP
	   Stdout.WriteLine " END   TASK" & Chr(32) & Now
		'call Echo(PROGRESS)

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
	Function Put_bom(part ,ins, lvl)
	
	
		'call Echo(part)
		sql_cmp =  Rcs_comp & Put_Sq(part) 
		set sql_c1  = cn.Execute( " SELECT COUNT( BC.PARENT_ITEM_CD ) MC FROM ( " & sql_cmp & ") BC" )		
		set sql_v1  = cn.Execute( sql_cmp )		
		Lv = lvl + 1
		'GH = CInt(sql_c1("MC").value)				
			   Do Until sql_v1.eof
				options = ins
				
					For Each fld In sql_v1.Fields
						options = options & Put_Sq(sql_v1(fld.Name).value) & ","
					Next				
						options = Mid(options,1,Len(options)-1)
						options = options & ","  & Put_Sq(Pd) & "," & Put_Sq(Line) & "," & Put_Sq(Upd) & "," & Put_Sq(Ref) & "," & lvl & " );"
						'call Echo (options)
						call Write_File(options, path & "\log\" & namelog & ".sql")
			
				ej.Execute(options)
				Cnt_ck = Cnt_Sq( sql_v1("COMP_ITEM_CD").value )
				Ng = 1
				
				IF	Cnt_ck > 0 Then
					'call Echo( sql_sv("COMP_ITEM_CD").value )
					'Lvl = 
					Ng = Put_bom( sql_v1("COMP_ITEM_CD").value, content, Lv )
				Else
					call Write_File( "LVL " & Ng, path & "\temp\" & namelog & ".sql")
				End If
															
				'MsgBox CInt(Ng)
				
				sql_v1.MoveNext		
			   LOOP	
		
		'cn.close
		'ej.close			   
		Put_bom = Lvl
	
	
	End Function





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
	
	Function Cnt_Sq(mydata)	
		sql_cmp =  Rcs_comp & Put_Sq( mydata )
		sql_tmp = cn.Execute( " SELECT COUNT( BC.PARENT_ITEM_CD ) MC FROM ( " & sql_cmp & ") BC" )
		
		cont = sql_tmp("MC").value
		'cn.close
		Cnt_Sq = CInt( cont )
	End Function

	Function Put_Sq(mydata)	
		Put_Sq = "'" & mydata & "'"
	End Function
	Function WhatEver(num)
		If(Len(num)=1) Then
			WhatEver=num & ".   "
		ElseIf(Len(num)=2) Then
			WhatEver=num & ".  "
		ElseIf(Len(num)=3) Then
			WhatEver=num & ". "
		Else
			WhatEver=num & "."
		End If
	End Function	
	Sub Echo(myStr)
		MsgBox myStr
		Wscript.Quit
	End Sub

	Function Read_File(myStr, myFname )
		Set file = fso.OpenTextFile( myStr & myFname, 1)
		Read_File = file.ReadAll 		
	End Function
	
	
	Sub Write_File(myStr, myFname)
		Set objFSO = CreateObject("Scripting.FileSystemObject")
			outFile = myFname
		Set	objFile = objFSO.CreateTextFile(outFile,True)
			objFile.WriteLine myStr
			objFile.Close
	End Sub
	
	
	
	