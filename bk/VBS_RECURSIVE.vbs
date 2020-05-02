 Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2)	
 Dim Chr_str(2)
				Step_p = 1
				GH = 7
				Digit = 50 / GH
				PROGRESS = 0
				
				If GH > 50 Then
				   Step_p = Cint(GH / 50)	
				   Digit = 1				   
				End If				
				
				
				 
				Chr_str(0) =  CInt( Digit )
				Chr_str(1) =   50 - ( GH * CInt( Chr_str(0) ) )
				
				If Chr_str(1) < 0 Then
					Chr_str(1) = 0				
				End If

				'call Echo( Chr_str(0) )
				
				'If FullRecord < 100 Then
				'	GH = 80*100			
				'	PG = ( GH * STEP_PRO ) / 100					
				'End If
				
				For i = 0 to 5 Step 1
				Stdout.Write " LEVEL " & (i + 1) & Chr(32)
				PROGRESS =  Recursive(GH, Chr_str, Step_p, 0)
				Stdout.Write " Complete! " &  Chr(33) & Chr(32) & PROGRESS & Chr(32) &  Chr(37) & Chr(32) & "Record data " & FormatNumber(GH,0) & " record " & vbLf
				'Stdout.WriteLine ""			
				NEXT
				
				
				Stdout.WriteLine ""
				'Stdout.WriteLine ""

	Function Recursive( FullRecord, Pros, Stp, ByRef Ind )
				'  Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2)	
				
						IF CDbl(Ind) < FullRecord Then
							Stdout.Write String(Pros(0),Chr(254))
							WScript.Sleep 100				
							'MsgBox PG
							Recursive = Recursive( FullRecord ,Pros, Stp, (Ind + Stp) )
						Else
							Stdout.Write String(Pros(1),Chr(254))
							Recursive = 100
						'Else
						'	Recursive = Recursive( FullRecord ,Pros,(Ind + 1), PG, STEP_PRO )
						End IF					
	End Function

	Sub Echo(myStr)
		MsgBox myStr
		Wscript.Quit
	End Sub




' Function GetAllSubFolders(RootFolder, ByRef pSubfoldersList)
'     Dim fso, SubFolder, root
'     Set fso = CreateObject("scripting.filesystemobject")
'     set root = fso.getfolder(RootFolder)
'     For Each Subfolder in root.SubFolders
'         If pSubFoldersList = "" Then
'             pSubFoldersList = Subfolder.Path         
'         Else
'             pSubFoldersList = pSubFoldersList & "|" & Subfolder.Path
'         End If
'         GetAllSubFolders Subfolder, pSubFoldersList
'     Next
'     GetAllSubFolders = pSubFoldersList
' End Function