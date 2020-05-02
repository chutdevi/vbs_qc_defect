On Error Resume Next

dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
dim path, log, filename, namefile, size

path = "G:\vbs_qc_defect"
filename = "EXPORT_QUERY_NG_DAILY"
namefile = "DATA_NG_DAILY.sql"

'MsgBox filename & Day(Date)




xHttp.Open "GET", "http://192.168.161.147/report/get_query/qc_query_export.php/Query_qc_daily.sql", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile path & "/work/" & namefile, 2 '//overwrite
end with




CONST bytesToKb = 1024
strFile = path & "/bin/" & namefile

SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")
SET objFile = objFSO.GetFile(strFile)

'wscript.echo "File Size: " & CINT(objFile.Size / bytesToKb) & "Kb"
	size = CINT(objFile.Size / bytesToKb)
	Myfile = path & "\log\" & "Log.log"
	If Err.Number <> 0 Or size < 2  Then
				
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set file = fso.OpenTextFile(Myfile, 1)
				conten = file.ReadAll
				file.Close	
			log = Now & " [ " & WeekdayName(DatePart("w", Date())) &" ]" & "[ " & "Error Export" & " ] " & "[ " & filename & " ] "
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

Sub Write_File(myStr, myFname)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
        outFile = myFname
	Set	objFile = objFSO.CreateTextFile(outFile,True)
		objFile.WriteLine myStr
		objFile.Close
End Sub	