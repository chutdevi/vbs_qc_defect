  
  On Error Resume Next
  
'Set fso = CreateObject ("Scripting.FileSystemObject")
'Set stdout = fso.GetStandardStream (1)
'Set stderr = fso.GetStandardStream (2)
'stdout.WriteLine "This will go to standard output."
'stderr.WriteLine "===="
  
  
  Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(2)

    ' WScript.Echo "Hello"
    ' WScript.StdOut.Write "Hello"
    ' WScript.StdOut.WriteLine "Hello"
    ' Stdout.WriteLine "Hello"
	GH = 3000
	STEP_PRO = 2
	PG = ( GH * STEP_PRO ) / 100
	PROGRESS = 0
    Stdout.Write " [ "
    For i = 0 to GH Step 1
		
		IF i >= PG THEN
			Stdout.Write "="
			PROGRESS = STEP_PRO / 2
			STEP_PRO = STEP_PRO + 2
			PG = (GH * STEP_PRO) / 100
			WScript.Sleep 100
		End IF
	Next
	Stdout.Write " ] > " & PROGRESS
	Stdout.WriteLine ""
	
	GH = 500
	STEP_PRO = 2
	PG = ( GH * STEP_PRO ) / 100
	PROGRESS = 0
    Stdout.Write " [ "
    For i = 0 to GH Step 1
		
		IF i >= PG THEN
			Stdout.Write "="
			PROGRESS = STEP_PRO / 2
			STEP_PRO = STEP_PRO + 2
			PG = (GH * STEP_PRO) / 100
			WScript.Sleep 100
		End IF
	For i = 0 to GH Step 1
	Stdout.Write " ] > " & PROGRESS
	Stdout.WriteLine ""
  
  'WScript.Echo(Hello")
  'WScript.Interactive = false
  'WScript.Echo("This wont display")
  'WScript.Interactive = true
  'WScript.Echo("This will display")