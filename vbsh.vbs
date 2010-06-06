' ------------------------------------------------------------------------------
' A simple interactive VBScript shell taken from:
'
'   <http://www.kryogenix.org/days/2004/04/01/interactiveVbscript>
'
' For my own convenience I added line-continuation and a help message.
' ------------------------------------------------------------------------------

Option Explicit

Dim line

Do While True
	WScript.StdOut.Write(">>> ")

	line = Trim(WScript.StdIn.ReadLine)
	Do While Right(line, 2) = " _" Or line = "_"
		line = RTrim(Left(line, Len(line)-1)) & " " & Trim(WScript.StdIn.ReadLine)
	Loop

	If LCase(line) = "exit" Then Exit Do

	If LCase(line) = "help" Then
		WScript.StdOut.Write "A simple interactive VBScript Shell." & vbNewLine & vbNewLine _
			& vbTab & "help = Print this help." & vbNewLine _
			& vbTab & "exit = Exit the shell." & vbNewLine _
			& vbNewLine
	Else
		On Error Resume Next
		Err.Clear
		Execute line
		If Err.Number <> 0 Then WScript.StdErr.WriteLine Err.Description & " (0x" & Hex(Err.Number) & ")"
		On Error Goto 0
	End If
Loop
