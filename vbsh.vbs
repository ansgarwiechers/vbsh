' ------------------------------------------------------------------------------
' A simple interactive VBScript shell taken from:
'
'   <http://www.kryogenix.org/days/2004/04/01/interactiveVbscript>
'
' For my own convenience I added line-continuation, and import function, and a
' help message.
' ------------------------------------------------------------------------------

Option Explicit

Main


Sub Main()
	Dim line

	Do While True
		WScript.StdOut.Write(">>> ")

		line = Trim(WScript.StdIn.ReadLine)
		Do While Right(line, 2) = " _" Or line = "_"
			line = RTrim(Left(line, Len(line)-1)) & " " & Trim(WScript.StdIn.ReadLine)
		Loop

		If LCase(line) = "exit" Then Exit Do

		If LCase(line) = "help" Or line = "?" Then
			WScript.StdOut.Write "A simple interactive VBScript Shell." & vbNewLine & vbNewLine _
				& vbTab & "help | ?                  Print this help." & vbNewLine _
				& vbTab & "exit                      Exit the shell." & vbNewLine _
				& vbTab & "import ""\PATH\TO\my.vbs""  Load and execute the contents of the VBScript." & vbNewLine _
				& vbNewLine
		Else
			On Error Resume Next
			Err.Clear
			Execute line
			If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
			On Error Goto 0
		End If
	Loop
End Sub

' ------------------------------------------------------------------------------
' Import the first occurrence of the given filename from the working directory
' or any directory in the %PATH%.
'
' <http://gazeek.com/coding/importing-vbs-files-in-your-vbscript-project/>
' ------------------------------------------------------------------------------
Private Sub Import(ByVal filename)
	Dim fso, sh, file, code, dir

	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")

	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
		' filename is not absolute
		If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
			' file doesn't exist in the working directory => iterate over the
			' directories in the %PATH% and take the first occurrence
			' if no occurrence is found => use filename as-is, which will result
			' in an error when trying to open the file
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				If fso.FileExists(fso.BuildPath(dir, filename)) Then
					filename = fso.BuildPath(dir, filename)
					Exit For
				End If
			Next
		End If
		filename = fso.GetAbsolutePathName(filename)
	End If

	Set file = fso.OpenTextFile(filename, 1, False)
	code = file.ReadAll
	file.Close

	ExecuteGlobal(code)
End Sub
