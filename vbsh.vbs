'! A simple interactive VBScript shell.
'!
'! @see http://www.kryogenix.org/days/2004/04/01/interactiveVbscript
'!
'! For my own convenience I added line-continuation, some helper functions (cd,
'! pwd, ls, import), and a help message.

Option Explicit

Private Const InitScript = "%USERPROFILE%\init.vbs"

Main


Sub Main()
	Dim line

	ImportInitScript
	Help

	Do While True
		WScript.StdOut.Write(">>> ")

		line = Trim(WScript.StdIn.ReadLine)
		Do While Right(line, 2) = " _" Or line = "_"
			line = RTrim(Left(line, Len(line)-1)) & " " & Trim(WScript.StdIn.ReadLine)
		Loop

		If LCase(line) = "exit" Then Exit Do

		If line = "?" Then
			Help
		Else
			On Error Resume Next
			Err.Clear
			Execute line
			If Err.Number <> 0 Then WScript.StdErr.WriteLine Trim(Err.Description & " (0x" & Hex(Err.Number) & ")")
			On Error Goto 0
		End If
	Loop
End Sub

'! Import initialization script if present.
Private Sub ImportInitScript
	Dim sh, fso, path, initScriptExists

	Set sh  = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")

	path = sh.ExpandEnvironmentStrings(InitScript)
	initScriptExists = fso.FileExists(path)

	Set sh  = Nothing
	Set fso = Nothing

	If initScriptExists Then Import path
End Sub

'! Print a help message.
Private Sub Help()
	WScript.StdOut.Write "A simple interactive VBScript Shell." & vbNewLine & vbNewLine _
		& vbTab & "help | ?                  Print this help." & vbNewLine _
		& vbTab & "import ""\PATH\TO\my.vbs""  Load and execute the contents of the VBScript." & vbNewLine _
		& vbTab & "exit                      Exit the shell." & vbNewLine _
		& vbNewLine & "Customize with an (optional) init script '" & InitScript & "'." & vbNewLine _
		& vbNewLine
End Sub

'! Import the first occurrence of the given filename from the working directory
'! or any directory in the %PATH%.
'!
'! @param  filename   Name of the file to import.
'!
'! @see http://gazeek.com/coding/importing-vbs-files-in-your-vbscript-project/
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

	Set fso = Nothing
	Set sh = Nothing
End Sub
