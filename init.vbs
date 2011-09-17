' Sample init script, save as %USERPROFILE%\init.vbs

Dim sh  : Set sh  = CreateObject("WScript.Shell")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

'! List the contents of the current directory. First subfolders, then files.
Sub Ls()
	Dim thisFolder, f

	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

	Set thisFolder = fso.GetFolder(Pwd)
	For Each f In thisFolder.SubFolders
		WScript.Echo FormatProperties(f)
	Next
	For Each f In thisFolder.Files
		WScript.Echo FormatProperties(f)
	Next
	Set thisFolder = Nothing

	Set fso = Nothing
End Sub

'! Print and return the name of the current directory.
Function Pwd()
	Dim sh : Set sh = CreateObject("WScript.Shell")
	WScript.Echo sh.CurrentDirectory
	Pwd = sh.CurrentDirectory
	Set sh = Nothing
End Function

'! Change the current directory to the given directory and return the new
'! current directory.
'!
'! @param  dir  The new working directory.
Function Cd(ByVal dir)
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim sh  : Set sh  = CreateObject("WScript.Shell")

	dir = fso.GetAbsolutePathName(dir)
	If fso.FolderExists(dir) Then
		sh.CurrentDirectory = dir
	Else
		WScript.Echo "Folder " & dir & " doesn't exist."
	End If
	Cd = sh.CurrentDirectory

	Set fso = Nothing
	Set sh  = Nothing
End Function

'! Return a string representation of attributes, size, last modified date and
'! name of the given object. The object must be a file or folder object. The
'! string representation has the form:
'!
'!   ATTRIBUTES  SIZE  LAST_MODIFIED  NAME
'!
'! @param  f  The file or folder object.
'! @return A string representation of the object's properties.
Private Function FormatProperties(obj)
	' Constants defined locally to make the function self-contained.
	Const ATTR_READONLY   = &h01
	Const ATTR_HIDDEN     = &h02
	Const ATTR_SYSTEM     = &h04
	Const ATTR_DIRECTORY  = &h10
	Const ATTR_ARCHIVE    = &h20
	Const ATTR_ALIAS      = &h40
	Const ATTR_COMPRESSED = &h80

	Dim order, attributes, attr, magnitude, size

	' SI order of magnitude prefixes for file sizes
	Set order = CreateObject("Scripting.Dictionary")
	order.Add 0, " "
	order.Add 1, "k"
	order.Add 2, "M"
	order.Add 3, "G"
	order.Add 4, "T"

	Set attributes = CreateObject("Scripting.Dictionary")
	attributes.Add 0              , "-"
	attributes.Add ATTR_READONLY  , "r"
	attributes.Add ATTR_HIDDEN    , "h"
	attributes.Add ATTR_SYSTEM    , "s"
	attributes.Add ATTR_DIRECTORY , "d"
	attributes.Add ATTR_ARCHIVE   , "a"
	attributes.Add ATTR_ALIAS     , "@"
	attributes.Add ATTR_COMPRESSED, "c"

	attr = attributes(obj.Attributes And ATTR_DIRECTORY) _
		& attributes(obj.Attributes And ATTR_READONLY) _
		& attributes(obj.Attributes And ATTR_HIDDEN) _
		& attributes(obj.Attributes And ATTR_SYSTEM) _
		& attributes(obj.Attributes And ATTR_ARCHIVE) _
		& attributes(obj.Attributes And ATTR_COMPRESSED) _
		& attributes(obj.Attributes And ATTR_ALIAS)

	magnitude = Int(log(obj.Size) / log(2) / 10) ' base 2, chunk size 1024
	size = FormatNumber(obj.Size / 1024^magnitude, 0) & order(magnitude)

	FormatProperties = attr & "  " & Right(String(5, " ") & size, 5) _
		& "  " & obj.DateLastModified & "  " & obj.Name
	If obj.Attributes And ATTR_DIRECTORY Then
		FormatProperties = FormatProperties & "\"
	End If
End Function
