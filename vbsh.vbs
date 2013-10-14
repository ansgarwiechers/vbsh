'! A simple interactive VBScript shell.
'!
'! @see http://www.kryogenix.org/days/2004/04/01/interactiveVbscript
'!
'! For my own convenience I added line-continuation, some helper functions (cd,
'! pwd, ls, import), and a help message.

Option Explicit

Private Const InitScript    = "%USERPROFILE%\init.vbs"
Private Const Documentation = "script56.chm"

Private helpKeywords, helpfile

Set helpKeywords = CreateObject("Scripting.Dictionary")
	helpKeywords.CompareMode = vbTextCompare
	helpKeywords.Add "-", "html/vsoprSubtract.htm"
	helpKeywords.Add "&", "html/vsoprConcatenation.htm"
	helpKeywords.Add "*", "html/vsoprMultiply.htm"
	helpKeywords.Add "/", "html/vsoprDivide.htm"
	helpKeywords.Add "\", "html/vsoprIntegerDivide.htm"
	helpKeywords.Add "^", "html/vsoprExponentiation.htm"
	helpKeywords.Add "+", "html/vsoprAdd.htm"
	helpKeywords.Add "=", "html/vsoprAssignment.htm"
	helpKeywords.Add "abs", "html/vsfctAbs.htm"
	helpKeywords.Add "add dictionary", "html/jsmthadddictionary.htm"
	helpKeywords.Add "add folder", "html/jsmthaddfolders.htm"
	helpKeywords.Add "addition operator", "html/vsoprAdd.htm"
	helpKeywords.Add "addition", "html/vsoprAdd.htm"
	helpKeywords.Add "and", "html/vsoprAnd.htm"
	helpKeywords.Add "arithmetic operators", "html/vsidxArithmetic.htm"
	helpKeywords.Add "array", "html/vsfctArray.htm"
	helpKeywords.Add "asc", "html/vsfctAsc.htm"
	helpKeywords.Add "ascii", "html/vsmscANSITable.htm"
	helpKeywords.Add "assignment operator", "html/vsoprAssignment.htm"
	helpKeywords.Add "assignment", "html/vsoprAssignment.htm"
	helpKeywords.Add "atendofline", "html/jsproatEndOfLine.htm"
	helpKeywords.Add "atendofstream", "html/jsproatEndOfStream.htm"
	helpKeywords.Add "atn", "html/vsfctAtn.htm"
	helpKeywords.Add "attributes", "html/jsproAttributes.htm"
	helpKeywords.Add "availablespace", "html/jsproAvailableSpace.htm"
	helpKeywords.Add "buildpath", "html/jsmthBuildPath.htm"
	helpKeywords.Add "call", "html/vsstmCall.htm"
	helpKeywords.Add "case", "html/vsstmSelectCase.htm"
	helpKeywords.Add "cbool", "html/vsfctCBool.htm"
	helpKeywords.Add "cbyte", "html/vsfctCByte.htm"
	helpKeywords.Add "ccur", "html/vsfctCCur.htm"
	helpKeywords.Add "cdate", "html/vsfctCDate.htm"
	helpKeywords.Add "cdbl", "html/vsfctCDbl.htm"
	helpKeywords.Add "chr", "html/vsfctChr.htm"
	helpKeywords.Add "cint", "html/vsfctCInt.htm"
	helpKeywords.Add "class", "html/vsstmClass.htm"
	helpKeywords.Add "class_initialize", "html/vsevtInitialize.htm"
	helpKeywords.Add "class_terminate", "html/vsevtTerminate.htm"
	helpKeywords.Add "clear", "html/vsmthClear.htm"
	helpKeywords.Add "clng", "html/vsfctCLng.htm"
	helpKeywords.Add "close", "html/jsmthClose.htm"
	helpKeywords.Add "column", "html/jsprocolumn.htm"
	helpKeywords.Add "comparemode", "html/jsprocompareMode.htm"
	helpKeywords.Add "comparison operators", "html/vsgrpComparison.htm"
	helpKeywords.Add "concatenation operator", "html/vsoprConcatenation.htm"
	helpKeywords.Add "concatenation operators", "html/vsidxConcatenation.htm"
	helpKeywords.Add "concatenation", "html/vsoprConcatenation.htm"
	helpKeywords.Add "const", "html/vsstmConst.htm"
	helpKeywords.Add "constants", "html/vsconVBScript.htm"
	helpKeywords.Add "conversion function", "html/vsidxConversion.htm"
	helpKeywords.Add "conversion functions", "html/vsidxConversion.htm"
	helpKeywords.Add "copy", "html/jsmthCopy.htm"
	helpKeywords.Add "copyfile", "html/jsmthCopyFile.htm"
	helpKeywords.Add "copyfolder", "html/jsmthCopyFolder.htm"
	helpKeywords.Add "cos", "html/vsfctCos.htm"
	helpKeywords.Add "count", "html/jsprocount.htm"
	helpKeywords.Add "createfolder", "html/jsmthCreateFolder.htm"
	helpKeywords.Add "createobject", "html/vsfctCreateObject.htm"
	helpKeywords.Add "createtextfile", "html/jsmthcreateTextFile.htm"
	helpKeywords.Add "csng", "html/vsfctCSng.htm"
	helpKeywords.Add "cstr", "html/vsfctCStr.htm"
	helpKeywords.Add "date", "html/vsfctDate.htm"
	helpKeywords.Add "dateadd", "html/vsfctDateAdd.htm"
	helpKeywords.Add "datecreated", "html/jsproDateCreated.htm"
	helpKeywords.Add "datediff", "html/vsfctDateDiff.htm"
	helpKeywords.Add "datelastaccessed", "html/jsproDateLastAccessed.htm"
	helpKeywords.Add "datelastmodified", "html/jsproDateLastModified.htm"
	helpKeywords.Add "datepart", "html/vsfctDatePart.htm"
	helpKeywords.Add "dateserial", "html/vsfctDateSerial.htm"
	helpKeywords.Add "datevalue", "html/vsfctDateValue.htm"
	helpKeywords.Add "day", "html/vsfctDay.htm"
	helpKeywords.Add "delete", "html/jsmthDelete.htm"
	helpKeywords.Add "deletefile", "html/jsmthDeleteFile.htm"
	helpKeywords.Add "deletefolder", "html/jsmthDeleteFolder.htm"
	helpKeywords.Add "description", "html/vsproDescription.htm"
	helpKeywords.Add "dictionary add", "html/jsmthadddictionary.htm"
	helpKeywords.Add "dictionary", "html/jsobjDictionary.htm"
	helpKeywords.Add "dim", "html/vsstmDim.htm"
	helpKeywords.Add "division operator", "html/vsoprDivide.htm"
	helpKeywords.Add "division", "html/vsoprDivide.htm"
	helpKeywords.Add "do loop", "html/vsstmDo.htm"
	helpKeywords.Add "do", "html/vsstmDo.htm"
	helpKeywords.Add "drive object", "html/jsobjDrive.htm"
	helpKeywords.Add "drive property", "html/jsproDrive.htm"
	helpKeywords.Add "drive", "html/jsobjDrive.htm"
	helpKeywords.Add "driveexists", "html/jsmthDriveExists.htm"
	helpKeywords.Add "driveletter", "html/jsproDriveLetter.htm"
	helpKeywords.Add "drives collection", "html/jscolDrives.htm"
	helpKeywords.Add "drives property", "html/jsproDrives.htm"
	helpKeywords.Add "drives", "html/jscolDrives.htm"
	helpKeywords.Add "drivetype", "html/jsproDriveType.htm"
	helpKeywords.Add "else", "html/vsstmIf.htm"
	helpKeywords.Add "elseif", "html/vsstmIf.htm"
	helpKeywords.Add "empty", "html/vskeyEmpty.htm"
	helpKeywords.Add "eqv", "html/vsoprEqv.htm"
	helpKeywords.Add "erase", "html/vsstmErase.htm"
	helpKeywords.Add "err", "html/vsobjErr.htm"
	helpKeywords.Add "error", "html/vtoriErrors.htm"
	helpKeywords.Add "eval", "html/vsfctEval.htm"
	helpKeywords.Add "execute regexp", "html/vsmthExecute.htm"
	helpKeywords.Add "execute", "html/vsstmExecute.htm"
	helpKeywords.Add "executeglobal", "html/vsstmExecuteGlobal.htm"
	helpKeywords.Add "exists", "html/jsmthExists.htm"
	helpKeywords.Add "exit", "html/vsstmExit.htm"
	helpKeywords.Add "exp", "html/vsfctExp.htm"
	helpKeywords.Add "explicit", "html/vsstmOptionExplicit.htm"
	helpKeywords.Add "exponent", "html/vsoprExponentiation.htm"
	helpKeywords.Add "false", "html/vskeyFalse.htm"
	helpKeywords.Add "file", "html/jsobjFile.htm"
	helpKeywords.Add "fileexists", "html/jsmthFileExists.htm"
	helpKeywords.Add "files collection", "html/jscolFiles.htm"
	helpKeywords.Add "files property", "html/jsproFiles.htm"
	helpKeywords.Add "files", "html/jscolFiles.htm"
	helpKeywords.Add "filesystem", "html/jsproFileSystem.htm"
	helpKeywords.Add "filesystemobject", "html/jsobjFileSystem.htm"
	helpKeywords.Add "filter", "html/vsfctFilter.htm"
	helpKeywords.Add "firstindex", "html/vsproFirstIndex.htm"
	helpKeywords.Add "fix", "html/vsfctInt.htm"
	helpKeywords.Add "folder add", "html/jsmthaddfolders.htm"
	helpKeywords.Add "folder", "html/jsobjFolder.htm"
	helpKeywords.Add "folderexists", "html/jsmthFolderExists.htm"
	helpKeywords.Add "folders", "html/jscolFolders.htm"
	helpKeywords.Add "for each next", "html/vsstmForEach.htm"
	helpKeywords.Add "for each", "html/vsstmForEach.htm"
	helpKeywords.Add "for next", "html/vsstmFor.htm"
	helpKeywords.Add "for", "html/vsstmFor.htm"
	helpKeywords.Add "formatcurrency", "html/vsfctFormatCurrency.htm"
	helpKeywords.Add "formatdatetime", "html/vsfctFormatDateTime.htm"
	helpKeywords.Add "formatnumber", "html/vsfctFormatNumber.htm"
	helpKeywords.Add "formatpercent", "html/vsfctFormatPercent.htm"
	helpKeywords.Add "freespace", "html/jsproFreeSpace.htm"
	helpKeywords.Add "function", "html/vsstmFunction.htm"
	helpKeywords.Add "get property", "html/vsstmPropertyGet.htm"
	helpKeywords.Add "get", "html/vsstmPropertyGet.htm"
	helpKeywords.Add "getabsolutepathname", "html/jsmthGetAbsolutePathname.htm"
	helpKeywords.Add "getbasename", "html/jsmthGetBaseName.htm"
	helpKeywords.Add "getdrive", "html/jsmthGetDrive.htm"
	helpKeywords.Add "getdrivename", "html/jsmthGetDriveName.htm"
	helpKeywords.Add "getextensionname", "html/jsmthGetExtensionName.htm"
	helpKeywords.Add "getfile", "html/jsmthGetFile.htm"
	helpKeywords.Add "getfilename", "html/jsmthGetFileName.htm"
	helpKeywords.Add "getfileversion", "html/jsmthGetFileVersion.htm"
	helpKeywords.Add "getfolder", "html/jsmthGetFolder.htm"
	helpKeywords.Add "getlocale", "html/vsfctGetLocale.htm"
	helpKeywords.Add "getobject", "html/vsfctGetObject.htm"
	helpKeywords.Add "getparentfoldername", "html/jsmthGetParentFolderName.htm"
	helpKeywords.Add "getref", "html/vsfctGetRef.htm"
	helpKeywords.Add "getspecialfolder", "html/jsmthGetSpecialFolder.htm"
	helpKeywords.Add "gettempname", "html/jsmthGetTempName.htm"
	helpKeywords.Add "global", "html/vsproGlobal.htm"
	helpKeywords.Add "helpcontext", "html/vsproHelpContext.htm"
	helpKeywords.Add "helpfile", "html/vsproHelpFile.htm"
	helpKeywords.Add "hex", "html/vsfctHex.htm"
	helpKeywords.Add "hour", "html/vsfctHour.htm"
	helpKeywords.Add "if else", "html/vsstmIf.htm"
	helpKeywords.Add "if then else", "html/vsstmIf.htm"
	helpKeywords.Add "if then", "html/vsstmIf.htm"
	helpKeywords.Add "if", "html/vsstmIf.htm"
	helpKeywords.Add "ignorecase", "html/vsproIgnoreCase.htm"
	helpKeywords.Add "imp", "html/vsoprImp.htm"
	helpKeywords.Add "initialize", "html/vsevtInitialize.htm"
	helpKeywords.Add "inputbox", "html/vsfctInputBox.htm"
	helpKeywords.Add "instr", "html/vsfctInStr.htm"
	helpKeywords.Add "instrrev", "html/vsfctInStrRev.htm"
	helpKeywords.Add "int", "html/vsfctInt.htm"
	helpKeywords.Add "integer division operator", "html/vsoprIntegerDivide.htm"
	helpKeywords.Add "integer division", "html/vsoprIntegerDivide.htm"
	helpKeywords.Add "is", "html/vsoprIs.htm"
	helpKeywords.Add "isarray", "html/vsfctIsArray.htm"
	helpKeywords.Add "isdate", "html/vsfctIsDate.htm"
	helpKeywords.Add "isempty", "html/vsfctIsEmpty.htm"
	helpKeywords.Add "isnull", "html/vsfctIsNull.htm"
	helpKeywords.Add "isnumeric", "html/vsfctIsNumeric.htm"
	helpKeywords.Add "isobject", "html/vsfctIsObject.htm"
	helpKeywords.Add "isready", "html/jsproIsReady.htm"
	helpKeywords.Add "isrootfolder", "html/jsproIsRootFolder.htm"
	helpKeywords.Add "item", "html/jsproitem.htm"
	helpKeywords.Add "items", "html/jsmthItems.htm"
	helpKeywords.Add "join", "html/vsfctJoin.htm"
	helpKeywords.Add "key", "html/jsprokey.htm"
	helpKeywords.Add "keys", "html/jsmthKeys.htm"
	helpKeywords.Add "lbound", "html/vsfctLBound.htm"
	helpKeywords.Add "lcase", "html/vsfctLCase.htm"
	helpKeywords.Add "left", "html/vsfctLeft.htm"
	helpKeywords.Add "len", "html/vsfctLen.htm"
	helpKeywords.Add "length", "html/vsproLength.htm"
	helpKeywords.Add "let property", "html/vsstmPropertyLet.htm"
	helpKeywords.Add "let", "html/vsstmPropertyLet.htm"
	helpKeywords.Add "line", "html/jsproline.htm"
	helpKeywords.Add "loadpicture", "html/vsfctLoadPicture.htm"
	helpKeywords.Add "log", "html/vsfctLog.htm"
	helpKeywords.Add "logical operators", "html/vsidxLogical.htm"
	helpKeywords.Add "loop", "html/vsstmDo.htm"
	helpKeywords.Add "ltrim", "html/vsfctLTrim.htm"
	helpKeywords.Add "match", "html/vsobjMatch.htm"
	helpKeywords.Add "matches", "html/vscolMatches.htm"
	helpKeywords.Add "mid", "html/vsfctMid.htm"
	helpKeywords.Add "minus", "html/vsoprSubtract.htm"
	helpKeywords.Add "minute", "html/vsfctMinute.htm"
	helpKeywords.Add "mod operator", "html/vsoprMod.htm"
	helpKeywords.Add "mod", "html/vsoprMod.htm"
	helpKeywords.Add "modulo operator", "html/vsoprMod.htm"
	helpKeywords.Add "modulo", "html/vsoprMod.htm"
	helpKeywords.Add "month", "html/vsfctMonth.htm"
	helpKeywords.Add "monthname", "html/vsfctMonthName.htm"
	helpKeywords.Add "move", "html/jsmthMove.htm"
	helpKeywords.Add "movefile", "html/jsmthMoveFile.htm"
	helpKeywords.Add "movefolder", "html/jsmthMoveFolder.htm"
	helpKeywords.Add "msgbox", "html/vsfctMsgBox.htm"
	helpKeywords.Add "multiplication operator", "html/vsoprMultiply.htm"
	helpKeywords.Add "multiply", "html/vsoprMultiply.htm"
	helpKeywords.Add "name", "html/jsproName.htm"
	helpKeywords.Add "not", "html/vsoprNot.htm"
	helpKeywords.Add "nothing", "html/vskeyNothing.htm"
	helpKeywords.Add "now", "html/vsfctNow.htm"
	helpKeywords.Add "null", "html/vskeyNull.htm"
	helpKeywords.Add "number", "html/vsproNumber.htm"
	helpKeywords.Add "oct", "html/vsfctOct.htm"
	helpKeywords.Add "on error goto 0", "html/vsstmOnError.htm"
	helpKeywords.Add "on error goto", "html/vsstmOnError.htm"
	helpKeywords.Add "on error resume next", "html/vsstmOnError.htm"
	helpKeywords.Add "on error resume", "html/vsstmOnError.htm"
	helpKeywords.Add "on error", "html/vsstmOnError.htm"
	helpKeywords.Add "openastextstream", "html/jsmthOpenAsTextStream.htm"
	helpKeywords.Add "opentextfile", "html/jsmthOpenTextFile.htm"
	helpKeywords.Add "operator precedence", "html/vsgrpOperatorPrecedence.htm"
	helpKeywords.Add "option explicit", "html/vsstmOptionExplicit.htm"
	helpKeywords.Add "option", "html/vsstmOptionExplicit.htm"
	helpKeywords.Add "or", "html/vsoprOr.htm"
	helpKeywords.Add "parentfolder", "html/jsproParentFolder.htm"
	helpKeywords.Add "path", "html/jsproPath.htm"
	helpKeywords.Add "pattern", "html/vsproPattern.htm"
	helpKeywords.Add "plus", "html/vsoprAdd.htm"
	helpKeywords.Add "precedence", "html/vsgrpOperatorPrecedence.htm"
	helpKeywords.Add "private", "html/vsstmPrivate.htm"
	helpKeywords.Add "property get", "html/vsstmPropertyGet.htm"
	helpKeywords.Add "property let", "html/vsstmPropertyLet.htm"
	helpKeywords.Add "property set", "html/vsstmPropertySet.htm"
	helpKeywords.Add "public", "html/vsstmPublic.htm"
	helpKeywords.Add "raise", "html/vsmthRaise.htm"
	helpKeywords.Add "randomize", "html/vsstmRandomize.htm"
	helpKeywords.Add "read", "html/jsmthRead.htm"
	helpKeywords.Add "readall", "html/jsmthReadAll.htm"
	helpKeywords.Add "readline", "html/jsmthReadLine.htm"
	helpKeywords.Add "redim", "html/vsstmRedim.htm"
	helpKeywords.Add "regexp execute", "html/vsmthExecute.htm"
	helpKeywords.Add "regexp replace", "html/vsmthReplace.htm"
	helpKeywords.Add "regexp", "html/vsobjRegExp.htm"
	helpKeywords.Add "rem", "html/vsstmRem.htm"
	helpKeywords.Add "remove", "html/jsmthRemove.htm"
	helpKeywords.Add "removeall", "html/jsmthRemoveAll.htm"
	helpKeywords.Add "replace regexp", "html/vsmthReplace.htm"
	helpKeywords.Add "replace", "html/vsfctReplace.htm"
	helpKeywords.Add "rgb", "html/vsfctRGB.htm"
	helpKeywords.Add "right", "html/vsfctRight.htm"
	helpKeywords.Add "rnd", "html/vsfctRnd.htm"
	helpKeywords.Add "rootfolder", "html/jsproRootFolder.htm"
	helpKeywords.Add "round", "html/vsfctRound.htm"
	helpKeywords.Add "rtrim", "html/vsfctLTrim.htm"
	helpKeywords.Add "runtime error", "html/vsmscSyntaxErrors.htm"
	helpKeywords.Add "scriptengine", "html/vsfctScriptEngine.htm"
	helpKeywords.Add "scriptenginebuildversion", "html/vsfctScriptEngineBuildVersion.htm"
	helpKeywords.Add "scriptenginemajorversion", "html/vsfctScriptEngineMajorVersion.htm"
	helpKeywords.Add "scriptengineminorversion", "html/vsfctScriptEngineMinorVersion.htm"
	helpKeywords.Add "second", "html/vsfctSecond.htm"
	helpKeywords.Add "select case", "html/vsstmSelectCase.htm"
	helpKeywords.Add "select", "html/vsstmSelectCase.htm"
	helpKeywords.Add "serialnumber", "html/jsproSerialNumberProp.htm"
	helpKeywords.Add "set property", "html/vsstmPropertySet.htm"
	helpKeywords.Add "set", "html/vsstmSet.htm"
	helpKeywords.Add "setlocale", "html/vsfctSetLocale.htm"
	helpKeywords.Add "sgn", "html/vsfctSgn.htm"
	helpKeywords.Add "sharename", "html/jsproShareName.htm"
	helpKeywords.Add "shortname", "html/jsproShortName.htm"
	helpKeywords.Add "shortpath", "html/jsproShortPath.htm"
	helpKeywords.Add "sin", "html/vsfctSin.htm"
	helpKeywords.Add "size", "html/jsproSize.htm"
	helpKeywords.Add "skip", "html/jsmthSkip.htm"
	helpKeywords.Add "skipline", "html/jsmthSkipLine.htm"
	helpKeywords.Add "source", "html/vsproSource.htm"
	helpKeywords.Add "space", "html/vsfctSpace.htm"
	helpKeywords.Add "split", "html/vsfctSplit.htm"
	helpKeywords.Add "sqr", "html/vsfctSqr.htm"
	helpKeywords.Add "strcomp", "html/vsfctStrComp.htm"
	helpKeywords.Add "string concatenation", "html/vsoprConcatenation.htm"
	helpKeywords.Add "string", "html/vsfctString.htm"
	helpKeywords.Add "strreverse", "html/vsfctStrReverse.htm"
	helpKeywords.Add "sub", "html/vsstmSub.htm"
	helpKeywords.Add "subfolders", "html/jsproSubFolders.htm"
	helpKeywords.Add "submatches", "html/vscolSubMatches.htm"
	helpKeywords.Add "subtraction operator", "html/vsoprSubtract.htm"
	helpKeywords.Add "subtraction", "html/vsoprSubtract.htm"
	helpKeywords.Add "syntax error", "html/vsmscRuntimeErrors.htm"
	helpKeywords.Add "tan", "html/vsfctTan.htm"
	helpKeywords.Add "terminate", "html/vsevtTerminate.htm"
	helpKeywords.Add "test", "html/vsmthTest.htm"
	helpKeywords.Add "textstream", "html/jsobjTextStream.htm"
	helpKeywords.Add "then", "html/vsstmIf.htm"
	helpKeywords.Add "time", "html/vsfctTime.htm"
	helpKeywords.Add "timer", "html/vsfctTimer.htm"
	helpKeywords.Add "timeserial", "html/vsfctTimeSerial.htm"
	helpKeywords.Add "timevalue", "html/vsfctTimeValue.htm"
	helpKeywords.Add "totalsize", "html/jsproTotalSize.htm"
	helpKeywords.Add "trim", "html/vsfctLTrim.htm"
	helpKeywords.Add "true", "html/vskeyTrue.htm"
	helpKeywords.Add "type", "html/jsproType.htm"
	helpKeywords.Add "typename", "html/vsfctTypeName.htm"
	helpKeywords.Add "ubound", "html/vsfctUBound.htm"
	helpKeywords.Add "ucase", "html/vsfctUCase.htm"
	helpKeywords.Add "value", "html/vsproValue.htm"
	helpKeywords.Add "vartype", "html/vsfctVarType.htm"
	helpKeywords.Add "volumename", "html/jsproVolumeName.htm"
	helpKeywords.Add "weekday", "html/vsfctWeekday.htm"
	helpKeywords.Add "weekdayname", "html/vsfctWeekdayName.htm"
	helpKeywords.Add "wend", "html/vsstmWhile.htm"
	helpKeywords.Add "while wend", "html/vsstmWhile.htm"
	helpKeywords.Add "while", "html/vsstmWhile.htm"
	helpKeywords.Add "with", "html/vsstmWith.htm"
	helpKeywords.Add "write", "html/jsmthWrite.htm"
	helpKeywords.Add "writeblanklines", "html/jsmthWriteBlankLines.htm"
	helpKeywords.Add "writeline", "html/jsmthWriteLine.htm"
	helpKeywords.Add "xor", "html/vsoprXor.htm"
	helpKeywords.Add "year", "html/vsfctYear.htm"

Main


Sub Main()
	Dim line

	ImportInitScript
	Usage

	Do While True
		WScript.StdOut.Write(">>> ")

		line = Trim(WScript.StdIn.ReadLine)
		Do While Right(line, 2) = " _" Or line = "_"
			line = RTrim(Left(line, Len(line)-1)) & " " & Trim(WScript.StdIn.ReadLine)
		Loop

		If LCase(line) = "exit" Then Exit Do

		If line = "?" Or line = "help" Then
			Usage
		ElseIf Left(line, 2) = "? " Then
			Help Trim(Replace(Mid(line, 3), """", ""))
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
Private Sub ImportInitScript()
	Dim sh, fso, path, initScriptExists

	Set sh  = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")

	path = sh.ExpandEnvironmentStrings(InitScript)
	initScriptExists = fso.FileExists(path)

	Set sh  = Nothing
	Set fso = Nothing

	If initScriptExists Then Import path
End Sub

'! Look up the given keyword in the VBScript documentation (if the help file is
'! located in the current working directory, the Windows help directory, or
'! somewhere in the %PATH%).
'!
'! @param  keyword  Keyword to look up.
Private Sub Help(keyword)
	Dim sh, fso, chm, dir

	Set sh  = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")

	If Not helpKeywords.Exists(keyword) Then
		WScript.StdOut.WriteLine "'" & keyword & "' not found in the help index."
		Exit Sub
	End If

	If isEmpty(helpfile) Then
		' check current directory
		chm = fso.BuildPath(sh.CurrentDirectory, Documentation)
		If Not fso.FileExists(chm) Then
			' check help directory
			chm = fso.BuildPath(sh.ExpandEnvironmentStrings("%SystemRoot%\Help"), Documentation)
		End If
		If Not fso.FileExists(chm) Then
			' check %PATH%
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				chm = fso.BuildPath(fso.GetAbsolutePathName(dir), Documentation)
				If fso.FileExists(chm) Then Exit For
			Next
		End If
		If fso.FileExists(chm) Then helpfile = chm
	End If

	If Not fso.FileExists(helpfile) Then
		WScript.StdErr.WriteLine "'" & Documentation & "' not found. Make sure the file is present in the current working working" & vbNewLine _
			& "directory or the Windows help directory, or add its location to the PATH environment" & vbNewLine _
			& "variable."
		Exit Sub
	End If

	sh.Run "hh.exe ""mk:@MSITStore:" & helpfile & "::/" & helpKeywords(keyword) & """", 1, False
End Sub

'! Print usage information.
Private Sub Usage()
	WScript.StdOut.Write "A simple interactive VBScript Shell." & vbNewLine & vbNewLine _
		& vbTab & "help | ?                  Print this help." & vbNewLine _
		& vbTab & "help | ? ""keyword""        If the VBScript documentation is installed, look" & vbNewLine _
		& vbTab & "                          up ""keyword"" in the helpfile (" & Documentation & ")." & vbNewLine _
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
