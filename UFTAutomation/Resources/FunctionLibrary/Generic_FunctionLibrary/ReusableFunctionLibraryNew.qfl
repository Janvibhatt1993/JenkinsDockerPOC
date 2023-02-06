' ================================================================================================
'  NAME 	: testCleanUp
'  DESCRIPTION 	: This function is used to write the summary part of the report & close all teh browsers.
'  PARAMETERS	:  
' ================================================================================================

Public Function funTestCleanUp()
	'systemutil.CloseProcessByName("FIREFOX.EXE")
	systemutil.CloseProcessByName("IEXPLORE.EXE")
	systemutil.CloseProcessByName("CHROME.EXE")
	'systemutil.CloseProcessByName("EXCEL.EXE")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
			objProcess.Terminate()
			'MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
	TestCaseExecutiveSummary ()
End Function


Public Function objectExistanceCheck(objPage, objObject , strPageName , strObjectName)
	objPage.sync
	objObject.highlight
	If objObject.Exist(2) Then
		LogResult micpass , UCASE(strObjectName) & " should be Found in the page " & UCASE(strPageName) , UCASE(strObjectName) & " is Found"
	Else
		LogResult micFail , UCASE(strObjectName) & " should be Found in the page " & UCASE(strPageName) , UCASE(strObjectName) & " is NOT Found"
	End If
End Function

'==============================================================================================
' Function/Sub: CountFiles(sFoldername)
' Purpose: How many files  in the folder.
'==============================================================================================
Public Function CountFiles(sFoldername)
	Dim oFS, oFolder, oFiles

	Set oFS = CreateObject("Scripting.FileSystemObject")
	CountFiles = oFS.GetFolder(sFoldername).Files.Count
	Set oFS = Nothing
End Function

'==============================================================================================
' Function/Sub: CountFiles(sFoldername)
' Purpose: How many lines in text file.
'==============================================================================================
Public Function CountFileLines(sFilename)
	Dim oFS, oFile
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFS.OpenTextFile(sFilename)	
	
	CountFileLines = Ubound(Split(oFile.ReadAll, vbNewLine))' - 1 'REM put out of here
	oFile.Close
	Set oFile = Nothing
	Set oFS = Nothing
End Function

'==============================================================================================
' Function/Sub: MoveFiles(sSourceFolder, sDestinationFolder)
' Purpose: 
'==============================================================================================
Public Sub MoveFiles(sSourceFolder, sDestinationFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	'Create the folder
	On Error resume next
	oFS.CreateFolder(sDestinationFolder)
	oFS.MoveFile sSourceFolder & "\*.*", sDestinationFolder
	On Error goto 0

	Set oFS = Nothing
End Sub

'==============================================================================================
' Function/Sub: CopyFiles(sSourceFolder, sDestinationFolder)
' Purpose: 
'==============================================================================================
Public Sub CopyFiles(sSourceFolder, sDestinationFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	'Create the folder
	On Error resume next
	oFS.CreateFolder(sDestinationFolder)
	oFS.CopyFile sSourceFolder & "\*.*", sDestinationFolder
	On Error goto 0

	Set oFS = Nothing
End Sub

'==============================================================================================
' Function/Sub: DeleteFiles(sSourceFolder)
' Purpose: 
'==============================================================================================
Public Sub DeleteFiles(sSourceFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	On Error resume next
	oFS.DeleteFile sSourceFolder & "\*.*"
	On Error goto 0

	Set oFS = Nothing
End Sub

'==============================================================================================
' Function/Sub: LatestFileDate(sFoldername, sFileType)
' Purpose: The latest file matching the filetype in the folder.
'==============================================================================================
Public Function LatestFileDate(sFoldername, sFileType)
	Dim oFS, oFolder, oFC, oFile
	Dim sLatestDate

	'on Error Resume Next

	Set oFS   = CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFS.GetFolder(sFoldername)
	Set oFC = oFolder.Files

	sLatestDate = "01/01/1900"
	For Each oFile in oFC
		If(instr(len(oFile)-len(sFileType)+ 1, ucase(oFile), ucase(sFileType),1)= len(oFile)-len(sFileType)+ 1)Then
			'Found matching file
			If CDate(oFile.datelastmodified) > CDate(sLatestDate) Then
				sLatestDate = oFile.datelastmodified
			End if
		End if
	Next

	Set oFile = Nothing
	Set oFC = Nothing
	Set oFolder = Nothing
	Set oFS = Nothing

	LatestFileDate = sLatestDate
End Function

Public Function Encrypt(sString)
	Dim sKey
	Dim iLenKey, iKeyPos, iLenStr, i, sNewStr

	sKey = "encryption"
	sNewStr = ""
	iLenKey = Len(sKey)
	iKeyPos = 1
	iLenStr = Len(sString)

	sString = StrReverse(sString)
	For i = 1 To iLenStr
		sNewStr = sNewStr & chr(asc(Mid(sString, i, 1)) + Asc(Mid(sKey, iKeyPos, 1)))
		iKeyPos = iKeyPos + 1
		If iKeyPos > iLenKey Then iKeyPos = 1
	Next
	encrypt = sNewStr
End Function

'==============================================================================================
' Function/Sub: Decrypt(sString)
' Purpose: 
'==============================================================================================
Public Function Decrypt(sString)
	Dim sKey
	Dim iLenKey, iKeyPos, iLenStr, i, sNewStr

	sKey = "encryption"
	sNewStr = ""
	iLenKey = Len(sKey)
	iKeyPos = 1
	iLenStr = Len(sString)

	sString=StrReverse(sString)
	For i = iLenStr To 1 Step -1
		sNewStr = sNewStr & chr(asc(Mid(sString, i, 1)) - Asc(Mid(sKey, iKeyPos, 1)))
		iKeyPos = iKeyPos + 1
		If iKeyPos > iLenKey Then iKeyPos = 1
	Next
	sNewStr=StrReverse(sNewStr)
	Decrypt = sNewStr
End Function

'==============================================================================================
' Function/Sub:  Returns the Full Path of the file if it is found
'==============================================================================================
Function FindFile(sFilename, sStartFolder)
	Dim sRetval
	Dim oFS, oFile, oFolder, oSubFolder

	sRetval = ""
	Set oFS = CreateObject("Scripting.FileSystemObject")

	'*PREREQ
	If Not oFS.FolderExists(sStartFolder) Then
		Exit Function
	End If
	'*
	
	Set oFolder = oFS.GetFolder(sStartFolder)
	For Each oFile in oFolder.Files
		If UCase(oFile.Name) = UCase(sFilename) Then
			sRetval = oFile.Path	' & "\" & sFilename
			Exit For
		End If
	Next

	If sRetval = "" Then
		For Each oSubFolder in oFolder.SubFolders 
			sRetval = FindFile(sFilename, oSubFolder.Path) ' & "\" & oSubFolder.Name)
			if sRetval <> "" then Exit For
		Next
	End If
			
	Set oFolder = Nothing
	Set oFS = Nothing

	FindFile = sRetval
End Function

'==============================================================================================
' Function/Sub: LoadObjectRepository()
' Purpose: 
'==============================================================================================
'REM QTP Specific function
Public Function LoadObjectRepository(sPath) 
	Dim oQTP, oQTPRepositories, sActionName
	Dim oFS, bRetval

	bRetVal = False
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	If  oFS.FileExists(sPath)= True Then				'Only run if file exists.
		bRetVal = True
		Set oQTP = CreateObject("QuickTest.Application")
		sActionName = Environment("ActionName") 'getting the action name
		Set oQTPRepositories = oQTP.Test.Actions(sActionName).ObjectRepositories

		If oQTPRepositories.Find(sPath) = -1 Then ' If the repository cannot be found in the collection
			oQTPRepositories.Add sPath ' Add the repository to the collection
		End If
		
		Set oQTPRepositories = Nothing
		Set oQTP = Nothing
	End If

	Set oFS = Nothing
	LoadObjectRepository = bRetVal
	
End Function

'==============================================================================================
' Function/Sub: RemoveBlankLines(sInfile, byref dicNewToOld))
' Purpose: 	
'==============================================================================================
Function RemoveBlankLines(sInfile, byref arrNewToOld)
	Dim iLineNum, sLine
	Dim iInLIne, iOutLIne

	'Read the files
	Dim sOutfile
	Dim oFS, oInFile, oOutfile
	sOutfile = sInFile & ".NoBlankS"
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oInFile = oFs.OpenTextFile(sInfile, 1, False)
    Set oOutfile = oFs.OpenTextFile(sOutfile, 2, True)

	iInLIne = 1 : iOutLIne = 1
	Do While Not oInFile.AtEndOfStream
		sLIne = oInFile.ReadLine
		If Trim(sline) <> "" Then
			oOutfile.WriteLine sLine
			arrNewToOld(iOutLIne) = iInLIne
			iOutLIne = iOutLIne + 1
		End If
		iInLIne = iInLIne+1
	Loop
	
	oOutfile.Close
	oInFile.Close

    Set oOutfile = Nothing
    Set oInFile = Nothing
    Set oFs = Nothing
	
	ReDim Preserve arrNewToOld(iOutLIne)

	RemoveBlankLines = sOutfile

End Function

'===================================================================================
' Function: DeleteZipFile(sZipFile)
'===================================================================================
Sub DeleteZipFile(sZipFile, sLog)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")
	'Does file to add exist?
	If oFS.FileExists(sZipFile) Then 
		oFS.DeleteFile sZipFile, True
		sLog = sLog & "D: File " & sZipFile & " deleted." &vbLf
	Else
		sLog = sLog & "W: File to delete doesn't exist: " & sZipFile & vbLf
	End If

	Set oFS = Nothing
End sub

'===================================================================================
' Function: AddFoldertoZip(sFilename, sZipFile)
'===================================================================================
Function AddFoldertoZip(sFilename, sZipFile, sLog)
	Dim iCount, iTimeout, iTemp
	Dim oApp, oFS, oFile, oZip
	Const ForWriting = 2
	
	iTimeout = 30	'Seconds
	Set oFS = CreateObject("Scripting.FileSystemObject")
	'Does file to add exist?
	If oFS.FolderExists(sFilename) Then
		'Ceate zip file?
		If Not oFS.FileExists(sZipFile) Then
			Set oZip = oFS.OpenTextFile(sZipFIle, ForWriting, True)
			oZip.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
			oZip.Close
			Set oZip = Nothing
			sLog = sLog & "D: Created zip file " & sZipFile & vblf
		End If
	Else
		sLog = sLog & "D: File " & sFilename & " not found." & vblf
		AddFoldertoZip = False : Exit Function
	End If

	'Add the file to the zip file
	'Set oFile = oFS.GetFile(sFilename)
    ' Create a Shell object
    Set oApp = CreateObject("Shell.Application")
    ' Copy the files to the compressed folder
    iCount = oApp.NameSpace(sZipFile).Items.Count
	
	oApp.NameSpace(sZipFile).CopyHere sFilename
	
	'WaitTime 0,200
	WaitTime(200)
	'WaitTime until the file is ready, otherwise seeing "oApp.NameSpace(sZipFile)" error
	On Error Resume Next
	iTemp = oApp.NameSpace(sZipFile).Items.Count
	Do While Err.Number <> 0
		Err.Clear
		'WaitTime 1
		WaitTime(100)
		iTemp = oApp.NameSpace(sZipFile).Items.Count
	Loop
	Err.Clear
	On Error Goto 0
	
	' Keep script waiting until compression is done
    Do Until oApp.NameSpace(sZipFile).Items.Count = iCount + 1 or iTimeout <= 0
        'WScript.Sleep 1000
		'WaitTime 1
		WaitTime(100)
		iTimeout = iTimeout - 1
    Loop

	sLog = sLog & "D: Added " & sFileName & " to zip file " & sZipFile & vbLf
	
	'Set oFile = Nothing
	Set oApp = Nothing
	Set oFS = Nothing

	AddFoldertoZip = True
End Function

'===================================================================================
' Function: AddFiletoZip(sFilename, sZipFile)
'===================================================================================
Function AddFiletoZip(sFilename, sZipFile, sLog)
	Dim iCount, iTimeout, iTemp
	Dim oApp, oFS, oFile, oZip
	Const ForWriting = 2
	
	iTimeout = 30	'Seconds
	Set oFS = CreateObject("Scripting.FileSystemObject")
	'Does file to add exist?
	If oFS.FileExists(sFilename) Then
		'Ceate zip file?
		If Not oFS.FileExists(sZipFile) Then
			Set oZip = oFS.OpenTextFile(sZipFIle, ForWriting, True)
			oZip.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
			oZip.Close
			Set oZip = Nothing
			sLog = sLog & "D: Created zip file " & sZipFile & vblf
		End If
	Else
		sLog = sLog & "D: File " & sFilename & " not found." & vblf
		AddFiletoZip = False : Exit Function
	End If

	'Add the file to the zip file
	'Set oFile = oFS.GetFile(sFilename)
    ' Create a Shell object
    Set oApp = CreateObject("Shell.Application")
    ' Copy the files to the compressed folder
    iCount = oApp.NameSpace(sZipFile).Items.Count
	
	oApp.NameSpace(sZipFile).CopyHere sFilename
	
	'WaitTime 0,200
	WaitTime(200)

	'WaitTime until the file is ready, otherwise seeing "oApp.NameSpace(sZipFile)" error
	On Error Resume Next
	iTemp = oApp.NameSpace(sZipFile).Items.Count
	Do While Err.Number <> 0
		Err.Clear
		'WaitTime 1
		WaitTime(100)
		iTemp = oApp.NameSpace(sZipFile).Items.Count
	Loop
	Err.Clear
	On Error Goto 0
	
	' Keep script waiting until compression is done
    Do Until oApp.NameSpace(sZipFile).Items.Count = iCount + 1 or iTimeout <= 0
        'WScript.Sleep 1000
		'WaitTime 1
		WaitTime(100)
		iTimeout = iTimeout - 1
    Loop

	sLog = sLog & "D: Added " & sFileName & " to zip file " & sZipFile & vbLf
	
	'Set oFile = Nothing
	Set oApp = Nothing
	Set oFS = Nothing

	AddFiletoZip = True
End Function

Public Function ZIPFolder(sFilepath, sZipExecPath, sZipOutput, sLog)
	Dim oWshScriptExec, oFS, oShell, oStdOut
	Dim sLastLine
	Set oShell = CreateObject("WScript.Shell")
	Set oFS = CreateObject("Scripting.FileSystemObject")
	
	if oFS.FolderExists(sFilepath) Then
		Set oWshScriptExec = oShell.Exec(sZipExecPath & " a -tzip -r " & sZipOutput & " " & sFilepath)
		Set oStdOut = oWshScriptExec.StdOut
		While Not oStdOut.AtEndOfStream
		   sLastLine = oStdOut.ReadLine
		   sLog = sLog & vbCrLf & sLastLine
		Wend
		If sLastLine = "Everything is Ok" Then
			sLog = sLog & "E: Folder successfuly packed"
			ZIPFolder = True
		Else
			sLog = sLog & "E: Zip doesn't exist"
			ZIPFolder = False
		End If
	Else
		sLog = sLog & "E: Folder doesn't exist"
		ZIPFolder = False
	end if	
End Function

'===================================================================================
' Function: EscapeString4Regex(sString)
'===================================================================================
Function EscapeString4Regex(sString)
	Dim sRetVal, i
	Dim sRegexChars, sRegexChar
		
	sRegexChars ="\^$*+?.()|{}[]"
	sRetVal = sString

	For i = 1 To Len(sRegexChars)
		sRegexChar = Mid(sRegexChars, i, 1)
		sRetVal = Replace(sRetVal, sRegexChar, "\" & sRegexChar)
	Next
	
	EscapeString4Regex = sRetVal

End Function

'===================================================================================
'Function : CheckFolderExists
'Description : Checks whether the folder exists in the path and returns True or False
'===================================================================================
Public Function CheckFolderExists(sPath)
	Dim oFSO, sExists
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
		
	'Check the path exists
	If (oFSO.FolderExists(sPath)) Then
		sExists = True
	Else
		sExists = False
	End If

	Set oFSO = Nothing
	CheckFolderExists = sExists
End Function
'===================================================================================
'Function : CheckFileExists
'Description : Checks whether the folder exists and if true then checks whether file exists in the path and 
'                     returns True or False
'===================================================================================
Public Function CheckFileExists(sPath, sFileName)
	Dim oFSO, sExists
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	'Check the path exists
	If (oFSO.FolderExists(sPath)) Then
		'Check the file exists
		If (oFSO.FileExists(sPath & sFileName)) Then
			sExists = True
		Else
			sExists = False
		End If
	Else
		sExists = False
	End If

    Set oFSO = Nothing
	CheckFileExists = sExists
End Function

'===================================================================================
'Function : CreateFolder
'Description : Checks whether the folder exists and if not it creates it.
'				does this recursively for its parent(s)
'===================================================================================
Public Sub CreateFolder(sFolder)
	Dim oFS

	Set oFS = CreateObject("Scripting.FileSystemObject")

	If not oFS.FolderExists(sFolder) Then
		If not oFS.FolderExists(oFS.GetParentFolderName(sFolder)) Then CreateFolder(oFS.GetParentFolderName(sFolder))
		oFS.CreateFolder(sFolder)
	End If

	Set oFS = Nothing
End Sub

'===================================================================================
'Function : FolderExists
'Description : True if the folder exists.
'===================================================================================
Public Function FolderExists(sFolder)
	Dim oFS, bRetVal
	
	bRetVal = False

	Set oFS = CreateObject("Scripting.FileSystemObject")
	If oFS.FolderExists(sFolder) Then bRetVal = True
	Set oFS = Nothing
	
	FolderExists = bRetVal

End Function

'==============================================================================================
' Function/Sub: IsRegexMatching(sPattern, sString, bIgnoreCase)
' Purpose: IsRegexMatching returns true, if given string matches given pattern
'==============================================================================================
Function IsRegexMatching(sPattern, sString, bIgnoreCase)
	Dim oRegEx
	
    Set oRegEx = New RegExp	
    
	oRegEx.Pattern = sPattern
	oRegEx.IgnoreCase = bIgnoreCase
	
	IsRegexMatching = oRegEx.Test(sString)
	
	Set oRegEx = Nothing
End Function

Function GetRegexMatches(ByRef arrResult, sPattern, sString, bIgnoreCase)
	Dim oRegEx, oMatches, oMatch
	Dim arrRes()
	Dim i, iCount
	
    Set oRegEx = New RegExp	
	
	oRegEx.Pattern = sPattern
	oRegEx.IgnoreCase = bIgnoreCase
	oRegEx.Global = True
	Set oMatches = oRegEx.Execute(sString)
	iCount = oMatches.Count
	If iCount > 0 Then
		ReDim arrRes(iCount - 1)
		i = 0
		For each oMatch  in oMatches
			arrRes(i) = oMatch.Value
			i = i + 1
		Next
		arrResult = arrRes
	End If
	GetRegexMatches = iCount
	
	Set oRegEx = Nothing
End Function

Private Function LocalFileExists(sFilePath)
	Dim oFS
	Set oFS = CreateObject("Scripting.FileSystemObject") 
	LocalFileExists = oFS.FileExists(sFilePath)
	Set oFS = Nothing
End Function

Private Function DeleteFile(sFilePath)
	Dim oFS
	Set oFS = CreateObject("Scripting.FileSystemObject") 
	If oFS.FileExists(sFilePath) Then
		oFS.DeleteFile sFilePath, True
	End If
	Set oFS = Nothing
End Function

Private Function FileRename(sFilePath, sNewName)
	Dim oFS, oFile
	Set oFS = CreateObject("Scripting.FileSystemObject") 
	If oFS.FileExists(sFilePath) Then
		Set oFile = oFS.GetFile(sFilePath)
		oFile.Name = sNewName
		FileRename = oFile.Path
		Set oFile = Nothing
	End If
	Set oFS = Nothing
End Function

'==============================================================================================
' Function/Sub: TimeDif(iEndTime, iStartTime)
' Purpose: Avoid giving negative values when midnight passes
'
' Parameters: end time, start time
'
' Returns: elapsed time
'==============================================================================================
Function TimeDif(iEndTime, iStartTime)
	If iStartTime < iEndTime Then
		TimeDif = iEndTime - iStartTime
	Else
		TimeDif = iEndTime - iStartTime + 86400
	End If
End Function

Function RegexReplace(sPattern, sSearchString, sReplaceString, bIgnoreCase)
	Dim oRegEx
	
    Set oRegEx = New RegExp	
    
	oRegEx.Pattern = sPattern
	oRegEx.IgnoreCase = bIgnoreCase
	oRegEx.Global = True
	
	RegexReplace = oRegEx.Replace(sSearchString, sReplaceString)
	
	Set oRegEx = Nothing
End Function

'*************************************************
'Function to report Errors
'*************************************************
	'Purpose: Report any errors encountered
	'I/P: None
	'O/P: ERR.NUMBER -- ERR.DESCRIPTION OR {keyword} Passed
Function fnErrorReport()

	If err.number <> 0 Then
		fnErrorReport = err.number & " -- " & Err.Description 
		Print "Error Occured - " & err.number & " -- " & err.description & " -- " & "TC Number: " & Datatable.Value ("TCID", "DT_TestSteps") & " - " & "step number: "& Datatable.Value ("StepNo.", "DT_TestSteps")
		Environment("Error") = fnErrorReport
		Else
		fnErrorReport = Keyword & " Passed"
		Print "No Error Occured " & "Step Number: " & DataTable.Value ("StepNo.", "DT_TestSteps")
	End If
	
err.clear
End Function
