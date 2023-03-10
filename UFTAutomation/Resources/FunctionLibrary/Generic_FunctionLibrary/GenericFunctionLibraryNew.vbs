'================================
'Function name - funToTakeDataFromExcel
'Description - Take excel value in Datatable
'================================
Public testCaseNumber,testCaseDescription,executionFlag,testEnvironment,testFunctionName,browserName
Dim iTotalPassed,iTotalFailed,iTotalOthers
Dim sTimeStamp
'Public strResultPath,strLibPath,strTestDataPath,htmlResultPath,strObjectRepositoryPath,strObjRepPath1
Set dictTestData = createObject("Scripting.Dictionary")


public Function funToTakeDataFromExcel(strResultPath,strLibPath,strTestDataPath,htmlResultPath,strObjectRepositoryPath)
	On error resume next
	'funToInitializeEnvironmentVariable()
	funToCreateLogFile
	funToRegitseruserfunction
	'Dim testCaseNumber,testCaseDescription,executionFlag,testEnvironment,testFunctionName
	Dim sheetExistFlag : sheetExistFlag = False
	Dim rowCount,colCount,testDirName,mainRow,Iterator,subFunctionRowCount,SubFunctionColCount,funRowData,subTestCaseNumber,funColData,funColName
	
	Datatable.ImportSheet strTestDataPath&"TestDataHybrid.xls","TestData",dtGlobalSheet
	rowCount=Datatable.GetSheet("Global").GetRowCount
	colCount=Datatable.GetSheet("Global").GetParametercount  
	
	If  Err.number =  0 Then
		For mainRow = 1 To rowCount Step 1
			Datatable.SetCurrentRow(mainRow)
			testCaseNumber = Datatable.Value("TestCaseNo","Global")
			testCaseDescription=Datatable.Value("TestCaseDescription","Global")
			executionFlag = Datatable.Value("ExecutionFlag","Global")
			testEnvironment =Datatable.Value("TestEnvironment","Global")
			testFunctionName = Datatable.Value("TestFunction","Global")
			browserName = Datatable.Value("Browser","Global")
			If Ucase(executionFlag)="YES" Then
				For Iterator = 1 To Datatable.GetSheetCount Step 1
					If Ucase(Datatable.GetSheet(Iterator).Name) = Ucase(testFunctionName) Then
						sheetExistFlag = True
						Exit For
					End If
				Next
				If sheetExistFlag<>True Then
					Datatable.AddSheet testFunctionName
					Datatable.ImportSheet strTestDataPath&"\TestDataHybrid.xls",testFunctionName,testFunctionName
				End If
				sheetExistFlag = False
				Reporter.ReportEvent micPass, "Test data sheet "&testFunctionName &" added ","TestDatasheet added successfully!"
				subFunctionRowCount=Datatable.GetSheet(testFunctionName).GetRowCount
				SubFunctionColCount=Datatable.GetSheet(testFunctionName).GetParameterCount
	
				If Instr(1,testFunctionName,"RegisterMultipleUser",1) Then
					Call funToAddMultipleDataInMultipleDict(subFunctionRowCount,SubFunctionColCount)
				else
					For funRowData = 1 To subFunctionRowCount Step 1
						Datatable.GetSheet(testFunctionName).SetCurrentRow(funRowData)
						subTestCaseNumber = Datatable.Value("TestCaseNo",testFunctionName)
						If Trim(subTestCaseNumber) = Trim(testCaseNumber) Then
							For funColData = 1 To SubFunctionColCount Step 1
								funColName=Datatable.GetSheet(testFunctionName).GetParameter(funColData).Name
								If not(dictTestData.Exists(funColName)) Then
									dictTestData.Add funColName,Datatable.Value(funColName,testFunctionName)
								End If
							Next
							Exit For
							Exit For	
						End If
					Next
				End  IF
				funToWriteLogsInFile "Calling Main Function"
				strResult=MainFunction()
				dictTestData.RemoveAll
				'msgbox strResult
				'Set dictTestData = Nothing
				call funToCountTestCaseStatus(strResult)
			End If
		Next
	else
		Reporter.ReportEvent micFail, "Test Data sheet path or details are not valid ","Kindly check and update Test data sheet as per requirement"
	End If
	If Err.number <> 0 Then
		Reporter.ReportEvent micFail, "Test case failed","Test case failed"
	else
		Reporter.ReportEvent micPass, "Test case passed","Test case passed"
	End If
	On error goto 0
End Function

'================================
'Function name - funToCountTestCaseStatus
'Description - Based on Test case execution it will count no of Pass and failed Test cases
'================================
Function funToCountTestCaseStatus(StrResult)
	Select Case StrResult
		Case "Passed"
			iTotalPassed = iTotalPassed + 1
			'Msgbox iTotalPassed
		Case "Failed"
			iTotalFailed = iTotalFailed + 1
			'Msgbox iTotalFailed
		Case Else
			iTotalOthers = iTotalOthers + 1
			'Msgbox iTotalOthers
	End Select
End Function

Function fnTimeStamp()

		sDay = Day(Now)
		sMonth = Month(Now)
		sYear = Year(Now)
		sHour = Hour(Now)
		sMin = Minute(Now)
		sSec = Second(Now)
		sTimeStamp = sDay & sMonth & sYear & "_" & sHour & "_" & sMin & "_" & sSec
		fnTimeStamp=sTimeStamp
End Function

' ================================================================================================
'  NAME			: Initialization
'  DESCRIPTION 	  	: This function is used to create global variables which stores location 
'			   path of TestResult, TestData, Scripts, AppLib, Browser CommonLib & ObjectRepo 
'			  Loads common repository
'  PARAMETERS		: nil
' ================================================================================================

Public Function funToInitializeEnvironmentVariableold()
	Dim sTestDir,arrPath,I
	sTestDir= Environment.Value ("TestDir")
	arrPath = Split (sTestDir, "\")
	Dim libOne,libTwo
	'Save Result Path to variable strResultPath
	arrPath(UBound(arrPath)) = "Results"
	For I=0 to UBound(arrPath)
		If (I=0) Then
			strResultPath = arrPath(I)
		Else
			strResultPath = strResultPath + "\" + arrPath(I)
		End If
	Next
	strResultPath = strResultPath & "\"
	'Save resources Path to variable sAppLibPath
	arrPath(UBound(arrPath)) = "Resources"
	For I=0 to UBound(arrPath)
		If (I=0) Then
			strLibPath = arrPath(I)
		Else
			strLibPath = strLibPath + "\" + arrPath(I)
		End If
	Next
	strLibPath = strLibPath & "\"		
	'Save TestData Path to variable strTestDataPath
	arrPath(UBound(arrPath)) = "DataSheet"
	For I=0 to UBound(arrPath)
		If (I=0) Then
			strTestDataPath = arrPath(I)
		Else
			strTestDataPath = strTestDataPath + "\" + arrPath(I)
		End If
	Next
	strTestDataPath = strTestDataPath & "\"	
	'------creating environment variable------------
	 Environment.value("strHtmlResultPath")=strResultPath&"DetailedHTMLResult\"
 	htmlResultPath = Environment.value("strHtmlResultPath")
 	 Environment.value("strObjectRepoPath") = strLibPath&"ObjectRepository\"
 	 strObjectRepositoryPath=Environment.value("strObjectRepoPath") & "CommonObjectRepository.tsr"
 	 'Loading the repository file
	If  strObjectRepositoryPath <>  "" Then
		RepositoriesCollection.RemoveAll() 
		RepositoriesCollection.Add(strObjectRepositoryPath)  
	End If
	 Environment.Value("strFunctionLibraryPath") = strLibPath&"FunctionLibrary"
 	 Environment.Value("strAppFunctionLibraryPath")=strLibPath&"FunctionLibrary\AppSpecific_FunctionLibrary\"
 	 Environment.Value("strGenFunctionLibraryPath")=strLibPath&"FunctionLibrary\Generic_FunctionLibrary\"
 	 'LoadFunctionLibrary Environment.Value("strGenFunctionLibraryPath")&"GenericFunctionLibraryNew.vbs",Environment.Value("strAppFunctionLibraryPath")&"ApplicationSpecificFunctionLibraryNew.vbs"
	 funToAddFunctionLibraryDynamically
' 	 Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
'		Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
'		For Each objProcess in colProcess
'		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
'		        objProcess.Terminate()
'		        MsgBox "- ACTION: " & objProcess.Name & " terminated"
'			End If
'		Next
'		For Each objProcess in colProcess
'			If LCase(objProcess.Name) = LCase("WerFault.exe") OR LCase(objProcess.Name) = LCase("WerFault.exe *32") Then
'	        objProcess.Terminate()
'	        MsgBox "- ACTION: " & objProcess.Name & " terminated"
'			End If
'		Next                                 
End Function

Function Column_Count_Retrieve(Rownumber)
     Datatable.Getsheet("Global").SetCurrentrow(Rownumber)
     Temp=0  
            Col_Count=Datatable.GetSheet("Global").GetParametercount      
                 for j=1 to Col_Count
                         Val=Datatable.GetSheet("Global").GetParameter(j).Value
                               if Val<> "" then
                                     Temp=Temp+1
                               end if
                  Next
      msgbox   "Row  " &Rownumber&"   :   "&Temp&" rows"
End Function


Function funToAddMultipleDataInMultipleDict(subFunctionRowCount,SubFunctionColCount)

	For funRowDataDict = 1 To subFunctionRowCount Step 1
		Datatable.GetSheet(testFunctionName).SetCurrentRow(funRowDataDict)
		subTestCaseNumber = Datatable.Value("TestCaseNo",testFunctionName)
		nestedDictName = subTestCaseNumber
		Set nestedDictName = CreateObject("Scripting.Dictionary")
		'If Trim(subTestCaseNumber) = Trim(testCaseNumber) Then
		If Instr(1,subTestCaseNumber,testCaseNumber,1) <> 0 Then
			For funColDataDict = 1 To SubFunctionColCount Step 1
				funColName=Datatable.GetSheet(testFunctionName).GetParameter(funColDataDict).Name
				If not(nestedDictName.Exists(funColDataDict)) Then
					nestedDictName.Add funColName,Datatable.Value(funColName,testFunctionName)
				End If
			Next
			I=i+1
			dictTestData.Add subTestCaseNumber,nestedDictName
			'Set nestedDictName = Nothing
			'Exit For
			'Exit For	
		End If
	Next
	'msgbox dictTestData.Count
End  Function
Function funToAddFunctionLibraryDynamicallyold()
Set fso = CreateObject("Scripting.FileSystemObject") 
Set dir = fso.GetFolder(Environment.Value("strFunctionLibraryPath")) 'get num files in parent dir 
total_files = dir.Files.Count 'get num files in each sub-sir 
For Each sub_dir In dir.SubFolders 'get num files in each sub-sub-dir 
	'msgbox sub_dir.Name
	If sub_dir.Name<>"" Then
		Set oFolder = fso.GetFolder(sub_dir)
		For Each oFile in oFolder.Files
			print oFile.Name
			splitVal=split (oFile.Name,".")
			If Ubound(splitVal) = 1 Then
				If Ucase(splitVal(1)) = Ucase("vbs") then
					'Msgbox oFile.Name
					Msgbox oFile.Path
					LoadFunctionLibrary oFile.Path
				End If
			End If
		Next
	End If
'	For Each sub_sub_dir In sub_dir.SubFolders 
'		msgbox sub_sub_dir.Name
'		curr_dir_files = sub_sub_dir.Files.Count 
'		print curr_dir_files
'		total_files = total_files+curr_dir_files 
'		print total_files
'	Next 
'	curr_dir_files = sub_dir.Files.Count
'	total_files = total_files + curr_dir_files
Next 

End Function


