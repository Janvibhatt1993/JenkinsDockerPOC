
Public sResultFile,sResultLogFile
Public iPassCount, iFailCount, iStartTime, iErrImageNumber
Public slNo
Public CAPTURE_ERROR_SCREENSHOT,gblTakePassFailScreenshot

gblTakePassFailScreenshot = True '''If the Value is True then script will take Pass and Fail both steps Screenshot
CAPTURE_ERROR_SCREENSHOT = True	  '''And IF value is False then Script will take only Fail steps screenshots
	
' ***********************************************************************************************
'
' 			R E P O R T   L I B R A R Y   F U N C T I O N S 
'
' ***********************************************************************************************
'1. CreateResultFile
'2. LogResult
'3. TestCaseExecutiveSummary
'4. funToResetCounter()
'5. funToCalculateStartTimeEndTime
'6. funToCreateLogFile
'7. funToWriteLogsInFile
' ================================================================================================
'  NAME			 : funToResetCounter
'  DESCRIPTION 	  	: This function keeps iPassCount,iFailCount and sr no counter to 0.
'  PARAMETERS		: NA
' ================================================================================================

Public Function funToResetCounter()
	iPassCount = 0 
	iFailCount = 0
	slNo = 0
	iErrImageNumber = 1
End  Function
' ================================================================================================
'  NAME			 : CreateResultFile
'  DESCRIPTION 	  	: This function creates HTML test script execution result. Result file name created is same as the test script name. 
'					  This function is called at the begining of the test script
'  PARAMETERS		: NA
' ================================================================================================
Public Function CreateResultFile()
	On error resume next
	 Dim fso,fFolder
	 iStartTime = Now
	 funToResetCounter
	 Set fso = CreateObject("Scripting.FileSystemObject")
	 If (fso.FolderExists(Environment.value("strHtmlResultPath")&testCaseNumber)) Then
	 	fso.DeleteFolder(Environment.value("strHtmlResultPath")&testCaseNumber)
	 End  If
	Set fFolder = fso.CreateFolder(Environment.value("strHtmlResultPath")&testCaseNumber)
	 ' Create Test Result file from TestName 
	 'msgbox Environment.value("strHtmlResultPath")
	sResultFile = Environment.value("strHtmlResultPath") & testCaseNumber&"\" &testCaseNumber &"_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".html"
	'msgbox sResultFile
	Set MyFile = fso.CreateTextFile(sResultFile,True)
		MyFile.WriteLine("<html><head><title>Test Script Execution Report</title></head>")
		MyFile.WriteLine("<body><table border='1' width='100%' bordercolorlight='#C0C0C0' cellspacing='0' cellpadding='0' bordercolordark='#C0C0C0' bordercolor='#C0C0C0' height='88'>")
		MyFile.WriteLine("<tr><td width='100%' colspan='4' height='28' bgcolor='#C0C0C0'><p align='center'><b><font face='Verdana' size='4' color='#000080'>Test Script Execution Report</font></b></td></tr>")
		MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Script Path</font></b></td>")
		MyFile.WriteLine("<td width='84%' height='25' colspan='3'><p style='margin-left: 5'><font face='Verdana' size='2'>")
		MyFile.WriteLine(Environment.Value("TestDir") & "</font></td></tr>")
		MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Case Name</font></b></td>")
		MyFile.WriteLine("<td width='84%' height='25' colspan='3'><font face='Verdana' size='2'>&nbsp;" & Environment.Value("TestName") & "</font></td></tr>	")
		MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test URL</font></b></td>")
		MyFile.WriteLine("<td width='84%' height='25' colspan='3'><font face='Verdana' size='2'>&nbsp;"& testEnvironment &"</font></td></tr>	")
		
		MyFile.WriteLine("</table>")
		MyFile.WriteLine("<p style='margin-left: 5'>&nbsp; </p>")
		
		MyFile.WriteLine("<table border='1' width='100%' bordercolordark='#C0C0C0' cellspacing='0' cellpadding='0' bordercolorlight='#C0C0C0' bordercolor='#C0C0C0' height='91'>")	
		
		'============== Result Column Header =====================
		MyFile.WriteLine("<tr><td width='5%' align='center' bgcolor='#000099' height='35'><b>")
		MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Sl. No.</font></b></td>")
		MyFile.WriteLine("<td width='45%' align='center' bgcolor='#000099' height='35'><b>")
		''MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Step Description</font></b></td>")
		MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Description/Expected Result</font></b></td>")
		MyFile.WriteLine("<td width='40%' align='center' bgcolor='#000099' height='35'><b>")
		''MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Details</font></b></td>")
		MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Actual Result</font></b></td>")
		MyFile.WriteLine("<td width='10%' bgcolor='#000099' height='35' align='center'><b>")
		''MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Status</font></b></td></tr>")
		MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Status</font></b></td></tr>")
		' Close the file
		MyFile.Close
		CreateResultFile = sResultFile
		On error goto 0
End  Function
'=================================================================================================
'  NAME			: LogResult
'  DESCRIPTION 	  	: This function is used to write test step status to QTP result file and also to the HTML test result file.
'  PARAMETERS		: sStatus 	- Status of the step (micPass / micFail / micGeneral)
'				  	  sTestStep 	- Test Step Name
'	       		  	  sDescription - Test Step Description
' ================================================================================================

Public Function LogResult(sStatus, sTestStep, sDescription)
'	 Set objBrowser = Description.Create
'	Set objWindow  = Description.Create
'	objBrowser("creationtime").value = "0"
'	objWindow("regexpwndtitle").value = "Microsoft Internet Explorer|Windows Internet Explorer|Mozilla Firefox |Google Chrome"
	On error resume next
	Const ForAppending = 8
	Const TristateUseDefault = -2
	Dim fso, f, ts
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FolderExists (Environment.value("strHtmlResultPath") & "ErrorSnapshot") = False) Then
		fso.CreateFolder (Environment.value("strHtmlResultPath") & "ErrorSnapshot")
	End If

	Set f = fso.GetFile(sResultFile)
	Set ts = f.OpenAsTextStream(ForAppending, TristateUseDefault)
	
'	If (sStatus = micGeneral) Then
'		Reporter.ReportEvent micGeneral, sTestStep, sDescription
'		ts.WriteLine("<tr><td width='12%' height='25' align='left' colspan='3'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#000099'> " & sTestStep & " " & sDescription & "</font></b></td></tr>")
'        ts.Close
'        Exit Function
'	End If
	slNo = slNo + 1
	ts.WriteLine("<tr><td width='5%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & slNo & "</font></td>")
	ts.WriteLine("<td width='45%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & sTestStep & "</font></td>")
	ts.WriteLine("<td width='40%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>")
	
	' Replace vbcrlf in sDescription to <br>
	arrDesc = Split (sDescription, vbcrlf)
	
	For Each arrElement In arrDesc
		ts.WriteLine(arrElement & "<br>")
	Next
	ts.WriteLine("</font></td>")
	If sStatus = micPass Then
	    	If (gblTakePassFailScreenshot = True ) Then
			sImgRelativePath = Environment.value("strHtmlResultPath")&"ErrorSnapshot\" & testCaseNumber & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_PASS_STEP_" & iErrImageNumber & ".png"
			print sImgRelativePath
			Desktop.CaptureBitmap sImgRelativePath, True				
			ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#05A251'><a href='" & sImgRelativePath & "'>PASS </a></font></td></tr>")	    		
			Reporter.ReportEvent micPass, sTestStep, sDescription, sImgRelativePath
		Else
			ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#05A251'> PASS </font></b></td></tr>")
			Reporter.ReportEvent micPass, sTestStep, sDescription
		End If
	    ts.Close
	    iPassCount = iPassCount + 1
	ElseIf sStatus = micFail Then
		'If (Browser(objBrowser).Exist(1)) Then
			If (CAPTURE_ERROR_SCREENSHOT = True) Then 
				sImgRelativePath = Environment.value("strHtmlResultPath")&"ErrorSnapshot\" & testCaseNumber & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_FAIL_STEP_" & iErrImageNumber & ".png"
				Desktop.CaptureBitmap sImgRelativePath, True		
				ts.WriteLine("<td width='10' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'><a href='" & sImgRelativePath & "'>FAIL </a></font></td></tr>")
				iErrImageNumber = iErrImageNumber + 1
				Reporter.ReportEvent micFail, sTestStep, sDescription, sImgRelativePath
			Else
				ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'>FAIL</font></td></tr>")
				Reporter.ReportEvent micFail, sTestStep, sDescription
			End If
		'Else
		'	ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'>FAIL</font></td></tr>")
		'End If
		iFailCount = iFailCount + 1
	    ts.Close
	ElseIF sStatus = micDone Then
	    	Reporter.ReportEvent micDone, sTestStep, sDescription
	    	ts.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#05A251'> PASS </font></b></td></tr>")
	        ts.Close
	        iPassCount = iPassCount + 1
	End If
	On error goto 0
End Function
' ================================================================================================
'  NAME			: TestCaseExecutiveSummary
'  DESCRIPTION 	  	: This function is used to create test script executive summary. This function 
'					  is called from test script at the end of the test script
'  PARAMETERS		: NA
' ================================================================================================
Public Function TestCaseExecutiveSummary ()
	On error resume next
	Const ForAppending = 8
	Const TristateUseDefault = -2
	Dim fso, f, ts
	Dim iEndTime,iTotalTime
	iEndTime = Now
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(sResultFile)
	Set ts = f.OpenAsTextStream(ForAppending, TristateUseDefault)

	ts.WriteLine("</table>")
	ts.WriteLine("<p>&nbsp;</p>")
	ts.WriteLine("<table border='1' width='52%' bordercolorlight='#C0C0C0' cellspacing='0' cellpadding='0' bordercolordark='#C0C0C0' bordercolor='#C0C0C0' height='88'>")
	ts.WriteLine("<tr><td width='53%' colspan='2' height='28' bgcolor='#C0C0C0'><p align='center'><b><font face='Verdana' size='4' color='#000080'>")
	ts.WriteLine("Test Script Execution Summary</font></b></td></tr>")
	
	ts.WriteLine("<tr><td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#05A251'>")
	ts.WriteLine("Total Pass Count</font></b></td>")
	
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'color='#05A251'>&nbsp;" & iPassCount & "</font></td></tr>")
	ts.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#FF0000'>")
	ts.WriteLine("Total Fail Count</font></b></td>")
	
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'color='#FF0000'>&nbsp;" & iFailCount & "</font></td></tr>")
	ts.WriteLine("<tr><td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>")
	ts.WriteLine("Start Time</font></b></td>")
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & iStartTime & "</font></td></tr>")
	ts.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>End Time</font></b></td>")
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & iEndTime  & "</font></td></tr>")
	ts.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Total Time</font></b></td>")
	iTotalTime = funToCalculateStartTimeEndTime(iStartTime,iEndTime)
	ts.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & iTotalTime  & "</font></td></tr></table>")
	ts.Close
	
	Dim fso_TestScriptStatus
	Set fso_TestScriptStatus = CreateObject("Scripting.FileSystemObject")
	Set TestScriptExecutionStatus = fso_TestScriptStatus.CreateTextFile(Environment.value("strResultPath") & "\TCStatus.txt", True)
	If (iFailCount = 0) Then
		TestScriptExecutionStatus.WriteLine ("Passed")
	Else
		TestScriptExecutionStatus.WriteLine ("Failed")
	End If
	TestScriptExecutionStatus.Close
	On error goto 0
End Function
' ================================================================================================
'  NAME			 : funToCalculateStartTimeEndTime
'  DESCRIPTION 	  	: This function calculates and returns difference between startTime and endTime in terms of seconds
'  PARAMETERS		: Start_Time - Test case execution start time
'					  End_Time - Test case execution end time
' ================================================================================================
Function funToCalculateStartTimeEndTime(Start_Time,End_Time)
    On error resume next
    TotalTime_Secs  = Datediff("s",Start_Time,End_Time)
    'convert  total  Seconds into "Seconds only/ Mins+Secs/ Hrs+Mins+Secs"
    If TotalTime_Secs < 60 Then
        TotalTime = "Total Time Taken For Complete Execution = " & TotalTime_Secs  & " Second(s) Approx."
    ElseIf   TotalTime_Secs >=60 and TotalTime_Secs < 3600 Then
        TotalTime = "Total Time Taken For Complete Execution = " & int(TotalTime_Secs/60) & " Minute(s) and "& TotalTime_Secs Mod 60 & " Second(s) Approx."
    End  If
    funToCalculateStartTimeEndTime=TotalTime
    On error goto 0
End Function

' ================================================================================================
'  NAME			 : funToCreateLogFile
'  DESCRIPTION 	  	: This function creates log file with timestamp details of each event
'  PARAMETERS		: NA
' ================================================================================================
Public Function funToCreateLogFile()
	On error resume next
	 Dim fso,fFolder,timeStamp
	 Set fso = CreateObject("Scripting.FileSystemObject")
	 If Not(fso.FolderExists(Environment.value("strResultPath")&"TestCaseLogFiles")) Then
	 	Set fFolder = fso.CreateFolder(Environment.value("strResultPath")&"TestCaseLogFiles")
	 End  If
	 timeStamp=fnTimeStamp
	sResultLogFile = Environment.value("strResultPath")&"TestCaseLogFiles\"&"TestExecutionLog"&timeStamp&".txt"
	'msgbox sResultLogFile
	Set MyFile = fso.CreateTextFile(sResultLogFile,True)
		MyFile.WriteLine("===================LOGS================")
		MyFile.Close
	Set MyFile = Nothing
	On error goto 0
End  Function
' ================================================================================================
'  NAME			 : funToWriteLogsInFile
'  DESCRIPTION 	  	: This function is used to append the logs inside created log file
'  PARAMETERS		: NA
' ================================================================================================	
Public Function funToWriteLogsInFile(strStatements)
	On error resume next
	Const ForAppending = 8
	Dim myFile1,fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set myFile1 = fso.OpenTextFile(sResultLogFile, ForAppending, True)
		myFile1.WriteLine(now &"  "&strStatements)
		myFile1.Close
	Set myFile1=Nothing
	Set fso=Nothing
	On error goto 0
End Function

'*********************************************End**************************************************
