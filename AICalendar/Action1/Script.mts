
funToOpenBrowser "https://opensource-demo.orangehrmlive.com/index.php/dashboard","CHROME"
funToLogin "Admin","admin123"

Function funToOpenBrowser(strURL, strBrowserName)
	Systemutil.CloseProcessByName("chrome.exe")
	WebUtil.LaunchBrowser strBrowserName
	AIUtil.SetContext Browser("creationtime:=0")
	Browser("creationtime:=0").Navigate strURL
	funToWaitExplicit(4)
End Function

Function funToLogin(strUserName,strPassword)
	funToLoopUntilExist (AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT"))
	On error resume next
	funToWaitExplicit(2)
	'AIUtil.Context.Freeze
	AIUtil.Context.SetBrowserScope(BrowserWindow)
	AIUtil("text_box", "LOGIN Panel").Type strUserName
	AIUtil.FindTextBlock("Username :").Click
	AIUtil("text_box", "", micFromBottom, 1).Type "admin123"
	'AIUtil("text_box", "Password").Highlight
	'AIUtil("text_box", "Password").Type strPassword
	AIUtil("button", "LOGIN").Click
	If AIUtil.FindTextBlock("Leave").Exist or Err.Number <> 0 Then
	funToWaitExplicit(3)
		Reporter.ReportEvent micDone ,"User should suceesfully logged in Orange HRM website","User logged in ornage HRM website"
		funToSearchLeave()
	Else
		Reporter.ReportEvent micFail ,"User should suceesfully logged in Orange HRM website","User failed to log in ornage HRM website"
		ExitTest
	End If
	On error goto 0
	'AIUtil.Context.UnFreeze
End Function

Function funToSearchLeave()
	If AIUtil.FindTextBlock("Leave").Exist Then
		Err.Clear
		On error resume next
		AIUtil.FindTextBlock("Leave").Click
		AIUtil("check_box", "Taken Oo").SetState "On"
		'print "Current OCR config:: "&AIUtil.RunSettings.OCR.GetConfigSet
		If AIUtil.RunSettings.OCR.GetConfigSet <> UFT_OCR Then
			AIUtil.RunSettings.OCR.UseConfigSet UFT_OCR
		End If
		AIUtil("calendar", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", "From")).Click
		If AIUtil.Calendar.Exist(2) then	
			AIUtil.Calendar.CaptureBitmap "D:\AI Mockup\imgCalendar.bmp",True
			AIUtil("calendar", micAnyText, micFromBottom, 1).Highlight
			AIUtil("text_box", "From").Type "2022-05-14"
			AIUtil("calendar", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", "From")).Click
			AIUtil("text_box", "Employee").Type "Admin A"
			AIUtil("text_box", "To").Type "2022-05-28"
			AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
			funToWaitExplicit(2)
			AIUtil("button", "Search").Click
			funToWaitExplicit(3)
			Reporter.ReportEvent Pass ,"User should search leave successfully ","Leave searched successfully"
			funToViewDirectory()
		Else
			Reporter.ReportEvent micFail ,"User should click on calendar ","Calendar is not displayed"
			'ExitTest	
		End If
		On Error goto 0
	'AIUtil("button", "", micFromTop, 2).Click
	End If	
End Function

Function funToViewDirectory()
	Err.Clear
	On error resume next
	'funToCheckCPUUsage
	AIUtil.ScrollOnObject AIUtil.FindTextBlock("Maintenance", micWithAnchorOnLeft, AIUtil.FindTextBlock("Directory")),"up",2
	If AIUtil.FindTextBlock("Maintenance", micWithAnchorOnLeft, AIUtil.FindTextBlock("Directory")).Exist  Then
		AIUtil.FindTextBlock("Maintenance", micWithAnchorOnLeft, AIUtil.FindTextBlock("Directory")).Hover
		AIUtil.FindTextBlock("Access Records").LongClick 1.5
		funToWaitExplicit(3)
		AIUtil("bell").MultiClick 4
		AIUtil("bell").CaptureBitmap "D:\AI Mockup\imgDirectory.bmp",True
		Reporter.ReportEvent micPass ,"Leave search and Directory view operation should completed successfully ","Leave search and Directory view operation completed successfully"	
		AIUtil("close", micAnyText, micFromTop, 1).Click
		
	Else
		Reporter.ReportEvent micFail ,"User should navigate on directory ","User unable to navigate on directory"	
	End If
	On error goto 0
End Function



Function funToLoopUntilExist(strObj)
	Err.Clear
	On error resume next
	Dim intCount, inCountMax
	inCount=0
	inCountMax=20
	Do
		inCount = inCount +1
	Loop until strObj.exist(1) or intCount > inCountMax
	If Err.Number<> 0 Then
		'print "Object :: "&strObj& "is not exist"
		Reporter.ReportEvent micFail, "Object exist", "Object ::"&strObj& " is not exist"
	End If
	On error goto 0
	Set inCount = Nothing
	Set inCountMax = Nothing
End Function

Function funToWaitExplicit(strTime)
	wait strTime
End Function

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

Function funToCheckCPUUsage()
	For Each Process in GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
	    'print Process.name & " " & CPUUSage(Process.Handle) & " %"   
	Next	
End Function

Function CPUUSage( ProcID )
    On Error Resume Next

    Set objService = GetObject("Winmgmts:{impersonationlevel=impersonate}!\Root\Cimv2")

    For Each objInstance1 in objService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process where IDProcess = '" & ProcID & "'")
        N1 = objInstance1.PercentProcessorTime
        D1 = objInstance1.TimeStamp_Sys100NS
        Exit For
    Next

     wait (20)

    For Each perf_instance2 in objService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process where IDProcess = '" & ProcID & "'")
        N2 = perf_instance2.PercentProcessorTime
        D2 = perf_instance2.TimeStamp_Sys100NS
        Exit For
    Next

    ' CounterType - PERF_100NSEC_TIMER_INV
    'Formula - (1- ((N2 - N1) / (D2 - D1))) x 100
    'print Formula
    Nd = (N2 - N1)
    Dd = (D2-D1)
    PercentProcessorTime = ( (Nd/Dd))  * 100
    'print PercentProcessorTime
    CPUUSage = Round(PercentProcessorTime ,0)
End Function

'Systemutil.CloseProcessByName("chrome.exe")
'	WebUtil.LaunchBrowser "CHROME"
'	AIUtil.SetContext Browser("creationtime:=0")
'	AIUtil.Context.SetBrowserScope(BrowserWindow)
'	Browser("creationtime:=0").Navigate "https://opensource-demo.orangehrmlive.com/index.php/dashboard"
'	wait 4
'	AIUtil("text_box", "LOGIN Panel").Type "Admin"
'	AIUtil.FindTextBlock("Username :").Click
'	'AIUtil("text_box", "").Type 
'	AIUtil("text_box", "", micFromBottom, 1).Type "admin123"
'	'AIUtil("text_box", "Password").Type "admin123"
'	AIUtil("button", "LOGIN").Click
