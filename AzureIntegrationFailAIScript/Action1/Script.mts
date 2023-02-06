
'-------------Execute with Chrome-----------------------
'----Scenario 1: Perform various AI methods like checkbox, toogle, calender etc using AI-----
funToOpenBrowser "https://opensource-demo.orangehrmlive.com/","CHROME"
funToLogin "Admin","admin123"
funToSearchLeave "2022-2-1" , "2022-11-30"


Function funToOpenBrowser(strURL, strBrowserName)
	On error resume next
	val =funToExtractString(strBrowserName)
	Systemutil.CloseProcessByName(val&".exe")
	funToCloseAllOpenBrowser
	WebUtil.LaunchBrowser strBrowserName
	AIUtil.SetContext Browser("creationtime:=0")
	Browser("creationtime:=0").Navigate strURL
	funToWaitExplicit(2)
	On error goto 0
End Function

Function funToLogin(strUserName,strPassword)
	funToLoopUntilExist (AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT"))
	Reporter.ReportEvent micDone,"User Should Navigate on OrangeHRM website","User successfully navigated on OrangeHRM website"
	On error resume next
	funToWaitExplicit(2)
	'AIUtil.Context.Freeze
	AIUtil.Context.SetBrowserScope(BrowserWindow)
	AIUtil("text_box", "Username").Type strUserName
	Reporter.ReportEvent micDone , "User Should Enter UserName","User has entered USerName::" & strUserName
	funToWaitExplicit(1)
	AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Click
	funToWaitExplicit(1)
	AIUtil("text_box", "Password").Type strPassword
	Reporter.ReportEvent micDone , "User Should Enter Password","User has entered Password::" & crypt.Encrypt(strPassword)
	
	AIUtil("button", "Login").Click
	Reporter.ReportEvent micDone , "User Should click on Login Button","User clicked on Login Button"
	If AIUtil.FindTextBlock("Dashboard", micFromTop, 1).Exist or Err.Number <> 0 Then
		funToWaitExplicit(1)
		Reporter.ReportEvent micFail ,"User should suceesfully logged in Orange HRM website","User logged in ornage HRM website"
	Else
		Reporter.ReportEvent micFail ,"User should suceesfully logged in Orange HRM website","User failed to log in ornage HRM website"
		ExitTest
	End If
	On error goto 0
	'AIUtil.Context.UnFreeze
End Function

Function funToSearchLeave(strStartDate,strEndDate)
	On Error resume next
	If AIUtil.FindTextBlock("Leave").Exist and Err.Number = 0 Then
		On error resume next
		AIUtil.FindTextBlock("Leave").Click
		funToWaitExplicit(2)
		AIUtil("calendar", micAnyText, micFromLeft, 1).Click
		If AIUtil.Calendar.Exist(2) Then	
			AIUtil("text_box", "From Date").Click
			funToDeleteValueFromWebedit
			AIUtil("text_box", "From Date").Type strStartDate
			AIUtil("text_box", "To Date").Click
			funToDeleteValueFromWebedit
			AIUtil("text_box", "To Date").Type strEndDate
			AIUtil("combobox", "Show Leave with Status*").Click
			AIUtil.FindTextBlock("Rejected").Click
			AIUtil("toggle_button", "Include Past Employees").SetState "On"
			AIUtil("button", "Search", micFromBottom, 1).Click
			funToWaitExplicit(1)
			Reporter.ReportEvent micPass ,"User should search leave successfully ","Leave searched successfully"
		Else
			Reporter.ReportEvent micFail ,"User should click on calendar ","Calendar is not displayed"
			ExitTest	
		End If
		On Error goto 0
		Err.Clear
	Else
		Reporter.ReportEvent micFail ,"Leave option should display","User failed to click on Leave option button"
	End If	
End Function

Function funToViewDirectory()
	On error resume next
	AIUtil.ScrollOnObject AIUtil.FindTextBlock("SPECIAL OFFER", micFromBottom, 1),"down",2
	If AIUtil("button", "SEE OFFER", micWithAnchorOnRight, AIUtil.FindTextBlock("Supremely thin, yet incredibly durable")).Exist and Err.Number = 0 Then
		AIUtil("button", "SEE OFFER", micWithAnchorOnRight, AIUtil.FindTextBlock("Supremely thin, yet incredibly durable")).LongClick 1.5
		AIUtil("profile", micAnyText, micWithAnchorOnLeft, AIUtil("search")).Hover
		AIUtil("help", micAnyText, micWithAnchorOnLeft, AIUtil("shopping_cart")).MultiClick 4
		Reporter.ReportEvent micPass ,"Leave search and Directory view operation should completed successfully ","Leave search and Directory view operation completed successfully"	
	Else
		'Msgbox "not exist"
		ExitTest
	End If
	On error goto 0
	Err.Clear
End Function



Function funToLoopUntilExist(strObj)
	On error resume next
	Dim intCount, inCountMax
	inCount=0
	inCountMax=20
	Do
		inCount = inCount +1
	Loop until strObj.exist(1) or Cint(intCount) = Cint(inCountMax)
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

Function funToWorkWithRegisterCustomClass(strUsername,strPassword,strClassName,strClassPath)
	funToLoopUntilExist (AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT"))
	On error resume next
	If AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Exist(2) and Err.Number = 0 Then
		AIUtil("text_box", "Username").Type strUsername
		funToWaitExplicit(2)
		AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Click
		funToWaitExplicit(2)
		AIUtil("text_box", "Password").Type strPassword
		
		AIUtil("button", "LOGIN").Click
		'-----see image or object ---????
		AIUtil.RegisterCustomClass strClassName, strClassPath
		
		AIUtil(strClassName).Click
		AIUtil(strClassName).Highlight
		AIUtil.Scroll "down" , 1
		AIUtil.FindTextBlock("Forgot your password?", micWithAnchorAbove, AIUtil(strClassName)).Highlight
		AIUtil.Scroll "up", 1
		
		AIUtil.FindTextBlock("Login", micFromTop, 1).Highlight
		AAIUtil.FindTextBlock("Login", micFromTop, 1).click 628, 318
		AIUtil.FindTextBlock("Login", micFromTop, 1).DoubleClick 628, 460
		
		print "Text Details 1 ::"&AIUtil.FindTextBlock("Login", micFromTop, 1).GetText
		print "Text Details 2 ::"& AIUtil.FindTextBlock("Login", micFromTop, 1).GetValue
		AIUtil.FindTextBlock("Login", micFromTop, 1).RightClick
		'AIUtil("button", "LOGIN").Click
	End If
	On error goto 0
	Err.Clear
End Function

Function funToExtractString(strStringName)
	On error resume next
	Dim obj,strExtractedVal,regEx
	strExtractedVal = ""
	Set regEx = New RegExp
	regEx.pattern="\D"
	regEx.Global = True
	Set obj = regEx.Execute(strStringName)
	'msgbox obj.count
	For i=0 to obj.count-1
		strExtractedVal= strExtractedVal&obj.item(i)
	Next
	funToExtractString=strExtractedVal
	Set obj = Nothing
	Set regEx = Nothing
	On error goto 0
End Function

Function funToCloseAllOpenBrowser()
	On error resume next
	Set brow = description.Create
	brow("micclass").Value = "Browser"
	num=Desktop.ChildObjects(brow).count
	If num > 0 Then
		For Iterator = num-1 To 0 Step -1
			Browser("creationtime:="&Iterator).Close
		Next
	End If
	On error goto 0
End Function




Function funToValidateDifferentLangu()
	On error resume next
	If AIUtil.RunSettings.OCR.GetConfigSet <> AI_OCR Then
		AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
	End  If
	funToWaitExplicit(2)
	AIUtil.Scroll "down" , 2
	AIUtil.RunSettings.Ocr.Languages="en,zhs,zht"
	AIUtil.FindTextBlock("热门搜索").Highlight
	strChinesText=AIUtil.FindTextBlock("热门搜索").GetValue
	Print "String in Chinese ::" &strChinesText
	AIUtil.FindTextBlock("English-Chinese Dictionary").Highlight
	strEnglishText=AIUtil.FindTextBlock("English-Chinese Dictionary").GetValue
	Print "String in English ::" &strEnglishText
	On error goto 0
	Err.Clear
End Function

Function funToWorkWithWebTable()
	On error resume next
	If AIUtil.FindTextBlock("SPEAKERS").Exist = True and Err.Number = 0 Then
		AIUtil.FindTextBlock("SPEAKERS").Click
		AIUtil("button", "BUY NOW").Click
		AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
		'AIUtil.Scroll "down", 2
		funToPageDown
		AIUtil.RunSettings.OCR.UseConfigSet UFT_OCR
		AIUtil.Table(micFromBottom,1).Highlight
		colCount=AIUtil.Table(micFromBottom,1).VisibleColumnCount
		Print "Total No of Column count in Webtable:: " &AIUtil.Table(micFromBottom,1).VisibleColumnCount
		rowCount=AIUtil.table(micFromBottom,1).VisibleRowCount
		Print "Total No of Row count in Webtable:: " & AIUtil.table(micFromBottom,1).VisibleRowCount
		If Cint(rowCount) <> 0 Then
			For Iterator = 0 To rowCount-1 Step 1
				For Iterator1 = 0 To colCount-1 Step 1
					AIUtil.Table(micFromBottom, 1).Cell(Iterator, Iterator1).Highlight
					val1=AIUtil.Table(micFromBottom, 1).Cell(Iterator, Iterator1).GetText
					print val1
				Next
			Next
		Else 
			Print "not able to find rows and columns"	
		End If
	Else
		'Msgbox "not exist"
		'ExitTest
	End If
	On Error goto 0
	Err.Clear
End Function


Function funToSearchLeaveOld()
	On Error resume next
	If AIUtil.FindTextBlock("Leave").Exist and Err.number = 0 Then
		AIUtil.FindTextBlock("Leave").Click
		funToWaitExplicit(2)
		AIUtil("calendar", micAnyText, micFromLeft, 1).Click
		If AIUtil.Calendar.Exist(2) then	
			AIUtil("text_box", "From Date").Click
			Dim myDeviceReplay
		   	Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
		     	For i = 1 to 11
		            myDeviceReplay.PressKey 14
		      	next
		      	AIUtil("text_box", "From Date").Type "2022-01-01"
'			wait 1
'			If AIUtil.Calendar.FindTextBlock("Clear").Exist(1) <> False Then
'				'AIUtil.Calendar.FindTextBlock("Clear").Highlight
'				AIUtil.Calendar.FindTextBlock("Clear").Click
'				AIutil.Scroll "up" , 1
'			Else
'				'AIUtil.FindTextBlock("Clear", micWithAnchorOnLeft, AIUtil.FindTextBlock("Today")).Highlight
'				AIUtil.FindTextBlock("Clear", micWithAnchorOnLeft, AIUtil.FindTextBlock("Today")).Click
'				AIutil.Scroll "up" , 1
'			End If


			
			
			'AIutil.Scroll "down" , 1
			'AIUtil("calendar", micAnyText, micFromRight, 1).Highlight
			
			'AIUtil("calendar", micAnyText, micFromRight, 1).Click

			'AIUtil.Calendar.FindTextBlock("Clear", micWithAnchorOnLeft, AIUtil.Calendar.FindTextBlock("Today")).Highlight
			'AIUtil.Calendar.FindTextBlock("Clear", micWithAnchorOnLeft, AIUtil.Calendar.FindTextBlock("Today")).Click
			'AIutil.Scroll "up" , 1
			'AIUtil.Calendar.FindTextBlock("Clear").Click
			AIUtil("text_box", "To Date").Click
			Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
		     	For i = 1 to 11
		            myDeviceReplay.PressKey 14
		      	next
			AIUtil("text_box", "To Date").Type "2022-11-30"
			'AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
			AIUtil("combobox", "Show Leave with Status*").Click
			AIUtil.FindTextBlock("Rejected").Click
			AIUtil("toggle_button", "Include Past Employees").SetState "On"
			AIUtil("button", "Search", micFromBottom, 1).Click
			funToWaitExplicit(2)
			Reporter.ReportEvent Pass ,"User should search leave successfully ","Leave searched successfully"
		Else
			Reporter.ReportEvent micFail ,"User should click on calendar ","Calendar is not displayed"
			ExitTest	
		End If
		On Error goto 0
		Err.Clear
	End If	
End Function

Function funToDeleteValueFromWebedit()
	Dim myDeviceReplay
	Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
     	For i = 1 to 11
            myDeviceReplay.PressKey 14
      	Next
End Function

Function funToPageDown()
	Dim myDeviceReplay
	Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
     	'For i = 1 to 11
        myDeviceReplay.PressKey 209
      	'Next
End Function








