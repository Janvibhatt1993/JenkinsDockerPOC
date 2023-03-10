Function funToLogin(Parameter2)
	funToLoopUntilExist (AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT"))
	On error resume next
	funToWaitExplicit(2)
	'AIUtil.Context.Freeze
	AIUtil.Context.SetBrowserScope(BrowserWindow)
	AIUtil("text_box", "Username").Highlight
	LogResult micDone, "User should enter username","User has entered username: "&Parameter2("UserName")
	AIUtil("text_box", "Username").Type Parameter2("UserName")
	funToWaitExplicit(1)
	LogResult micDone, "User should enter password","User has entered password: "&crypt.Encrypt(Parameter2("UserName"))
	AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Click
	funToWaitExplicit(1)
	AIUtil("text_box", "Password").Type Parameter2("Password")
	LogResult micDone, "User should click on login button","User clicked on login button "
	AIUtil("button", "Login").Click
	If AIUtil.FindTextBlock("Dashboard", micFromTop, 1).Exist or Err.Number <> 0 Then
		funToWaitExplicit(1)
		LogResult micDone, "User should suceesfully logged in Orange HRM website","User logged in ornage HRM website"
	Else
		LogResult micFail, "User should suceesfully logged in Orange HRM website","User failed to logged in ornage HRM website error details :: "&err.description
		'ExitTest
	End If
	On error goto 0
	'AIUtil.Context.UnFreeze
End Function


Function funToSearchLeave(Parameter2)
	On error resume next
	If AIUtil.FindTextBlock("Leave").Exist and Err.Number = 0 Then
		LogResult micDone, "User should click on leave","User clicked on leave"
		AIUtil.FindTextBlock("Leave").Click
		funToWaitExplicit(2)
		LogResult micDone, "User should enter start date and end date in calender option","User entered start date and end date in calender option"
		AIUtil("calendar", micAnyText, micFromLeft, 1).Click
		If AIUtil.Calendar.Exist(2) Then	
			AIUtil("text_box", "From Date").Click
			funToDeleteValueFromWebedit
			AIUtil("text_box", "From Date").Type Parameter2("StartDate")
			AIUtil("text_box", "To Date").Click
			funToDeleteValueFromWebedit
			AIUtil("text_box", "To Date").Type Parameter2("EndDate")
			LogResult micDone, "User should select rejected leave status","User selected rejected leave status"
			AIUtil("combobox", "Show Leave with Status*").Click
			AIUtil.FindTextBlock("Rejected").Click
			AIUtil("toggle_button", "Include Past Employees").SetState "On"
			AIUtil("button", "Search", micFromBottom, 1).Click
			funToWaitExplicit(1)
			If Err.number <> 0 Then
				LogResult micFail, "User should search leave successfully","Functionality failed ::"& Err,description
			Else
				LogResult micDone, "User should search leave successfully","Leave searched successfully"
			End If
		Else
			LogResult micFail, "User should search leave successfully","Leave failed to searched"
			'ExitTest	
		End If
	Else
		LogResult micFail, "User should search leave successfully","User unable to search leave "&err.description
	End If
	On Error goto 0
	Err.Clear	
End Function

Function funToLoopUntilExist(strObj)
	On error resume next
	Dim intCount, inCountMax
	inCount=0
	inCountMax=20
	Do
		inCount = inCount +1
	Loop until strObj.exist(1) or Cint(inCount) = Cint(inCountMax)
	If Err.Number<> 0 Then
		LogResult micFail, "Object exist", "Object ::"&strObj& " is not exist"
	End If
	On error goto 0
	Set inCount = Nothing
	Set inCountMax = Nothing
End Function

Function funToWaitExplicit(strTime)
	wait strTime
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
        myDeviceReplay.PressKey 209
End Function

Function funToWorkWithRegisterCustomClass(Parameter3)
	funToLoopUntilExist (AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT"))
	On error resume next
	If AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Exist(2) and Err.Number = 0 Then
		LogResult micDone, "User should enter username","User entered username"
		AIUtil("text_box", "Username").Type Parameter3("UserName")
		funToWaitExplicit(2)
		AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Click
		funToWaitExplicit(2)
		LogResult micDone, "User should enter password","User entered password"
		AIUtil("text_box", "Password").Type Parameter3("Password")
		AIUtil("button", "LOGIN").Click
		LogResult micDone, "User should click on Login button","User clicked on Login button"
		'-----see image or object ---????
		LogResult micDone, "User should register custom class using class image","User has registered custom class using class image"
		AIUtil.RegisterCustomClass Parameter3("ClassName"), Parameter3("ClassImagePath")
		
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
		LogResult micDone, "User should scroll up, down, doble click, right click and extract the value from object ","User has scroll up, down, doble click, right click and extracted the value from object"	
	Else
		LogResult micFail, "User should be able to navigate on Login screen","User unable to navigate on Login screen "&err.description
		'AIUtil("button", "LOGIN").Click
	End If
	On error goto 0
	Err.Clear
End Function

Function funToViewDirectory()
	On error resume next
	AIUtil.ScrollOnObject AIUtil.FindTextBlock("SPECIAL OFFER", micFromBottom, 1),"down",4
	If AIUtil("button", "SEE OFFER", micWithAnchorOnRight, AIUtil.FindTextBlock("Supremely thin, yet incredibly durable")).Exist and Err.Number = 0 Then
		LogResult micDone, "User should click on SEE OFFER button","User clicked on SEE OFFER button"
		AIUtil("button", "SEE OFFER", micWithAnchorOnRight, AIUtil.FindTextBlock("Supremely thin, yet incredibly durable")).LongClick 1.5
		AIUtil("profile", micAnyText, micWithAnchorOnLeft, AIUtil("search")).Hover
		AIUtil("help", micAnyText, micWithAnchorOnLeft, AIUtil("shopping_cart")).MultiClick 4
		LogResult micDone, "User should perform longclick, mouse hover, multiclick and scroll on object successfully","User performed longclick, mouse hover, multiclick and scroll on object successfully"
	Else
		LogResult micFail, "User should navigate on SEE OFFER button","User failed to navigate on SEE OFFER button"	
		'ExitTest
	End If
	On error goto 0
	Err.Clear
End Function

Function funToValidateDifferentLangu()
	On error resume next
	If AIUtil.FindTextBlock("English-Chinese Dictionary").Exist = True Then
		If AIUtil.RunSettings.OCR.GetConfigSet <> AI_OCR Then
			AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
		End  If
		funToWaitExplicit(2)
		AIUtil.Scroll "down" , 2
		AIUtil.RunSettings.Ocr.Languages="en,zhs,zht"
		LogResult micDone, "User should change OCR setting to add multiple languages","User has successfully changeed OCR setting to add multiple languages"	
		If AIUtil.FindTextBlock("热门搜索").Exist = True Then
			AIUtil.FindTextBlock("热门搜索").Highlight
			strChinesText=AIUtil.FindTextBlock("热门搜索").GetValue
			Print "String in Chinese ::" &strChinesText
			LogResult micDone, "User should extract chinese string", "User has extracted chinese string"
			AIUtil.FindTextBlock("English-Chinese Dictionary").Highlight
			strEnglishText=AIUtil.FindTextBlock("English-Chinese Dictionary").GetValue
			Print "String in English ::" &strEnglishText
		Else
			LogResult micFail, "User should extract chinese string","User failed to extract chinese string value"
		End If
	Else
		LogResult micFail, "User should navigate on wordpress website","User failed to navigate on wordpress website"	
	End If
	On error goto 0
	Err.Clear
End Function

Function funToWorkWithWebTable()
	On error resume next
	If AIUtil.FindTextBlock("SPEAKERS").Exist = True and Err.Number = 0 Then
		AIUtil.FindTextBlock("SPEAKERS").Click
		LogResult micDone, "User should click on SPEAKER link","User clicked on SPEAKER link"
		AIUtil("button", "BUY NOW").Click
		funToWaitExplicit(4)
		LogResult micDone, "User should click on BUY NOW link","User clicked on BUY NOW link"
		AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
		'AIUtil.Scroll "down", 2
		funToPageDown
		AIUtil.RunSettings.OCR.UseConfigSet UFT_OCR
		funToWaitExplicit(3)
		AIUtil.Table(micFromBottom,1).Highlight
		funToWaitExplicit(3)
		colCount=AIUtil.Table(micFromBottom,1).VisibleColumnCount
		Print "Total No of Column count in Webtable:: " &AIUtil.Table(micFromBottom,1).VisibleColumnCount
		rowCount=AIUtil.table(micFromBottom,1).VisibleRowCount
		Print "Total No of Row count in Webtable:: " & AIUtil.table(micFromBottom,1).VisibleRowCount
		If Cint(rowCount) <> 0 Then
			LogResult micDone, "User should extract webtable total number of rows and columns","User has extracted webtable rows and columns"&" Row count ::"&rowCount& " Column count ::"&colCount
			For Iterator = 0 To rowCount-1 Step 1
				For Iterator1 = 0 To colCount-1 Step 1
					AIUtil.Table(micFromBottom, 1).Cell(Iterator, Iterator1).Highlight
					val1=AIUtil.Table(micFromBottom, 1).Cell(Iterator, Iterator1).GetText
					print val1
				Next
			Next
		Else 
			LogResult micFail, "User should extract webtable total number of rows and columns","User failed to extract webtable rows and columns"
		End If
	Else
		LogResult micFail, "User should click on SPEAKER link","User failed to click on SPEAKER link"
		'ExitTest
	End If
	On Error goto 0
	Err.Clear
End Function



