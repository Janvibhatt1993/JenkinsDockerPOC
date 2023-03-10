Set newDict = CreateObject("Scripting.Dictionary")

'Set webUI = Browser("Advantage Shopping").Page("Advantage Shopping")

'------------------------register user functions-----------------------

''RegisterUserFunc "WebEdit", "fnSetValueInWebEdit", "fnSetValue"
'RegisterUserFunc "Link", "fnToClickOnLink", "fnClick"
'RegisterUserFunc "WebList", "fnToSelectValueFromWebList", "fnSetValueDropDwn"
'RegisterUserFunc "WebCheckBox", "fnSetValueInWebCheckBox", "fnSetValue"
'RegisterUserFunc "WebButton", "fnToClickOnWebButton", "fnClick"
''
' ***********************************************************************************************
'
' 			A P P L I C A T I O N   S P E C I F I C   F U N C T I O N S 
'
' ***********************************************************************************************
'1. MainFunction()
'2. funToRegisterUser(Parameter1)
'3. funToLogOut()

'=========================================================================================
'  NAME			 : invokeBrowser
'  DESCRIPTION 	  	: This function is designed to invoke the browser as per Test data
'  PARAMETERS		: NA
' =========================================================================================
public Function invokeBrowser()
	'Browser("Advantage Shopping").DeleteCookies "Advantage Shopping"
	On Error Resume Next
	systemutil.CloseProcessByName("iexplore.exe")
	systemutil.CloseProcessByName("chrome.exe")
	systemutil.CloseProcessByName("firefox.exe")
	funToCloseAllOpenBrowser																						
	Dim mode_Maximized, mode_Minimized,strURL
	mode_Maximized = 3 'Open in maximized mode
	mode_Minimized = 2 'Open in minimized mode
	
	strURL = split(testEnvironment,"=")(1)
	If Lcase(Left(testFunctionName,2)) <>  Lcase("AI") Then
		If UCASE(browserName) = "IE" Then						'open Browser according to XLS sheet
			SystemUtil.Run "iexplore.exe", strURL , , ,mode_Maximized 
		End If
		
		If UCASE(browserName) = "CHROME" Then
			SystemUtil.Run "chrome.exe", strURL , , ,mode_Maximized
		End If
		
		If UCASE(browserName) = "FIREFOX" Then
			SystemUtil.Run "firefox,.exe", strURL , , ,mode_Maximized
		End If
		
		'Browser("Advantage Shopping").DeleteCookies
		'Browser("Advantage Shopping").ClearCache
		If Browser("Advantage Shopping").Page("Advantage Shopping").Exist(15) Then
			invokeBrowser = true
			Browser("Advantage Shopping").Page("Advantage Shopping").Sync
			LogResult micDone, "User should navigate on Advantage Shopping website" , "User successfully navigated on Advantage Shopping website"
		Else
			invokeBrowser = false
			LogResult micFail, "User should navigate on Advantage Shopping website " , "User unable to navigate on Advantage Shopping website"
		End If
	Else
		WebUtil.LaunchBrowser browserName
		AIUtil.SetContext Browser("creationtime:=0")
		Browser("creationtime:=0").Navigate strURL
		funToWaitExplicit(8)
	End If
	On error goto 0
End Function

'=========================================================================================
'  NAME			 : MainFunction
'  DESCRIPTION 	  	: This is main function to drive all the Test cases as per Test function mentioned in data sheet.
'  PARAMETERS		: NA
' =========================================================================================

public Function MainFunction()
	On error resume next
	CreateResultFile
	funToWriteLogsInFile "User is inside main function"
	
	Dim akeys,key,strTestCaseResult
	invokeBrowser
	funToRegitseruserfunction
	Select Case testFunctionName
		Case "Login"
			LogResult micDone, "User should able to perform Login functionality " , "User Login menu appeared" 
			funToRegisterUser dictTestData 
		Case "AddToCart"
			LogResult micDone, "User should perform Add to cart and checkout functionality" , "User has added Instruments in cart and proceed to checkout"
			funToAddProductsInCart
		Case "RegisterMultipleUser"
			LogResult micDone, "User should perform RegisterMultipleUser Function" , "User Login menu appeared" 
			Dim intI,Iterator,innerkey: intI = 0
			For Each Key in dictTestData.Keys
		  		'Print "Key - " & Key
				For each innerkey in dictTestData(key)
				    'print innerkey
				    'print dictTestData(key)(innerkey)	
				    newDict.Add innerkey, dictTestData(key)(innerkey)
				Next	    
				 For Iterator = intI To dictTestData.Count Step 1
						funToRegisterUser(newDict)
						newDict.RemoveAll
						intI = intI+1
						Exit For
				  Next
			 Next
		Case "AIMethods"
			LogResult micDone, "User should able to perform various actions like checkbox, toogle, calender etc using AI " , "User navigated on login page of OrangeHRM website" 
			funToLogin dictTestData
			funToSearchLeave dictTestData
		Case "AICustomFunction"
			LogResult micDone, "User should able to perform register custom function of AI " , "User navigated on login page of OrnageHRM website " 
			funToWorkWithRegisterCustomClass dictTestData
		Case "AICrossBrowser"
			LogResult micDone, "User should able to perform various click methods using AI " , "User navigated on login Advantage online shopping website " 
			funToViewDirectory
		Case "AILanguage"
			LogResult micDone, "User should validate multi languages using AI " , "User navigated on word reference website "
			funToValidateDifferentLangu
		Case "AIWebTable"
			LogResult micDone, "User should validate and extract details of webtable using AI" , "User navigated on Advantage online website "
			funToWorkWithWebTable

	End Select
	
	TestCaseExecutiveSummary
	funToDeRegitseruserfunction
	If Err.Number<> 0 Then
		print Err.Number
		print Err.description
		strTestCaseResult="Failed"
		funToWriteLogsInFile "Error in a Test case execution"&vbcrlf
	else
		strTestCaseResult="Passed"
		funToWriteLogsInFile "No error in a Test case execution"&vbcrlf
	End If
	MainFunction=strTestCaseResult
	On error goto 0
End Function


Function funToRegisterUser(Parameter1)
	On error resume next
	Browser("Advantage Shopping").Page("Advantage Shopping").Highlight
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu").Highlight
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu").fnToClickOnLink
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("CREATE NEW ACCOUNT").fnToClickOnLink
	wait 2
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("usernameRegisterPage").fnSetValueInWebEdit Parameter1("UserName")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("passwordRegisterPage").fnSetValueInWebEdit Parameter1("Password")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("confirm_passwordRegisterPage").fnSetValueInWebEdit Parameter1("ConfirmPassword")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("emailRegisterPage").fnSetValueInWebEdit Parameter1("Email")
	wait 2
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("first_nameRegisterPage").fnSetValueInWebEdit Parameter1("FirstName")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("last_nameRegisterPage").fnSetValueInWebEdit Parameter1("LastName")
	wait 2
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("phone_numberRegisterPage").fnSetValueInWebEdit Parameter1("PhoneNumber")
	wait 2
	Browser("Advantage Shopping").Page("Advantage Shopping").WebList("countryListboxRegisterPage").fnToSelectValueFromWebList Parameter1("Country")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("cityRegisterPage").fnSetValueInWebEdit Parameter1("City")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("addressRegisterPage").fnSetValueInWebEdit Parameter1("Address")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("state_/_province_/_regionRegis").fnSetValueInWebEdit Parameter1("State")
	wait 2
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("postal_codeRegisterPage").fnSetValueInWebEdit Parameter1("PostalCode")
	Browser("Advantage Shopping").Page("Advantage Shopping").WebCheckBox("i_agree").fnSetValueInWebCheckBox "ON"
	wait 2
	Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("register_btnundefined").Highlight
	Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("register_btnundefined").fnToClickOnWebButton
	If Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("innertext:=User name already exists ").Exist(4) = True Then
		wait 4
		If Instr(1,testCaseDescription,"already a user status",1) Then
			LogResult micPass, "Allready a user status validation " , "Allready a user status validated successfully"
		else
			LogResult micFail, "User All-ready exists" , "User All-ready exists, enter new user details"
		End If
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME").fnToClickOnLink
	Else
		funToLogOut
		'Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME").fnToClickOnLink
		LogResult micDone, "User should added successfully" , "User added successfully"
	End If
	On error goto 0
End Function

Function funToLogOut()
	On error resume next
	wait 5
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("UserMenu").fnToClickOnLink
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("Sign out").Highlight
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("Sign out").fnToClickOnLink
	
	If Err.number <> 0 Then
		LogResult micFail, "User clicks on Logout button" , "User unable to click on logout button and detail error is ::"&err.description
	else
		LogResult micDone, "User clicks on Logout button" , "User logged out successfully"
	End If
	On error goto 0
End Function

Function funToAddProductsInCart()
	If  Browser("Advantage Shopping").Page("Advantage Shopping").Link("innertext:="&dictTestData("Instruments")).Exist = True Then
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("innertext:="&dictTestData("Instruments")).highlight
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("innertext:="&dictTestData("Instruments")).fnToClickOnLink
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("buy_now").fnToClickOnWebButton
		wait 3
		funToPerformActionBasedOnProperty dictTestData("Color"),"title"
		wait 3
		funToSelectProductQuantity dictTestData("Quantity")
		Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("save_to_cart").fnToClickOnWebButton
		wait 2
		If Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("check_out_btn").Exist = True Then
			Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("check_out_btn").fnToClickOnWebButton
		Else
			Browser("Advantage Shopping").Page("Advantage Shopping").Link("ShoppingCart").fnToClickOnLink
			funToPerformActionBasedOnProperty "checkOutButton","html id"
		End If
		wait 3
		LogResult micDone, "User should add "&dictTestData("Instruments") &" with color"&dictTestData("Color")&" and quantity "&dictTestData("Quantity") &" in cart successfully" , "User has added "&dictTestData("Instruments") &" with color"&dictTestData("Color")&" and quantity "&dictTestData("Quantity") &" in cart successfully"
		Browser("Advantage Shopping").Page("Advantage Shopping").Link("HOME").fnToClickOnLink
	Else
		LogResult micFail, "User should add "&dictTestData("Instruments") &" with color"&dictTestData("Color")&" and quantity "&dictTestData("Quantity") &" in cart successfully" , "The browser page is not visible"
	End If
End Function

Function funToPerformActionBasedOnProperty(sVar,sObjProperty)
	Dim oDesc,i
	Set oDesc = Description.Create
	oDesc(sObjProperty).value = Trim(Ucase(sVar))
	Set obj = Browser("Advantage Shopping").Page("Advantage Shopping").ChildObjects(oDesc)
	'msgbox obj.count
	For i = 0 to obj.Count -1			
   		sWebRadioBtn = obj(i).GetROProperty(sObjProperty) 
   		'print sWebRadioBtn
	   	If Trim(Ucase(sWebRadioBtn)) = Trim(Ucase(sVar)) Then
	   		LogResult micDone, "Click on "&sVar&" button ", "Clicked on "&sVar&" button "
	   		obj(i).Highlight
	   		obj(i).click
	   		Exit For
	   	else
	   		LogResult micFail, sVar&" button should Exist", sVar&" button does not Exist"
	   	End If
	Next
End Function

Function funToSelectProductQuantity(totalQuantityOrder)
	Dim sQuantityMax,sQuantityVal : sQuantityMax= 100
	sQuantityVal=Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("type:=text","name:=quantity","html tag:=INPUT").GetROProperty("value")
	If Cint(sQuantityVal) <> " " Then
		If cint(sQuantityVal) < cint(totalQuantityOrder) Then
			Do
				funToPerformActionBasedOnProperty "plus","class"
				sQuantityVal = cint(sQuantityVal)+1
			Loop until cint(totalQuantityOrder)=cint(sQuantityVal)  or sQuantityVal=sQuantityMax
		Else
			Do
				funToPerformActionBasedOnProperty "minus","class"
				sQuantityVal = cint(sQuantityVal)-1
			Loop until cint(totalQuantityOrder)=cint(sQuantityVal)  or sQuantityVal=sQuantityMax
		End If
		LogResult micDone, "User should add "&totalQuantityOrder&" in quanity field", "User has added "&totalQuantityOrder&" in quanity field"
	Else
		LogResult micFail, "User should update quantity", "User unable to update the quantity because quantity field is blank"
	End If
End Function

Public Function funToRegitseruserfunction()
	RegisterUserFunc "WebEdit", "fnSetValueInWebEdit", "fnSetValue"
	RegisterUserFunc "Link", "fnToClickOnLink", "fnClick"
	RegisterUserFunc "WebList", "fnToSelectValueFromWebList", "fnSetValueDropDwn"
	RegisterUserFunc "WebCheckBox", "fnSetValueInWebCheckBox", "fnSetValue"
	RegisterUserFunc "WebButton", "fnToClickOnWebButton", "fnClick"
End Function

Public Function funToDeRegitseruserfunction()
	UnRegisterUserFunc "WebEdit", "fnSetValueInWebEdit"
	UnRegisterUserFunc "Link", "fnToClickOnLink"
	UnRegisterUserFunc "WebList", "fnToSelectValueFromWebList"
	UnRegisterUserFunc "WebCheckBox", "fnSetValueInWebCheckBox"
	UnRegisterUserFunc "WebButton", "fnToClickOnWebButton"
End Function


'Function funToLogin(Parameter2)
'	funToLoopUntilExist (AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT"))
'	On error resume next
'	funToWaitExplicit(2)
'	'AIUtil.Context.Freeze
'	AIUtil.Context.SetBrowserScope(BrowserWindow)
'	AIUtil("text_box", "Username").Highlight
'	LogResult micDone, "User should enter username","User has entered username: "&Parameter2("UserName")
'	AIUtil("text_box", "Username").Type Parameter2("UserName")
'	funToWaitExplicit(1)
'	LogResult micDone, "User should enter password","User has entered password: "&crypt.Encrypt(Parameter2("UserName"))
'	AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Click
'	funToWaitExplicit(1)
'	AIUtil("text_box", "Password").Type Parameter2("Password")
'	LogResult micDone, "User should click on login button","User clicked on login button "
'	AIUtil("button", "Login").Click
'	If AIUtil.FindTextBlock("Dashboard", micFromTop, 1).Exist or Err.Number <> 0 Then
'		funToWaitExplicit(1)
'		LogResult micDone, "User should suceesfully logged in Orange HRM website","User logged in ornage HRM website"
'	Else
'		LogResult micFail, "User should suceesfully logged in Orange HRM website","User failed to logged in ornage HRM website error details :: "&err.description
'		'ExitTest
'	End If
'	On error goto 0
'	'AIUtil.Context.UnFreeze
'End Function
'
'
'Function funToSearchLeave(Parameter2)
'	On error resume next
'	If AIUtil.FindTextBlock("Leave").Exist and Err.Number = 0 Then
'		LogResult micDone, "User should click on leave","User clicked on leave"
'		AIUtil.FindTextBlock("Leave").Click
'		funToWaitExplicit(2)
'		LogResult micDone, "User should enter start date and end date in calender option","User entered start date and end date in calender option"
'		AIUtil("calendar", micAnyText, micFromLeft, 1).Click
'		If AIUtil.Calendar.Exist(2) Then	
'			AIUtil("text_box", "From Date").Click
'			funToDeleteValueFromWebedit
'			AIUtil("text_box", "From Date").Type Parameter2("StartDate")
'			AIUtil("text_box", "To Date").Click
'			funToDeleteValueFromWebedit
'			AIUtil("text_box", "To Date").Type Parameter2("EndDate")
'			LogResult micDone, "User should select rejected leave status","User selected rejected leave status"
'			AIUtil("combobox", "Show Leave with Status*").Click
'			AIUtil.FindTextBlock("Rejected").Click
'			AIUtil("toggle_button", "Include Past Employees").SetState "On"
'			AIUtil("button", "Search", micFromBottom, 1).Click
'			funToWaitExplicit(1)
'			If Err.number <> 0 Then
'				LogResult micFail, "User should search leave successfully","Functionality failed ::"& Err,description
'			Else
'				LogResult micDone, "User should search leave successfully","Leave searched successfully"
'			End If
'		Else
'			LogResult micFail, "User should search leave successfully","Leave failed to searched"
'			'ExitTest	
'		End If
'	Else
'		LogResult micFail, "User should search leave successfully","User unable to search leave "&err.description
'	End If
'	On Error goto 0
'	Err.Clear	
'End Function
'
'Function funToLoopUntilExist(strObj)
'	On error resume next
'	Dim intCount, inCountMax
'	inCount=0
'	inCountMax=20
'	Do
'		inCount = inCount +1
'	Loop until strObj.exist(1) or Cint(inCount) = Cint(inCountMax)
'	If Err.Number<> 0 Then
'		LogResult micFail, "Object exist", "Object ::"&strObj& " is not exist"
'	End If
'	On error goto 0
'	Set inCount = Nothing
'	Set inCountMax = Nothing
'End Function
'
'Function funToWaitExplicit(strTime)
'	wait strTime
'End Function
'
'
'Function funToCloseAllOpenBrowser()
'	On error resume next
'	Set brow = description.Create
'	brow("micclass").Value = "Browser"
'	num=Desktop.ChildObjects(brow).count
'	If num > 0 Then
'		For Iterator = num-1 To 0 Step -1
'			Browser("creationtime:="&Iterator).Close
'		Next
'	End If
'	On error goto 0
'End Function
'
'Function funToDeleteValueFromWebedit()
'	Dim myDeviceReplay
'	Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
'     	For i = 1 to 11
'            myDeviceReplay.PressKey 14
'      	Next
'End Function
'
'Function funToPageDown()
'	Dim myDeviceReplay
'	Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
'        myDeviceReplay.PressKey 209
'End Function
'
'Function funToWorkWithRegisterCustomClass(Parameter3)
'	funToLoopUntilExist (AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT"))
'	On error resume next
'	If AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Exist(2) and Err.Number = 0 Then
'		LogResult micDone, "User should enter username","User entered username"
'		AIUtil("text_box", "Username").Type Parameter3("UserName")
'		funToWaitExplicit(2)
'		AIUtil.FindTextBlock("QrangeHrRM OPEN SOURCE HR MANAGEMENT").Click
'		funToWaitExplicit(2)
'		LogResult micDone, "User should enter password","User entered password"
'		AIUtil("text_box", "Password").Type Parameter3("Password")
'		AIUtil("button", "LOGIN").Click
'		LogResult micDone, "User should click on Login button","User clicked on Login button"
'		'-----see image or object ---????
'		LogResult micDone, "User should register custom class using calss image","User has registered custom class using calss image"
'		AIUtil.RegisterCustomClass Parameter3("ClassName"), Parameter3("ClassImagePath")
'		
'		AIUtil(strClassName).Click
'		AIUtil(strClassName).Highlight
'		AIUtil.Scroll "down" , 1
'		AIUtil.FindTextBlock("Forgot your password?", micWithAnchorAbove, AIUtil(strClassName)).Highlight
'		AIUtil.Scroll "up", 1		
'		AIUtil.FindTextBlock("Login", micFromTop, 1).Highlight
'		AAIUtil.FindTextBlock("Login", micFromTop, 1).click 628, 318
'		AIUtil.FindTextBlock("Login", micFromTop, 1).DoubleClick 628, 460
'		
'		print "Text Details 1 ::"&AIUtil.FindTextBlock("Login", micFromTop, 1).GetText
'		print "Text Details 2 ::"& AIUtil.FindTextBlock("Login", micFromTop, 1).GetValue
'		AIUtil.FindTextBlock("Login", micFromTop, 1).RightClick
'		LogResult micDone, "User should scroll up, down, doble click, right click and extract the value from object ","User has scroll up, down, doble click, right click and extracted the value from object"	
'	Else
'		LogResult micFail, "User should be able to navigate on Login screen","User unable to navigate on Login screen "&err.description
'		'AIUtil("button", "LOGIN").Click
'	End If
'	On error goto 0
'	Err.Clear
'End Function
'
'Function funToViewDirectory()
'	On error resume next
'	AIUtil.ScrollOnObject AIUtil.FindTextBlock("SPECIAL OFFER", micFromBottom, 1),"down",4
'	If AIUtil("button", "SEE OFFER", micWithAnchorOnRight, AIUtil.FindTextBlock("Supremely thin, yet incredibly durable")).Exist and Err.Number = 0 Then
'		LogResult micDone, "User should click on SEE OFFER button","User clicked on SEE OFFER button"
'		AIUtil("button", "SEE OFFER", micWithAnchorOnRight, AIUtil.FindTextBlock("Supremely thin, yet incredibly durable")).LongClick 1.5
'		AIUtil("profile", micAnyText, micWithAnchorOnLeft, AIUtil("search")).Hover
'		AIUtil("help", micAnyText, micWithAnchorOnLeft, AIUtil("shopping_cart")).MultiClick 4
'		LogResult micDone, "User should perform longclick, mouse hover, multiclick and scroll on object successfully","User performed longclick, mouse hover, multiclick and scroll on object successfully"
'	Else
'		LogResult micFail, "User should navigate on SEE OFFER button","User failed to navigate on SEE OFFER button"	
'		'ExitTest
'	End If
'	On error goto 0
'	Err.Clear
'End Function
'
'Function funToValidateDifferentLangu()
'	On error resume next
'	If AIUtil.FindTextBlock("English-Chinese Dictionary").Exist = True Then
'		If AIUtil.RunSettings.OCR.GetConfigSet <> AI_OCR Then
'			AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
'		End  If
'		funToWaitExplicit(2)
'		AIUtil.Scroll "down" , 2
'		AIUtil.RunSettings.Ocr.Languages="en,zhs,zht"
'		LogResult micDone, "User should change OCR setting to add multiple languages","User has successfully changeed OCR setting to add multiple languages"	
'		If AIUtil.FindTextBlock("热门搜索").Exist = True Then
'			AIUtil.FindTextBlock("热门搜索").Highlight
'			strChinesText=AIUtil.FindTextBlock("热门搜索").GetValue
'			Print "String in Chinese ::" &strChinesText
'			LogResult micDone, "User should extract chinese string", "User has extracted chinese string"
'			AIUtil.FindTextBlock("English-Chinese Dictionary").Highlight
'			strEnglishText=AIUtil.FindTextBlock("English-Chinese Dictionary").GetValue
'			Print "String in English ::" &strEnglishText
'		Else
'			LogResult micFail, "User should extract chinese string","User failed to extract chinese string value"
'		End If
'	Else
'		LogResult micFail, "User should navigate on wordpress website","User failed to navigate on wordpress website"	
'	End If
'	On error goto 0
'	Err.Clear
'End Function
'
'Function funToWorkWithWebTable()
'	On error resume next
'	If AIUtil.FindTextBlock("SPEAKERS").Exist = True and Err.Number = 0 Then
'		AIUtil.FindTextBlock("SPEAKERS").Click
'		LogResult micDone, "User should click on SPEAKER link","User clicked on SPEAKER link"
'		AIUtil("button", "BUY NOW").Click
'		funToWaitExplicit(4)
'		LogResult micDone, "User should click on BUY NOW link","User clicked on BUY NOW link"
'		AIUtil.RunSettings.OCR.UseConfigSet AI_OCR
'		'AIUtil.Scroll "down", 2
'		funToPageDown
'		AIUtil.RunSettings.OCR.UseConfigSet UFT_OCR
'		funToWaitExplicit(3)
'		AIUtil.Table(micFromBottom,1).Highlight
'		funToWaitExplicit(3)
'		colCount=AIUtil.Table(micFromBottom,1).VisibleColumnCount
'		Print "Total No of Column count in Webtable:: " &AIUtil.Table(micFromBottom,1).VisibleColumnCount
'		rowCount=AIUtil.table(micFromBottom,1).VisibleRowCount
'		Print "Total No of Row count in Webtable:: " & AIUtil.table(micFromBottom,1).VisibleRowCount
'		If Cint(rowCount) <> 0 Then
'			LogResult micDone, "User should extract webtable total number of rows and columns","User has extracted webtable rows and columns"&" Row count ::"&rowCount& " Column count ::"&colCount
'			For Iterator = 0 To rowCount-1 Step 1
'				For Iterator1 = 0 To colCount-1 Step 1
'					AIUtil.Table(micFromBottom, 1).Cell(Iterator, Iterator1).Highlight
'					val1=AIUtil.Table(micFromBottom, 1).Cell(Iterator, Iterator1).GetText
'					print val1
'				Next
'			Next
'		Else 
'			LogResult micFail, "User should extract webtable total number of rows and columns","User failed to extract webtable rows and columns"
'		End If
'	Else
'		LogResult micFail, "User should click on SPEAKER link","User failed to click on SPEAKER link"
'		'ExitTest
'	End If
'	On Error goto 0
'	Err.Clear
'End Function
'
'
'
'
