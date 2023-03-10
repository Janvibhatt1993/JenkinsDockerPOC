'
'RegisterUserFunc "WebEdit", "fnSetValueInWebEdit", "fnSetValue"
'RegisterUserFunc "Link", "fnToClickOnLink", "fnClick"
'RegisterUserFunc "WebList", "fnToSelectValueFromWebList", "fnSetValueDropDwn"
'RegisterUserFunc "WebCheckBox", "fnSetValueInWebCheckBox", "fnSetValue"
'RegisterUserFunc "WebButton", "fnToClickOnWebButton", "fnClick"

' ***********************************************************************************************
'
' 			U S E R    D E F I N E D   F U N C T I O N S 
'
' ***********************************************************************************************
'1. fnSetValue
'2. fnClick
'3. fnSetValueDropDwn

'=========================================================================================
'  NAME			 : fnSetValue
'  DESCRIPTION 	  	: This function set value in a text field
'  PARAMETERS		: objControl - Parameter to perform action
'					 strTestData - Dictionary value - Test data
' =========================================================================================
Public Function fnSetValue(objControl, strTestData)
	On Error Resume Next
    	If objControl.Exist(1) and Err.Number =  0 Then
	        LogResult micDone, objControl.ToString() &" should Exist" , objControl.ToString() &" found" 
	        If objControl.GetROProperty("visible") = True and strTestData <> " " Then
	            Reporter.ReportEvent micDone,objControl.ToString() &" should be visible",objControl.ToString() &" is not visible"
	        Else
	            Reporter.ReportEvent micFail, objControl.ToString() &" should be visible",objControl.ToString() &" is not visible"
	            LogResult micDone, objControl.ToString() &" should be visible" , objControl.ToString() &" is not visible"
	            Exit Function
	        End If
	        objControl.Set strTestData
	        Reporter.ReportEvent micDone,objControl.ToString() &" should Enter the data",objControl.ToString() &" was Entered the data"
	Else
	        Reporter.ReportEvent micFail,objControl.ToString() &"object should Exist",objControl.ToString() &" object does not exist"
	        LogResult micDone, objControl.ToString() &" object should Exist" , objControl.ToString() &" object does not exist"
    	End If
    	On error goto 0
End Function

'=========================================================================================
'  NAME			 : fnClick
'  DESCRIPTION 	  	: This function perform click operation
'  PARAMETERS		: objControl - Parameter to perform action
' =========================================================================================
Public Function fnClick(objControl)
	On Error Resume Next
    	If objControl.Exist(1) and Err.Number =  0 Then
        'Reporter.ReportEvent micGeneral,"TestObject Should Exist","TestObject object found"
	        LogResult micDone, objControl.ToString() &" should Exist" , objControl.ToString() &" found"
	        If objControl.GetROProperty("visible") = True and strTestData <> " " Then
	            Reporter.ReportEvent micDone,objControl.ToString() &" should be visible",objControl.ToString() &" is not visible"
	        Else
	            Reporter.ReportEvent micFail, objControl.ToString() &" should be visible",objControl.ToString() &" is not visible"
	            LogResult micFail, objControl.ToString() &" should be visible" , objControl.ToString() &" is not visible"
	            Exit Function
	        End If
	         objControl.Click
	        Reporter.ReportEvent micDone,objControl.ToString() &" should click",objControl.ToString() &" clicked successfully"
	    Else
	        Reporter.ReportEvent micFail,objControl.ToString() &"object should Exist",objControl.ToString() &" object does not exist"
	        LogResult micFail, objControl.ToString() &" object should Exist" , objControl.ToString() &"object does not exist"
    	End If
    	On error goto 0
End Function
'=========================================================================================
'  NAME			 : fnSetValueDropDwn
'  DESCRIPTION 	  	: This function select a value from drop down list
'  PARAMETERS		: objControl - Parameter to perform action
'					 strTestData - Dictionary value - Test data
' =========================================================================================
Public Function fnSetValueDropDwn(objControl, strTestData)

	On Error Resume Next
    	If objControl.Exist(1) and Err.Number =  0 Then
	        LogResult micDone, objControl.ToString() &" should Exist" , objControl.ToString() &" found"
	        If objControl.GetROProperty("visible") = True and strTestData <> " " Then
	            Reporter.ReportEvent micDone,objControl.ToString() &" should be visible",objControl.ToString() &" is not visible"
	        Else
	            Reporter.ReportEvent micFail, objControl.ToString() &" should be visible",objControl.ToString() &" is not visible"
	            LogResult micFail, objControl.ToString() &" should be visible" , objControl.ToString() &" is not visible"
	            Exit Function
	        End If
	        'objControl.Set strTestData
	        
	        objControl.Select strTestData
	        Reporter.ReportEvent micDone,objControl.ToString() &" should select value from dropdown",objControl.ToString() &" selected vlaue from dropdown successfully"  
	    Else
	    	Reporter.ReportEvent micDone,objControl.ToString() &" should select value from dropdown",objControl.ToString() &" not selected vlaue from dropdown successfully"  
	   	LogResult micFail, objControl.ToString() &" should select value from dropdown" , objControl.ToString() &" was unabvle to select the value from dropdown"
    	End If
End Function
'
'UnRegisterUserFunc "WebEdit", "fnSetValueInWebEdit"
'UnRegisterUserFunc "Link", "fnToClickOnLink"
'UnRegisterUserFunc "WebList", "fnToSelectValueFromWebList"
'UnRegisterUserFunc "WebCheckBox", "fnSetValueInWebCheckBox"
'UnRegisterUserFunc "WebButton", "fnToClickOnWebButton"

