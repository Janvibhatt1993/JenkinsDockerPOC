
' ================================================================================================
'  NAME			: Initialization
'  DESCRIPTION 	  	: This function is used to create global variables which stores location 
'			   path of TestResult, TestData, Scripts, AppLib, Browser CommonLib & ObjectRepo 
'			  Loads common repository
'  PARAMETERS		: nil
' ================================================================================================
public strResultPath,strLibPath,strTestDataPath,htmlResultPath,strObjectRepositoryPath
funToInitializeEnvironmentVariable

Public Function funToInitializeEnvironmentVariable()
	On error resume next
	Dim sTestDir,arrPath,I
	sTestDir= Environment.Value ("TestDir")
	arrPath = Split (sTestDir, "\")
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
	Environment.value("strResultPath") = strResultPath
	Environment.value("strHtmlResultPath")=strResultPath&"DetailedHTMLResult\"
 	htmlResultPath = Environment.value("strHtmlResultPath")
 	Environment.value("strObjectRepoPath") = strLibPath&"ObjectRepository\"
 	strObjectRepositoryPath=Environment.value("strObjectRepoPath") & "Repository1.tsr"
 	 'Loading the repository file
	If  strObjectRepositoryPath <>  "" Then
		RepositoriesCollection.RemoveAll() 
		RepositoriesCollection.Add(strObjectRepositoryPath)  
	End If
	wait 3
	MyObjRepFiles = RepositoriesCollection.Count 
	'msgbox"repo count:::"&MyObjRepFiles
	'msgbox RepositoriesCollection.Item(1)
	Environment.Value("strFunctionLibraryPath") = strLibPath&"FunctionLibrary"
 	Environment.Value("strAppFunctionLibraryPath") = strLibPath&"FunctionLibrary\AppSpecific_FunctionLibrary\"
 	Environment.Value("strGenFunctionLibraryPath") = strLibPath&"FunctionLibrary\Generic_FunctionLibrary\"
 	
 	funToAddFunctionLibraryDynamically      
	On error goto 0
 	
End Function

Function funToAddFunctionLibraryDynamically()
	On error resume next
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set dir = fso.GetFolder(Environment.Value("strFunctionLibraryPath")) 'get num files in parent dir 
	total_files = dir.Files.Count 'get num files in each sub-sir 
	For Each sub_dir In dir.SubFolders 'get num files in each sub-sub-dir 
		If sub_dir.Name<>"" Then
			Set oFolder = fso.GetFolder(sub_dir)
			For Each oFile in oFolder.Files
				print oFile.Name
				splitVal=split (oFile.Name,".")
				If Ubound(splitVal) = 1 Then
					If Ucase(splitVal(1)) = Ucase("vbs") then
						print oFile.Path
						LoadFunctionLibrary oFile.Path
					End If
				End If
			Next
		End If
	Next 
	On error goto 0
End Function

Call funToTakeDataFromExcel(strResultPath,strLibPath,strTestDataPath,htmlResultPath,strObjectRepositoryPath)

'strToaddress = "janvi.bhatt@lntinfotech.com"
'strSubject = "Sample-Subject"
'strBody = "Sample-Body"
'strAttachPath = "C:\UFTAutomation\Results\TCStatus.txt"
'
'SendMailFromOutlook strToaddress,strSubject,strBody,strAttachPath
'
'Function SendMailFromOutlook(strToaddress,strSubject,strBody,strAttachPath)
'      Dim objOut, Objmail
'      Set objOut = CreateObject("Outlook.Application")
'      Set Objmail = objOut.CreateItem(0)        
'
'        With Objmail
'        .To = strToaddress
'        .Subject = strSubject
'        .Body = strBody
'        If strAttachPath = "" Then                  
'            .Attachments.Add(strAttachPath)
'        End If
'        .Save
'        .Send
'      End With
'    
'      'Clear the memory
'      Set objOut = Nothing
'      Set Objmail = Nothing
'End Function
'
'ZipAttach = ""
'
'    ' Declare the required variables
'    Dim objZip, objSA, objFolder, zipFile, FolderToZip
'    
'   FolderToZip = "C:\UFTAutomation\Results\DetailedHTMLResult"
'   'FolderToZip = strFoldertoZip
'    zipFile = FoldertoZip & ".zip"
'    
'    'Create the basis of a zip file.
'    CreateObject("Scripting.FileSystemObject") _
'    .CreateTextFile(zipFile, True) _
'    .Write "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
'
'    ' Create the object of the Shell
'    Set objSA = CreateObject("Shell.Application")
'   
'    ' Add the folder to the Zip
'    Set objZip = objSA.NameSpace(zipFile)
'    Set objFolder = objSA.NameSpace(FolderToZip)
'    objZip.CopyHere(objFolder.Items)
'
'    wait(10)
'
'    ZipAttach = zipFile
'    
'    
'Set objMail = CreateObject("CDO.Message")
'Set objConf = CreateObject("CDO.Configuration")
'Set objFlds = objConf.Fields
' 
''Set various parameters and properties of CDO object
'objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing Jump ") = 2 'cdoSendUsingPort
''your smtp server domain or IP address goes here such as smtp.yourdomain.com
'objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver Jump ") = "115.113.131.85"
'objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport Jump ") = 25 'default port for email
''uncomment next three lines if you need to use SMTP Authorization
''objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendusername Jump ") = "your-username"
''objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword Jump ") = "your-password"
''objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate Jump ") = 1 'cdoBasic
'objFlds.Update
'objMail.Configuration = objConf
'objMail.From = "janvi.bhatt@lntinfotech.com"
'objMail.To = "jjanvi.bhatt@lntinfotech.com"
'objMail.Subject = "Put your email's subject line here"
'objMail.TextBody = "Your email body content goes here"
'objMail.AddAttachment "C:\UFTAutomation\Results\DetailedHTMLResult.zip" 'Don't use = after AddAttachment, just provide the path
'objMail.Send
' 
''Set all objects to nothing after sending the email
'Set objFlds = Nothing
'Set objConf = Nothing
'Set objMail = Nothing
'
'
'Call Office365_Email_Test
'Sub Office365_Email_Test()
'    Dim objMessage, objConfig, fields
'    Set objMessage = CreateObject("CDO.Message")
'    Set objConfig =  CreateObject("CDO.Configuration")
'    Set fields = objConfig.fields
'    With fields
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "115.113.131.85"
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "janvi.bhatt@lntinfotech.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Reeva@2022"
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendtls") = True
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
'        .Update
'    End With
'    Set objMessage.Configuration = objConfig
'    
'    With objMessage
'        .Subject = "Test Message"
'        .From = "janvi.bhatt@lntinfotech.com"
'        .To = "janvi.bhatt@lntinfotech.com"
'        .HTMLBody = "Test Message"
'    End With
'    objMessage.Send
'    
'    '''''The transport failed to connect to the server.'''
'End Sub
'
'call Encrypt("janvibhatt")
'Public Function Encrypt(sString)
'	Dim sKey
'	Dim iLenKey, iKeyPos, iLenStr, i, sNewStr
'
'	sKey = "encryption"
'	sNewStr = ""
'	iLenKey = Len(sKey)
'	iKeyPos = 1
'	iLenStr = Len(sString)
'
'	sString = StrReverse(sString)
'	For i = 1 To iLenStr
'		sNewStr = sNewStr & chr(asc(Mid(sString, i, 1)) + Asc(Mid(sKey, iKeyPos, 1)))
'		iKeyPos = iKeyPos + 1
'		If iKeyPos > iLenKey Then iKeyPos = 1
'	Next
'	encrypt = sNewStr
'	
'	print sNewStr
'	valu = sNewStr
'	val = Decrypt(valu)
'	msgbox val
'End Function
'
'Public Function Decrypt(sString)
'	Dim sKey
'	Dim iLenKey, iKeyPos, iLenStr, i, sNewStr
'
'	sKey = "encryption"
'	sNewStr = ""
'	iLenKey = Len(sKey)
'	iKeyPos = 1
'	iLenStr = Len(sString)
'
'	sString=StrReverse(sString)
'	For i = iLenStr To 1 Step -1
'		sNewStr = sNewStr & chr(asc(Mid(sString, i, 1)) - Asc(Mid(sKey, iKeyPos, 1)))
'		iKeyPos = iKeyPos + 1
'		If iKeyPos > iLenKey Then iKeyPos = 1
'	Next
'	sNewStr=StrReverse(sNewStr)
'	Decrypt = sNewStr
'	print sNewStr
'End Function
'
''a="janvibhatt"
''Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("passwordRegisterPage").Set a
'
	
	
'Set fso = CreateObject("Scripting.FileSystemObject") 
'Set dir = fso.GetFolder("D:\Automation\UFTAutomation\Resources\FunctionLibrary") 'get num files in parent dir 
'total_files = dir.Files.Count 'get num files in each sub-sir 
'For Each sub_dir In dir.SubFolders 'get num files in each sub-sub-dir 
'	msgbox sub_dir.Name
'	If sub_dir.Name<>"" Then
'		Set oFolder = fso.GetFolder(sub_dir)
'		For Each oFile in oFolder.Files
'			'msgbox oFile.Name
'			print oFile.Name
'			splitVal=split (oFile.Name,".")
'			msgbox Ubound(splitVal)
'			If Ubound(splitVal) = 1 Then
'				If Ucase(splitVal(1)) = Ucase("vbs") then
'					Msgbox oFile.Name
'					Msgbox oFile.Path
'					LoadFunctionLibrary oFile.Path
'				End If
'				
'				'msgbox "pass"
'			End If
''			If UCase(oFile.Name) = UCase(sFilename) Then
''				sRetval = oFile.Path	' & "\" & sFilename
''			Exit For
'			'End If
'		Next
'	End If
'	For Each sub_sub_dir In sub_dir.SubFolders 
'		msgbox sub_sub_dir.Name
'		curr_dir_files = sub_sub_dir.Files.Count 
'		print curr_dir_files
'		total_files = total_files+curr_dir_files 
'		print total_files
'	Next 
'	curr_dir_files = sub_dir.Files.Count
'	total_files = total_files + curr_dir_files
'Next 
'MsgBox "Total Number Of Files In All Folders = "&total_files 


'Browser("Advantage Shopping").Page("Advantage Shopping").WebList("countryListboxRegisterPage").Set "India"


'-------------Execute with Chrome-----------------------
'----Scenario 1: Perform various AI methods like checkbox, toogle, calender etc using AI-----

