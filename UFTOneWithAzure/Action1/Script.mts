Print "======== Hi All... we are from NAUT Test automation Team ========"
Systemutil.CloseProcessByName("chrome.exe")
Systemutil.Run "http://advantageonlineshopping.com/#/"
wait 5
AIUtil.SetContext Browser("creationtime:=0")
AIUtil.Context.SetBrowserScope(BrowserWindow)
wait 3
AIUtil.FindTextBlock("SPEAKERS").CheckExists True
AIUtil("search").Search "17t"
wait 3
If AIUtil.FindTextBlock("HP ENVY - 17t Touch", micFromBottom, 1).Exist = True Then
	AIUtil.FindTextBlock("HP ENVY - 17t Touch", micFromBottom, 1).Click
	AIUtil("plus", micAnyText, micFromBottom, 1).Click
	If AIUtil("button", "ADD TO CART").Exist = True then
		AIUtil("button", "ADD TO CART").Click
		wait 2
		AIUtil("shopping_cart").Click
		wait 2
		AIUtil.FindTextBlock("REMOVE").Click
		wait 2
		AIUtil.FindTextBlock("dvantageDEMO").Click
	End If
End If

