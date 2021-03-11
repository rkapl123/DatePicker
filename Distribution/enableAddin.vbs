On Error Resume Next
Set XLApp = GetObject(,"Excel.Application")
If err <> 0 Then
	Set XLApp = CreateObject("Excel.Application")
	WScript.Sleep 1000
End If
On Error goto 0
XLApp.Visible = true
For each ai in XLApp.AddIns
	If ai.name = "DatePicker.xll" then
		ai.Installed = True
	end if
next
Wscript.Echo ("Please restart Excel to make Installation effective ...")