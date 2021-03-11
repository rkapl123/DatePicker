Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Integration
Imports System.Runtime.InteropServices

Public Module Menu
    Public theRibbon As CustomUI.IRibbonUI
    Public theMenuHandler As MenuHandler
End Module

''' <summary>handles all Ribbon related aspects</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As CustomUI.IRibbonUI)
        Menu.theRibbon = theRibbon
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        ' Ribbon definition XML
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='CustomTab' label='DatePicker'>" +
        "<group id='DatePicker' label='Date Picker'>" +
            "<button id='Button1' label='Date Picker' imageMso='' size='large' onAction='displayDatepicker' screentip='shows Datepicker'/>" +
            "<dialogBoxLauncher><button id='Button2' label='getInfo' onAction='getInfo' screentip='Info about Date Picker'/></dialogBoxLauncher>" +
        "</group>" +
        "</tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function

#Disable Warning IDE0060 ' Hide not used Parameter warning as this is very often the case with the below callbacks from the ribbon

    ''' <summary>get Addin Info</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getInfo(control As CustomUI.IRibbonControl) As String

    End Function

    ''' <summary>display Datepicker</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function displayDatepicker(control As CustomUI.IRibbonControl) As String

    End Function

#Enable Warning IDE0060

End Class
