Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
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
            "<button id='Button1' label='Date Picker' imageMso='DateAndTimeInsert' size='large' onAction='displayDatepicker' screentip='shows Datepicker'/>" +
            "<dialogBoxLauncher><button id='Button2' onAction='getInfo' screentip='Info about Date Picker'/></dialogBoxLauncher>" +
        "</group>" +
        "</tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function

#Disable Warning IDE0060 ' Hide not used Parameter warning as this is very often the case with the below callbacks from the ribbon

    ''' <summary>get Addin Info</summary>
    ''' <param name="control"></param>
    Public Sub getInfo(control As CustomUI.IRibbonControl)
        Dim sModule As String
        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            sModule = tModule.FileName
            If sModule.ToUpper.Contains("DATEPICKER") Then
                If MsgBox("Datepicker provides a replacement for the MSCOMCT2 Datepicker abandoned by Microsoft." + vbCrLf +
                    "You can either use it from the Ribbon to place a date value or a date range (if more than one cell was selected) into the selection" + vbCrLf +
                     "or within VBA Userforms to place Date Input Fields (see TestVBA.xlsm for demonstration code)." + vbCrLf + vbCrLf +
                    "Addin File: " + sModule + vbCrLf + "Built: " + FileDateTime(sModule).ToString(), MsgBoxStyle.Information) Then
                End If
            End If
        Next
    End Sub

    ''' <summary>display Datepicker</summary>
    ''' <param name="control"></param>
    Public Sub displayDatepicker(control As CustomUI.IRibbonControl)
        Dim currentSelection As Excel.Range = ExcelDnaUtil.Application.Selection
        If currentSelection Is Nothing Then Exit Sub
        Dim theDatepicker As DatePicker = New DatePicker()
        ' for multiple cells, enable date range with large calendar (year) otherwise only 1 month and single date
        If currentSelection.Cells.Count() > 1 Then
            theDatepicker.Calendar.SetCalendarDimensions(4, 3)
            theDatepicker.Calendar.MaxSelectionCount = 366
        Else
            theDatepicker.Calendar.SetCalendarDimensions(1, 1)
            theDatepicker.Calendar.MaxSelectionCount = 1
        End If
        If currentSelection.Rows.Count() > 1 AndAlso InStr(currentSelection.Cells(1, 1).NumberFormat, "mm") AndAlso InStr(currentSelection.Cells(2, 1).NumberFormat, "mm") Then
            theDatepicker.Calendar.SetSelectionRange(Date.FromOADate(currentSelection.Cells(1, 1).Value2), Date.FromOADate(currentSelection.Cells(2, 1).Value2))
        ElseIf currentSelection.Columns.Count() > 1 AndAlso InStr(currentSelection.Cells(1, 1).NumberFormat, "mm") AndAlso InStr(currentSelection.Cells(1, 2).NumberFormat, "mm") Then
            theDatepicker.Calendar.SetSelectionRange(Date.FromOADate(currentSelection.Cells(1, 1).Value2), Date.FromOADate(currentSelection.Cells(1, 2).Value2))
        ElseIf InStr(currentSelection.Cells(1, 1).NumberFormat, "mm") Then
            theDatepicker.Calendar.SetDate(Date.FromOADate(currentSelection.Cells(1, 1).Value2))
        End If
        theDatepicker.Calendar.ShowWeekNumbers = True
        theDatepicker.Calendar.ShowTodayCircle = False
        theDatepicker.ShowDialog()
        If currentSelection.Cells.Count() = 1 Then
            currentSelection.Value = theDatepicker.startDate
            currentSelection.Cells(1, 1).NumberFormat = "dd/mm/yyyy"
        ElseIf currentSelection.Rows.Count() > 1 Then
            currentSelection.Cells(1, 1).Value = theDatepicker.startDate
            currentSelection.Cells(1, 1).NumberFormat = "dd/mm/yyyy"
            currentSelection.Cells(2, 1).Value = theDatepicker.endDate
            currentSelection.Cells(2, 1).NumberFormat = "dd/mm/yyyy"
        ElseIf currentSelection.Columns.Count() > 1 Then
            currentSelection.Cells(1, 1).Value = theDatepicker.startDate
            currentSelection.Cells(1, 1).NumberFormat = "dd/mm/yyyy"
            currentSelection.Cells(1, 2).Value = theDatepicker.endDate
            currentSelection.Cells(1, 2).NumberFormat = "dd/mm/yyyy"
        End If
    End Sub

#Enable Warning IDE0060

End Class
