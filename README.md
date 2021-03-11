[![License](https://img.shields.io/github/license/rkapl123/DatePicker.svg)](https://github.com/rkapl123/DatePicker/blob/master/LICENSE)

# DatePicker

A Datepicker independent of MSCOMCT2, also for 64bit Office.

You can either use it from the Ribbon to place a date value or a date range (if more than one cell was selected a 1 Year calendar is displayed) into the selection or within VBA Userforms to place Date Input Fields (see also TestVBA.xlsm for demonstration code).

Selection of a Date is either done with OK or by double clicking a date in the calendar widget.

Adding to VBA works by setting a label field and utilizing the click event on it to display the associated DatePicker created with CreateObject("DatePicker.DatePicker").
The internals of the VB.NET [MonthCalendar Widget](https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.monthcalendar) are passed through in member variable Calendar and can be set accordingly before the ShowDialog (avoid method Show as it won't be modal and close immediately).
The results are returned in member variables StartDate and EndDate, they are passed in Excel's internal Julian date format (of type double).

Following code is placed within a VBA Userform:

```vb
Dim theDatepicker2 As Object

' single Datefield setting
Private Sub Datefield_Click()
    Dim theDatepicker As Object
    Set theDatepicker = CreateObject("DatePicker.DatePicker")
    theDatepicker.Calendar.MaxSelectionCount = 1
    theDatepicker.ShowDialog
    Me.Datefield.Caption = Format(theDatepicker.StartDate, "dd.mm.yyyy")
End Sub

' Datefield range setting, first date field
Private Sub Label1_Click()
    theDatepicker2.ShowDialog
    Me.Label1.Caption = Format(theDatepicker2.StartDate, "dd.mm.yyyy")
    Me.Label2.Caption = Format(theDatepicker2.EndDate, "dd.mm.yyyy")
End Sub

' Datefield range setting, second date field
Private Sub Label2_Click()
    theDatepicker2.ShowDialog
    Me.Label1.Caption = Format(theDatepicker2.StartDate, "dd.mm.yyyy")
    Me.Label2.Caption = Format(theDatepicker2.EndDate, "dd.mm.yyyy")
End Sub

' close userform
Private Sub OKButton_Click()
    Me.Hide
End Sub

' for Datefield ranges, need to initialize the datepicker before fields are clicked
Private Sub UserForm_Initialize()
    Set theDatepicker2 = CreateObject("DatePicker.DatePicker")
    theDatepicker2.Calendar.SetCalendarDimensions 4, 3
    theDatepicker2.Calendar.MaxSelectionCount = 366
End Sub

```

# Install

To install Datepicker, unzip the latest release and run deployAddin.cmd in the Distribution Folder. This copies the 32 or 64bit version of the Addin to your %appdata%\Microsoft\AddIns\ folder and enables the Ribbon inside Excel by running enableAddin.vbs