Imports System.Runtime.InteropServices
Imports System.Windows.Forms

<ClassInterface(ClassInterfaceType.AutoDispatch)>
<ProgId("DatePicker.DatePicker")>
<ComVisible(True)>
Public Class DatePicker
    Public startDate As Double
    Public endDate As Double
    Public Calendar As MonthCalendar

    Public Sub New()
        InitializeComponent()
        Calendar = theCalendar
    End Sub

    Private Sub OK_Click(sender As Object, e As EventArgs) Handles OK.Click
        startDate = theCalendar.SelectionStart.Date.ToOADate()
        endDate = theCalendar.SelectionEnd.Date.ToOADate()
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub
    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        startDate = 0
        endDate = 0
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DatePicker_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Width = Me.theCalendar.Width + 15
        Me.Height = Me.theCalendar.Height + 75
    End Sub

    Private lastClickTick As Integer = 0
    Private hasChanged As Boolean

    ''' <summary>emulate double click on date value</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub theCalendar_DateSelected(sender As Object, e As DateRangeEventArgs) Handles theCalendar.DateSelected
        If (Environment.TickCount - lastClickTick) <= SystemInformation.DoubleClickTime And Not hasChanged Then
            startDate = theCalendar.SelectionStart.Date.ToOADate()
            endDate = theCalendar.SelectionEnd.Date.ToOADate()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            lastClickTick = Environment.TickCount
        End If
        hasChanged = False
    End Sub

    Private Sub theCalendar_DateChanged(sender As Object, e As DateRangeEventArgs) Handles theCalendar.DateChanged
        lastClickTick = Environment.TickCount
        hasChanged = True
    End Sub

End Class