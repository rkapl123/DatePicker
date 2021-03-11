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
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
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
        Me.Width = Me.theCalendar.Width + 10
        Me.Height = Me.theCalendar.Height + 60
    End Sub
End Class