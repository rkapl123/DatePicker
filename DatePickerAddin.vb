Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports ExcelDna.ComInterop

''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
        ComServer.DllRegisterServer()
        ' Ribbon setup
        Menu.theMenuHandler = New MenuHandler
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Menu.theMenuHandler = Nothing
        ComServer.DllUnregisterServer()
    End Sub

    ''' <summary>open workbook: update Ribbon</summary>
    ''' <param name="Wb"></param>
    Private Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles Application.WorkbookOpen
        If Not Wb.IsAddin Then Menu.theRibbon.Invalidate()
    End Sub

    ''' <summary>WorkbookActivate: update Ribbon</summary>
    Private Sub Application_WorkbookActivate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        Menu.theRibbon.Invalidate()
    End Sub

End Class

