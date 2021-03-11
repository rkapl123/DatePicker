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

End Class

