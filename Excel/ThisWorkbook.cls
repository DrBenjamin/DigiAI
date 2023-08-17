Option Explicit
Private Sub Workbook_Open() 
    '  Workaround when Macro is opened in protected view
    If HANDLER_ENABLED Then
        Set OpenHandler = New WorkbookOpenHandler
        Set OpenHandler.oApp = Application
    Else
        ActiveWindow.DisplayHeadings = False
    End If
End Sub
