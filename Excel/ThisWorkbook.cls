Option Explicit
Private Sub Workbook_Open() 

    '  Workaround when Macro is opened in protected view
    Set OpenHandler = New WorkbookOpenHandler
    Set OpenHandler.oApp = Application

End Sub
