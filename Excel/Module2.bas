Attribute VB_Name = "Module2"
Option Explicit
' Test routine
Sub Test()
    Dim X As Integer, Y As Integer, c As Integer
    Dim pro As String
    Dim word As Variant
    Dim words() As String

    ' Download Help File
    Call DownloadFileFromURL(FileUrl:= "https://www.benbox.org/R/DigiAI.chm")
    
    ' Create sheet
    If Not WorksheetExists("Heat Map") Then Sheets.Add.Name = "Heat Map"
    
    ' Hide sheets
    Worksheets("Project description").Visible = True
    Worksheets("Project keywords").Visible = True
    Worksheets("Digital landscape GIZ").Visible = True
    
    'Assigning Digital landscape values as array
    For Y = 1 To 51
        For X = 1 To 13
            arr_landscape(Y, X) = Worksheets("Digital landscape GIZ").Cells(Y + 1, X).Value
        Next X
        arr_landscape(Y, 14) = 0
    Next Y

    'Assigning keywords values as array
    pro = Worksheets("Project keywords").Range("A2").Value
    pro = Replace(pro, ".", "")
    project = Split(pro, ", ")
    
    ' Calculating Hits
    arr_len = UBound(project) - LBound(project) + 1
    For Y = 1 To 51
        For X = 1 To 13
            c = 0
            Do Until c = arr_len
                If Len(arr_landscape(Y, X)) > 0 Then
                    ' Split string into array of string
                    words = Split(arr_landscape(Y, X), " ")
                    For Each word In words
                        If InStr(1, project(c), word, vbTextCompare) > 0 Then
                            arr_landscape(Y, 14) = Int(arr_landscape(Y, 14)) + 1
                            'Debug.Print "InStr: " & " Project: " & project(c) & " landscape: " & arr_landscape(Y, X) & " " & arr_landscape(Y, 14)
                        End If
                    Next word
                End If
                c = c + 1
            Loop
        Next X
    Next Y
    
    ' Getting the initiative with the most keyword matches
    Hits = 1
    For Y = 1 To 51
        If Int(arr_landscape(Y, 14)) > Int(arr_landscape(Hits, 14)) Then
            Hits = Y
        End If
    Next Y
    Debug.Print "Position: " & Hits & " Initiative: " & arr_landscape(Hits, 1) & " Value: " & arr_landscape(Hits, 14)
    
    ' Create HeatMap
    

    ' Show UserForm
    With UserForm1
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        '.Show
    End With

    ' Test MsgBox
    Dim Result As Integer
    Dim helpFile As String
    helpFile = ThisWorkbook.Path & Application.PathSeparator & "DigiAI.chm"
    Debug.Print helpFile
    Shell "HH " & ThisWorkbook.Path & Application.PathSeparator & "DigiAI.chm", vbMaximizedFocus
    Result = MsgBox(Prompt:= "Should an Outlook template Mail be created with this text (draft):" & vbNewLine & vbNewLine & "Test" & vbNewLine & vbNewLine & "Check Information on Link: https://www.benbox.org", Buttons:= vbYesNo + vbQuestion + vbMsgBoxHelpButton, Title:= "Send a Mail?", HelpFile:= helpFile, Context:= 1012)
    If Result = vbYes Then
        Debug.Print("Yes")
    Else
        Debug.Print ("No Mail will be send")
    End If
End Sub
