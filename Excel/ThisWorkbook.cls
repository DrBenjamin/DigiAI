Option Explicit
Private Sub Workbook_Open()
    Dim X As Integer, Y As Integer
    Dim pro As String
    Dim word As Variant
    Dim words() As String
    Dim c As Integer
    
    ' Create sheet
    If Not WorksheetExists("Heat Map") Then Sheets.Add.Name = "Heat Map"
    
    ' Hide sheets
    Worksheets("Project description").Visible = False
    Worksheets("Project keywords").Visible = False
    Worksheets("Digital landscape GIZ").Visible = False
    Worksheets("Wallpaper").Visible = False
    
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
                            Debug.Print "InStr: " & " Project: " & project(c) & " landscape: " & arr_landscape(Y, X) & " " & arr_landscape(Y, 14)
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

    ' Show UserForm
    With UserForm1
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub
