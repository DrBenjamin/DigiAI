Attribute VB_Name = "Module2"
Option Explicit
' Test routine
Sub Test()
    Dim X As Integer, Y As Integer, c As Integer
    Dim pro As String
    Dim word As Variant
    Dim words() As String
    Dim coeff As Double
    Dim col As Integer
    Dim row As Integer

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
    For Y = 1 To 51
        col = Int(Y Mod 2) + 1
        row = Int((Y + 1) / 2)
        With Worksheets("Heat Map") 
            .Hyperlinks.Add Anchor:= .Cells(row, col), Address:= arr_landscape(Y, 3), ScreenTip:= arr_landscape(Y, 1), TextToDisplay:= arr_landscape(Y, 1)
            .Cells(row, col).ColumnWidth = 80
        End With

        If arr_landscape(Hits, 14) > 20 Then 
            coeff = arr_landscape(Hits, 14) / 20
        Else
            coeff = 1.0
        End If

        If arr_landscape(Y, 14) > 0 and arr_landscape(Y, 14) <= (2 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 255, 204)
        ElseIf arr_landscape(Y, 14) > (2 * coeff) and arr_landscape(Y, 14) <= (4 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 255, 153)
        ElseIf arr_landscape(Y, 14) > (4 * coeff) and arr_landscape(Y, 14) <= (6 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 255, 102)
        ElseIf arr_landscape(Y, 14) > (6 * coeff) and arr_landscape(Y, 14) <= (8 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 255, 51)
        ElseIf arr_landscape(Y, 14) > (8 * coeff) and arr_landscape(Y, 14) <= (10 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 255, 0)
        ElseIf arr_landscape(Y, 14) > (10 * coeff) and arr_landscape(Y, 14) <= (12 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 204, 0)
        ElseIf arr_landscape(Y, 14) > (12 * coeff) and arr_landscape(Y, 14) <= (14 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 153, 0)
        ElseIf arr_landscape(Y, 14) > (14 * coeff) and arr_landscape(Y, 14) <= (16 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 102, 0)
        ElseIf arr_landscape(Y, 14) > (16 * coeff) and arr_landscape(Y, 14) <= (18 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 51, 0)
        ElseIf arr_landscape(Y, 14) > (18 * coeff) Then
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 0, 0)
        Else
            Worksheets("Heat Map").Cells(row, col).Interior.Color = RGB(255, 255, 255)
        End If
    Next Y
    With Worksheets("Heat Map").Range("A1:B" & row).Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With

    ' Show UserForm
    With UserForm2
        .Label1.Caption = arr_landscape(Hits, 1)
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub
