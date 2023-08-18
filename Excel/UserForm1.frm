VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Digitalization Advisor"
   ClientHeight    =   9795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21630
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CreateHeatmap_Click()
    Dim Y As Integer
    Dim X As Integer
    Dim dataRange As Range
    Dim chartRange As Range
    Dim chartObj As ChartObject
    
    ' Create Heat Map
    For Y = 1 To 51
        For X = 1 To 13
            Worksheets("Heat Map").Cells(Y, X).Value = arr_landscape(Y, X)
        Next X
    Next Y

    ' Hide UserForm
    UserForm1.Hide
    
    ' Do the LLM
    'prompt = arr_landscape(1, 1) & arr_landscape(1, 4) & arr_landscape(1, 5) & arr_landscape(1, 6) & arr_landscape(1, 7) & arr_landscape(1, 8) & arr_landscape(1, 9) & arr_landscape(1, 10) & arr_landscape(1, 11) & arr_landscape(1, 12) & arr_landscape(1, 13)
    'Call OpenAI_Completion

    ' Show UserForm 2
    With UserForm2
        .Label1.Caption = arr_landscape(Hits, 1)
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub
Private Sub UpdateButton_Click()
    Dim X As Integer, Y As Integer
    Dim l As String
    Dim lines() As String
    Dim i As String
    Dim items() As String
    
    ' Updating array with changes
    lines = Split(TextBox.Value, ";")
    X = 1
    Y = 1
    For Each l In lines
        l = Replace(l, vbNewLine, "")
        items = Split(l, ", ")
        For Each i In items
            If X < 14 Then arr_landscape(Y, X) = i
            X = X + 1
        Next i
        X = 1
        Y = Y + 1
    Next l
End Sub
Sub SaveRangeAsPicture()
    ' Save a selected cell range as a JPG file to computer's desktop
    Dim cht As ChartObject
    Dim ActiveShape As Shape

    'Copy/Paste Cell Range as a Picture
    Selection.Copy
    ActiveSheet.Pictures.Paste(link:=False).Select
    Set ActiveShape = ActiveSheet.Shapes(ActiveWindow.Selection.Name)
    
    'Create a temporary chart object (same size as shape)
    Set cht = ActiveSheet.ChartObjects.Add(Left:=ActiveCell.Left, Width:=ActiveShape.Width, Top:=ActiveCell.Top, Height:=ActiveShape.Height)

    'Format temporary chart to have a transparent background
    cht.ShapeRange.Fill.Visible = msoFalse
    cht.ShapeRange.Line.Visible = msoFalse
        
    'Copy/Paste Shape inside temporary chart
    ActiveShape.Copy
    cht.Activate
    ActiveChart.Paste
    
    'Save chart to User's Desktop as PNG File
    cht.Chart.Export Environ("USERPROFILE") & "\Desktop\" & ActiveShape.Name & ".jpg"

    'Delete temporary Chart
    cht.Delete
    ActiveShape.Delete

    'Re-Select Shape (appears like nothing happened!)
    ActiveShape.Select
End Sub
Sub Pict(PicName)
    Dim Ans As String
    Dim cht As Excel.ChartObject
    Dim ActiveShape As Shape
    Dim Strpath As String
    
    Worksheets("Wallpaper").Shapes(PicName).Copy
    Set ActiveShape = Worksheets("Wallpaper").Shapes(PicName)
    Set cht = ActiveSheet.ChartObjects.Add(Left:=ActiveCell.Left, Width:=ActiveShape.Width, Top:=ActiveCell.Top, Height:=ActiveShape.Height)

    'Format temporary chart to have a transparent background
    cht.ShapeRange.Fill.Visible = msoFalse
    cht.ShapeRange.Line.Visible = msoFalse
    
    'Copy/Paste Shape inside temporary chart
    ActiveShape.Copy
    cht.Activate
    ActiveChart.Paste
    
    Strpath = ThisWorkbook.Path & "\Temp.jpg"
    cht.Chart.Export Strpath
    cht.Delete
    Set cht = Nothing
    
    ' Insert image
    Me.Image.Picture = LoadPicture(Strpath)
End Sub
Private Sub UserForm_Initialize()
    Dim Pic As Object
    Dim PicName As String
    
    For Each Pic In Sheets("Wallpaper").Pictures
        If TypeName(Pic) = "Picture" Then
            PicName = Pic.Name
        End If
    Next Pic

    Call Pict(PicName)
End Sub
Private Sub UserForm_Activate()
    Dim X As Integer, Y As Integer
    Dim Ans As String
    Dim keywords() As String
    Dim clr As Control
    
    ' Getting keywords
    sentence = Worksheets("Project keywords").Range("A2").Value
    sentence = Replace(sentence, ", Keywords: ", "")
    sentence = Replace(sentence, ".", "")
    keywords = Split(sentence, ", ")
    
    ' Merging array values to a string
    For Y = 1 To 51
        For X = 1 To 13
        If X < 13 Then
            Ans = Ans & arr_landscape(Y, X) & ", "
        Else
            Ans = Ans & arr_landscape(Y, X)
        End If
        Next X
        If Y < 51 Then Ans = Ans & "; " & vbNewLine
    Next Y
    
    ' UserForm
    With UserForm1
        .TextBox = Ans
        .TextBox.TextAlign = fmTextAlignCenter
        .TextBox.BackStyle = fmBackStyleTransparent
        .Keywords_Label = sentence
        .Keywords_Label.TextAlign = fmTextAlignCenter
        .Keywords_Label.BackStyle = fmBackStyleTransparent
        .Label1.BackStyle = fmBackStyleTransparent
        .Label2.BackStyle = fmBackStyleTransparent
        If Worksheets("Wallpaper").Range("A2").Value = "White" Then
            .TextBox.ForeColor = vbWhite
            .Keywords_Label.ForeColor = vbWhite
            .Label1.ForeColor = vbWhite
            .Label2.ForeColor = vbWhite
        End If
        If Worksheets("Wallpaper").Range("A2").Value = "Black" Then
            .TextBox.ForeColor = vbBlack
            .Keywords_Label.ForeColor = vbBlack
            .Label1.ForeColor = vbBlack
            .Label2.ForeColor = vbBlack
        End If
    End With
End Sub