VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12945
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
    Dim helpFile As String
    helpFile = ThisWorkbook.Path & Application.PathSeparator & "DigiAI.chm::/Html/about.htm" 'e.g. DigiAI.chm::/Html/Initiatives/Make-IT Initiative/index.htm
    Shell "HH " & helpFile, vbMaximizedFocus
End Sub
Private Sub CommandButton2_Click()
    'Call SendMail(text:= CStr(Left(output(2), len(output(2)) - 7)), recipients:= arr_landscape(Hits, 2))
    Call SendMail(text:= "[Mail text...]", recipients:= arr_landscape(Hits, 2))
End Sub
Private Sub Label1_Click()
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:= arr_landscape(Hits, 3), NewWindow:= True
    Unload Me
Exit Sub
NoCanDo:
    Debug.Print "Can't open " & arr_landscape(Hits, 3)
End Sub
Private Sub UserForm_Initialize()
    ' Load custom image
    Me.Image.Picture = LoadPicture(ThisWorkbook.Path & "\Temp.jpg")
End Sub
Private Sub UserForm_Activate()
    ' UserForm 2
    With UserForm2
        .Label1.BackStyle = fmBackStyleTransparent
        .Label2.BackStyle = fmBackStyleTransparent
        .Label3.BackStyle = fmBackStyleTransparent
        If Worksheets("Wallpaper").Range("A2").Value = "White" Then
            .Label2.ForeColor = vbWhite
            .Label3.ForeColor = vbWhite
        End If
        If Worksheets("Wallpaper").Range("A2").Value = "Black" Then
            .Label2.ForeColor = vbBlack
            .Label3.ForeColor = vbBlack
        End If
    End With
End Sub