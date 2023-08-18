Attribute VB_Name = "Module1"
Option Explicit
Public bDeferredOpen As Boolean
Public OpenHandler As WorkbookOpenHandler
Public Const HANDLER_ENABLED = True
Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Const API_KEY As String = "K3iE2nh1Rex94RBUPj0wJFkblB3TS6XPtImSpgzGyscKcAwH"
Const API_ENDPOINT As String = "https://api.openai.com/v1/completions"
Const MODELL As String = "text-davinci-003"
Const MAX_TOKENS As String = "1024"
Const TEMPERATURE As String = "0.5"
Global arr_landscape(1 To 51, 1 To 14) As String
Global prompt As String
Global sentence As String
Global project() As String
Global Hits As Integer
Global arr_len As Integer
Global output As Variant
Function RangeToString(ByVal myRange As Range) As String
    RangeToString = ""
    If Not myRange Is Nothing Then
        Dim myCell As Range
        For Each myCell In myRange
            RangeToString = RangeToString & "," & myCell.Value
        Next myCell
        
        ' Remove extra comma
        RangeToString = Right(RangeToString, Len(RangeToString) - 1)
    End If
End Function
Sub OpenAI_Completion()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' Check if API key is available
    If API_KEY = "<API_KEY>" Then
        MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Check if there is anything in the prompt
    If Trim(prompt) <> "" Then
    ' Clean prompt to avoid parsing error in JSON payload
    prompt = CleanJSONString(prompt)
    Else
        MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Show status in status bar
    Application.StatusBar = "Processing OpenAI request..."

    ' Create XMLHTTP object
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")

    ' Define request body
    Dim requestBody As String
    Dim command As String
    prompt = "Please formulize a summary out of these keywords: " & prompt
    requestBody = "{" & """model"": """ & MODELL & """," & """prompt"": """ & prompt & """," & """max_tokens"": " & MAX_TOKENS & "," & """temperature"": " & TEMPERATURE & "}"
              
    ' Open and send the HTTP request
    With httpRequest
        .Open "POST", API_ENDPOINT, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & "sk-" & StrReverse(API_KEY)
        .send (requestBody)
    End With

    ' Check if the request is successful
    If httpRequest.Status = 200 Then
        ' Parse the JSON response
        Dim response As String
        response = httpRequest.responseText

        ' Get the completion and clean it up
        Dim completion As String
        completion = ParseResponse(response)
        
        ' Split the completion into lines
        output = Split(completion, "\n")

        ' MsgBox to aks for Mail template creation
        Dim Result As Integer
        Dim objShell As Object
        Set objShell = CreateObject("Wscript.Shell")
        With UserForm2
            .StartUpPosition = 0
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
            .Show
        End With
    Else
        MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
    End If
          
    Application.StatusBar = False
    Application.ScreenUpdating = True      
    Exit Sub
          
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
' Helper function to check if worksheet exists
Function WorksheetExists(worksheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
    On Error GoTo 0
End Function
' Helper function to parse the reponse text
Function ParseResponse(ByVal response As String) As String
    On Error Resume Next
    Dim startIndex As Long
    startIndex = InStr(response, """text"":""") + 8
    Dim endIndex As Long
    endIndex = InStr(response, """index"":") - 2
    ParseResponse = Mid(response, startIndex, endIndex - startIndex)
    On Error GoTo 0
End Function
' Helper function to clean text
Function CleanJSONString(inputStr As String) As String
    On Error Resume Next
    ' Remove line breaks
    CleanJSONString = Replace(inputStr, vbCrLf, "")
    CleanJSONString = Replace(CleanJSONString, vbCr, "")
    CleanJSONString = Replace(CleanJSONString, vbLf, "")

    ' Replace all double quotes with single quotes
    CleanJSONString = Replace(CleanJSONString, """", "'")
    On Error GoTo 0
End Function
' Replaces the backslash character only if it is immediately followed by a double quote.
Function ReplaceBackslash(text As Variant) As String
    On Error Resume Next
    Dim i As Integer
    Dim newText As String
    newText = ""
    For i = 1 To Len(text)
        If Mid(text, i, 2) = "\" & Chr(34) Then
            newText = newText & Chr(34)
            i = i + 1
        Else
            newText = newText & Mid(text, i, 1)
        End If
    Next i
    ReplaceBackslash = newText
    On Error GoTo 0
End Function
Sub SendMail(text As String, recipients As String)
    ' Send an outlook mail
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    
    ' Create a new instance of Outlook
    Set olApp = New Outlook.Application
    
    ' Create a new email message
    Set olMail = olApp.CreateItem(olMailItem)
    
    ' Set the email properties
    With olMail
        .To = recipients
        .Subject = "Request for help on a Digitalization project"
        .Body = "Dear" & vbNewLine & vbNewLine & text & vbNewLine & vbNewLine & "Best regards" & vbNewLine & vbNewLine & "Your Digitalization Team"
        .Attachments.Add ThisWorkbook.Path & Application.PathSeparator & "Temp.jpg"
        .Display
        '.Send
    End With
    
    ' Clean up
    Set olMail = Nothing
    Set olApp = Nothing
End Sub
Sub ClearCached()
Application.ScreenUpdating = False
    Dim originalSetting As Long

    Shell "RunDll32.exe InetCpl.Cpl, ClearMyTracksByProcess 2"

    'Browsing_History
    Shell "RunDll32.exe InetCpl.Cpl, ClearMyTracksByProcess 8"
    originalSetting = Application.RecentFiles.Maximum
    Application.RecentFiles.Maximum = 0
    Application.RecentFiles.Maximum = originalSetting
    Debug.Print "Cache cleared"
End Sub
Sub DownloadFileFromURL(fileUrl As String)
    Dim objXmlHttpReq As Object
    Dim objStream As Object

    ' Clear cache
    Call ClearCached()

    ' Print url
    Debug.Print FileUrl

    Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
    objXmlHttpReq.Open "GET", FileUrl, False, "username", "password"
    objXmlHttpReq.send

    If objXmlHttpReq.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
        objStream.Write objXmlHttpReq.responseBody
        objStream.SaveToFile ThisWorkbook.Path & Application.PathSeparator & "DigiAI.chm", 2 ' 2 = overwriting existing files
        objStream.Close
        Debug.Print "File downloaded"
    Else
        Debug.Print "Error: " & objXmlHttpReq.Status & " " & objXmlHttpReq.statusText
    End If
End Sub