Attribute VB_Name = "Module1"
Option Explicit
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
10        On Error GoTo ErrorHandler
20        Application.ScreenUpdating = False

          ' Check if API key is available
30        If API_KEY = "<API_KEY>" Then
40            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
50            Application.ScreenUpdating = True
60            Exit Sub
70        End If

          ' Check if there is anything in the prompt
90        If Trim(prompt) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
100           prompt = CleanJSONString(prompt)
110       Else
120           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
130           Application.ScreenUpdating = True
140           Exit Sub
150       End If

          ' Show status in status bar
200       Application.StatusBar = "Processing OpenAI request..."

          ' Create XMLHTTP object
          Dim httpRequest As Object
210       Set httpRequest = CreateObject("MSXML2.XMLHTTP")

          ' Define request body
          Dim requestBody As String
          Dim command As String
          prompt = "Please formulize a summary out of these keywords: " & prompt
220       requestBody = "{" & _
              """model"": """ & MODELL & """," & _
              """prompt"": """ & prompt & """," & _
              """max_tokens"": " & MAX_TOKENS & "," & _
              """temperature"": " & TEMPERATURE & _
              "}"
              
          ' Open and send the HTTP request
230       With httpRequest
240           .Open "POST", API_ENDPOINT, False
250           .SetRequestHeader "Content-Type", "application/json"
260           .SetRequestHeader "Authorization", "Bearer " & "sk-" & StrReverse(API_KEY)
270           .send (requestBody)
280       End With

          ' Check if the request is successful
290       If httpRequest.Status = 200 Then
              ' Parse the JSON response
              Dim response As String
300           response = httpRequest.responseText

              ' Get the completion and clean it up
              Dim completion As String
310           completion = ParseResponse(response)
              
              ' Split the completion into lines
              Dim lines As Variant
320           lines = Split(completion, "\n")

              ' Output the lines
              Dim i As Long
330           For i = LBound(lines) To UBound(lines)
331               If i > 1 Then
340                 MsgBox lines(i)
342               End If
350           Next i
              
430       Else
440           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
450       End If
          
460       Application.StatusBar = False
470       Application.ScreenUpdating = True
          
480       Exit Sub
          
ErrorHandler:
490       MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
500       Application.StatusBar = False
510       Application.ScreenUpdating = True
End Sub
' Helper function to check if worksheet exists
Function WorksheetExists(worksheetName As String) As Boolean
520       On Error Resume Next
530       WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
540       On Error GoTo 0
End Function
' Helper function to parse the reponse text
Function ParseResponse(ByVal response As String) As String
550       On Error Resume Next
          Dim startIndex As Long
560       startIndex = InStr(response, """text"":""") + 8
          Dim endIndex As Long
570       endIndex = InStr(response, """index"":") - 2
580       ParseResponse = Mid(response, startIndex, endIndex - startIndex)
590       On Error GoTo 0
End Function
' Helper function to clean text
Function CleanJSONString(inputStr As String) As String
600       On Error Resume Next
          ' Remove line breaks
610       CleanJSONString = Replace(inputStr, vbCrLf, "")
620       CleanJSONString = Replace(CleanJSONString, vbCr, "")
630       CleanJSONString = Replace(CleanJSONString, vbLf, "")

          ' Replace all double quotes with single quotes
640       CleanJSONString = Replace(CleanJSONString, """", "'")
650       On Error GoTo 0
End Function
' Replaces the backslash character only if it is immediately followed by a double quote.
Function ReplaceBackslash(text As Variant) As String
660       On Error Resume Next
          Dim i As Integer
          Dim newText As String
670       newText = ""
680       For i = 1 To Len(text)
690           If Mid(text, i, 2) = "\" & Chr(34) Then
700               newText = newText & Chr(34)
710               i = i + 1
720           Else
730               newText = newText & Mid(text, i, 1)
740           End If
750       Next i
760       ReplaceBackslash = newText
770       On Error GoTo 0
End Function
