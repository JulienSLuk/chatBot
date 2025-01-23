Sub VBAChatBot()
    Dim userInput As Variant ' Use Variant to handle both string and False
    Dim response As String
    
    ' Display available topics in a user-friendly format
    ShowAvailableTopics
    
    ' Get user input
    userInput = Application.InputBox("Enter a command from the available topics:", "Excel VBA Chatbot", Type:=2)
    
    ' Exit if the user cancels
    If userInput = False Or Trim(CStr(userInput)) = "" Then
        MsgBox "No command entered. Exiting chatbot.", vbExclamation, "Chatbot Exit"
        Exit Sub
    End If
    
    ' Get the chatbot's response
    response = GetChatbotResponse(CStr(userInput)) ' Ensure passing as string
    
    ' Display the response or run the associated macro
    If response <> "" Then
        MsgBox response, vbInformation, "Chatbot Response"
    Else
        MsgBox "Command not recognized. Please check the available topics.", vbExclamation, "Unrecognized Command"
    End If
End Sub

Private Sub ShowAvailableTopics()
    Dim topicsArr As Variant ' Explicitly declare as Variant to handle the array

    ' Predefined topics stored in an array
    topicsArr = Array( _
        "Append pd button", _
        "Cleanup Output button", _
        "dys1", _
        "D.NR" _
    )

    ' Display available topics in a message box
    MsgBox "Available Topics:" & vbCrLf & vbCrLf & Join(topicsArr, vbCrLf), vbInformation, "Topics List"
End Sub

Function GetChatbotResponse(userInput As String) As String
    Dim wsKB As Worksheet
    Dim wsChatLog As Worksheet
    Dim lastRow As Long
    Dim foundRow As Long
    
    ' Set worksheets
    On Error Resume Next
    Set wsKB = ThisWorkbook.Sheets("KnowledgeBase")
    Set wsChatLog = ThisWorkbook.Sheets("ChatLog")
    On Error GoTo 0
    
    ' Ensure KnowledgeBase exists
    If wsKB Is Nothing Then
        MsgBox "KnowledgeBase sheet not found. Please ensure the KnowledgeBase sheet exists.", vbCritical, "Error"
        Exit Function
    End If
    
    ' Ensure ChatLog exists, or create it
    If wsChatLog Is Nothing Then
        Set wsChatLog = ThisWorkbook.Sheets.Add
        wsChatLog.Name = "ChatLog"
        wsChatLog.Range("A1:B1").Value = Array("User Input", "Response")
    End If
    
    ' Search for the user input in the KnowledgeBase
    lastRow = wsKB.Cells(wsKB.Rows.Count, 1).End(xlUp).Row
    foundRow = 0
    On Error Resume Next
    foundRow = Application.Match(userInput, wsKB.Columns(1), 0)
    On Error GoTo 0
    
    If foundRow > 0 Then
        ' Log the query
        wsChatLog.Cells(wsChatLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = userInput
        wsChatLog.Cells(wsChatLog.Rows.Count, 2).End(xlUp).Offset(1, 0).Value = wsKB.Cells(foundRow, 2).Value
        
        ' Return the corresponding response
        GetChatbotResponse = wsKB.Cells(foundRow, 2).Value
        
        ' Run macro if specified
        If Left(GetChatbotResponse, 1) = "=" Then
            Application.Run Mid(GetChatbotResponse, 2) ' Remove '=' and run macro
            GetChatbotResponse = ""
        End If
    Else
        ' Log unrecognized command
        wsChatLog.Cells(wsChatLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = userInput
        wsChatLog.Cells(wsChatLog.Rows.Count, 2).End(xlUp).Offset(1, 0).Value = "Unrecognized Command"
        GetChatbotResponse = ""
    End If
End Function


