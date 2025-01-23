Sub VBAChatBot()
    Dim userInput As String
    Dim response As String
    Dim promptText As String
    Dim topics As String
    Dim result As Variant

    ' Predefined topics for dropdown
    topics = "Cleanup Output,Copy and Format Dates,Check Whitelist,Find Missing NRIC,Append Data,NRIC Dynamic Fill"
    
    ' Welcome message
    MsgBox "Welcome to the Excel VBA Chatbot!" & vbNewLine & _
           "You can ask about various macros or type custom questions." & vbNewLine & _
           "For common topics, use the dropdown for guidance.", vbInformation, "Chatbot Introduction"
    
    ' Prompt user with dropdown options
    result = Application.InputBox("Ask a question about the macros or select a topic:" & vbNewLine & topics, _
                                  "VBA Chatbot", Type:=2)
    
    ' Handle Cancel button or empty input
    If VarType(result) = vbBoolean Then
        If result = False Then Exit Sub ' Exit if Cancel
    End If
    
    If Trim(result) = "" Then
        MsgBox "Please enter a valid question.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Cast result explicitly to String
    userInput = CStr(result)
    
    ' Process the input and generate a response
    response = GetChatbotResponse(userInput)
    
    ' Display the response
    If response <> "" Then
        MsgBox response, vbInformation, "Chatbot Response"
    Else
        MsgBox "I'm sorry, I couldn't understand your question. Try selecting a topic or rephrasing.", vbExclamation, "Chatbot Response"
    End If
    
    ' Log the chat history
    Call LogChatHistory(userInput, response)
End Sub

Function GetChatbotResponse(userInput As String) As String
    Dim wsKnowledgeBase As Worksheet
    Dim lastRow As Long
    Dim commandRange As Range
    Dim matchRow As Variant
    
    ' Set the KnowledgeBase sheet
    Set wsKnowledgeBase = ThisWorkbook.Sheets("KnowledgeBase")
    
    ' Find the last row in the KnowledgeBase sheet
    lastRow = wsKnowledgeBase.Cells(wsKnowledgeBase.Rows.Count, 1).End(xlUp).Row
    
    ' Set the range for commands (Column A)
    Set commandRange = wsKnowledgeBase.Range("A2:A" & lastRow)
    
    ' Look for a match with the user input
    matchRow = Application.Match(LCase(userInput), commandRange, 0)
    
    If Not IsError(matchRow) Then
        ' If a match is found, return the response from Column B
        GetChatbotResponse = wsKnowledgeBase.Cells(matchRow + 1, 2).Value
    Else
        ' If no match is found, return a default response
        GetChatbotResponse = "I'm sorry, I don't understand that command. Please check the KnowledgeBase sheet."
    End If
End Function


Sub LogChatHistory(question As String, response As String)
    Dim wsLog As Worksheet
    Dim lastRow As Long

    ' Check if the ChatLog sheet exists, if not create it
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("ChatLog")
    On Error GoTo 0
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "ChatLog"
        wsLog.Range("A1:B1").Value = Array("Question", "Response")
    End If

    ' Log the chat history
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(lastRow, 1).Value = question
    wsLog.Cells(lastRow, 2).Value = response
End Sub


