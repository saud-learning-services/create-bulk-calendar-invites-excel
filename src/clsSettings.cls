VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member Variables
Private m_senderName As String
Private m_sendImmediate As Boolean
Private m_isDebugging As Boolean
Private m_firstSendRow As Integer
Private m_lastSendRow As Integer
Private m_errMsg As String
Private m_errNum As Long

Private m_logger As clsLogger

'Event - Initialize
Private Sub Class_Initialize()
    m_senderName = "Tester"
    m_sendImmediate = False
    m_isDebugging = True
    m_firstSendRow = 2
    m_lastSendRow = 2
    m_errNum = 42002
    m_errMsg = "An error occurred while using property, method or function in" _
        & " class module 'clsSettings'"
End Sub

Public Sub ConstructSettings( _
    ExamSheet As clsExamSheet, Logger As clsLogger)
    Askers ExamSheet, Logger
End Sub


'Getter - senderName
Property Get SenderName() As String
    SenderName = m_senderName
End Property

'Getter - sendImmediate
Property Get SendImmediate() As Boolean
    SendImmediate = m_sendImmediate
End Property

'Getter - isDebugging
Property Get isDebugging() As Boolean
    isDebugging = m_isDebugging
End Property

'Getter - firstSendRow
Property Get FirstSendRow() As Integer
    FirstSendRow = m_firstSendRow
End Property

'Getter - lastSendRow
Property Get LastSendRow() As Integer
    LastSendRow = m_lastSendRow
End Property

'Method - Askers - asks for settings
'IDEA - maybe pass the column slice of courses so users can preview which courses are affected
Public Sub Askers(ExamSheet As clsExamSheet, Logger As clsLogger)
    Set m_logger = Logger
    Dim InputStr As String: InputStr = ""
    Dim Asker As String: Asker = ""

    Dim Confirmed As Boolean: Confirmed = False
    Do While Confirmed = False
        m_logger.LogInfo "Start asking for settings"
        Asker = "Please enter your name (for record keeping):"
        InputStr = InputAsker(Asker, "Enter Name")
        If InputStr = "_qds" Then
            GoTo QuickDebugSettings
        End If
        m_senderName = InputStr

        InputStr = ""
        Do While Not (InputOneOf(InputStr, "yes", "no"))
            Asker = _
                "Are you sending / drafting actual invites?" & _
                vbNewLine & vbNewLine & "Enter:" & vbNewLine & _
                "'yes' - sending / drafting invites for actual exams" & vbNewLine & _
                "'no' - testing / debugging only" & vbNewLine & vbNewLine & _
                "(Enter answer exactly without quotes)"
            InputStr = InputAsker(Asker, "Check debugging")
        Loop
        If InputStr = "yes" Then
            m_isDebugging = False
        Else
            m_isDebugging = True
        End If

        InputStr = ""
        Do While Not (InputOneOf(InputStr, "send", "draft"))
            Asker = _
                "Send invites immediately or save as draft?" & _
                vbNewLine & vbNewLine & "Enter:" & vbNewLine & _
                "'send' - send invites immediately" & vbNewLine & _
                "'draft' - save invite drafts to send later" & vbNewLine & vbNewLine & _
                "(Enter answer without quotes)"
            InputStr = InputAsker(Asker, "Send or Draft")
        Loop
        If InputStr = "send" Then
            m_sendImmediate = True
        Else
            m_sendImmediate = False
        End If

        InputStr = ""
        Do While Not (InputOneOf(InputStr, "selected", "assign"))
            Asker = _
                "Send invites for exam(s) in selected range or" & _
                " assign range manually?" & _
                vbNewLine & vbNewLine & "Enter:" & vbNewLine & _
                "'selected' - send / draft invites for exams in selected range" & vbNewLine & _
                "'assign' - assign which rows to send / draft invites" & vbNewLine & vbNewLine & _
                "(Enter answer without quotes)"
            InputStr = InputAsker(Asker, "Selected or assigned range")
        Loop
        If InputStr = "selected" Then
            m_firstSendRow = Application.Selection.row
            m_lastSendRow = m_firstSendRow + Application.Selection.Rows.Count - 1
        Else
            InputStr = "0"
            Do Until IsNumeric(InputStr) And CInt(InputStr) > 1
                Asker = _
                    "Enter row number of the FIRST EXAM to send / draft invite." & _
                    vbNewLine & vbNewLine & "Please enter FIRST EXAM's row as an integer:"
                InputStr = InputAsker(Asker, "First Exam Row Number")
            Loop
            m_firstSendRow = CInt(InputStr)

            InputStr = "0"
            Do Until IsNumeric(InputStr) And CInt(InputStr) >= m_lastSendRow
                Asker = _
                    "Enter row number of the LAST EXAM to send / draft invite." & _
                    vbNewLine & vbNewLine & "Please enter LAST EXAM's row as an integer:"
                InputStr = InputAsker(Asker, "Last Exam Row Number")
            Loop
            m_lastSendRow = CInt(InputStr)
        End If
        
        Asker = _
            "Proceed with following settings?" & vbNewLine & vbNewLine & _
            "Sender Name: " & m_senderName & vbNewLine & _
            "Send invites immediately: " & m_sendImmediate & vbNewLine & _
            "Debugging: " & m_isDebugging & vbNewLine & _
            "First row to send / draft invite: " & _
                ExamSheet.CourseInfo(m_firstSendRow) & vbNewLine & _
            "Last row to send / draft invite: " & _
                ExamSheet.CourseInfo(m_lastSendRow) & vbNewLine & _
            vbNewLine & "Enter:" & vbNewLine & _
            "'yes'  - to proceed with above settings" & vbNewLine & _
            "'no' - to restart" & vbNewLine & _
            "(anything else) - to stop macro"
        InputStr = InputAsker(Asker, "Input Confirmation")
        
        If InputStr = "yes" Then
            Confirmed = True
            m_logger.LogInfo "Settings received and confirmed:"
            m_logger.LogInfo "Sender Name: " & m_senderName, indt:=1
            m_logger.LogInfo "Send immediately: " & m_sendImmediate, indt:=1
            m_logger.LogInfo "Debugging: " & m_isDebugging, indt:=1
            m_logger.LogInfo "First row to send / draft invite: " & _
                ExamSheet.CourseInfo(m_firstSendRow), indt:=1
            m_logger.LogInfo "Last row to send / draft invite: " & _
                ExamSheet.CourseInfo(m_lastSendRow), indt:=1
            m_logger.LogInfo "Proceeding with above settings"
        ElseIf InputStr = "no" Then
            Confirmed = False
            m_logger.LogDebug "Settings rejected, trying again"
        Else
            m_logger.LogInfo "Operation cancelled"
            End
        End If
    Loop
    Exit Sub
QuickDebugSettings:
    m_senderName = "_qds"
    m_isDebugging = True
    m_sendImmediate = False
    m_firstSendRow = Application.Selection.row
    m_lastSendRow = m_firstSendRow + Application.Selection.Rows.Count - 1
    m_logger.LogInfo "Quick Debug Settings used:"
    m_logger.LogInfo "Sender Name: " & m_senderName, indt:=1
    m_logger.LogInfo "Send immediately: " & m_sendImmediate, indt:=1
    m_logger.LogInfo "Debugging: " & m_isDebugging, indt:=1
    m_logger.LogInfo "First row to send / draft invite: " & _
        ExamSheet.CourseInfo(m_firstSendRow), indt:=1
    m_logger.LogInfo "Last row to send / draft invite: " & _
        ExamSheet.CourseInfo(m_lastSendRow), indt:=1
    m_logger.LogInfo "Proceeding with above settings"
End Sub

'Tests if input is one of two valid inputs
Private Function InputOneOf( _
    ByVal receivedInput As String, _
    Optional ByVal ValidStr1 As String = "yes", _
    Optional ByVal ValidStr2 As String = "no") As Boolean
    
    m_logger.LogDebug "Check if input '" & receivedInput & "' exactly one of " _
        & "'" & ValidStr1 & "' or '" & ValidStr2
    If receivedInput = ValidStr1 Or receivedInput = ValidStr2 Then
        InputOneOf = True
        m_logger.LogDebug "Checked: nput validated and passes", indt:=1
    Else
        InputOneOf = False
        m_logger.LogDebug "Checked: input invalid", indt:=1
    End If
End Function

'Tests if input string is empty, raise error and stop running
Private Sub StopInputEmpty(ByVal receivedInput As String)
    m_logger.LogDebug "Check if string '" & receivedInput & "' is empty"
    If receivedInput = "" Then
        m_errMsg = "Checked: invalid, received empty input. Cancelling operation"
        m_logger.LogFatal m_errMsg, Err, m_errNum, indt:=1
        Err.Raise m_errNum, Description:=m_errMsg
    Else
        m_logger.LogDebug "Checked: input string is non-empty", indt:=1
    End If
End Sub

Public Function InputAsker(ByVal askerStr As String, _
    Optional ByVal inputBoxName As String = "Input") As String
    
    Dim receivedInput As String
    m_logger.LogDebug "Asking following:" _
        & vbNewLine & "'''" & vbNewLine & askerStr & vbNewLine & "'''"
    receivedInput = InputBox(askerStr, inputBoxName)
    m_logger.LogDebug "Received input '" & receivedInput & "'", indt:=1
    StopInputEmpty (receivedInput)
    InputAsker = receivedInput
End Function
