VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAttendee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member variables
Private m_attFirstName As String
Private m_attMidPrefName As String
Private m_attLastName As String
Private m_attFullName As String
Private Const m_attFInitLen As Integer = 1
Private Const m_attLInitLen As Integer = 1
Private m_attFNInit As String
Private m_attMPNInit As String
Private m_attLNInit As String
Private m_attHasMidPref As Boolean
Private m_attEmail As String
Private m_attRole As String
Private m_attReColl As Collection
Private m_attReCollLead As Collection
Private m_attReCollTriage As Collection
Private m_isDebugging As Boolean
Private m_errMsg As String
Private m_errNum As Long

Private m_logger As clsLogger

Private Sub Class_Initialize()
    m_attFirstName = "Firstname"
    m_attLastName = "Lastname"
    m_attMidPrefName = ""
    m_attEmail = ""
    m_errMsg = "An error occurred while using property, method or function in" _
        & " class module 'clsAttendee'"
    m_errNum = 42005
End Sub

Public Sub ConstructAttendee( _
    ByVal Firstname As String, _
    ByVal Lastname As String, _
    ByVal Email As String, _
    Logger As clsLogger, _
    Optional isDebugging As Boolean = True, _
    Optional ByVal MidPrefname As String = "")
    
    Set m_logger = Logger
    m_logger.LogDebug "Start constructing attendee"
    m_isDebugging = isDebugging
    m_attFirstName = Firstname
    m_attLastName = Lastname
    m_attMidPrefName = MidPrefname
    If m_attMidPrefName <> "" Then
        m_attHasMidPref = True
        m_attFullName = m_attFirstName & " " & m_attMidPrefName & " " & m_attLastName
    Else
        m_attHasMidPref = False
        m_attFullName = m_attFirstName & " " & m_attLastName
    End If
    m_logger.LogDebug "Name of attendee '" & m_attFullName & "'", indt:=1
    m_attEmail = Email
    m_logger.LogDebug "Email of attendee '" & m_attEmail & "'", indt:=1
    m_attFNInit = VBA.Left(m_attFirstName, m_attFInitLen)
    m_logger.LogDebug "First initial '" & m_attFNInit & "'", indt:=1
    m_attMPNInit = VBA.Left(m_attMidPrefName, m_attFInitLen)
    m_logger.LogDebug "Mid initial '" & m_attMPNInit & "'", indt:=1
    m_attLNInit = VBA.Left(m_attLastName, m_attLInitLen)
    m_logger.LogDebug "Mid initial '" & m_attLNInit & "'", indt:=1
    Set m_attReCollLead = New Collection
    Set m_attReCollTriage = New Collection
    Set m_attReColl = New Collection
    ConstructReColls
End Sub

Public Sub ConstructReColls()
    'Suppose person:
        'Firstname - Jean
        'Mid/Prefname - Alex
        'Lastname - Smith
        'ie. Jean Alex Smith
    'Below: Re test true for "Jean S", "Jean Smith", "Jean Smth" etc.
    m_logger.LogDebug "Initializing regular expression patterns for '" _
        & m_attFirstName & " " & m_attMidPrefName & " " & m_attLastName & "'"
    m_attReColl.Add _
        m_attFirstName & "\s*" & m_attLNInit
    'Below: Re test true for "Jean S (lead)", "Jean Smith (LEAD)", "Jean Smth (Lead)" etc.
    m_attReCollLead.Add _
        m_attFirstName & "\s*" & m_attLNInit & "\w*?\s*\(*[Ll][Ee][Aa][Dd]\)*"
    'Below: Re test true for "Jean S (triage)", "Jean Smith (TRIAGE)", "Jean Smth (Triage)" etc.
    m_attReCollTriage.Add _
        m_attFirstName & "\s*" & m_attLNInit & "\w*?\s*\(*[Tt][Rr][Ii][Aa][Gg][Ee]\)*"
    
    'Below: Re test true for "J Smith", "Jen Smith", "Jaen Smith"
    m_attReColl.Add m_attFNInit & "\w*?\s*" & m_attLastName
    'Below: Re test true for "J Smith (lead)", "Jen Smith (LEAD)", "Jaen Smith (Lead)"
    m_attReCollLead.Add m_attFNInit & "\w*?\s*" & m_attLastName & "\s*\(*[Ll][Ee][Aa][Dd]\)*"
    'Below: Re test true for "J Smith (lead)", "Jen Smith (LEAD)", "Jaen Smith (Lead)"
    m_attReCollTriage.Add m_attFNInit & "\w*?\s*" & m_attLastName & "\s*\(*[Tt][Rr][Ii][Aa][Gg][Ee]\)*"
    
    If m_attHasMidPref Then
        'Below: Re test true for "Jean S", "Jean (Alex) S" etc, "Jean Alex S".
        m_attReColl.Add _
            m_attFirstName & "\s*\(*\s*" & m_attMidPrefName & "\w*?\s*\)*\s*" & m_attLNInit
        'Below: Re test true for "Jean S (lead)", "Jean (Alex) Smith (LEAD)", "Jean Alex Smth (Lead)" etc.
        m_attReCollLead.Add _
            m_attFirstName & "\s*\(*\s*" & m_attMidPrefName & "\w*?\s*\)*\s*" & m_attLNInit & _
            "\w*?\s*\(*[Ll][Ee][Aa][Dd]\)*"
        'Below: Re test true for "Jean S (triage)", "Jean (Alex) Smith (TRIAGE)", "Jean Alex Smth (Triage)" etc.
        m_attReCollTriage.Add _
            m_attFirstName & "\s*\(*\s*" & m_attMidPrefName & "\w*?\s*\)*\s*" & m_attLNInit & _
            "\w*?\s*\(*[Tt][Rr][Ii][Aa][Gg][Ee]\)*"
        
        'Below: Re test true for "Alex S", "Alex Smith", "Alex Smth"
        m_attReColl.Add m_attMidPrefName & "\s*" & m_attLNInit
        'Below: Re test true for "Alex S (lead)", "Alex Smith (LEAD)", "Alex Smth (Lead)"
        m_attReCollLead.Add m_attMidPrefName & "\s*" & m_attLNInit & _
            "\w*?\s*\(*[Ll][Ee][Aa][Dd]\)*"
        'Below: Re test true for "Alex S (triage)", "Alex Smith (TRIAGE)", "Alex Smth (Triage)"
        m_attReCollTriage.Add m_attMidPrefName & "\s*" & m_attLNInit & _
            "\w*?\s*\(*[Tt][Rr][Ii][Aa][Gg][Ee]\)*"
        
        'Below: Re test true for "A Smith", "Alex Smith", "Aelx Smith"
        m_attReColl.Add m_attMPNInit & "\w*?\s*" & m_attLastName
        'Below: Re test true for "A Smith (lead)", "Alx Smith (LEAD)", "Alxe Smith (Lead)"
        m_attReCollLead.Add m_attMPNInit & "\w*?\s*" & m_attLastName & "\s*\(*[Ll][Ee][Aa][Dd]\)*"
        'Below: Re test true for "A Smith (triage)", "Alx Smith (TRIAGE)", "Alxe Smith (Triage)"
        m_attReCollTriage.Add m_attMPNInit & "\w*?\s*" & m_attLastName & "\s*\(*[Tt][Rr][Ii][Aa][Gg][Ee]\)*"
    End If
    If m_isDebugging Then
        m_logger.LogDebug "The following regular expressions were initialized", indt:=1
        Dim ReStr
        For Each ReStr In m_attReCollLead
            m_logger.LogDebug ReStr, indt:=2
        Next ReStr
        For Each ReStr In m_attReCollTriage
            m_logger.LogDebug ReStr, indt:=2
        Next ReStr
        For Each ReStr In m_attReColl
            m_logger.LogDebug ReStr, indt:=2
        Next ReStr
    End If
    
    m_logger.LogDebug "Successfully initialized regular expression patterns for '" _
        & m_attFullName & "'"
End Sub

Public Function CheckNameRe(ByVal SourceText As String) As String
    m_logger.LogDebug "Checking regex for '" & m_attFullName & _
        "' in text '" & _
        vbNewLine & "'''" & vbNewLine _
        & SourceText & vbNewLine & "'''"
    Dim ReStr As Variant
    Dim NameRegex As New RegExp
    For Each ReStr In m_attReCollLead
        NameRegex.Pattern = ReStr
        m_logger.LogDebug "Checking regex '" & ReStr & "'", indt:=1
        If NameRegex.Test(SourceText) Then
            m_logger.LogDebug "Found regex '" & ReStr & "'"
            CheckNameRe = "Lead"
            m_logger.LogDebug "Attending '" & CheckNameRe & "'", indt:=1
            Exit Function
        End If
        m_logger.LogDebug "Did not find regex '" & ReStr & "'", indt:=1
    Next ReStr
    For Each ReStr In m_attReCollTriage
        NameRegex.Pattern = ReStr
        m_logger.LogDebug "Checking regex '" & ReStr & "'", indt:=1
        If NameRegex.Test(SourceText) Then
            m_logger.LogDebug "Found regex '" & ReStr & "'"
            CheckNameRe = "Triage"
            m_logger.LogDebug "Attending '" & CheckNameRe & "'", indt:=1
            Exit Function
        End If
        m_logger.LogDebug "Did not find regex '" & ReStr & "'", indt:=1
    Next ReStr
    For Each ReStr In m_attReColl
        NameRegex.Pattern = ReStr
        m_logger.LogDebug "Checking regex '" & ReStr & "'", indt:=1
        If NameRegex.Test(SourceText) Then
            m_logger.LogDebug "Found regex '" & ReStr & "'"
            CheckNameRe = "True"
            m_logger.LogDebug "Attending '" & CheckNameRe & "'", indt:=1
            Exit Function
        End If
        m_logger.LogDebug "Did not find regex '" & ReStr & "'", indt:=1
    Next ReStr
    CheckNameRe = "False"
    m_logger.LogDebug "Did not find '" & m_attFullName & _
        "' in source tested"
    m_logger.LogDebug "Attending '" & CheckNameRe & "'", indt:=1
End Function

Property Get Firstname() As String
    Firstname = m_attFirstName
End Property

Property Get Lastname() As String
    Lastname = m_attLastName
End Property

Property Get MidPrefname() As String
    MidPrefname = m_attMidPrefName
End Property

Property Get Fullname() As String
    Fullname = m_attFullName
End Property

Property Get Email() As String
    Email = m_attEmail
End Property
