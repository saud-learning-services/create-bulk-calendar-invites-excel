VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMailSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member variables
Private m_mlDataSheet As Worksheet
Private m_mlLastRow As Integer
Private m_mlLastCol As Integer
Private m_mlFrNameCol As Integer
Private m_mlLNameCol As Integer
Private m_mlMPNameCol As Integer
Private m_mlFlNameCol As Integer
Private m_mlEmailCol As Integer
Private m_mlAttColl As Collection
Private m_isDebugging As Boolean
Private m_errMsg As String
Private m_errNum As Long

Private m_logger As clsLogger

Private Sub Class_Initialize()
    m_errMsg = "An error occurred while using property, method or function in" _
        & " class module 'clsMailSheet'"
    m_errNum = 42006
End Sub

Public Sub ConstructMailSheet(MailDataWS As Worksheet, _
    Logger As clsLogger, Optional isDebugging = True)
    Set m_mlDataSheet = MailDataWS
    Set m_logger = Logger
    m_isDebugging = isDebugging
    m_logger.LogInfo "Start initializing Mail Sheet properties"
    m_mlLastRow = m_mlDataSheet.UsedRange.Rows.Count
    m_mlLastCol = m_mlDataSheet.UsedRange.Columns.Count
    m_mlFrNameCol = HeaderNum("First")
    m_mlLNameCol = HeaderNum("Last")
    m_mlMPNameCol = HeaderNum("Middle / Preferred")
    m_mlEmailCol = HeaderNum("E-mail")
    Set m_mlAttColl = New Collection
    m_logger.LogInfo "Mail Sheet properties initialized"
End Sub

'Function - get value at (Row, Col) of Exam Sheet
Public Function GetVal(ByVal RowNum As Integer, ByVal ColNum As Integer) As String
    m_logger.LogDebug "Getting value from Mail Sheet row '" & RowNum _
        & "', col '" & ColNum & "'"
    m_logger.LogDebug "Under '" & m_mlDataSheet.Cells(1, ColNum) & "' column", _
        indt:=1
    GetVal = m_mlDataSheet.Cells(RowNum, ColNum)
    m_logger.LogDebug "Successfully got '" & GetVal & "'"
End Function

Property Get FirstnameInfo(ByVal RowNum As Integer) As String
    FirstnameInfo = GetVal(RowNum, m_mlFrNameCol)
End Property

Property Get LastnameInfo(ByVal RowNum As Integer) As String
    LastnameInfo = GetVal(RowNum, m_mlLNameCol)
End Property

Property Get MidPrefnameInfo(ByVal RowNum As Integer) As String
    MidPrefnameInfo = GetVal(RowNum, m_mlMPNameCol)
End Property

Property Get EmailInfo(ByVal RowNum As Integer) As String
    EmailInfo = GetVal(RowNum, m_mlEmailCol)
End Property

Property Get FullnameInfo(ByVal RowNum As Integer) As String
    FullnameInfo = GetVal(RowNum, m_mlFlNameCol)
End Property

Property Get Attendees() As Collection
    Set Attendees = m_mlAttColl
End Property

Public Sub ConstructAttendeeColl()
    m_logger.LogDebug "Constructing collection of attendees & contacts"
    Dim RowNum As Integer
    For RowNum = 2 To m_mlLastRow
        m_logger.LogDebug "Scanning Mail Sheet row '" & RowNum & "'", _
            indt:=1
        Dim newAtt As clsAttendee
        Set newAtt = New clsAttendee
        newAtt.ConstructAttendee _
            Firstname:=FirstnameInfo(RowNum), _
            Lastname:=LastnameInfo(RowNum), _
            Email:=EmailInfo(RowNum), _
            Logger:=m_logger, _
            isDebugging:=m_isDebugging, _
            MidPrefname:=MidPrefnameInfo(RowNum)
        m_logger.LogDebug "Try adding to collection of attendees", indt:=1
        m_mlAttColl.Add Item:=newAtt
        m_logger.LogDebug "Finished scanning Mail Sheet row '" & RowNum & "'", _
            indt:=1
    Next RowNum
    m_logger.LogDebug "Successfully constructed collection of attendees & contacts"
End Sub

'Function - get column number of a header by name
Public Function HeaderNum(ColName As String) As Integer
    On Error GoTo CannotFind
    With m_mlDataSheet
    HeaderNum = .Range(.Cells(1, 1), .Cells(1, m_mlLastCol)) _
        .Find(ColName, LookIn:=xlValues, MatchCase:=False).Column
    End With
    m_logger.LogDebug "Found header '" & ColName & "' at column number '" & HeaderNum & "'"
    Exit Function
CannotFind:
    m_errMsg = "Error when finding column with name of '" & ColName & "'." _
        & "Try checking spelling of column names"
    m_logger.LogFatal m_errMsg, Err
End Function


