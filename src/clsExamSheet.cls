VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExamSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member variables
Private m_exDataSheet As Worksheet
Private m_exLastRow As Integer
Private m_exLastCol As Integer
Private m_exCourseCol As Integer
Private m_exSecCol As Integer
Private m_exProfCol As Integer
Private m_exDateCol As Integer
Private m_exTimeCol As Integer
Private m_exFormCol As Integer
Private m_exDurCol As Integer
Private m_exNumStudCol As Integer
Private m_exCalCol As Integer
Private m_exT1Col As Integer
Private m_exT2Col As Integer
Private m_exRoomCol As Integer
Private m_exTeamTextCol As Integer
Private m_exPplInvitedCol As Integer
Private m_exAutoDraftCol As Integer
Private m_exLinkCol As Integer
Private m_exTypeCol As Integer
Private m_errMsg As String
Private m_errNum As Long
Private m_mergeRecord As Dictionary
Private m_isUnermged As Boolean
Private m_exSpecialCols As Dictionary

'Member variable - require class(es)
' - clsLogger
Private m_logger As clsLogger

'IDEA - will have something like RawSheetData.Time(RowNum)

'Event - Initialize
Private Sub Class_Initialize()
m_errMsg = "An error occurred while using property, method or function in" _
    & " class module 'clsExamSheet'"
m_errNum = 42003
Set m_mergeRecord = New Dictionary
End Sub

'Setter - Exam Sheet
'This should be called first and immediately after constructing new class obj
'Property Set Data(xlWS As Worksheet)
'   Set m_exDataSheet = xlWS
'End Property

'Getter - Exam Sheet
Property Get Data() As Worksheet
Set Data = m_exDataSheet
End Property

'Class Method - initialize various properties of the class
'this initialization must be called manually
'this method should be called early, right after setting Data
Public Sub ConstructExamSheet(ExamDataWS As Worksheet, Logger As clsLogger)
    Set m_exDataSheet = ExamDataWS
    Set m_logger = Logger
    m_logger.LogInfo "Start initializing Exam Sheet properties"
    m_exLastRow = m_exDataSheet.UsedRange.Rows.Count
    m_exLastCol = m_exDataSheet.UsedRange.Columns.Count
    m_exCourseCol = HeaderNum("COURSE")
    m_exSecCol = HeaderNum("SECTIONS")
    m_exProfCol = HeaderNum("INSTRUCTOR")
    m_exDateCol = HeaderNum("DATE")
    m_exTimeCol = HeaderNum("TIME")
    m_exFormCol = HeaderNum("FORMAT")
    m_exDurCol = HeaderNum("DURATION")
    m_exNumStudCol = HeaderNum("# STUDENTS")
    m_exCalCol = HeaderNum("CALENDAR INVITE")
    m_exPplInvitedCol = HeaderNum("PEOPLE INVITED")
    m_exT1Col = HeaderNum("TIER 1")
    m_exT2Col = HeaderNum("TIER 2")
    m_exRoomCol = HeaderNum("SUPPORT ROOM")
    m_exTeamTextCol = HeaderNum("TEAMS MESSAGE TEMPLATE")
    m_exAutoDraftCol = HeaderNum("AUTO DRAFT")
    m_exLinkCol = HeaderNum("EXAM LINK")
    m_exTypeCol = HeaderNum("MIDTERM / FINAL / SD?")
    m_logger.LogInfo "Exam Sheet properties initialized"
End Sub

'Function - get value at (Row, Col) of Exam Sheet
Public Function GetVal(ByVal RowNum As Integer, ByVal ColNum As Integer) As String
    m_logger.LogDebug "Getting value from Exam Sheet row '" & RowNum _
        & "', col '" & ColNum & "'"
    m_logger.LogDebug "From course '" & m_exDataSheet.Cells(RowNum, m_exCourseCol) & "'", 1
    m_logger.LogDebug "Under '" & m_exDataSheet.Cells(1, ColNum) & "' column", 1
    GetVal = m_exDataSheet.Cells(RowNum, ColNum)
    m_logger.LogDebug "Successfully got value"
End Function

'Class Method - set value at (Row, Col) of Exam Sheet
Public Sub SetVal(ByVal RowNum As Integer, ByVal ColNum As Integer, ByVal NewVal As String)
    m_logger.LogDebug "Setting Exam Sheet row '" & RowNum _
        & "', col '" & ColNum & _
        "' as: '" & NewVal & "'"
    m_logger.LogDebug "For course '" & m_exDataSheet.Cells(RowNum, m_exCourseCol) & "'", 1
    m_logger.LogDebug "Under '" & m_exDataSheet.Cells(1, ColNum) & "' column", 1
    m_exDataSheet.Cells(RowNum, ColNum).Value = NewVal
    m_logger.LogDebug "Successfully set value"
End Sub

'Getter - last row in exam sheet
Property Get LastRow() As Integer
    LastRow = m_exLastRow
End Property

'Getter - last column of exam sheet
Property Get LastCol() As Integer
    LastCol = m_exLastCol
End Property

'Function - get column number of a header by name
Public Function HeaderNum(ColName As String, Optional RecordCol As Boolean = True) As Integer
    On Error GoTo CannotFind
    With m_exDataSheet
        HeaderNum = .Range(.Cells(1, 1), .Cells(1, m_exLastCol)) _
            .Find(ColName, LookIn:=xlValues, MatchCase:=False).Column
    End With
    'If column is of particular interest, the header name and column number is ...
    '... also stored, for the purposes of determining which columns can be unmerged etc.
    If RecordCol Then
        If m_exSpecialCols Is Nothing Then
            Set m_exSpecialCols = New Dictionary
        End If
        m_exSpecialCols.Add Key:=ColName, Item:=HeaderNum
    End If
    m_logger.LogDebug "Found header '" & ColName & "' at column number '" & HeaderNum & "'"
    Exit Function
CannotFind:
    m_errMsg = "Error when finding column with name of '" & ColName & "'." _
        & "Try checking spelling of column names."
    m_logger.LogFatal m_errMsg, Err, raiseFatal:=True
End Function

Property Get CourseInfo(ByVal RowNum As Integer) As String
    CourseInfo = GetVal(RowNum, m_exCourseCol)
End Property

Property Get SectionInfo(ByVal RowNum As Integer) As String
    SectionInfo = GetVal(RowNum, m_exSecCol)
End Property

Property Get InstructorInfo(ByVal RowNum As Integer) As String
    InstructorInfo = GetVal(RowNum, m_exProfCol)
End Property

Property Get DateInfo(ByVal RowNum As Integer) As String
    DateInfo = GetVal(RowNum, m_exDateCol)
End Property

Property Get TimeInfo(ByVal RowNum As Integer) As String
    TimeInfo = GetVal(RowNum, m_exTimeCol)
End Property

Property Get FormatInfo(ByVal RowNum As Integer) As String
    FormatInfo = GetVal(RowNum, m_exFormCol)
End Property

Property Get DurationInfo(ByVal RowNum As Integer) As String
    DurationInfo = GetVal(RowNum, m_exDurCol)
End Property

Property Get NumStudentsInfo(ByVal RowNum As Integer) As String
    NumStudentsInfo = GetVal(RowNum, m_exNumStudCol)
End Property

Property Get CalStatusInfo(ByVal RowNum As Integer) As String
    CalStatusInfo = GetVal(RowNum, m_exCalCol)
End Property

Property Let CalStatusInfo(ByVal RowNum As Integer, ByVal newCalStatus As String)
    Call SetVal(RowNum, m_exCalCol, newCalStatus)
End Property

Property Get T1Info(ByVal RowNum As Integer) As String
    T1Info = GetVal(RowNum, m_exT1Col)
End Property

Property Get T2Info(ByVal RowNum As Integer) As String
    T2Info = GetVal(RowNum, m_exT2Col)
End Property

Property Get T2Col() As Integer
    T2Col = m_exT2Col
End Property

Property Get RoomInfo(ByVal RowNum As Integer) As String
    RoomInfo = GetVal(RowNum, m_exRoomCol)
End Property

Property Get TeamTextInfo(ByVal RowNum As Integer) As String
    TeamTextInfo = GetVal(RowNum, m_exTeamTextCol)
End Property

Property Let TeamTextInfo(ByVal RowNum As Integer, ByVal newTeamText As String)
    Call SetVal(RowNum, m_exTeamTextCol, newTeamText)
End Property

Property Get AutoDraftInfo(ByVal RowNum As Integer) As String
    AutoDraftInfo = GetVal(RowNum, m_exAutoDraftCol)
End Property

Property Get ExamLinkInfo(ByVal RowNum As Integer) As String
    ExamLinkInfo = GetVal(RowNum, m_exLinkCol)
End Property

Property Get ExamTypeInfo(ByVal RowNum As Integer) As String
    ExamTypeInfo = GetVal(RowNum, m_exTypeCol)
End Property

Property Get ExamPeopleInvited(ByVal RowNum As Integer) As String
    ExamPeopleInvited = GetVal(RowNum, m_exPplInvitedCol)
End Property

Property Let ExamPeopleInvited(ByVal RowNum As Integer, ByVal invitedPerson As String)
    Call SetVal(RowNum, m_exPplInvitedCol, invitedPerson)
End Property

Public Sub UnmergeCol(ByVal ColNum As Integer)
    m_isUnermged = True
    m_logger.LogDebug "Unmerging cells at column '" _
        & ColNum & "' header '" & GetVal(1, ColNum) & "'"
    Dim MergedCells As New Dictionary
    Dim RowNum As Integer
    For RowNum = 1 To m_exLastRow
        With m_exDataSheet
            If .Cells(RowNum, ColNum).MergeCells Then
                MergedCells.Add Key:=CStr(RowNum), Item:=.Cells(RowNum, ColNum).MergeArea.Count
                .Cells(RowNum, ColNum).UnMerge
                m_logger.LogDebug "Successfully unmerged cell at row '" & RowNum & _
                    "' with size '" & MergedCells.Item(CStr(RowNum)) & "'", indt:=1
                Dim RowInMerge As Integer
                For RowInMerge = RowNum + 1 To RowNum + MergedCells.Item(CStr(RowNum)) - 1
                    .Cells(RowInMerge, ColNum) = .Cells(RowInMerge - 1, ColNum)
                Next RowInMerge
            End If
        End With
    Next RowNum
    m_logger.LogDebug "Cells successfully unmerged", indt:=1
    m_mergeRecord.Add Key:=CStr(ColNum), Item:=MergedCells
    m_logger.LogDebug "Successfully unmerged cells at column '" & ColNum & "'"
End Sub

Public Sub RemergeCol(ByVal ColNum As Integer)
    If Not (m_mergeRecord.Exists(CStr(ColNum))) Then
        Exit Sub
    End If
    m_logger.LogDebug "Remerging cells at column '" _
        & ColNum & "' header '" & GetVal(1, ColNum) & "'"
    Dim MergedRow As Variant
    Dim MergedCells As Dictionary
    Set MergedCells = m_mergeRecord.Item(CStr(ColNum))
    m_logger.LogDebug "Found unmerge record for column '" & ColNum & "'", indt:=1
    Application.DisplayAlerts = False
    For Each MergedRow In MergedCells
        Dim MergeToRow As Integer: MergeToRow = MergedRow + MergedCells.Item(CStr(MergedRow)) - 1
        m_logger.LogDebug "Remerging cell from row '" & MergedRow & "'" & _
            " to row '" & MergeToRow & "'" & _
            " size '" & MergedCells.Item(CStr(MergedRow)) & "'", indt:=1
        With m_exDataSheet
            .Range( _
                .Cells(MergedRow, ColNum), _
                .Cells(MergeToRow, ColNum)).Merge
            m_logger.LogDebug "Successfully remerged cell at '" & _
                MergedRow & "'", indt:=1
        End With
    Next MergedRow
    Application.DisplayAlerts = True
    m_logger.LogDebug "Successfully remerged cells in '" & ColNum & "'"
End Sub

Public Sub UnmergeAll(ParamArray colList() As Variant)
    Dim toUnmerge
    m_logger.LogInfo "Unmerging given columns"
    For Each toUnmerge In colList
        If m_exSpecialCols.Exists(CStr(toUnmerge)) Then
            UnmergeCol (m_exSpecialCols.Item(CStr(toUnmerge)))
        End If
    Next toUnmerge
    m_logger.LogInfo "Successfully unmerged given columns"
End Sub

Public Sub RemergeAll()
m_logger.LogInfo "Remerging all unmerged columns"
Dim MergedCol As Variant
For Each MergedCol In m_mergeRecord
    RemergeCol CInt(MergedCol)
Next MergedCol
m_logger.LogInfo "Successfully remerged all unmerged columns"
m_isUnermged = False
End Sub

Private Sub Class_Terminate()
    If m_isUnermged Then
        RemergeAll
    End If
End Sub
