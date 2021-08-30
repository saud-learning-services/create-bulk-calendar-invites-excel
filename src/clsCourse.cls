VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member variables
Private m_crName As String
Private m_crSec As String
Private m_crExLink As String
Private m_crProf As String
Private m_crExStart As String
Private m_crExEnd As String
Private m_crExTime As String
Private m_crExDur As String
Private m_crExForm As String
Private m_crExType As String
Private m_crNumStud As String
Private m_crOriginRow As Integer
Private Const m_defaultExDur As Integer = 225 'minutes
Private Const m_defaultExStart As String = "07:00"
Private m_errMsg As String
Private m_errNum As Long

Private m_exSheet As clsExamSheet
Private m_logger As clsLogger

Private Sub Class_Initialize()
    m_crName = "CCCC 000"
    m_crSec = "000"
    m_crExLink = "https://canvas.ubc.ca"
    m_crProf = "Firstname Lastname"
    m_crExTime = "--:-- - --:--"
    m_crExStart = "--:--"
    m_crExEnd = "--:--"
    m_crExDur = m_defaultExDur
    m_crExForm = "Exam"
    m_crNumStud = 1
    m_crOriginRow = 1
    m_errMsg = "An error occurred while using property, method or function in" _
        & " class module 'clsCourse'"
    m_errNum = 42004
End Sub

Public Sub ConstructCourse(ExamSheet As clsExamSheet, CourseRow As Integer, _
    Logger As clsLogger)
    
    m_crOriginRow = CourseRow
    Set m_logger = Logger
    Set m_exSheet = ExamSheet
    Logger.LogDebug "Constructing course from row '" & CourseRow & "'"
    m_crName = m_exSheet.CourseInfo(CourseRow)
    Logger.LogDebug "Found course name '" & m_crName & "'", indt:=1
    m_crSec = m_exSheet.SectionInfo(CourseRow)
    Logger.LogDebug "Found section name '" & m_crSec & "'", indt:=1
    m_crProf = m_exSheet.InstructorInfo(CourseRow)
    Logger.LogDebug "Found instructor(s) '" & m_crProf & "'", indt:=1
    m_crExTime = m_exSheet.TimeInfo(CourseRow)
    Logger.LogDebug "Found raw exam time string '" & m_crExTime & "'", _
        indt:=1
    SetStartEnd m_crExTime
    m_crExDur = m_exSheet.DurationInfo(CourseRow)
    If m_crExDur = "" Then
        m_crExDur = WorksheetFunction.Round( _
            DateDiff("n", _
            FormatDateTime(m_crExStart, vbShortTime), _
            FormatDateTime(m_crExEnd, vbShortTime)), 0)
        m_crExDur = CStr(m_crExDur) & " mins"
    End If
    m_crExForm = m_exSheet.FormatInfo(CourseRow)
    m_crNumStud = m_exSheet.NumStudentsInfo(CourseRow)
    m_crExLink = m_exSheet.ExamLinkInfo(CourseRow)
    m_crExType = m_exSheet.ExamTypeInfo(CourseRow)
    
    Logger.LogDebug "Finished constructing course from row '" & CourseRow & "', '" _
        & m_crName & "' section '" & m_crSec & "'"
End Sub

Public Sub SetStartEnd(ByVal TimeStr As String)
    Dim RegexTimes As RegExp
    Set RegexTimes = New RegExp
    
    m_logger.LogDebug _
        "Searching for exam start and end times in string '" & TimeStr & "'"
    RegexTimes.Pattern = "\b\d*\d:\d\d"
    RegexTimes.Global = True
    Dim TimeAllMatches As Object
    Set TimeAllMatches = RegexTimes.Execute(TimeStr)
    
    Dim NumTimesFound As Integer: NumTimesFound = TimeAllMatches.Count
    m_logger.LogDebug "Found '" & NumTimesFound & "' results matching time format", _
        indt:=1
    Select Case NumTimesFound
        Case Is >= 2
            m_crExStart = FormatDateTime(TimeAllMatches.Item(0), vbShortTime)
            m_logger.LogDebug "Exam start time set '" & m_crExStart & "'", _
                indt:=1
            m_crExEnd = FormatDateTime(TimeAllMatches.Item(1), vbShortTime)
            m_logger.LogDebug "Exam end time set '" & m_crExEnd & "'", _
                indt:=1
        Case 1
            m_crExStart = FormatDateTime(TimeAllMatches.Item(0), vbShortTime)
            m_logger.LogDebug "Exam start time set '" & m_crExStart & "'", _
                indt:=1
            m_logger.LogError "Did not find exam end time in string '" _
                & TimeStr & "', assume end at start + '" & m_defaultExDur & "' minutes", _
                Err, errNumMods:=m_errNum, indt:=1
            m_crExEnd = FormatDateTime( _
                DateAdd("n", m_defaultExDur, FormatDateTime(m_crExStart, vbShortTime)), _
                vbShortTime)
            m_logger.LogWarning "Exam end time set '" & m_crExEnd & "'", _
                indt:=1
        Case 0
            On Error GoTo NoTimeFound
            m_logger.LogError "Did not find exam times in string '" _
                & TimeStr & "' with regular expression, " & _
                "now trying to format cell value directly", _
                Err, errNumMods:=m_errNum, indt:=1
            m_crExStart = FormatDateTime(TimeStr, vbShortTime)
            m_logger.LogWarning "Exam start time set '" & m_crExStart & "'", _
                indt:=1
            m_logger.LogError "Did not find exam end time in string '" _
                & TimeStr & "', assume end at start + '" & m_defaultExDur & "' minutes", _
                Err, errNumMods:=m_errNum, indt:=1
            m_crExEnd = FormatDateTime( _
                DateAdd("n", m_defaultExDur, FormatDateTime(m_crExStart, vbShortTime)), _
                vbShortTime)
            m_logger.LogWarning "Exam end time set '" & m_crExEnd & "'", _
                indt:=1
    End Select
    Exit Sub
NoTimeFound:
    m_logger.LogError "Did not find any time at all in string '" & TimeStr & "'", _
        Err, errNumMods:=m_errNum, indt:=1
    m_logger.LogError "Setting start as '" & m_defaultExStart & "', " _
        & "please check if further error occurs", _
        Err, errNumMods:=m_errNum, indt:=1
    m_crExStart = FormatDateTime(m_defaultExStart, vbShortTime)
    m_crExEnd = FormatDateTime( _
        DateAdd("n", m_defaultExDur, FormatDateTime(m_defaultExStart, vbShortTime)), _
        vbShortTime)
End Sub

'Private m_crName As String
Property Get Name() As String
    Name = m_crName
End Property

'Private m_crSec As String
Property Get Section() As String
    Section = m_crSec
End Property

'Private m_crExLink As String
Property Get ExamLink() As String
    ExamLink = m_crExLink
End Property

'Private m_crProf As String
Property Get Instructor() As String
    Instructor = m_crProf
End Property

'Private m_crExStart As String
Property Get ExamStartTime() As String
    ExamStartTime = FormatDateTime(m_crExStart, vbShortTime)
End Property

'Private m_crExEnd As String
Property Get ExamEndTime() As String
    ExamEndTime = FormatDateTime(m_crExEnd, vbShortTime)
End Property

'Private m_crExDur As String
Property Get ExamDuration() As String
    ExamDuration = m_crExDur
End Property

'Private m_crExForm As String
Property Get ExamFormat() As String
    ExamFormat = m_crExForm
End Property

Property Get ExamType() As String
    ExamType = m_crExType
End Property

'Private m_crNumStud As String
Property Get ExamNumStudents() As String
    ExamNumStudents = m_crNumStud
End Property
