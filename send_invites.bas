'Attribute VB_Name = "Send_Invite_Module"
'testyank
Option Explicit

'Initialize global variables
Dim WB As Workbook

'Initialize global variables concerning ExamSheet
Dim ExamSheet As Worksheet
Dim ExLastRow As Integer
Dim ExLastCol As Integer
Dim ExCourseCol As Integer
Dim ExSecCol As Integer
Dim ExProfCol As Integer
Dim ExDateCol As Integer
Dim ExTimeCol As Integer
Dim ExFormCol As Integer
Dim ExCalCol As Integer
Dim ExT1Col As Integer
Dim ExT2Col As Integer
Dim ExRoomCol As Integer
Dim ExTeamTextCol As Integer
Dim ExMaxCourseCol As Integer

Dim ExSpecials As Variant

Dim ExCourseKey As String
Dim ExSecKey As String
Dim ExProfKey As String
Dim ExDateKey As String
Dim ExTimeKey As String
Dim ExPreMeetKey As String
Dim ExT2MeetKey As String
Dim ExEndKey As String
Dim ExFormKey As String
Dim ExCalKey As String
Dim ExT1Key As String
Dim ExT2Key As String
Dim ExRoomKey As String
Dim ExTeamTextKey As String

'Initialize global variables concerning MailSheet
Dim MailSheet As Worksheet
Dim T1LastRow As Integer

'Initialize global variables for error raising
Dim ErrorMsg As String
Dim ErrorNumMod As Long

'Initialize / set values for global variables
Private Sub InitMod()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ErrorNumMod = 42024
    ErrorMsg = "An error occurred in the custom cal invite module..."

    Set WB = ThisWorkbook
    Set ExamSheet = WB.Sheets("Exam Sheet")
    Set MailSheet  = WB.Sheets("Mail List")

    Call MakeRefRange(MailSheet, T1LastRow)
    Call MakeRefRange(ExamSheet, ExLastRow, ExLastCol)

    ExCourseKey = "COURSE"
    ExSecKey = "SECTIONS"
    ExProfKey = "INSTRUCTOR"
    ExDateKey = "DATE"
    ExTimeKey = "TIME"
    ExPreMeetKey = "PRE-MEETING"
    ExT2MeetKey = "TIER 2 MEET"
    ExEndKey = "END TIME"
    ExFormKey = "FORMAT"
    ExCalKey = "CALENDAR INVITE"
    ExT1Key = "TIER 1"
    ExT2Key = "TIER 2"
    ExRoomKey = "SUPPORT ROOM"
    ExTeamTextKey = "TEAMS MESSAGE TEMPLATE"

    'Call FindCol(ExamSheet, <col_name_str>, <col_integer>, <last_col_integer>)
    Call FindCol(ExamSheet, ExCourseKey, ExCourseCol, ExLastCol)
    Call FindCol(ExamSheet, ExSecKey, ExSecCol, ExLastCol)
    Call FindCol(ExamSheet, ExProfKey, ExProfCol, ExLastCol)
    Call FindCol(ExamSheet, ExDateKey, ExDateCol, ExLastCol)
    Call FindCol(ExamSheet, ExTimeKey, ExTimeCol, ExLastCol)
    Call FindCol(ExamSheet, ExFormKey, ExFormCol, ExLastCol)
    Call FindCol(ExamSheet, ExCalKey, ExCalCol, ExLastCol)
    Call FindCol(ExamSheet, ExT1Key, ExT1Col, ExLastCol)
    Call FindCol(ExamSheet, ExT2Key, ExT2Col, ExLastCol)
    Call FindCol(ExamSheet, ExRoomKey, ExRoomCol, ExLastCol)
    Call FindCol(ExamSheet, ExTeamTextKey, ExTeamTextCol, ExLastCol)

    ExSpecials = Array( _
        ExCourseCol, _
        ExSecCol, _
        ExProfCol, _
        ExDateCol, _
        ExTimeCol, _
        ExFormCol, _
        ExRoomCol _
        )

    Call UnmergeCol(ExamSheet, ExDateCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExFormCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExRoomCol, ExLastRow)
End Sub

Private Sub EndMod()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Call RemergeCol(ExamSheet, ExDateCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExFormCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExRoomCol, ExLastRow)
End Sub

Private Sub UnmergeCol( _
    RefSheet As Worksheet, _
    NumCol As Integer, _
    LastRow As Integer)

    Dim Row As Integer
    For Row = 2 To LastRow Step 1
        With RefSheet
            If .Cells(Row, NumCol).MergeCells Then
                .Cells(Row, NumCol).UnMerge
            ElseIf .Cells(Row, NumCol).Value = "" Then
                .Cells(Row, NumCol).Value = .Cells(Row-1, NumCol).Value
            End If
        End With
    Next Row
End Sub

Private Sub RemergeCol( _
    RefSheet As Worksheet, _
    NumCol As Integer, _
    LastRow As Integer)

    Application.DisplayAlerts = False
    Dim Row As Integer
    Dim FirstSameRow As Integer

    For Row = 2 To LastRow Step 1
        With RefSheet
            If .Cells(Row, NumCol).Value <> _
                .Cells(Row-1, NumCol).Value Then
                FirstSameRow = Row
            End If

            If .Cells(Row, NumCol).Value <> _
                .Cells(Row+1, NumCol).Value Then

                .Range(.Cells(FirstSameRow, NumCol), _
                    .Cells(Row, NumCol)).Merge
            End If
        End With
    Next Row
    Application.DisplayAlerts = True
End Sub

Private Sub FindCol( _
    RefSheet As Worksheet, _
    ColName As String, _
    NumCol As Integer, _
    LastCol As Integer, _
    Optional AnchorRow As Integer = 1, _
    Optional AnchorCol As Integer = 1)

    On Error GoTo CannotFind

    FindName:
        With RefSheet
            NumCol = .Range(.Cells(AnchorRow, AnchorCol), _
                .Cells(AnchorRow, LastCol)).Find _
                (ColName, LookIn:=xlValues, MatchCase:=False).Column
        End With
        Exit Sub
    CannotFind:
        ErrorMsg = "Error when finding column with name of '" & ColName & "'." _
            & vbNewLine & vbNewLine & "Try checking spelling of column names." _
            & vbNewLine & "Try deleting empty rows."
        Err.Raise ErrorNumMod, Description:= ErrorMsg
End Sub


'Finds last row and column if applicable
'Can produce a specific reference range also
Private Sub MakeRefRange( _
    RefSheet As Worksheet, _
    NumRow As Integer, _
    Optional NumCol As Integer = 0, _
    Optional RefRange As Range, _
    Optional AnchorRow As Integer = 1, _
    Optional AnchorCol As Integer = 1)

    Call FindLastRC(RefSheet, NumRow, NumCol)

    If Not(RefRange Is Nothing) Then
        With RefSheet
            Set RefRange = .Range(.Cells(AnchorRow, AnchorCol), _
                .Cells(NumRow, NumCol))
        End With
    End If
End Sub

'Finds last row and column in a range
'Will not find row / column if number already exists (non-zero)
Private Sub FindLastRC( _
    FindInSheet As Worksheet, _
    LastRow As Integer, _
    Optional LastCol As Integer = 0)

    If Not(LastRow > 0) Then
        LastRow = FindInSheet.Cells(Rows.Count, 1).End(xlUp).Row
    End If

    If Not(LastCol > 0) Then
        LastCol = FindInSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    End If
End Sub

'This adds recipients by email to the exam meeting
Private Sub FetchMails( _
    MailSheet As Worksheet, _
    LastRow As Integer, _
    SourceRow As Integer, _
    SourceCol As Integer, _
    Optional FNameCol As Integer = 1, _
    Optional LNameCol As Integer = 2, _
    Optional MailCol As Integer = 3, _
    Optional PrefNameCol As Integer = 4)

    Dim SourceVal As String
    SourceVal = ExamSheet.Cells(SourceRow, SourceCol).Value

    Dim ScanRow As Integer
    'Iterate over mailing list, if name in mail list found in SourceVal, add Email
    For ScanRow = 2 To LastRow Step 1
        Dim FNameStr As String
        Dim LNameStr As String
        Dim FInitStr As String
        Dim LInitStr As String
        Dim HasPrefName As Boolean

        'Extract names from ScanRow in MailRange
        With MailSheet
            FNameStr = .Cells(ScanRow, FNameCol)
            LNameStr = .Cells(ScanRow, LNameCol)
            FInitStr = Left(FNameStr, 1)
            LInitStr = Left(LNameStr, 1)
            HasPrefName = (.Cells(ScanRow, PrefNameCol) <> "")

            If HasPrefName Then
                Dim PrefNameStr As String
                Dim PrefInitStr As String
                PrefNameStr = .Cells(ScanRow, PrefNameCol)
                PrefInitStr = Left(PrefNameStr, 1)
            End If
        End With

        'Regular expression objects for searching names
        Dim RegexFirstLInit As Object
        Dim RegexFInitLast As Object
        Set RegexFirstLInit = CreateObject("VBScript.RegExp")
        Set RegexFInitLast = CreateObject("VBScript.RegExp")
        
        'Define patterns used in regular expressions for search
        'One regex for "Firstname L" and one for "F<letters> Lastname"
        If HasPrefName Then
            RegexFirstLInit.Pattern = PrefNameStr & "\s+?" & LInitStr
            RegexFInitLast.Pattern = PrefInitStr & "\w+?\s+?" & LNameStr
        Else
            RegexFirstLInit.Pattern = FNameStr & "\s+?" & LInitStr
            RegexFInitLast.Pattern = FInitStr & "\w+?\s+?" & LNameStr
        End If
        
        If RegexFirstLInit.Test(SourceVal) _
            Or RegexFInitLast.Test(SourceVal) Then

            Debug.Print MailSheet.Cells(ScanRow, MailCol)
        ElseIf HasPrefName Then
            'If has preferred name but can't find by preferred name, try official name
            RegexFirstLInit.Pattern = FNameStr & "\s+?" & LInitStr
            RegexFInitLast.Pattern = FInitStr & "\w+?\s+?" & LNameStr

            If RegexFirstLInit.Test(SourceVal) _
                Or RegexFInitLast.Test(SourceVal) Then

                Debug.Print MailSheet.Cells(ScanRow, MailCol)
            End If
        End If
    Next ScanRow
End Sub

Private Sub InviteCourseAtts()

End Sub

Public Sub DraftIndivInvite( _
    Optional InviteRow As Integer = 0, _
    Optional StepRows As Integer, _
    Optional PreMeetingMins As Integer = 30, _
    Optional T2MeetingMins As Integer = 15, _
    Optional OnlyOneInvite As Boolean = True, _
    Optional testRunning As Boolean = True)

    StepRows = 1
    If InviteRow = 0 Then
        InviteRow = ActiveCell.Row
    End If

    If OnlyOneInvite Then
        Call InitMod
    End If

    'Check if the T1 cell is merged, check for length and assign accordingly to nextrow
    With ExamSheet
        If .Cells(InviteRow, ExT1Col).MergeCells Then
            StepRows = .Cells(InviteRow, ExT1Col).MergeArea.Count
        End If
    End With

    'Create dicionary of dictionaries, key is course name
    'Each course dictionary is further keyed by course attributes (eg. "SECTIONS")
    Dim Courses As Object
    Dim CourseInfo As Object
    Dim CourseInInvite As Integer

    Set Courses = CreateObject("Scripting.Dictionary")
    
    Dim RegexTimes As Object
    Set RegexTimes = CreateObject("VBScript.RegExp")
    RegexTimes.Pattern = "\d\d:\d\d"
    RegexTimes.Global = True
    Dim TimeAllMatches
    Dim TimeFound As Integer
    Dim TimePreMeet As String
    Dim TimeT2Meet As String
    Dim TimeExam As String
    Dim TimeEnd As String

    For CourseInInvite = 1 To StepRows Step 1
        Set CourseInfo = CreateObject("Scripting.Dictionary")
        Dim CourseAtt As Integer
        With ExamSheet
            For CourseAtt = LBound(ExSpecials) To UBound(ExSpecials) Step 1
                If ExSpecials(CourseAtt) <> ExTimeCol _
                    And ExSpecials(CourseAtt) <> ExDateCol Then
                    CourseInfo.Add  _
                        Key:= .Cells(1, ExSpecials(CourseAtt)).Value, _
                        Item:= .Cells(InviteRow + CourseInInvite - 1, ExSpecials(CourseAtt)).Value
                ElseIf ExSpecials(CourseAtt) = ExDateCol Then
                    CourseInfo.Add _
                        Key:= ExDateKey, _
                        Item:= FormatDateTime( _
                            .Cells(InviteRow + CourseInInvite - 1, ExSpecials(CourseAtt)).Value, _
                            vbShortDate)
                ElseIf ExSpecials(CourseAtt) = ExTimeCol Then
                    Set TimeAllMatches = RegexTimes.Execute( _
                        .Cells(InviteRow + CourseInInvite - 1, ExSpecials(CourseAtt)).Value)

                    TimeExam = Left(TimeAllMatches.Item(0), 5)

                    If CourseInInvite = 1 Then
                        TimePreMeet = FormatDateTime( _
                            DateAdd("n", PreMeetingMins * -1, TimeExam), vbShortTime)
                        TimePreMeet = Left(TimePreMeet, 5)
                        TimeT2Meet = FormatDateTime( _
                            DateAdd("n", T2MeetingMins * -1, TimePreMeet), vbShortTime)
                        TimeT2Meet = Left(TimeT2Meet, 5)
                    End If

                    If TimeAllMatches.Count >= 2 Then
                        TimeEnd = Left(TimeAllMatches.Item(1), 5)
                    Else
                        TimeEnd = ""
                    End If

                    With CourseInfo
                        .Add Key:= ExT2MeetKey, Item:= TimeT2Meet
                        .Add Key:= ExPreMeetKey, Item:= TimePreMeet
                        .Add Key:= ExTimeKey, Item:= TimeExam
                        .Add Key:= ExEndKey, Item:= TimeEnd
                    End With
                End If
            Next CourseAtt

            Courses.Add _
                Key:= .Cells(InviteRow + CourseInInvite - 1, 1).Value, _
                Item:= CourseInfo
        End With
    Next CourseInInvite

    If testRunning Then
        Dim cours As Variant
        Dim attDeCours As Variant
        For Each cours In Courses.Keys
            Debug.Print cours
            For Each attDeCours In Courses(cours)
                Debug.Print "    ", attDeCours, Courses(cours).Item(attDeCours)
            Next attDeCours
            Debug.Print "------"
        Next cours
    End If

    If OnlyOneInvite Then
        Call EndMod
    End If

End Sub
