'Attribute VB_Name = "Send_Invite_Module"
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
Dim ExDurCol As Integer
Dim ExNumStudCol As Integer
Dim ExCalCol As Integer
Dim ExT1Col As Integer
Dim ExT2Col As Integer
Dim ExRoomCol As Integer
Dim ExTeamTextCol As Integer
Dim ExAutoDraftCol As Integer
Dim ExAutoUpdateCol As Integer

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
Dim ExDurKey As String
Dim ExNumStudKey As String
Dim ExCalKey As String
Dim ExT1Key As String
Dim ExT2Key As String
Dim ExRoomKey As String
Dim ExTeamTextKey As String
Dim ExAutoDraftKey As String
Dim ExAutoUpdateKey As String

Dim InvNameKey As String
Dim InvRoleKey As String
Dim CapAlphs As Variant

'Initialize global variables concerning MailSheet
Dim MailSheet As Worksheet
Dim MailLastRow As Integer

'Initialize global variables for error raising
Dim ErrorMsg As String
Dim ErrorNumMod As Long

'Initialize / set values for global variables
'Tune default settings to make computation faster (less updating etc)
'Unmerge cells in select columns
Private Sub InitMod()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ErrorNumMod = 42024
    ErrorMsg = "An error occurred in the custom cal invite module..."

    Set WB = ThisWorkbook
    Set ExamSheet = WB.Sheets("Exam Sheet")
    Set MailSheet  = WB.Sheets("Mail List")

    Call MakeRefRange(MailSheet, MailLastRow)
    Call MakeRefRange(ExamSheet, ExLastRow, ExLastCol)

    'Keys for dictionary lookups, most follows spreadsheet column names
    ExCourseKey = "COURSE"
    ExSecKey = "SECTIONS"
    ExProfKey = "INSTRUCTOR"
    ExDateKey = "DATE"
    ExTimeKey = "TIME"
    ExPreMeetKey = "PRE-MEETING"
    ExT2MeetKey = "TIER 2 MEET"
    ExEndKey = "END TIME"
    ExFormKey = "FORMAT"
    ExDurKey = "DURATION"
    ExNumStudKey = "# STUDENTS"
    ExCalKey = "CALENDAR INVITE"
    ExT1Key = "TIER 1"
    ExT2Key = "TIER 2"
    ExRoomKey = "SUPPORT ROOM"
    ExTeamTextKey = "TEAMS MESSAGE TEMPLATE"
    ExAutoDraftKey = "AUTO DRAFT"
    ExAutoUpdateKey = "AUTO UPDATE"
    InvNameKey = "FULL NAME"
    InvRoleKey = "ROLE"

    'Call FindCol(ExamSheet, <col_name_str>, <col_integer>, <last_col_integer>)
    Call FindCol(ExamSheet, ExCourseKey, ExCourseCol, ExLastCol)
    Call FindCol(ExamSheet, ExSecKey, ExSecCol, ExLastCol)
    Call FindCol(ExamSheet, ExProfKey, ExProfCol, ExLastCol)
    Call FindCol(ExamSheet, ExDateKey, ExDateCol, ExLastCol)
    Call FindCol(ExamSheet, ExTimeKey, ExTimeCol, ExLastCol)
    Call FindCol(ExamSheet, ExFormKey, ExFormCol, ExLastCol)
    Call FindCol(ExamSheet, ExDurKey, ExDurCol, ExLastCol)
    Call FindCol(ExamSheet, ExNumStudKey, ExNumStudCol, ExLastCol)
    Call FindCol(ExamSheet, ExCalKey, ExCalCol, ExLastCol)
    Call FindCol(ExamSheet, ExT1Key, ExT1Col, ExLastCol)
    Call FindCol(ExamSheet, ExT2Key, ExT2Col, ExLastCol)
    Call FindCol(ExamSheet, ExRoomKey, ExRoomCol, ExLastCol)
    Call FindCol(ExamSheet, ExTeamTextKey, ExTeamTextCol, ExLastCol)
    Call FindCol(ExamSheet, ExAutoUpdateKey, ExAutoUpdateCol, ExLastCol)
    Call FindCol(ExamSheet, ExAutoDraftKey, ExAutoDraftCol, ExLastCol)

    ExSpecials = Array( _
        ExCourseCol, _
        ExSecCol, _
        ExProfCol, _
        ExDateCol, _
        ExTimeCol, _
        ExFormCol, _
        ExDurCol, _
        ExNumStudCol, _
        ExRoomCol _
        )

    CapAlphs = Array( _
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
        "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

    Call SaveOutput

    Call UnmergeCol(ExamSheet, ExCourseCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExSecCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExProfCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExDateCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExTimeCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExFormCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExDurCol, ExLastRow)
    Call UnmergeCol(ExamSheet, ExRoomCol, ExLastRow)

    Call CheckDates
End Sub

'To conclude process, restore default spreadsheet settings, remerge cells
Private Sub EndMod()
    Call RemergeCol(ExamSheet, ExCourseCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExSecCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExProfCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExDateCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExTimeCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExFormCol, ExLastRow)
    Call RemergeCol(ExamSheet, ExDurCol, ExLastRow)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub WriteMultipleInvites()
    Call InitMod
    
    Dim Sender As String
    Dim SendImmediate As Boolean
    Dim IsDebugging As Boolean
    Dim FirstExRow As Integer
    Dim LastExRow As Integer

    Call Askers(Sender, _
        SendImmediate, _
        IsDebugging, _
        FirstExRow, _
        LastExRow)

    Dim InviteRow As Integer
    InviteRow = FirstExRow
    Dim StepThruCourses As Integer

    Do While InviteRow <= LastExRow
        Call WriteSingleInvite( _
            SendImmediate, _
            Sender, _
            InviteRow, _
            StepThruCourses, _
            IsDebugging, _
            OnlyOneInvite:=False)
        InviteRow = InviteRow + StepThruCourses
    Loop

    Call EndMod
End Sub

'Confirms settings with macro user
Private Sub Askers( _
    Sender As String, _
    SendImmediate As Boolean, _
    IsDebugging As Boolean, _
    FirstExRow As Integer, _
    LastExRow As Integer)

    Dim InputStr As String
    Dim RunSelected As Boolean

    Dim SelectedFirst As Integer
    Dim SelectedLast As Integer
    SelectedFirst = Application.Selection.Row
    SelectedLast = SelectedFirst + Application.Selection.Rows.Count - 1

    Dim Confirmed As Boolean
    Confirmed = False
    Do While Confirmed = False
        Sender = InputBox("Please enter your name (for record keeping):", "Enter Name")

        InputStr = ""
        Do While InputStr <> "yes" And InputStr <> "no"
            InputStr = InputBox( _
                "Are you sending / drafting actual invites?" & _
                vbNewLine & vbNewLine & "Enter:" & vbNewline & _
                "'yes' - sending / drafting invites for actual exams" & vbNewLine & _
                "'no' - testing / debugging only" & vbNewline & vbNewline & _
                "(Enter answer without quotes)", "Check debugging")

            If InputStr = "" Then
                Call EndMod
                End
            End If
        Loop
        If InputStr = "yes" Then
            IsDebugging = False
        Else
            IsDebugging = True
        End If

        InputStr = "0"
        Do While InputStr <> "send" And InputStr <> "save"
            InputStr = InputBox("Send invites immediately or save as draft?" & _
                vbNewLine & vbNewLine & "Enter:" & vbNewline & _
                "'send' - send invites immediately" & vbNewLine & _
                "'save' - save invite drafts to send later" & vbNewline & vbNewline & _
                "(Enter answer without quotes)", "Send or Save")

            If InputStr = "" Then
                Call EndMod
                End
            End If
        Loop
        If InputStr = "send" Then
            SendImmediate = True
        Else
            SendImmediate = False
        End If

        InputStr = ""
        Do While InputStr <> "selected" And InputStr <> "assign"
            InputStr = InputBox("Send invites for exam(s) in selected range or" & _
                " assign range manually?" & _
                vbNewLine & vbNewLine & "Enter:" & vbNewline & _
                "'selected' - send / draft invites for exams in selected range" & vbNewLine & _
                "'assign' - assign which rows to send / draft invites" & vbNewline & vbNewline & _
                "(Enter answer without quotes)", "Selected range or assign new")

            If InputStr = "" Then
                Call EndMod
                End
            End If
        Loop
        If InputStr = "selected" Then
            RunSelected = True
        Else
            RunSelected = False
        End If

        If RunSelected Then
            FirstExRow = SelectedFirst
            LastExRow = SelectedLast
        Else
            InputStr = "0"
            Do Until IsNumeric(InputStr) And CInt(InputStr) > 1
                InputStr = InputBox( _
                    "Enter row number of the FIRST EXAM to send / draft invite." & _
                    vbNewline & vbNewline & "Please enter FIRST EXAM's row as an integer:", _
                    "First Exam Row Number")
                If InputStr = "" Then
                    Call EndMod
                    End
                End If
            Loop
            FirstExRow = CInt(InputStr)

            InputStr = "0"
            Do Until IsNumeric(InputStr) And CInt(InputStr) >= FirstExRow
                InputStr = InputBox( _
                    "Enter row number of the LAST EXAM to send / draft invite." & _
                    vbNewline & vbNewline & "Please enter LAST EXAM's row as an integer:", _
                    "Last Exam Row Number")
                If InputStr = "" Then
                    Call EndMod
                    End
                End If
            Loop
            LastExRow = CInt(InputStr)
        End If

        InputStr = InputBox( _
            "Proceed with following settings?" & vbNewline & vbNewline & _
            "Sender Name: " & Sender & vbNewline & _
            "Send invites immediately: " & SendImmediate & vbNewline & _
            "Debugging: " & IsDebugging & vbNewline & _
            "First course to send / draft invite: " & _
                ExamSheet.Cells(FirstExRow, ExCourseCol).Value & vbNewline & _
            "Last course to send / draft invite: " & _
                ExamSheet.Cells(LastExRow, ExCourseCol).Value & vbNewline & _
            vbNewline & "Enter:" & vbNewline & _
            "'yes'  - to proceed with above settings" & vbNewline & _
            "'no' - to stop macro" & vbNewline & _
            "(anything else) - to start over", _
            "Input Confirmation")

        If InputStr = "yes" Then
            Confirmed = True
        ElseIf InputStr = "no" Then
            End
        End If
    Loop
End Sub

'Goes through cells in a column, unmerge cells; empty cells gets above cell value
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

'Goes through cells in a column, merge cells with identical values
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

Private Sub InvNameRoleAtts( _
    InviteeAtts As Object, _
    RoleInd As Integer, _
    HasPrefName As Boolean, _
    InvIsLead As Boolean, _
    InvIsTriage As Boolean, _
    FNameStr As String, _
    LNameStr As String, _
    PrefNameStr As String, _
    SourceCol As Integer)

    With InviteeAtts
        If HasPrefName Then
            .Add Key:= InvNameKey, Item:= FNameStr & " " & PrefNameStr & " " & LNameStr
        Else
            .Add Key:= InvNameKey, Item:= FNameStr & " " & LNameStr
        End If

        If SourceCol = ExT1Col Then
            If InvIsTriage Then
                .Add Key:= InvRoleKey, Item:="Triage"
            Else
                .Add Key:= InvRoleKey, Item:="T1" & CapAlphs(RoleInd)
                RoleInd = RoleInd + 1
            End If
        ElseIf SourceCol = ExT2Col Then
            If InvIsLead Then
                .Add Key:= InvRoleKey, Item:="Lead"
            Else
                .Add Key:= InvRoleKey, Item:="T2" & CapAlphs(RoleInd)
                RoleInd = RoleInd + 1
            End If
        Else
            .Add Key:= InvRoleKey, Item:=ExProfKey
        End If
    End With

End Sub

'This adds recipients by email to the exam meeting
Private Sub FetchMails( _
    LastRow As Integer, _
    SourceCol As Integer, _
    SourceVal As String, _
    InviteeList As Object, _
    Optional FNameCol As Integer = 2, _
    Optional LNameCol As Integer = 3, _
    Optional MailCol As Integer = 4, _
    Optional PrefNameCol As Integer = 5, _
    Optional TestRunning As Boolean = True)


    Dim TestPrefix As String
    If TestRunning Then
        TestPrefix = "testingemailonly"
    Else
        TestPrefix = ""
    End If


    Dim RoleInd As Integer: RoleInd = 0

    Dim ScanRow As Integer
    'Iterate over mailing list, if name in mail list found in SourceVal, add Email
    For ScanRow = 2 To LastRow Step 1
        Dim FNameStr As String
        Dim LNameStr As String
        Dim FInitStr As String
        Dim LInitStr As String
        Dim HasPrefName As Boolean
        Dim PrefNameStr As String
        Dim PrefInitStr As String

        'Extract names from ScanRow in MailRange
        With MailSheet
            FNameStr = .Cells(ScanRow, FNameCol)
            LNameStr = .Cells(ScanRow, LNameCol)
            FInitStr = Left(FNameStr, 2)
            If SourceCol = ExProfCol Then
                LInitStr = Left(LNameStr, 3)
            Else
                LInitStr = Left(LNameStr, 1)
            End If
            HasPrefName = (.Cells(ScanRow, PrefNameCol) <> "")

            If HasPrefName Then
                PrefNameStr = .Cells(ScanRow, PrefNameCol)
                PrefInitStr = Left(PrefNameStr, 1)
            Else
                PrefNameStr = ""
                PrefInitStr = ""
            End If
        End With

        'Regular expression objects for searching names
        Dim RegexFirstLInit As Object
        Dim RegexFInitLast As Object
        Dim RegexFirstLInitLead As Object
        Dim RegexFInitLastLead As Object
        Dim RegexFirstLInitTriage As Object
        Dim RegexFInitLastTriage As Object
        Set RegexFirstLInit = CreateObject("VBScript.RegExp")
        Set RegexFInitLast = CreateObject("VBScript.RegExp")
        Set RegexFirstLInitLead = CreateObject("VBScript.RegExp")
        Set RegexFInitLastLead = CreateObject("VBScript.RegExp")
        Set RegexFirstLInitTriage = CreateObject("VBScript.RegExp")
        Set RegexFInitLastTriage = CreateObject("VBScript.RegExp")
        
        'Define patterns used in regular expressions for search
        'One regex for "Firstname L" and one for "F<letters> Lastname"
        Dim InviteeAtts As Object
        Set InviteeAtts = CreateObject("Scripting.Dictionary")

        If HasPrefName Then
            RegexFirstLInit.Pattern = PrefNameStr & "\s*?" & LInitStr
            RegexFInitLast.Pattern = PrefInitStr & "\w*?\s*?" & LNameStr
            RegexFirstLInitLead.Pattern = PrefNameStr & "\s*?" & LInitStr & _
                "\s*,*\s*\(*(Lead)+"
            RegexFInitLastLead.Pattern = PrefInitStr & "\w*?\s*?" & LNameStr & _
                "\s*,*\s*\(*(Lead)+"
            RegexFirstLInitTriage.Pattern = PrefNameStr & "\s*?" & LInitStr & _
                "\s*,*\s*\(*(Triage)+"
            RegexFInitLastTriage.Pattern = PrefInitStr & "\w*?\s*?" & LNameStr & _
                "\s*,*\s*\(*(Triage)+"
        Else
            RegexFirstLInit.Pattern = FNameStr & "\s*?" & LInitStr
            RegexFInitLast.Pattern = FInitStr & "\w*?\s*?" & LNameStr
            RegexFirstLInitLead.Pattern = FNameStr & "\s*?" & LInitStr & _
                "\s*,*\s*\(*(Lead)+"
            RegexFInitLastLead.Pattern = FInitStr & "\w*?\s*?" & LNameStr & _
                "\s*,*\s*\(*(Lead)+"
            RegexFirstLInitTriage.Pattern = FNameStr & "\s*?" & LInitStr & _
                "\s*,*\s*\(*(Triage)+"
            RegexFInitLastTriage.Pattern = FInitStr & "\w*?\s*?" & LNameStr & _
                "\s*,*\s*\(*(Triage)+"
        End If

        Dim InvIsLead As Boolean
        Dim InvIsTriage As Boolean
        
        If RegexFirstLInit.Test(SourceVal) _
            Or RegexFInitLast.Test(SourceVal) Then

                InvIsLead = (RegexFirstLInitLead.Test(SourceVal)) Or _
                    (RegexFInitLastLead.Test(SourceVal))
                InvIsTriage = (RegexFirstLInitTriage.Test(SourceVal)) Or _
                    (RegexFInitLastTriage.Test(SourceVal))

                Call InvNameRoleAtts( _
                    InviteeAtts, _
                    RoleInd, _
                    HasPrefName, _
                    InvIsLead, _
                    InvIsTriage, _
                    FNameStr, _
                    LNameStr, _
                    PrefNameStr, _
                    SourceCol)

                InviteeList.Add _
                    Key:= TestPrefix & MailSheet.Cells(ScanRow, MailCol).Value, _
                    Item:= InviteeAtts
        ElseIf HasPrefName Then
            'If has preferred name but can't find by preferred name, try official name
            RegexFirstLInit.Pattern = FNameStr & "\s+?" & LInitStr
            RegexFInitLast.Pattern = FInitStr & "\w+?\s+?" & LNameStr

            If RegexFirstLInit.Test(SourceVal) _
                Or RegexFInitLast.Test(SourceVal) Then

                InvIsLead = (RegexFirstLInitLead.Test(SourceVal)) Or _
                    (RegexFInitLastLead.Test(SourceVal))
                Call InvNameRoleAtts( _
                    InviteeAtts, _
                    RoleInd, _
                    HasPrefName, _
                    InvIsLead, _
                    InvIsTriage, _
                    FNameStr, _
                    LNameStr, _
                    PrefNameStr, _
                    SourceCol)

                InviteeList.Add _
                    Key:= TestPrefix & MailSheet.Cells(ScanRow, MailCol).Value, _
                    Item:= InviteeAtts
            End If
        End If
    Next ScanRow
    If TestRunning Then
        Dim TestPerson
        Dim TestAtt
        For Each TestPerson in InviteeList
            Debug.Print TestPerson
            For Each TestAtt In InviteeList(TestPerson)
                Debug.Print "    ", TestAtt, InviteeList(TestPerson).Item(TestAtt)
            Next TestAtt
        Next TestPerson
    End If
End Sub

'Create dicionary of dictionaries, key is course name
'Each course dictionary is further keyed by course attributes (eg. "SECTIONS")
Private Sub InviteAtts( _
    StepThruCourses As Integer, _
    Courses As Object, _
    InviteRow As Integer, _
    Optional PreMeetingMins As Integer = 30, _
    Optional T2MeetingMins As Integer = 30)

    Dim CourseInInvite As Integer
    Dim CourseInfo As Object

    Dim RegexTimes As Object
    Set RegexTimes = CreateObject("VBScript.RegExp")
    RegexTimes.Pattern = "\b\d*\d:\d\d"
    RegexTimes.Global = True
    Dim TimeAllMatches
    Dim TimeFound As Integer
    Dim TimePreMeet As String
    Dim TimeT2Meet As String
    Dim TimeExam As String
    Dim TimeEnd As String

    For CourseInInvite = 1 To StepThruCourses Step 1
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

                    If TimeAllMatches.Count = 0 Then
                        TimeExam = FormatDateTime( _
                            ExamSheet.Cells(InviteRow, ExTimeCol), vbShortTime)
                    Else
                        TimeExam = Left(TimeAllMatches.Item(0), 5)
                    End If

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
                        TimeEnd = FormatDateTime( _
                            DateAdd("n", 180, TimeExam), vbShortTime)
                    End If

                    With CourseInfo
                        .Add Key:= ExT2MeetKey, Item:= TimeT2Meet
                        .Add Key:= ExPreMeetKey, Item:= TimePreMeet
                        .Add Key:= ExTimeKey, Item:= TimeExam
                        .Add Key:= ExEndKey, Item:= TimeEnd
                    End With
                End If
            Next CourseAtt

            If .Cells(InviteRow + CourseInInvite - 1, ExCourseCol).Value _
                = .Cells(InviteRow + CourseInInvite - 2, ExCourseCol).Value _
                Or _
                .Cells(InviteRow + CourseInInvite - 1, ExCourseCol).Value _
                = .Cells(InviteRow + CourseInInvite, ExCourseCol).Value _
                Then
                Courses.Add _
                    Key:= .Cells(InviteRow + CourseInInvite - 1, ExCourseCol).Value _
                        & " " & .Cells(InviteRow + CourseInInvite - 1, ExSecCol).Value, _
                    Item:= CourseInfo
            Else
                Courses.Add _
                    Key:= .Cells(InviteRow + CourseInInvite - 1, ExCourseCol).Value, _
                    Item:= CourseInfo
            End If
        End With
    Next CourseInInvite
End Sub

'Drafts invite for one set of course(s) that share a meeting time & room
Public Sub WriteSingleInvite( _
    Optional SendImmediate As Boolean = False, _
    Optional Sender As String = "Sauder LS", _
    Optional InviteRow As Integer = 0, _
    Optional StepThruCourses As Integer, _
    Optional IsDebugging As Boolean = True, _
    Optional OnlyOneInvite As Boolean = True)

    StepThruCourses = 1
    If InviteRow = 0 Then
        InviteRow = ActiveCell.Row
    End If

    If OnlyOneInvite Then
        Call InitMod
    End If

    'Check if the T1 cell is merged, check for length and assign accordingly to nextrow
    With ExamSheet
        If .Cells(InviteRow, ExT2Col).MergeCells Then
            StepThruCourses = .Cells(InviteRow, ExT2Col).MergeArea.Count
        End If
    End With

    Dim IsUpdating As Boolean
    Dim HaveSentInv As Boolean
    Dim HaveSavedInv As Boolean
    Dim NoUpdate As Boolean
    Dim NoDraft As Boolean
    Dim NewStatus As String

    Call CheckIfUpdating( _
        HaveSentInv, _
        HaveSavedInv, _
        NoDraft, _
        NoUpdate, _
        InviteRow)

    Dim DatesBack As Integer
    If IsDebugging Then
        DatesBack = 548
    Else
        DatesBack = 1
    End If

    Dim InvSendTimeStamp As String
    InvSendTimeStamp = FormatDateTime(Now, vbShortDate)

    If NoDraft _
        Or DateDiff("d", _
            InvSendTimeStamp, _
            ExamSheet.Cells(InviteRow, ExDateCol)) _
            <= -1*DatesBack _
        Then
        Exit Sub
    ElseIf NoUpdate And (HaveSentInv Or HaveSavedInv) Then
        Exit Sub
    ElseIf Not(NoUpdate) And (HaveSentInv Or HaveSavedInv) Then
        IsUpdating = True
    Else
        IsUpdating = False
    End If

    If HaveSentInv Or SendImmediate Then
        NewStatus = "sent"
    ElseIf Not(SendImmediate) Then
        NewStatus = "saved"
    End If

    Dim IsOnCall As Boolean
    Call CheckOnCall(IsOnCall, InviteRow)

    Dim Courses As Object
    Set Courses = CreateObject("Scripting.Dictionary")
    
    Call InviteAtts( _
        StepThruCourses, _
        Courses, _
        InviteRow)
    Dim FirstCourse As String
    FirstCourse = Courses.Keys()(0)
    Dim InvSubject As String
    InvSubject = ""
    Dim SubPrefix As String
    If IsDebugging Then
        SubPrefix = "IGNORE TESTING ONLY - "
    Else
        SubPrefix = ""
    End If
    Dim InvPreMeetTime As String
    Dim MaxEndTime As String
    MaxEndTime = FormatDateTime( _
        Courses(FirstCourse).Item(ExEndKey), _
        vbShortTime)

    InvPreMeetTime = FormatDateTime( _
        Courses(FirstCourse).Item(ExPreMeetKey), _
        vbShortTime)

    Dim cours As Variant
    For Each cours In Courses.Keys
        If FormatDateTime(Courses(cours).Item(ExEndKey), vbShortTime) _
            > MaxEndTime Then

            MaxEndTime = _
                FormatDateTime(Courses(cours).Item(ExEndKey), vbShortTime)
        End If
        InvSubject = InvSubject & cours & " " & Courses(cours).Item(ExProfKey) & " "
        If IsDebugging Then
            Dim coursAtt As Variant
            Debug.Print cours
            For Each coursAtt In Courses(cours)
                Debug.Print "    ", coursAtt, Courses(cours).Item(coursAtt)
            Next coursAtt
            Debug.Print "------"
        End If
    Next cours
    MaxEndTime = FormatDateTime(DateAdd("n", 30, MaxEndTime), vbShortTime)
    InvSubject = InvSubject & "EXAM(S) - " & InvPreMeetTime

    'Build an Outlook Invite
    Dim Ot As Outlook.Application
    Set Ot = New Outlook.Application
    Dim OtAppointT2 As Outlook.AppointmentItem
    Dim OtAppointMain As Outlook.AppointmentItem
    Dim OtNameSpace As Outlook.Namespace
    Dim OtAppSender As Outlook.Recipient
    Dim OtFolder As Outlook.MAPIFolder

    Set OtNamespace = Ot.GetNameSpace("MAPI")
    Set OtAppSender = OtNameSpace.CreateRecipient("learning.services@sauder.ubc.ca")
    If OtAppSender.Resolve Then
        Set OtFolder = OtNameSpace.GetSharedDefaultFolder(OtAppSender, olFolderCalendar)
    End If

    Dim InvStatus As String
    InvStatus = ExamSheet.Cells(InviteRow, ExCalCol).Value

    If IsUpdating Then
        Dim OtObj As Object

        Dim OtAppSubject As String
        Dim OtAppDate As String
        Dim OtAppTime As String

        For Each OtObj in OtFolder.Items
            OtAppSubject = OtObj.Subject
            OtAppDate = FormatDateTime(OtObj.Start, vbShortDate)
            OtAppTime = FormatDateTime(OtObj.Start, vbShortTime)


            Dim TotRecip As Integer
            Dim recip As Integer

            If DateDiff("d", InvSendTimeStamp, OtAppDate) >= -1*DatesBack And _
                DateDiff("d", InvSendTimeStamp, OtAppDate) <= 90 Then
                    If DateDiff("n",OtAppDate & " " & OtAppTime,InvPreMeetTime)<=60 Then
                        If OtAppSubject = SubPrefix & "Tier 2 Block - " & InvSubject Then
                            Set OtAppointT2 = OtObj
                            With OtAppointT2
                                TotRecip = .Recipients.Count - 1
                                For recip = 1 to TotRecip
                                    .Recipients.Remove(2)
                                Next recip
                                .Body = ""
                            End With
                        ElseIf Not(IsOnCall) _
                            And OtAppSubject = _
                            SubPrefix & "Pre-exam Meeting - " & InvSubject Then
                            Set OtAppointMain = OtObj
                            With OtAppointMain
                                TotRecip = .Recipients.Count - 1
                                For recip = 1 to TotRecip
                                    .Recipients.Remove(2)
                                Next recip
                                .Body = ""
                            End With
                        End If
                    End If
            End If
        Next OtObj
    End If

    If Not(IsUpdating) Or (OtAppointMain Is Nothing) Then
        IsUpdating = False
        InvStatus = ""
        If Not IsOnCall Then
            Set OtAppointMain = OtFolder.Items.Add(olAppointmentItem)
            OtAppointMain.MeetingStatus = olMeeting
        End If
        Set OtAppointT2 = OtFolder.Items.Add(olAppointmentItem)
        OtAppointT2.MeetingStatus = olMeeting
    End If

    Dim SourceVal As String

    SourceVal = ExamSheet.Cells(InviteRow, ExT2Col).Value
    Dim T2Invitees As Object
    Set T2Invitees = CreateObject("Scripting.Dictionary")

    Call FetchMails( _
        MailLastRow, _
        ExT2Col, _
        SourceVal, _
        T2Invitees)

    Dim T2InvHTML As String
    If Not IsOnCall Then
        Dim MainInvHTML As String
        SourceVal = ExamSheet.Cells(InviteRow, ExT1Col).Value
        Dim T1Invitees As Object
        Set T1Invitees = CreateObject("Scripting.Dictionary")

        Call FetchMails( _
            MailLastRow, _
            ExT1Col, _
            SourceVal, _
            T1Invitees)

        Dim ProfInvitees As Object
        Set ProfInvitees = CreateObject("Scripting.Dictionary")

        Dim Prof As Integer
        SourceVal = ""
        For Prof = 1 To StepThruCourses Step 1
            SourceVal = SourceVal & ExamSheet.Cells(InviteRow + Prof - 1, ExProfCol).Value & " "
        Next Prof

        Call FetchMails( _
            MailLastRow, _
            ExProfCol, _
            SourceVal, _
            ProfInvitees)

        Call WriteHTMLInvBody( _
            MainInvHTML, _
            Courses, _
            IsOnCall, _
            FirstCourse, _
            T2Invitees, _
            T1Invitees)

        Call WriteHTMLInvBody( _
            T2InvHTML, _
            Courses, _
            IsOnCall, _
            FirstCourse, _
            T2Invitees, _
            T1Invitees, _
            IsT2 := True)
    Else
        Call WriteHTMLInvBody( _
            T2InvHTML, _
            Courses, _
            IsOnCall, _
            FirstCourse, _
            T2Invitees, _
            IsT2 := True)
    End If

    Dim HTMLMail As Outlook.MailItem
    Set HTMLMail = Ot.CreateItem(olMailItem)

    With HTMLMail
        .BodyFormat = olFormatHTML
        .HTMLBody = T2InvHTML
        .GetInspector().WordEditor.Range.FormattedText.Copy
    End With
    OtAppointT2.GetInspector().WordEditor.Range.FormattedText.Paste

    Dim PersonMail

    If Not IsOnCall Then
        With HTMLMail
            .BodyFormat = olFormatHTML
            .HTMLBody = MainInvHTML
            .GetInspector().WordEditor.Range.FormattedText.Copy
        End With

        With OtAppointMain
            .GetInspector().WordEditor.Range.FormattedText.Paste
            .Subject = SubPrefix & "Pre-exam Meeting - " & InvSubject
            .Start = _
                FormatDateTime(Courses(FirstCourse).Item(ExDateKey), vbShortDate) & " " & _
                FormatDateTime(Courses(FirstCourse).Item(ExPreMeetKey), vbShortTime)
            .End = _
                FormatDateTime(Courses(FirstCourse).Item(ExDateKey), vbShortDate) & " " & _
                FormatDateTime(Courses(FirstCourse).Item(ExTimeKey), vbShortTime)
            .Location = Courses(FirstCourse).Item(ExRoomKey)
            For Each PersonMail In T1Invitees.Keys
                .Recipients.Add (PersonMail)
            Next PersonMail
            For Each PersonMail In ProfInvitees.Keys
                .Recipients.Add (PersonMail)
            Next PersonMail
            For Each PersonMail In T2Invitees.Keys
                .Recipients.Add (PersonMail)
            Next PersonMail
            If IsDebugging Then
                .Recipients.Add("michael.xiong@sauder.ubc.ca")
            End If
            .Display
            If HaveSentInv Or SendImmediate Then
                .Save
                .Send
            Else
                .Save
                .Close(olSave)
            End If
        End With
        InvStatus = InvStatus & " T1, "
    End If

    With OtAppointT2
        .Subject = SubPrefix & "Tier 2 Block - " & InvSubject
        .Start = _
            FormatDateTime(Courses(FirstCourse).Item(ExDateKey), vbShortDate) & " " & _
            FormatDateTime(Courses(FirstCourse).Item(ExPreMeetKey), vbShortTime)
        .End = _
            FormatDateTime(Courses(FirstCourse).Item(ExDateKey), vbShortDate) & " " & _
            FormatDateTime(MaxEndTime, vbShortTime)
            If IsOnCall Then
                .Location = "(On-Call) " & Courses(FirstCourse).Item(ExRoomKey)
            Else
                .Location = "Pre-meeting at " & Courses(FirstCourse).Item(ExRoomKey)
            End If
        For Each PersonMail In T2Invitees.Keys
            .Recipients.Add (PersonMail)
        Next PersonMail
        If IsDebugging Then
            .Recipients.Add("michael.xiong@sauder.ubc.ca")
        End If
        .Display
        If HaveSentInv Or SendImmediate Then
            .Save
            .Send
        Else
            .Save
            .Close(olSave)
        End If
    End With
    If IsUpdating Then
        InvStatus = InvStatus & "T2 updated " & NewStatus & " by "& Sender & InvSendTimeStamp & "; "
    Else
        InvStatus = InvStatus & "T2 " & NewStatus & " by " & Sender & " " & InvSendTimeStamp & "; "
    End If
    ExamSheet.Cells(InviteRow, ExCalCol).Value = InvStatus

    Dim MSTeamsStr As String
    Call MakeTeamsMsg(MSTeamsStr, InvPreMeetTime, InviteRow, Courses)
    ExamSheet.Cells(InviteRow, ExTeamTextCol).Value = MSTeamsStr

    If OnlyOneInvite Then
        Call EndMod
    End If

End Sub

Private Sub WriteHTMLInvBody ( _
    InvHTML As String, _
    Courses As Object, _
    IsOnCall As Boolean, _
    FirstCourse As String, _
    T2Invitees As Object, _
    Optional T1Invitees As Object, _
    Optional IsT2 As Boolean = False, _
    Optional IsDebugging As Boolean = True)

    Dim T2Lead As String
    Dim Person
    For Each Person In T2Invitees.Keys
        If T2Invitees(Person).Item(InvRoleKey) = "Lead" Then
            T2Lead = T2Invitees(Person).Item(InvNameKey)
        End If
    Next Person

    If Not IsOnCall Then
        If IsT2 Then
            InvHTML = "<p>" & _
                "<a style='font-weight: bold'" & _
                "href='https://teams.microsoft.com/_#/school/files/General?threadId=19%3Af2532fc7eed14adc89b8847d75351c57%40thread.tacv2&amp;ctx=channel&amp;context=Documentation&amp;rootfolder=%252Fteams%252FubcSAUD-gr-LSExamSupport%252FShared%2520Documents%252FGeneral%252FDocumentation' target='_blank' rel='noopener'>" & _
                "Tier 2 Support Guide Here </a></p><p><b>LS Exam Lead: " & _
                T2Lead & "</b></p> <p>" & _
                "The lead will open the Zoom Room and set-up the breakout rooms. " & _
                "They will also create the group chats in MS Teams for T1, T2 and " & _
                "instructor(s) to communicate outside Zoom. </p> <p>"
        Else
            InvHTML = _
                "<p>Hello Everyone,</p> <p>" & _
                "Looking forward to seeing everyone for a brief meeting at: " & _
                "<b>" & Courses(FirstCourse).Item(ExPreMeetKey) & " ~ " & _
                Courses(FirstCourse).Item(ExTimeKey) & "</b></p> <p>"
        End If
        InvHTML = InvHTML & _
        "Support Room: <b>" & Courses(FirstCourse).Item(ExRoomKey) & "</b></p>" & _
        "<p>We will be supporting this exam in a Zoom Room. MS Teams will be used as " & _
        "a communication backchannel for the support staff and the instructors. "
    Else
        InvHTML = _
            "<p><b>This is an on-call exam, you only need to be on MS Teams.</b></p>" & _
            "<p>PLEASE NOTE: On-Call means you aren't required to provide live support." & _
            " Instead you will be backup contact for the actual exam admins. " & _
            "If applicable, on-call staff will create the MS Team chat for " & _
            "communication w/ instructor(s) if applicable. 90% of the time LS " & _
            "is on-call for Siobhan & team at DAP/BUSI, so please add her to chat " & _
            "with the instructor. </p>"
    End If
    InvHTML = InvHTML & _
        "To log into MS Teams, please " & _
        "<a style='font-weight:bold' href='https://teams.microsoft.com/go#'>" & _
        "CLICK HERE</a></p> <p>&nbsp;</p> <p> <a " & _
        "style='background-color: lightyellow; font-weight:bold' " & _
        "href='https://teams.microsoft.com/_#/school/files/General?threadId=19%3Af2532fc7eed14adc89b8847d75351c57%40thread.tacv2&amp;ctx=channel&amp;context=Documentation&amp;rootfolder=%252Fteams%252FubcSAUD-gr-LSExamSupport%252FShared%2520Documents%252FGeneral%252FDocumentation'>" & _
        "EXAM DOCUMENTATION FOR SUPPORT (T1 & T2) CLICK HERE" & _
        "</a> </p> <p>&nbsp;</p> <p><b>EXAM COURSE(S)</b></p> <p>"

    Dim cours
    For Each cours in Courses.Keys
        InvHTML = InvHTML & _
            "</p> <p><b>" & cours & "</b></p> <p>" & _
            "Instructor(s): <b>" & Courses(cours).Item(ExProfKey) & _
            " (" & Courses(cours).Item(ExSecKey) & " Section(s))</b></p>" & _
            "<p>Time: <b>" & Courses(cours).Item(ExTimeKey) & _
            " ~ " & Courses(cours).Item(ExEndKey) & "</b></p>" & _
            "<p>Duration: <b>" & Courses(cours).Item(ExDurKey) & "</b></p>" & _
            "<p>Format: <b>" & Courses(cours).Item(ExFormKey) & "</b></p>" & _
            "<p>Room: <b>" & Courses(cours).Item(ExRoomKey) & "</b></p>" & _
            "<p>Number of Students: <b>" & Courses(cours).Item(ExNumStudKey) & _
            "</b></p> <p>&nbsp;</p>"
    Next cours

    InvHTML = InvHTML & "<p><b>Support Team</b></p> <p><u>" & _
        "<i>Tier 2 Lead - </i>" & T2Lead & "</u></p>"

    For Each Person In T2Invitees.Keys
        If T2Invitees(Person).Item(InvRoleKey) <> "Lead" Then
            InvHTML = InvHTML & _
                "<p>" & _
                T2Invitees(Person).Item(InvRoleKey) & _
                " - " & T2Invitees(Person).Item(InvNameKey) & _
                "</p>"
        End If
    Next Person

    If Not(IsOnCall) Then
        InvHTML = InvHTML & "<p><u>"
        For Each Person In T1Invitees.Keys
            If T1Invitees(Person).Item(InvRoleKey) = "Triage" Then
                InvHTML = InvHTML & _
                    "<i>" & T1Invitees(Person).Item(InvRoleKey) & "</i>" & _
                    " - " & T1Invitees(Person).Item(InvNameKey)
            End If
        Next Person

        InvHTML = InvHTML & "</u></p>"

        For Each Person In T1Invitees.Keys
            If T1Invitees(Person).Item(InvRoleKey) <> "Triage" Then
                InvHTML = InvHTML & _
                    "<p>" & _
                    T1Invitees(Person).Item(InvRoleKey) & _
                    " - " & T1Invitees(Person).Item(InvNameKey) & _
                    "</p>"
            End If
        Next Person
    End If

    InvHTML = InvHTML & "<p><u>"

    If IsDebugging Then
        Debug.Print InvHTML
    End If
End Sub

Private Sub CheckOnCall(IsOnCall As Boolean, _
    InviteRow As Integer)

    Dim RegexOnCall As Object
    Set RegexOnCall = CreateObject("VBScript.RegExp")
    RegexOnCall.Pattern = "on\s*-*\s*call"
    RegexOnCall.IgnoreCase = True

    IsOnCall = RegexOnCall.Test(ExamSheet.Cells(InviteRow, ExT2Col).Value)
End Sub

Private Sub CheckIfUpdating(HaveSentInv As Boolean, _
    HaveSavedInv As Boolean, _
    NoDraft As Boolean, _
    NoUpdate As Boolean, _
    InviteRow As Integer)

    Dim RegexHaveSent As Object
    Set RegexHaveSent = CreateObject("VBScript.RegExp")
    RegexHaveSent.Pattern = "sent"
    RegexHaveSent.IgnoreCase = True
    HaveSentInv = RegexHaveSent.Test(ExamSheet.Cells(InviteRow, ExCalCol))

    Dim RegexHaveSaved As Object
    Set RegexHaveSaved = CreateObject("VBScript.RegExp")
    RegexHaveSaved.Pattern = "saved"
    RegexHaveSaved.IgnoreCase = True
    HaveSavedInv = RegexHaveSaved.Test(Examsheet.Cells(InviteRow, ExCalCol))

    Dim RegexNoUpdate As Object
    Set RegexNoUpdate = CreateObject("VBScript.RegExp")
    RegexNoUpdate.Pattern = "no*"
    RegexNoUpdate.IgnoreCase = True
    NoUpdate = RegexNoUpdate.Test(ExamSheet.Cells(InviteRow, ExAutoUpdateCol))
    NoDraft = RegexNoUpdate.Test(ExamSheet.Cells(InviteRow, ExAutoDraftCol))
End Sub

Public Sub NukeDebugInvites
    Dim Ot As Outlook.Application
    Set Ot = New Outlook.Application
    Dim OtNameSpace As Outlook.Namespace
    Dim OtAppSender As Outlook.Recipient
    Dim OtFolder As Outlook.MAPIFolder

    Set OtNamespace = Ot.GetNameSpace("MAPI")
    Set OtAppSender = OtNameSpace.CreateRecipient("learning.services@sauder.ubc.ca")
    If OtAppSender.Resolve Then
        Set OtFolder = OtNameSpace.GetSharedDefaultFolder(OtAppSender, olFolderCalendar)
    End If

    Dim SubPrefix As String
    SubPrefix = "IGNORE TESTING ONLY - "

    Dim OtObj As Object

    Dim OtAppSubject As String
    Dim OtAppDate As String
    Dim OtAppTime As String

    Dim HasLeftovers As Boolean
    HasLeftovers = True
    Do While HasLeftovers = True
        For Each OtObj in OtFolder.Items
            OtAppSubject = OtObj.Subject
            If InStr(OtAppSubject, SubPrefix) Then
                Debug.Print OtObj.Subject
                OtObj.Delete
            End If
        Next OtObj
        Debug.Print "------"
        HasLeftovers = False
        For Each OtObj in OtFolder.Items
            OtAppSubject = OtObj.Subject
            If InStr(OtAppSubject, SubPrefix) Then
                HasLeftovers = True
            End If
        Next OtObj
    Loop
End Sub

Private Sub CheckDates()
    Dim DateCell As Integer

    For DateCell = 2 To ExLastRow Step 1
        With ExamSheet
            If Not(IsDate(.Cells(DateCell, ExDateCol))) Then
                Call EndMod
                ErrorMsg = "Date format error at: " _
                    & .Cells(DateCell, ExDateCol).Address _
                    & ". Please check to ensure date is written as: " _
                    & "'yyyy-mm-dd' (no quotes) and try again." _
                    & vbNewLine & vbNewLine _
                    & "Check other date cells to ensure correct format as well."
                Err.Raise ErrorNumMod, Description:= ErrorMsg
            End If
        End With
    Next DateCell
End Sub

Private Sub MakeTeamsMsg( _
    MSTeamsStr As String, _
    InvPreMeetTime As String, _
    InviteRow As Integer, _
    Courses As Object)

    MSTeamsStr = "Greetings all! Looking forward to everyone supporting "
    Dim cours As Variant
    For Each cours in Courses.Keys
        MSTeamsStr = MSTeamsStr _
            & cours & " exam w/ " & Courses(cours).Item(ExProfKey) _
            & " using " + Courses(cours).Item(ExFormKey) + ", "
    Next cours
    MSTeamsStr = MSTeamsStr _
        & " hope to see <INSERT NAMES HERE> on " & ExamSheet.Cells(InviteRow, ExDateCol) _
        & " at " & InvPreMeetTime + " for a quick pre-meeting. " _
        & "Let us know if you have any questions. Thank you everyone!"
End Sub

Private Sub SaveOutput()
    Dim SaveTime As String
    SaveTime = FormatDateTime(Now, vbLongDate) _
        & "_" & FormatDateTime(Now, vbShortTime)
    SaveTime = Replace(SaveTime, "/", "")
    SaveTime = Replace(SaveTime, " ", "_")
    SaveTime = Replace(SaveTime, ":", "")
    SaveTime = Replace(SaveTime, ",", "")
    SaveTime = Replace(SaveTime, ".", "")
    Dim SaveDir As String
    SaveDir = ActiveWorkbook.Path & "\output\"

    If DIR(SaveDir, vbDirectory) = "" Then
        MkDir(SaveDir)
    End If

    Dim SaveName As String
    SaveName = SaveDir & "output_" & SaveTime
    WB.SaveAs FileName:=SaveName, FileFormat:=52
End Sub
