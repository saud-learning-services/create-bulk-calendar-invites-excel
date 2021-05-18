'Attribute VB_Name = "Send_Invite_Module"
Option Explicit

'Initialize Global Variables
Dim WB As Workbook
Dim ExamSheet As Worksheet
Dim T1Mails As Worksheet
Dim T2Mails As Worksheet
Dim T1Range As Range
Dim T1FNames As Range
Dim ZoomRooms As Worksheet
Dim EST1ColN As Integer
Dim EST2ColN As Integer
Dim ESLColN As Integer
Dim ESLRowN As Integer
Dim T1LColN As Integer
Dim T1LRowN As Integer
Dim T2LColN As Integer
Dim T2LRowN As Integer


Private Sub InitVars()
    'Initialize Workbook and Sheets
    Set WB = ThisWorkbook
    Set ExamSheet = WB.Sheets("Exam Sheet")
    Set T1Mails  = WB.Sheets("Tier 1 Email List")
    Set T2Mails = WB.Sheets("Tier 2 Email List")
    Set ZoomRooms = WB.Sheets("Zoom Rooms")

    EST1ColN = ExamSheet.Range("A1:AD1").Find _
        ("TIER 1", LookIn:=xlValues, MatchCase:=False).Column

    ESLRowN = ExamSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ESLColN = ExamSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    T1LRowN = T1Mails.Cells(Rows.Count, 1).End(xlUp).Row
    T1LColN = T1Mails.Cells(1, Columns.Count).End(xlToLeft).Column
    With T1Mails
        Set T1Range = T1Mails.Range(.Cells(1, 1), _
            .Cells(T1LRowN, T1LColN))
        Set T1FNames = T1Mails.Range(.Cells(1, 1), _
            .Cells(T1LRowN, 1))
    End With
End Sub

'This Sub sends one invite, for the row where the active cell is
'OR to be used as part of loop
Public Sub SendInvite(Optional CurrRow As Integer = 0)
    Call InitVars
    
    'If CurrRow has default val: means sending current row, assign current row
    If (CurrRow = 0) Then
        CurrRow = ActiveCell.Row
    End If

    'Declare and build appointment
    Dim Otlook As Outlook.Application
    Set Otlook = New Outlook.Application
    Dim OtAppoint As Outlook.AppointmentItem
    Set OtAppoint = Otlook.CreateItem(olAppointmentItem)
    OtAppoint.MeetingStatus = olMeeting

    'Pull From ExamSheet CurrRow Cell with T1
    Dim T1Source As String
    T1Source = ExamSheet.Cells(CurrRow, EST1ColN).Value
    
    'Array with each T1 Role, Person
    Dim T1List As Variant
    T1List = Split(T1Source, ";")

    'Iterate over each T1 Role, Person
    Dim T1Per As Variant
    For Each T1Per In T1List
        Dim T1RoleName As Variant 'Array of a single role, person
        Dim T1PerName As String 'String of a single name
        Dim T1PerMail As String 'Mail to pull from T1Mails

        T1Per = Replace(T1Per, " ","")
        T1RoleName = Split(T1Per, ",")
        T1PerName = T1RoleName(1)
        
        'Find Email of current person
        With Application.WorksheetFunction
            T1PerMail = .Index(T1Range, _
                .Match(T1PerName, T1FNames, 0), T1LColN)
            Debug.Print T1PerMail
        End With
        OtAppoint.Recipients.Add(T1PerMail) 'Add recipient to Invite
        'To add emails to invite
    Next T1Per

    'Next
        'Automate Time
        'Automate Invite meeting link
        'More sophisticated Body
            'Course Name
            'Instructor
            'Time
            'Zoom room
            'Support Staff List
    With OtAppoint
        .Subject = "IGNORE - Testing Only"
        .Body = "Please ignore, this is for testing only :]"
        .Start = #5/19/2021 5:00:00 AM#
        .Duration = 1
        .Save
        '.Send 'This will send the invite right-away
    End With
End Sub

Public Sub CallSubs()
    Call SendInvite
End Sub
