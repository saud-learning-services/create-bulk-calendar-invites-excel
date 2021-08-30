VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInviteSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member variables
Private m_invDate As String
Private m_invStart As String
Private m_invPMEnd As String
Private m_t As String
Private Const m_invPMDur As Integer = 30 'minutes
Private m_invEnd As String
Private Const m_invEndFlex As Integer = 15 'minutes
Private m_ot As Outlook.Application
Private m_invHTMLFormatter As Outlook.MailItem
Private m_otAptT2 As Outlook.AppointmentItem
Private m_otAptPM As Outlook.AppointmentItem
Private m_otNameSpace As Outlook.Namespace
Private m_otAptMaker As Outlook.Recipient
Private m_otFolder As Outlook.Folder
Private Const m_otFrom As String = "learning.services@sauder.ubc.ca"
Private m_invDrafting As Boolean
Private m_invUpdating As Boolean
Private m_invIsOnline As Boolean
Private m_invIsOnCall As Boolean
Private m_invIsDAP As Boolean
Private m_invHasPM As Boolean
Private m_fromRow As Integer
Private m_numCourses As Integer
Private m_alphArr
Private m_invT2Lett As Integer
Private m_invT1Lett As Integer
Private m_errMsg As String
Private m_errNum As Long
Private m_invT1 As Dictionary
Private m_invT2 As Dictionary
Private m_invCourses As Collection
Private m_invLoc As String
Private m_isDebugging As Boolean
Private m_skip As Boolean
Private m_invTitle As String
Private m_invPMTitle As String
Private m_invDraftTime As String
Private m_invDatesBack As Integer

Private m_exSheet As clsExamSheet
Private m_mlSheet As clsMailSheet
Private m_logger As clsLogger
Private m_settings As clsSettings

Private Sub Class_Initialize()
    m_errMsg = "An error occurred while using property, method or function in" _
        & " class module 'clsInviteSession'"
    m_errNum = 42007
    m_invStart = "--:--"
    m_invEnd = "--:--"
    m_invPMEnd = "--:--"
    m_invT1Lett = 0
    m_invT2Lett = 0
    m_alphArr = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", _
        "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
End Sub

Public Sub ConstructProperties( _
    Logger As clsLogger, _
    ExamSheet As clsExamSheet, _
    MailSheet As clsMailSheet, _
    Settings As clsSettings, _
    RowNum As Integer, _
    Optional CoursesInInv As Integer = 1, _
    Optional AllowUpdates As Boolean = False)

    'Basic Settings
    Set m_logger = Logger
    Set m_settings = Settings
    m_isDebugging = Settings.isDebugging
    m_fromRow = RowNum
    m_numCourses = CoursesInInv
    Set m_exSheet = ExamSheet
    m_logger.LogInfo "Constructing properties of invite session for '" & _
        m_exSheet.CourseInfo(m_fromRow) & "' to '" & _
        m_exSheet.CourseInfo(m_fromRow + m_numCourses - 1) & "'"
        m_logger.LogDebug "Constructing basic settings"
    m_invDrafting = CheckIfDrafting()
    If AllowUpdates Then
        m_invUpdating = CheckIfUpdating()
    Else
        m_invUpdating = False
    End If
    If Not (m_invDrafting) And Not (m_invUpdating) Then
        m_skip = True
        m_logger.LogInfo "Skipping this invite as not updating nor drafting"
        Exit Sub
    End If
    m_invIsDAP = CheckIfDAP()
    m_invIsOnline = CheckIfOnline()
    m_invIsOnCall = CheckIfOnCall()
    m_invHasPM = CheckIfHasPM()
    m_logger.LogDebug "Successfully configured basic settings"
    
    m_logger.LogDebug "Constructing courses for time & place info"
    Set m_invCourses = New Collection
    AddCourses
    GetInviteTimes
    GetInviteDate
    m_invDraftTime = FormatDateTime(Now(), vbShortDate)
    If m_isDebugging Then
        m_t = "pleaseIgnoreTestingOnly"
        m_invDatesBack = 150
    Else
        m_invDatesBack = 1
    End If
    m_invLoc = m_exSheet.RoomInfo(m_fromRow)
    m_logger.LogDebug "Raw invite location from exam sheet '" & m_invLoc & "'"
    m_invTitle = MakeInviteTitle()
    If m_invHasPM Then
        m_invPMTitle = MakeInviteTitle(ForTier2:=False)
    End If
    m_logger.LogDebug "Successfully constructed courses, time & space info"
    If m_invUpdating Then
        m_logger.LogInfo "Finding previous invites to update"
        GetPrevAppointments
    End If
    
    'Attendee properties from exam sheet and mail sheet
    m_logger.LogDebug "Constructing attendees to invite"
    Dim T2Str As String: T2Str = m_exSheet.T2Info(m_fromRow)
    m_logger.LogDebug "Raw tier 2 attendees information from exam sheet '" _
        & T2Str & "'"
    Dim T1Str As String: T1Str = m_exSheet.T1Info(m_fromRow)
    m_logger.LogDebug "Raw tier 1 attendees information from exam sheet '" _
        & T1Str & "'"
    Set m_mlSheet = MailSheet
    Set m_invT1 = New Dictionary
    Set m_invT2 = New Dictionary
    GetAttendees T1Str, T2Str 'who
    m_logger.LogDebug "Successfully constructed attendees to invite"
    m_logger.LogInfo "Successfully constructed invite properties"
End Sub

Public Sub WriteAppointments()
    If m_otAptT2 Is Nothing Then
        m_logger.LogInfo "To draft new tier 2 invite, drafting with address '" _
            & m_otFrom & "' for '" & m_invTitle & "'"
        m_invUpdating = False
        Set m_ot = New Outlook.Application
        Set m_otNameSpace = m_ot.GetNamespace("MAPI")
        Set m_otAptMaker = m_otNameSpace.CreateRecipient(m_otFrom)
        If m_otAptMaker.Resolve Then
            m_logger.LogDebug "Address valid, getting shared calendar folder"
            Set m_otFolder = _
                m_otNameSpace.GetSharedDefaultFolder(m_otAptMaker, olFolderCalendar)
            m_logger.LogDebug "Sucessfully got shared calendar folder"
        End If
        m_logger.LogDebug "Adding new blank Tier 2 invite to folder"
        Set m_otAptT2 = m_otFolder.Items.Add(olAppointmentItem)
        m_otAptT2.MeetingStatus = olMeeting
        m_logger.LogDebug "Successfully added new blank Tier 2 invite to folder"
        
        If m_invHasPM Then
            m_logger.LogInfo "To draft new pre-exam invite for '" & m_invPMTitle & "'"
            m_logger.LogDebug "Adding new blank Pre-exam invite to folder"
            Set m_otAptPM = m_otFolder.Items.Add(olAppointmentItem)
            m_otAptPM.MeetingStatus = olMeeting
            m_logger.LogDebug "Successfully added new blank Pre-exam invite to folder"
        End If
    End If

    m_exSheet.ExamPeopleInvited(m_fromRow) = ""
    
    Dim htmlBody As String
    m_logger.LogDebug "Using mail item to format tier 2 invite body"
    htmlBody = GetInviteBody(ForTier2:=True)
    Set m_invHTMLFormatter = m_ot.CreateItem(olMailItem)
    With m_invHTMLFormatter
        .BodyFormat = olFormatHTML: m_logger.LogDebug "Set format", indt:=1
        .htmlBody = htmlBody
        .GetInspector().WordEditor.Range.FormattedText.Copy
        m_logger.LogDebug "Body copied", indt:=1
    End With
    m_otAptT2.GetInspector().WordEditor.Range.FormattedText.Paste
    m_logger.LogDebug "Successfully pasted formatted body to tier 2 invite"
    
    Dim role
    Dim newInvStatus As String
    newInvStatus = m_invDraftTime & " by " & _
        m_settings.SenderName & " - "
    If m_invHasPM Then
        m_logger.LogDebug "Using mail item to format pre-exam invite body"
        htmlBody = GetInviteBody(ForTier2:=False)
        With m_invHTMLFormatter
            .BodyFormat = olFormatHTML: m_logger.LogDebug "Set format", indt:=1
            .htmlBody = GetInviteBody(ForTier2:=False)
            .GetInspector().WordEditor.Range.FormattedText.Copy
            m_logger.LogDebug "Body copied", indt:=1
        End With
        With m_otAptPM
            .GetInspector().WordEditor.Range.FormattedText.Paste
            m_logger.LogDebug "Successfully pasted formatted body to pre-exam invite"
            m_logger.LogInfo "Constructing remaining pre-exam invite parts"
            .Subject = m_invPMTitle
            m_logger.LogInfo "Title set '" & .Subject & "'", indt:=1
            .Start = FormatDateTime(m_invDate, vbShortDate) & " " & _
                FormatDateTime(m_invStart, vbShortTime)
            m_logger.LogInfo "Start time '" & .Start & "'", indt:=1
            .End = FormatDateTime(m_invDate, vbShortDate) & " " & _
                FormatDateTime(m_invPMEnd, vbShortTime)
            m_logger.LogInfo "End time '" & .End & "'", indt:=1
            .Location = m_invLoc
            m_logger.LogInfo "Location '" & .Location & "'", indt:=1
            m_logger.LogDebug "Adding tier 1 attendees", indt:=1
            For Each role In m_invT1
                .Recipients.Add (m_t & m_invT1.Item(role).Email)
                m_logger.LogInfo "Added '" & m_t & m_invT1.Item(role).Email & "'", indt:=1
                m_exSheet.ExamPeopleInvited(m_fromRow) = m_exSheet.ExamPeopleInvited(m_fromRow) & _
                    m_invT1.Item(role).Fullname & "; "
            Next role
            m_logger.LogDebug "Successfully added tier 1 attendees", indt:=1
            m_logger.LogDebug "Adding tier 2 attendees", indt:=1
            For Each role In m_invT2
                .Recipients.Add (m_t & m_invT2.Item(role).Email)
                m_logger.LogInfo "Added '" & m_t & m_invT2.Item(role).Email & "'", indt:=1
                m_exSheet.ExamPeopleInvited(m_fromRow) = m_exSheet.ExamPeopleInvited(m_fromRow) & _
                    m_invT2.Item(role).Fullname & "; "
            Next role
            m_logger.LogDebug "Successfully added tier 2 attendees", indt:=1
            If m_isDebugging Then
                .Recipients.Add ("michael.xiong@sauder.ubc.ca")
            End If
            .Display
            If m_settings.SendImmediate Or CheckPrevCalInv = "sent" Then
                .Save
                .Send
                m_logger.LogInfo "Successfully sent '" & m_invTitle & "'"
            Else
                .Save
                .Close (olSave)
                m_logger.LogInfo "Successfully drafted and saved '" & m_invTitle & "'"
            End If
        End With
        newInvStatus = newInvStatus & "T1 "
    End If
    With m_otAptT2
        m_logger.LogInfo "Constructing remaining tier 2 invite parts"
        .Subject = m_invTitle
        m_logger.LogInfo "Title set '" & .Subject & "'", indt:=1
        .Start = FormatDateTime(m_invDate, vbShortDate) & " " & _
            FormatDateTime(m_invStart, vbShortTime)
        m_logger.LogInfo "Start time '" & .Start & "'", indt:=1
        .End = FormatDateTime(m_invDate, vbShortDate) & " " & _
            FormatDateTime(m_invEnd, vbShortTime)
        m_logger.LogInfo "End time '" & .End & "'", indt:=1
        .Location = m_invLoc
        m_logger.LogInfo "Location '" & .Location & "'", indt:=1
        m_logger.LogDebug "Adding tier 2 attendees", indt:=1
        For Each role In m_invT2
            .Recipients.Add (m_t & m_invT2.Item(role).Email)
            m_logger.LogInfo "Added '" & m_t & m_invT2.Item(role).Email & "'", indt:=1
            If Not (m_invHasPM) Then
                m_exSheet.ExamPeopleInvited(m_fromRow) = m_exSheet.ExamPeopleInvited(m_fromRow) & _
                    m_invT2.Item(role).Fullname & "; "
            End If
        Next role
        m_logger.LogDebug "Successfully added tier 2 attendees", indt:=1
        If m_isDebugging Then
            .Recipients.Add ("michael.xiong@sauder.ubc.ca")
        End If
        .Display
        If m_settings.SendImmediate Or CheckPrevCalInv = "sent" Then
            .Save
            .Send
            newInvStatus = newInvStatus & "T2 sent; "
            m_logger.LogInfo "Successfully sent '" & m_invTitle & "'"
        Else
            .Save
            .Close (olSave)
            newInvStatus = newInvStatus & "T2 drafted / saved; "
            m_logger.LogInfo "Successfully drafted and saved '" & m_invTitle & "'"
        End If
    End With
    m_exSheet.CalStatusInfo(m_fromRow) = newInvStatus
    m_logger.LogInfo "Updated calendar invite status '" & newInvStatus & "'"
End Sub

Public Sub GetPrevAppointments()
    Set m_ot = New Outlook.Application
    Set m_otNameSpace = m_ot.GetNamespace("MAPI")
    Set m_otAptMaker = m_otNameSpace.CreateRecipient(m_otFrom)
    If m_otAptMaker.Resolve Then
        Set m_otFolder = _
            m_otNameSpace.GetSharedDefaultFolder(m_otAptMaker, olFolderCalendar)
    End If
    Dim OtCalInv As Outlook.AppointmentItem
    Dim OtCalSub As String
    Dim OtCalStart As String
    Dim OtCalDate As String
    m_logger.LogDebug "Scanning calendar folder '" & m_otFolder.Name & "'"
    For Each OtCalInv In m_otFolder.Items
        OtCalSub = OtCalInv.Subject
        m_logger.LogDebug "Checking '" & OtCalSub & "'", indt:=1
        OtCalStart = FormatDateTime(OtCalInv.Start, vbShortTime)
        OtCalDate = FormatDateTime(OtCalInv.Start, vbShortDate)
        
        Dim recip
        If DateDiff("d", m_invDraftTime, OtCalDate) >= -1 * m_invDatesBack _
            And DateDiff("d", m_invDraftTime, OtCalDate) <= 90 _
            And DateDiff("n", OtCalStart, m_invStart) <= 60 Then
            If OtCalSub = m_invTitle Then
                Set m_otAptT2 = OtCalInv
                m_logger.LogInfo "Found '" & OtCalSub & "'", indt:=1
                With m_otAptT2
                    For recip = 1 To .Recipients.Count - 1
                        .Recipients.Remove (2)
                    Next recip
                End With
                m_logger.LogDebug "Reset recipeints", indt:=1
            ElseIf Not (m_invIsOnCall) And OtCalSub = m_invPMTitle Then
                m_logger.LogInfo "Found '" & OtCalSub & "'", indt:=1
                Set m_otAptPM = OtCalInv
                With m_otAptPM
                    For recip = 1 To .Recipients.Count - 1
                        .Recipients.Remove (2)
                    Next recip
                End With
                m_logger.LogDebug "Reset recipeints", indt:=1
            End If
        End If
    Next OtCalInv
End Sub

Public Sub GetAttendees(T1Str As String, T2Str As String)
    m_logger.LogDebug "Checking which attendees are in the invite"
    Dim att As clsAttendee
    Dim attRole As String
    Dim reResult As String
    For Each att In m_mlSheet.Attendees
        reResult = att.CheckNameRe(T2Str)
        Select Case reResult
            Case "Lead"
                attRole = "Lead"
                m_invT2.Add Key:=attRole, Item:=att
                m_logger.LogDebug "Noting attendee '" & att.Fullname & _
                    "' with role '" & attRole & "'"
            Case "True"
                attRole = GetAttNum(2)
                m_invT2.Add Key:=attRole, Item:=att
                m_logger.LogDebug "Noting attendee '" & att.Fullname & _
                    "' with role '" & attRole & "'"
        End Select
        reResult = att.CheckNameRe(T1Str)
        Select Case reResult
            Case "Triage"
                attRole = "Triage"
                m_invT1.Add Key:=attRole, Item:=att
                m_logger.LogDebug "Noting attendee '" & att.Fullname & _
                    "' with role '" & attRole & "'"
            Case "True"
                attRole = GetAttNum(1)
                m_invT1.Add Key:=attRole, Item:=att
                m_logger.LogDebug "Noting attendee '" & att.Fullname & _
                    "' with role '" & attRole & "'"
        End Select
    Next att
    Dim role
    m_logger.LogInfo "Following Tier 2 attendees will be invited: "
    For Each role In m_invT2
        m_logger.LogInfo "'" & role & "' - '" & m_invT2.Item(role).Fullname & _
            "' - '" & m_invT2.Item(role).Email & "'", _
            indt:=1
    Next role
    m_logger.LogInfo "Following Tier 1 attendees will be invited: "
    For Each role In m_invT1
        m_logger.LogInfo "'" & role & "' - '" & m_invT1.Item(role).Fullname & _
            "' - '" & m_invT1.Item(role).Email & "'", _
            indt:=1
    Next role
End Sub

Public Sub AddCourses()
    m_logger.LogInfo "Constructing courses starting row '" & m_fromRow _
    & "' to row '" & CStr(m_fromRow + m_numCourses - 1) & "'"
    Dim RowStep As Integer
    For RowStep = m_fromRow To m_fromRow + m_numCourses - 1
        m_logger.LogDebug "Adding course to invite session from row '" & _
        m_exSheet.CourseInfo(RowStep) & "'"
        Dim newCourse As clsCourse
        Set newCourse = New clsCourse
        newCourse.ConstructCourse m_exSheet, RowStep, m_logger
        m_invCourses.Add newCourse
        m_logger.LogInfo "Successfully added '" & newCourse.Name & "' from row '" & _
        m_exSheet.CourseInfo(RowStep) & "'", indt:=1
    Next RowStep
    m_logger.LogInfo "Finished constructing courses"
End Sub

Public Sub GetInviteTimes()
    Dim StartTime As String: StartTime = FormatDateTime("23:00", vbShortTime)
    Dim PMEndTime As String: PMEndTime = FormatDateTime("23:00", vbShortTime)
    Dim EndTime As String: EndTime = FormatDateTime("01:00", vbShortTime)
    Dim CourseInInv As clsCourse
    m_logger.LogDebug "Computing invite times from exam time(s)"
    For Each CourseInInv In m_invCourses
        With CourseInInv
            m_logger.LogDebug "Checking course '" & .Name _
                & "' exam start '" & .ExamStartTime & "'"
            Select Case .ExamStartTime
                Case Is < FormatDateTime(DateAdd("n", m_invPMDur, StartTime), vbShortTime)
                    m_logger.LogDebug "Exam start time earlier than previously " & _
                        "scanned start", _
                        indt:=1
                    StartTime = FormatDateTime( _
                        DateAdd("n", -1 * m_invPMDur, .ExamStartTime), vbShortTime)
                    PMEndTime = .ExamStartTime
                    m_logger.LogDebug "New invite start '" & StartTime & "', '" & _
                        "pre-meeting until '" & PMEndTime & "'", indt:=1
            End Select
            m_logger.LogDebug "Checking course '" & .Name _
                & "' exam end '" & .ExamEndTime & "'"
            Select Case .ExamEndTime
                Case Is > FormatDateTime(DateAdd("n", -1 * m_invEndFlex, EndTime), vbShortTime)
                    m_logger.LogDebug "Exam end time later than previously " & _
                        "scanned end", _
                        indt:=1
                    EndTime = FormatDateTime( _
                        DateAdd("n", m_invEndFlex, .ExamEndTime), vbShortTime)
                    m_logger.LogDebug "New invite end '" & EndTime & "'", _
                        indt:=1
            End Select
        End With
    Next CourseInInv
    If StartTime = "" Or EndTime = "" Then
        m_errMsg = "One of the time is empty"
        m_logger.LogError m_errMsg, Err, m_errNum
        Err.Raise m_errNum, Description:=m_errMsg
    End If
    If FormatDateTime(StartTime, vbShortTime) > _
        FormatDateTime(EndTime, vbShortTime) Then
        m_errMsg = "Start time is later than end time"
        m_logger.LogError m_errMsg, Err, m_errNum
        Err.Raise m_errNum, Description:=m_errMsg
    End If
    m_invStart = StartTime
    m_invEnd = EndTime
    m_invPMEnd = PMEndTime
    m_logger.LogDebug "Invite starts at: '" & m_invStart & "'"
    m_logger.LogDebug "Pre-meeting ends at: '" & m_invPMEnd & "'"
    m_logger.LogDebug "Invite ends at: '" & m_invEnd & "'"
End Sub

Public Sub GetInviteDate()
    On Error GoTo CannotGetDate
    Dim InvDate 'Assume date in sheet is formatted eg "Monday, July 6, 2021"
    m_logger.LogDebug "Getting date from '" & m_exSheet.DateInfo(m_fromRow) & "'"
    InvDate = Split(m_exSheet.DateInfo(m_fromRow), ",")
    m_invDate = FormatDateTime(InvDate(1) & " " & InvDate(2), vbShortDate)
    m_logger.LogDebug "Successfully got date '" & m_invDate & "'"
    Exit Sub
CannotGetDate:
    m_errMsg = "Unable to get date from '" & m_exSheet.DateInfo(m_fromRow) & "', " _
        & "expect dates of format 'dddd, mmmm dd yyyy' e.g. 'Monday, July 6, 2021'"
    m_logger.LogError m_errMsg, Err
    Err.Raise m_errNum, Description:=m_errMsg
End Sub

Public Function GetAttNum(TierNum As Integer) As String
    m_logger.LogDebug "Generating attendee number of tier '" & TierNum & "'"
    Select Case TierNum
        Case 1
            GetAttNum = "Tier" & TierNum & m_alphArr(m_invT1Lett)
            m_invT1Lett = m_invT1Lett + 1
        Case 2
            GetAttNum = "Tier" & TierNum & m_alphArr(m_invT2Lett)
            m_invT2Lett = m_invT2Lett + 1
    End Select
    m_logger.LogDebug "Generated '" & GetAttNum & "'"
End Function

Public Function htmlPreMeetOnline() As String
    m_logger.LogDebug "Making online exam pre-meeting html body"
    Dim Txt As String
    Txt = "<p>Hello,</p>"
    Txt = Txt & "<p>Looking forward to seeing you all for a brief meeting at " & _
        m_invStart & "</p>"
    Txt = Txt & "<p>Support Room Link: " & m_invLoc & "</p>"
    Txt = Txt & "<p>" & _
        "We will be supporting this exam in Zoom. MS Teams will be used as a " & _
        "communication backchannel for the support staff and the instructors. " & _
        "To log into MS Teams, please " & _
        "<a href='https://teams.microsoft.com/go#'>" & _
        "CLICK HERE</a></p>"
    Txt = Txt & "<p>" & _
        "<a href='https://teams.microsoft.com/_#/school/files/General?threadId=19%3Af2532fc7eed14adc89b8847d75351c57%40thread.tacv2&ctx=channel&context=Documentation&rootfolder=%252Fteams%252FubcSAUD-gr-LSExamSupport%252FShared%2520Documents%252FGeneral%252FDocumentation'>" & _
        "EXAM SUPPORT DOCUMENTATION FOR TIER 1 / TIER 2 — CLICK HERE</a></p>"

    m_logger.LogDebug "Adding course information to body", indt:=1
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        Txt = Txt & "<p><b>" & CourseInInv.Name & "</b><br>"
        Txt = Txt & "Section(s): " & CourseInInv.Section & "<br>"
        Txt = Txt & "Instructor(s): " & CourseInInv.Instructor & "<br>"
        Txt = Txt & "Start: " & CourseInInv.ExamStartTime & "<br>"
        Txt = Txt & "End: " & CourseInInv.ExamEndTime & "<br>"
        Txt = Txt & "Duration: " & CourseInInv.ExamDuration & "<br>"
        Txt = Txt & "Format: " & CourseInInv.ExamFormat & "</p>"
    Next CourseInInv
    
    m_logger.LogDebug "Adding attendee information to body", indt:=1
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    Dim role
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    If m_invT1.Exists("Triage") Then
        Txt = Txt & "<p>Triage: " & m_invT1.Item("Triage").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    For Each role In m_invT1
        If role <> "Triage" Then
            Txt = Txt & role & ": " & m_invT1.Item(role).Fullname & "<br>"
        End If
    Next role
    Txt = Txt & "</p>"
    
    htmlPreMeetOnline = Txt
    m_logger.LogDebug "Successfully made html body for online exam pre-meeting"
    m_logger.LogDebug "Body as follows" & _
        vbNewLine & "'''" & vbNewLine & Txt & vbNewLine & "'''", indt:=1
End Function

Public Function htmlT2Online() As String
    m_logger.LogDebug "Making online exam tier 2 html body"
    Dim Txt As String
    Txt = "<p>The Tier 2 Guide can be found " & _
        "<a href='https://teams.microsoft.com/_#/school/files/General?threadId=19%3Af2532fc7eed14adc89b8847d75351c57%40thread.tacv2&ctx=channel&context=Documentation&rootfolder=%252Fteams%252FubcSAUD-gr-LSExamSupport%252FShared%2520Documents%252FGeneral%252FDocumentation'>" & _
        "here</a>.</p>"
    
    Txt = Txt & "<p>Support room link: " & m_invLoc & "</p>"
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>LS Exam Lead: " & m_invT2.Item("Lead").Fullname & "</p>"
    End If
    Txt = Txt & _
        "<p>The exam lead will open the Zoom room, as well as set up the breakout rooms. <br>" & _
        "They will also create the group chats in MS Teams for Tier 1, Tier 2, " & _
        "and the instructor(s) to communicate outside of Zoom.</p>"
    
    m_logger.LogDebug "Adding course information to body", indt:=1
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        Txt = Txt & "<p><b>" & CourseInInv.Name & "</b><br>"
        Txt = Txt & "Section(s): " & CourseInInv.Section & "<br>"
        Txt = Txt & "Instructor(s): " & CourseInInv.Instructor & "<br>"
        Txt = Txt & "Start: " & CourseInInv.ExamStartTime & "<br>"
        Txt = Txt & "End: " & CourseInInv.ExamEndTime & "<br>"
        Txt = Txt & "Duration: " & CourseInInv.ExamDuration & "<br>"
        Txt = Txt & "Format: " & CourseInInv.ExamFormat & "<br>"
        Txt = Txt & "Number of students: " & CourseInInv.ExamNumStudents & "<br>"
        Txt = Txt & "Exam link: " & _
            "<a href='" & CourseInInv.ExamLink & "'>" & "exam</a></p>"
    Next CourseInInv
    
    Txt = Txt & "<p>CfAs: </p>"
    Txt = Txt & "<p>Alternates: </p>"
    
    m_logger.LogDebug "Adding attendee information to body", indt:=1
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    Dim role
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    If m_invT1.Exists("Triage") Then
        Txt = Txt & "<p>Triage: " & m_invT1.Item("Triage").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    For Each role In m_invT1
        If role <> "Triage" Then
            Txt = Txt & role & ": " & m_invT1.Item(role).Fullname & "<br>"
        End If
    Next role


    Txt = Txt & "</p>"
    htmlT2Online = Txt
    m_logger.LogDebug "Successfully made html body for tier 2 online exam invite"
    m_logger.LogDebug "Body as follows" & _
        vbNewLine & "'''" & vbNewLine & Txt & vbNewLine & "'''", indt:=1
End Function

Public Function htmlPreMeetInperson()
    m_logger.LogDebug "Making in-person exam pre-meeting html body"
    Dim Txt As String
    Txt = "<p>Hello,</p>"
    Txt = Txt & "<p>Looking forward to seeing you all for a brief meeting at " & _
        m_invStart & "</p>"
    Txt = Txt & "<p>Support Room: " & m_invLoc & "</p>"

    m_logger.LogDebug "Adding course information to body", indt:=1
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        Txt = Txt & "<p><b>" & CourseInInv.Name & "</b><br>"
        Txt = Txt & "Section(s): " & CourseInInv.Section & "<br>"
        Txt = Txt & "Instructor(s): " & CourseInInv.Instructor & "<br>"
        Txt = Txt & "Start: " & CourseInInv.ExamStartTime & "<br>"
        Txt = Txt & "End: " & CourseInInv.ExamEndTime & "<br>"
        Txt = Txt & "Duration: " & CourseInInv.ExamDuration & "<br>"
        Txt = Txt & "Format: " & CourseInInv.ExamFormat & "</p>"
    Next CourseInInv
    
    m_logger.LogDebug "Adding attendee information to body", indt:=1
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    Dim role
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    If m_invT1.Exists("Triage") Then
        Txt = Txt & "<p>Triage: " & m_invT1.Item("Triage").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    For Each role In m_invT1
        If role <> "Triage" Then
            Txt = Txt & role & ": " & m_invT1.Item(role).Fullname & "<br>"
        End If
    Next role
    Txt = Txt & "</p>"
    
    htmlPreMeetInperson = Txt
    m_logger.LogDebug "Successfully made html body for in-person exam pre-meeting"
    m_logger.LogDebug "Body as follows" & _
        vbNewLine & "'''" & vbNewLine & Txt & vbNewLine & "'''", indt:=1
End Function

Public Function htmlT2Inperson() As String
    m_logger.LogDebug "Making in-person exam tier 2 html body"
    Dim Txt As String
    Txt = Txt & "<p>Support room: " & m_invLoc & "</p>"
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>LS Exam Lead: " & m_invT2.Item("Lead").Fullname & "</p>"
    End If
    
    m_logger.LogDebug "Adding course information to body", indt:=1
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        Txt = Txt & "<p><b>" & CourseInInv.Name & "</b><br>"
        Txt = Txt & "Section(s): " & CourseInInv.Section & "<br>"
        Txt = Txt & "Instructor(s): " & CourseInInv.Instructor & "<br>"
        Txt = Txt & "Start: " & CourseInInv.ExamStartTime & "<br>"
        Txt = Txt & "End: " & CourseInInv.ExamEndTime & "<br>"
        Txt = Txt & "Duration: " & CourseInInv.ExamDuration & "<br>"
        Txt = Txt & "Format: " & CourseInInv.ExamFormat & "<br>"
        Txt = Txt & "Number of students: " & CourseInInv.ExamNumStudents & "<br>"
        Txt = Txt & "Exam link: " & _
            "<a href='" & CourseInInv.ExamLink & "'>" & "exam</a></p>"
    Next CourseInInv
    
    Txt = Txt & "<p>CfAs: </p>"
    Txt = Txt & "<p>Alternates: </p>"
    
    m_logger.LogDebug "Adding attendee information to body", indt:=1
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    Dim role
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    If m_invT1.Exists("Triage") Then
        Txt = Txt & "<p>Triage: " & m_invT1.Item("Triage").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    For Each role In m_invT1
        If role <> "Triage" Then
            Txt = Txt & role & ": " & m_invT1.Item(role).Fullname & "<br>"
        End If
    Next role

    Txt = Txt & "</p>"
    htmlT2Inperson = Txt
    m_logger.LogDebug "Successfully made html body for tier 2 online exam invite"
    m_logger.LogDebug "Body as follows" & _
        vbNewLine & "'''" & vbNewLine & Txt & vbNewLine & "'''", indt:=1
End Function

Public Function htmlDAPOnCallOnline() As String
    m_logger.LogDebug "Making DAP online on-call Exam html body"
    Dim Txt As String
    Txt = "<p>LS Staff on Call:</p>"
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    End If
    Dim role
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    Txt = Txt & "</p>"
    
    Txt = Txt & "<p>DAP Tier 1: " & m_exSheet.T1Info(m_fromRow) & "</p>"
    Txt = Txt & "<p> PLEASE NOTE: On-call means you won’t be required to provide live support. " & _
        "Instead, you will simply be a backup contact for whoever is actually administering the exam. </p>"
    Txt = Txt & "<p><b>You are only required to be available via MS teams. " & _
        "Please set up a MS teams chat with the DAP Tier 1 staff person listed above.</b></p>"
    
    m_logger.LogDebug "Adding course information to body", indt:=1
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        Txt = Txt & "<p><b>" & CourseInInv.Name & "</b><br>"
        Txt = Txt & "Section(s): " & CourseInInv.Section & "<br>"
        Txt = Txt & "Instructor(s): " & CourseInInv.Instructor & "<br>"
        Txt = Txt & "Start: " & CourseInInv.ExamStartTime & "<br>"
        Txt = Txt & "End: " & CourseInInv.ExamEndTime & "<br>"
        Txt = Txt & "Duration: " & CourseInInv.ExamDuration & "<br>"
        Txt = Txt & "Format: " & CourseInInv.ExamFormat & "<br>"
        Txt = Txt & "Number of students: " & CourseInInv.ExamNumStudents & "<br>"
        Txt = Txt & "Exam link: " & _
            "<a href='" & CourseInInv.ExamLink & "'>" & "exam</a></p>"
    Next CourseInInv
    
    Txt = Txt & "<p>CfAs: </p>"
    Txt = Txt & "<p>Alternates: </p>"
    
    m_logger.LogDebug "Adding attendee information to body", indt:=1
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    
    Txt = Txt & "<p><b>How to create a MS Teams chat</b></p>"
    Txt = Txt & "<p>You start one-on-one and group chats the same way: " & _
        "by selecting <b>New chat</b> at the top of your chat list. "
    Txt = Txt & "<a href='https://support.content.office.net/en-us/media/05f0ae37-24a8-4d25-a660-3bba5b54a6db.png'>" & _
        "See Image for New Chat Icon</a></p>"
    Txt = Txt & "<p>Once you've selected New chat and entered a person’s name, " & _
        "compose your message in the box at the bottom of the chat. </p>"
    htmlDAPOnCallOnline = Txt
    m_logger.LogDebug "Successfully made html body for DAP online on-call exam"
    m_logger.LogDebug "Body as follows" & _
        vbNewLine & "'''" & vbNewLine & Txt & vbNewLine & "'''", indt:=1
End Function

Public Function htmlOnCallOnline() As String
    m_logger.LogDebug "Making online on-call Exam html body"
    Dim Txt As String
    Txt = "<p>LS Staff on Call:</p>"
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    Dim role
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    Txt = Txt & "</p>"
    
    Txt = Txt & "<p> PLEASE NOTE: On-call means you won’t be required to provide live support. " & _
        "Instead, you will simply be a backup contact for whoever is actually administering the exam. </p>"
    Txt = Txt & "<p><b>You are only required to be available via MS teams. " & _
        "Please set up a MS teams chat with the instructor listed below.</b></p>"
    
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        Txt = Txt & "<p><b>" & CourseInInv.Name & "</b><br>"
        Txt = Txt & "Section(s): " & CourseInInv.Section & "<br>"
        Txt = Txt & "Instructor(s): " & CourseInInv.Instructor & "<br>"
        Txt = Txt & "Start: " & CourseInInv.ExamStartTime & "<br>"
        Txt = Txt & "End: " & CourseInInv.ExamEndTime & "<br>"
        Txt = Txt & "Duration: " & CourseInInv.ExamDuration & "<br>"
        Txt = Txt & "Format: " & CourseInInv.ExamFormat & "<br>"
        Txt = Txt & "Number of students: " & CourseInInv.ExamNumStudents & "<br>"
        Txt = Txt & "Exam link: " & _
            "<a href='" & CourseInInv.ExamLink & "'>" & "exam</a></p>"
    Next CourseInInv
    
    Txt = Txt & "<p>CfAs: </p>"
    Txt = Txt & "<p>Alternates: </p>"
    
    m_logger.LogDebug "Adding attendee information to body", indt:=1
    If m_invT2.Exists("Lead") Then
        Txt = Txt & "<p>Lead: " & m_invT2.Item("Lead").Fullname & "<br>"
    Else
        Txt = Txt & "<p>"
    End If
    For Each role In m_invT2
        If role <> "Lead" Then
            Txt = Txt & role & ": " & m_invT2.Item(role).Fullname & "<br>"
        End If
    Next role
    
    Txt = Txt & "<p><b>How to create a MS Teams chat</b></p>"
    Txt = Txt & "<p>You start one-on-one and group chats the same way: " & _
        "by selecting <b>New chat</b> at the top of your chat list. "
    Txt = Txt & "<a href='https://support.content.office.net/en-us/media/05f0ae37-24a8-4d25-a660-3bba5b54a6db.png'>" & _
        "See Image for New Chat Icon</a></p>"
    Txt = Txt & "<p>Once you've selected New chat and entered a person’s name, " & _
        "compose your message in the box at the bottom of the chat. </p>"
    htmlOnCallOnline = Txt
    m_logger.LogDebug "Successfully made html body for online on-call exam"
    m_logger.LogDebug "Body as follows" & _
        vbNewLine & "'''" & vbNewLine & Txt & vbNewLine & "'''", indt:=1
End Function

Public Function GetInviteBody(Optional ForTier2 As Boolean = True) As String
    Dim Txt As String
    Select Case m_invIsOnline
        Case True
            If m_invIsDAP Then
                Txt = htmlDAPOnCallOnline()
            ElseIf m_invIsOnCall Then
                Txt = htmlOnCallOnline()
            ElseIf ForTier2 Then
                Txt = htmlT2Online()
            Else
                Txt = htmlPreMeetOnline()
            End If
        Case False
            If ForTier2 Then
                Txt = htmlT2Inperson()
            Else
                Txt = htmlPreMeetInperson()
            End If
    End Select
    GetInviteBody = Txt
End Function

Public Function CheckIfUpdating() As Boolean
    m_logger.LogDebug "Check if course allows auto-updating in the exam sheet"
    Dim ReNo As New RegExp
    ReNo.Pattern = "no"
    ReNo.IgnoreCase = True
    If ReNo.Test(m_exSheet.AutoUpdateInfo(m_fromRow)) Then
        CheckIfUpdating = False
        m_logger.LogDebug "Course does not allow auto-updating as per the exam sheet"
    ElseIf CheckPrevCalInv = "" Then
        CheckIfUpdating = False
        m_logger.LogDebug "Course does not allow auto-updating as per the exam sheet"
    Else
        CheckIfUpdating = True
        m_logger.LogDebug "Course allows auto-updating as per the exam sheet"
    End If
End Function

Public Function CheckIfDrafting() As Boolean
    Dim ReNo As New RegExp
    ReNo.Pattern = "no"
    ReNo.IgnoreCase = True
    If ReNo.Test(m_exSheet.AutoDraftInfo(m_fromRow)) Then
        CheckIfDrafting = False
        m_logger.LogDebug "Course does not allow auto drafting"
    ElseIf CheckPrevCalInv <> "" Then
        CheckIfDrafting = False
        m_logger.LogDebug "Course does not allow auto drafting"
    Else
        CheckIfDrafting = True
         m_logger.LogDebug "Course does allow auto drafting"
    End If
End Function



Public Function CheckPrevCalInv() As String
    m_logger.LogDebug "Checking previous calendar invite status"
    Dim ReSent As New RegExp: ReSent.Pattern = "sent"
    ReSent.IgnoreCase = True
    Dim ReDrafted As New RegExp: ReDrafted.Pattern = "drafte*d*"
    ReDrafted.IgnoreCase = True
    Dim ReSaved As New RegExp: ReSaved.Pattern = "saved*"
    ReSaved.IgnoreCase = True
    
    Dim CalStatus As String: CalStatus = m_exSheet.CalStatusInfo(m_fromRow)
    If ReSent.Test(CalStatus) Then
        CheckPrevCalInv = "sent"
        Exit Function
    ElseIf ReDrafted.Test(CalStatus) Then
        CheckPrevCalInv = "drafted"
        Exit Function
    ElseIf ReSaved.Test(CalStatus) Then
        CheckPrevCalInv = "saved"
        Exit Function
    End If
    CheckPrevCalInv = ""
    m_logger.LogDebug "Found previous calendar invite status '" & CheckPrevCalInv & _
        "' from exam sheet using regular expressions"
End Function

Public Function CheckIfDAP() As Boolean
    m_logger.LogDebug "Check if exam is DAP (on-call)"
    Dim ReDAP As New RegExp
    ReDAP.Pattern = "\(*DAP\)*"
    CheckIfDAP = ReDAP.Test(m_exSheet.CourseInfo(m_fromRow))
    m_logger.LogDebug "Checked: exam DAP is " & CheckIfDAP
End Function

Public Function CheckIfOnline() As Boolean
    m_logger.LogDebug "Check if exam is online"
    Dim ReInPerson As New RegExp
    ReInPerson.IgnoreCase = True
    ReInPerson.Pattern = "in\s*-*\s*person"
    CheckIfOnline = Not (ReInPerson.Test(m_exSheet.FormatInfo(m_fromRow)))
    ReInPerson.Pattern = "paper"
    CheckIfOnline = CheckIfOnline And _
        Not (ReInPerson.Test(m_exSheet.FormatInfo(m_fromRow)))
    m_logger.LogDebug "Checked: exam online is " & CheckIfOnline
End Function

Public Function CheckIfOnCall() As Boolean
    m_logger.LogDebug "Check if exam is on-call"
    Dim ReOnCall As New RegExp
    ReOnCall.Pattern = "on\s*-*\s*call"
    CheckIfOnCall = ReOnCall.Test(m_exSheet.T2Info(m_fromRow))
    m_logger.LogDebug "Checked: exam on-call is " & CheckIfOnCall
End Function

Public Function CheckIfHasPM() As Boolean
    m_logger.LogDebug "Check if exam has a pre-meeting"
    Dim ReNoT1Needed As New RegExp
    ReNoT1Needed.Pattern = "No\s*T(ier)*\s*-*\s*1\s*Need"
    ReNoT1Needed.IgnoreCase = True
    If Not (m_invIsDAP) _
        And Not (m_invIsOnCall) _
        And Not (ReNoT1Needed.Test(m_exSheet.T1Info(m_fromRow))) Then
        CheckIfHasPM = True
    Else
        CheckIfHasPM = False
    End If
    m_logger.LogDebug "Checked: exam has pre-meeting " & CheckIfHasPM
End Function

Public Function MakeInviteTitle(Optional ForTier2 As Boolean = True) As String
    m_logger.LogDebug "Making title for invite"
    Dim InvPrefix As String
    If m_invIsDAP Then
        m_logger.LogDebug "Title is for Tier 2 DAP", indt:=1
        InvPrefix = m_t & "Tier 2 - ON CALL (DAP) - "
    ElseIf m_invIsOnCall Then
        m_logger.LogDebug "Title is for Tier 2 On-Call", indt:=1
        InvPrefix = m_t & "Tier 2 - ON CALL - "
    ElseIf ForTier2 Then
        m_logger.LogDebug "Title is for Tier 2", indt:=1
        InvPrefix = m_t & "Tier 2 - "
    ElseIf m_invHasPM Then
        m_logger.LogDebug "Title is for Pre-Meeting ", indt:=1
        InvPrefix = m_t & "Pre-exam Meeting - "
    End If
    
    Dim InvSubject As String: InvSubject = ""
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        InvSubject = InvSubject & " " & _
            CourseInInv.Name & " " & _
            CourseInInv.Section & " " & _
            CourseInInv.ExamFormat & " " & _
            CourseInInv.ExamType & " " & _
            CourseInInv.Instructor & "; "
    Next CourseInInv
    
    MakeInviteTitle = InvPrefix & InvSubject
    m_logger.LogDebug "Made title '" & MakeInviteTitle & "'"
End Function

Public Sub WriteTeamsMsgTempate()
    m_logger.LogDebug "Constructing teams message template"
    Dim TeamsMsg As String
    TeamsMsg = "Greetings all! Looking forward to everyone supporting "
    Dim CourseInInv As clsCourse
    For Each CourseInInv In m_invCourses
        With CourseInInv
            TeamsMsg = TeamsMsg & _
                .Name & " " & .Section & " " & "w/ " & .Instructor & "; "
        End With
    Next CourseInInv
    TeamsMsg = TeamsMsg & "hope to see <INSERT @NAMES HERE>"
    TeamsMsg = TeamsMsg & " at " & m_invDate & " " & m_invStart & " in " & _
        m_exSheet.RoomInfo(m_fromRow) & " " & _
        " for a quick pre-meeting. Let us know if you have any questions. Thank you everyone!"
    m_exSheet.TeamTextInfo(m_fromRow) = TeamsMsg
    m_logger.LogDebug "Constructed teams message template as '" & TeamsMsg & "'"
End Sub

Property Get Skip() As Boolean
    Skip = m_skip
End Property


