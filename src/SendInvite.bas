Attribute VB_Name = "SendInvite"
Option Explicit

Public Sub Main()
    'SaveAsOutputFile
    Dim Logger As New clsLogger
    Logger.StartLogging LogLevel:=1, LogToFile:=True, LogToSheet:=True
    Dim Exams As New clsExamSheet
    On Error GoTo ConcludeMain
    Exams.ConstructExamSheet _
        ExamDataWS:=ThisWorkbook.Sheets("Exam Sheet"), Logger:=Logger
    Dim Settings As New clsSettings
    Settings.ConstructSettings _
        ExamSheet:=Exams, Logger:=Logger
    RuntimeOptimize
    Exams.UnmergeAll "COURSE", "SECTIONS", "MIDTERM / FINAL / SD?", "INSTRUCTOR", _
        "DATE", "TIME", "FORMAT", "DURATION", "# STUDENTS", "CALENDAR INVITE", _
        "SUPPORT ROOM", "AUTO UPDATE", "AUTO DRAFT", "TEAMS MESSAGE TEMPLATE", "EXAM LINK"
    Dim Mails As New clsMailSheet
    Mails.ConstructMailSheet _
        MailDataWS:=ThisWorkbook.Sheets("Mail List"), _
        Logger:=Logger, isDebugging:=Settings.isDebugging
    Mails.ConstructAttendeeColl
    Dim RowInd As Integer: RowInd = Settings.FirstSendRow
    Dim RowStep As Integer
    Do While RowInd <= Settings.LastSendRow
        On Error GoTo PrintErrExam
        Logger.LogInfo "", indt:=5
        Dim Invite As clsInviteSession
        Set Invite = New clsInviteSession
        With Exams
            If .Data.Cells(RowInd, .T2Col).MergeCells Then
                RowStep = .Data.Cells(RowInd, .T2Col).MergeArea.Count
            Else
                RowStep = 1
            End If
        End With
        Invite.ConstructProperties Logger, Exams, Mails, Settings, RowInd, RowStep, True
        If Not (Invite.Skip) Then
            Invite.WriteAppointments
        End If
        Invite.WriteTeamsMsgTempate
        GoTo NextExam
PrintErrExam:
        Exams.CalStatusInfo(RowInd) = Now() & " - FAILED"
        Logger.LogError "FAILED to construct/save/send invite for '" & _
            Exams.CourseInfo(RowInd) & "' at row '" & RowInd & "'", _
            errVBA:=Err
        Logger.LogWarning "Going to next exam"
        Resume NextExam
NextExam:
        Logger.LogInfo "", indt:=5
        RowInd = RowInd + RowStep
    Loop
    StopRuntimeOptimize
    MsgBox "Program completed, please check sheet and log, " & _
    "and please double check Outlook"
    Exit Sub
ConcludeMain:
    Logger.LogError "Fatal error encountered, please check log and sheet", Err
    Exams.RemergeAll
    StopRuntimeOptimize
    Logger.LogFatal "Fatal error", Err, raiseFatal:=True
End Sub

Public Sub RuntimeOptimize()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Public Sub StopRuntimeOptimize()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub SaveAsOutputFile()
    Dim OutputFileName As String
    OutputFileName = EnsureDirPath("output") & TimeStampFile("output", ".xlsm")
    ThisWorkbook.SaveAs (OutputFileName)
End Sub

Public Sub NukeTestInvites()
    Dim Logger As New clsLogger
    Logger.StartLogging "nuke", LogLevel:=1, LogToSheet:=False
    Dim Ot As Outlook.Application: Set Ot = New Outlook.Application
    Dim OtNameSpace As Outlook.Namespace
    Set OtNameSpace = Ot.GetNamespace("MAPI")
    Dim OtFrom As String: OtFrom = "learning.services@sauder.ubc.ca"
    Dim OtFolder As Outlook.Folder
    Dim OtAptMaker As Outlook.Recipient
    Set OtAptMaker = OtNameSpace.CreateRecipient(OtFrom)
    If OtAptMaker.Resolve Then
        Set OtFolder = _
            OtNameSpace.GetSharedDefaultFolder(OtAptMaker, olFolderCalendar)
    End If
    Dim ReTestStr As New RegExp
    ReTestStr.Pattern = "testingonly"
    ReTestStr.IgnoreCase = True
    Dim OtCalInv As Outlook.AppointmentItem
    Dim leftovers As Boolean: leftovers = True
    Dim nukeCount As Long: nukeCount = 0

    Logger.LogInfo "Dropping nuke on test only invites with '" & _
        ReTestStr.Pattern & "' in title"
    Do While leftovers = True
        Logger.LogDebug "Scanning calendar folder '" & OtFolder.Name & "'"
        For Each OtCalInv In OtFolder.Items
            Logger.LogDebug "Scanning '" & OtCalInv.Subject & "'"
            If ReTestStr.Test(OtCalInv.Subject) Then
                Logger.LogInfo "Nuking '" & OtCalInv.Subject & "'"
                OtCalInv.Delete
                nukeCount = nukeCount + 1
                Logger.LogInfo "Successfully nuked", indt:=1
            End If
        Next OtCalInv
        leftovers = False
        Logger.LogDebug "Double checking '" & OtFolder.Name & "'"
        For Each OtCalInv In OtFolder.Items
            Logger.LogDebug "Scanning '" & OtCalInv.Subject & "'"
            If ReTestStr.Test(OtCalInv.Subject) Then
                leftovers = True
                Exit For
            End If
        Next OtCalInv
    Loop
    Logger.LogInfo "Successfully nuked '" & nukeCount & "' testing invites"
End Sub

Public Function EnsureDirPath(ByVal DirName As String) As String
    Dim WorkbookDir As String: WorkbookDir = ActiveWorkbook.Path
    Dim SaveDir As String: SaveDir = WorkbookDir & "\" & DirName & "\"
    If VBA.Right(WorkbookDir, Len(DirName)) = DirName Then
        EnsureDirPath = WorkbookDir & "\"
    ElseIf Dir(SaveDir, vbDirectory) <> "" Then
        EnsureDirPath = SaveDir
    ElseIf Dir(SaveDir, vbDirectory) = "" Then
        MkDir (SaveDir)
        EnsureDirPath = SaveDir
    Else
        EnsureDirPath = Application.DefaultFilePath
    End If
End Function

Public Function TimeStampFile(ByVal FileName As String, _
    Optional ByVal FileExt As String = ".txt") As String
    Dim TimeStamp As String
    TimeStamp = FormatDateTime(Now, vbShortDate) & "_" & _
        FormatDateTime(Now, vbLongTime) & "_"
    TimeStamp = Replace(TimeStamp, "/", "")
    TimeStamp = Replace(TimeStamp, " ", "_")
    TimeStamp = Replace(TimeStamp, ":", "")
    TimeStamp = Replace(TimeStamp, ",", "")
    TimeStamp = Replace(TimeStamp, ".", "")
    TimeStampFile = VBA.Right(TimeStamp & FileName & FileExt, 30)
End Function

Public Sub ExportModules()
    'Can only be run if "Trust Access to the VBA Object Model" checked
    Const CompModule = 1
    Const CompClassModule = 2
    Dim srcDir As String
    srcDir = EnsureDirPath("src")
    Dim fileExtension As String
    Dim xlProjComp As Object
    For Each xlProjComp In ThisWorkbook.VBProject.VBComponents
        'Debug.Print xlProjComp.Name
        If xlProjComp.Type = CompModule Or xlProjComp.Type = CompClassModule Then
            Select Case xlProjComp.Type
                Case CompClassModule
                    fileExtension = ".cls"
                Case CompModule
                    fileExtension = ".bas"
            End Select
            
            xlProjComp.Export (srcDir & "\" & xlProjComp.Name & fileExtension)
            Debug.Print "Exported '" & xlProjComp.Name & "' to '" & srcDir & "'"
        End If
    Next xlProjComp
End Sub

