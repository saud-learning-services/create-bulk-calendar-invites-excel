VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Member variables
Private m_logLevel As Integer
Private Const m_logLevelDefault As Integer = 1
Private m_logLevelNames(4) As String
Private m_log As String
Private m_logSheet As Worksheet
Private m_logToConsole As Boolean
Private m_logToFile As Boolean
Private m_logToSheet As Boolean
Private m_logFileName As String
Private m_errMsg As String
Private m_errNum As Long
Private m_logFileClosed As Boolean
Private m_hasModules As Boolean
Private m_logFileNum As Integer
Private m_logLine As Long

Private Sub Class_Initialize()
    m_logLevel = m_logLevelDefault
    m_logLevelNames(0) = "DEBUG" 'Investigative details
    m_logLevelNames(1) = "INFO" 'General statuses
    m_logLevelNames(2) = "WARNING" 'Anomaly deserving of attention if results unexpected
    m_logLevelNames(3) = "ERROR" 'Could impede program but may be able to handle
    m_logLevelNames(4) = "FATAL" 'Program breaking errors
    m_errMsg = "An error occurred while using property, method or function in" _
        & " class module 'clsLogger'"
    m_errNum = 42001
    m_logToConsole = True
    m_logToFile = True
    m_logToSheet = False
    m_logFileClosed = True
End Sub

Private Sub Class_Terminate()
    StopLogging
End Sub

Public Sub StartLogging( _
    Optional logFileName As String = "log", _
    Optional logFileExt As String = ".txt", _
    Optional LogToConsole As Boolean = True, _
    Optional LogToFile As Boolean = True, _
    Optional LogToSheet As Boolean = False, _
    Optional LogLevel As Integer = m_logLevelDefault)
    
    'If Not Exists
    m_logFileName = VBA.Right(TimeStampLog(logFileName, logFileExt), 30)
    
    'TimeStamp from Modules
    m_logToConsole = LogToConsole
    m_logToFile = LogToFile
    m_logToSheet = LogToSheet
    m_logLine = 1
    'Don't log to sheet if there're too many sheets i.e. >= 5
    If ThisWorkbook.Worksheets.Count >= 3 Then
        m_logToSheet = False
    End If
    If m_logToFile Then
        m_logFileNum = FreeFile
        m_log = EnsureLogPath("logs") & m_logFileName
        Open m_log For Output As #m_logFileNum
        m_logFileClosed = False
    End If
    If m_logToSheet Then
        Dim SelSheet As Worksheet: Set SelSheet = ActiveSheet
        Dim SelRange As Range: Set SelRange = Selection
        'Following line created and selects new logging sheet
        Set m_logSheet = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        m_logSheet.Name = VBA.Left(TimeStampLog(logFileName, ""), 30)
        'Go back to previous user sheet & range selection
        SelSheet.Select
        SelRange.Select
    End If
    m_logLevel = EnforceLvlLim(LogLevel)
    Me.Log "Logger initialized with settings: "
    Me.Log "Log file name: " & logFileName, logMsgIndents:=1
    Me.Log "Log to console (VBA immediate window): " & m_logToConsole, logMsgIndents:=1
    Me.Log "Log to file: " & m_logToFile, logMsgIndents:=1
    Me.Log "Log to sheet: " & m_logToSheet, logMsgIndents:=1
    Me.Log "Logging messages of level '" & m_logLevel & _
        "', i.e. '" & m_logLevelNames(m_logLevel) _
        & "' and above", logMsgIndents:=1
End Sub

Public Sub StopLogging()
    If m_logFileClosed = False Then
        Close #m_logFileNum
    End If
End Sub

Public Sub Log(ByVal logContent As String, _
    Optional ByVal logMsgLevel As Integer = m_logLevelDefault, _
    Optional errFromVBA As ErrObject, _
    Optional errNumFromMods As Long = 42000, _
    Optional ByVal logMsgIndents As Integer = 0, _
    Optional raiseErr As Boolean = False)
    
    Dim logMsg As String
    logMsgLevel = EnforceLvlLim(logMsgLevel)
    If logMsgLevel >= m_logLevel Then
        If m_logLevelNames(logMsgLevel) = "FATAL" _
            Or m_logLevelNames(logMsgLevel) = "ERROR" Then
            Dim logMsgSysErr
            Set errFromVBA = Err
            logMsgSysErr = LogPrefix(logMsgLevel, logMsgIndents) & _
                "From VBA - error number '" & errFromVBA.Number & "', " & _
                "description '" & errFromVBA.Description & "'."
            logMsg = LogPrefix(logMsgLevel, logMsgIndents) & _
                "From Logger error log - error number '" & errNumFromMods & "', " & _
                "description '" & logContent & "'"
            If m_logToConsole Then
                Debug.Print logMsgSysErr
                Debug.Print logMsg
            End If
            If m_logToFile Then
                Print #m_logFileNum, logMsgSysErr
                Print #m_logFileNum, logMsg
            End If
            If m_logToSheet Then
                m_logSheet.Cells(m_logLine, 1) = logMsgSysErr
                m_logLine = m_logLine + 1
                m_logSheet.Cells(m_logLine, 1) = logMsg
                m_logLine = m_logLine + 1
            End If
            If m_logLevelNames(logMsgLevel) = "FATAL" _
                And raiseErr Then
                Close 'All Open Files
                Err.Raise errNumFromMods, Description:=logContent
            End If
        Else
            logMsg = LogPrefix(logMsgLevel, logMsgIndents) & logContent
            If m_logToConsole Then
                Debug.Print logMsg
            End If
            If m_logToFile Then
                Print #m_logFileNum, logMsg
            End If
            If m_logToSheet Then
                m_logSheet.Cells(m_logLine, 1) = logMsg
                m_logLine = m_logLine + 1
            End If
        End If
    End If
End Sub

Public Sub LogDebug(ByVal debugContent As String, _
    Optional ByVal indt As Integer = 0)
    
    Me.Log debugContent, logMsgLevel:=0, logMsgIndents:=indt
End Sub

Public Sub LogInfo(ByVal infoContent As String, _
    Optional ByVal indt As Integer = 0)
    
    Me.Log infoContent, logMsgLevel:=1, logMsgIndents:=indt
End Sub

Public Sub LogWarning(ByVal warningContent As String, _
    Optional ByVal indt As Integer = 0)
    
    Me.Log warningContent, logMsgLevel:=2, logMsgIndents:=indt
End Sub

Public Sub LogError(ByVal errorContent As String, _
    errVBA As ErrObject, _
    Optional errNumMods As Long = 42000, _
    Optional ByVal indt As Integer = 0)
    
    Me.Log errorContent, _
        logMsgLevel:=3, _
        errFromVBA:=errVBA, _
        errNumFromMods:=errNumMods, _
        logMsgIndents:=indt
End Sub

Public Sub LogFatal(ByVal fatalContent As String, _
    errVBA As ErrObject, _
    Optional errNumMods As Long = 42000, _
    Optional ByVal indt As Integer = 0, _
    Optional ByVal raiseFatal As Boolean = False)
    
    Me.Log fatalContent, _
        logMsgLevel:=4, _
        errFromVBA:=errVBA, _
        errNumFromMods:=errNumMods, _
        logMsgIndents:=indt, _
        raiseErr:=raiseFatal
End Sub

'Function - add prefix to log message, incl timestamp
Public Function LogPrefix( _
    Optional level As Integer = 1, _
    Optional indents As Integer = 0) As String
    
    Dim TimeStamp As String
    TimeStamp = FormatDateTime(Now(), vbShortDate) _
        & " " & FormatDateTime(Now(), vbLongTime)
    Dim prefixIndents As String: prefixIndents = ""
    Dim ind As Integer
    For ind = 0 To indents Step 1
        prefixIndents = prefixIndents & " ... "
    Next ind
    level = EnforceLvlLim(level)
    
    'Ensure text after log level string is flush
    'Assumes longest level name is 7 chars ('WARNING'), change if more lvls added
    Dim MaxLvlLen As Integer: MaxLvlLen = 7
    Dim WhiteSpaces As String
    WhiteSpaces = ""
    If Len(m_logLevelNames(level)) < MaxLvlLen Then
        Dim wtSpace As Integer
        For wtSpace = 1 To MaxLvlLen - Len(m_logLevelNames(level))
            WhiteSpaces = WhiteSpaces & " "
        Next wtSpace
    End If
    
    LogPrefix = TimeStamp & " " & m_logLevelNames(level) & WhiteSpaces & prefixIndents
End Function

Property Get LogToConsole() As Boolean
    LogToConsole = m_logToConsole
End Property

Property Get LogToFile() As Boolean
    LogToFile = m_logToFile
End Property

Property Get LogToSheet() As Boolean
    LogToSheet = m_logToSheet
End Property

Property Get LogLevel() As Integer
    LogLevel = m_logLevel
End Property

Property Let LogLevel(level As Integer)
    m_logLevel = EnforceLvlLim(level)
End Property

Property Get LogLevelByName() As String
    LogLevelByName = m_logLevelNames(m_logLevel)
End Property

Property Let LogLevelByName(levelName As String)
    Dim level As Integer
    For level = LBound(m_logLevelNames) To UBound(m_logLevelNames)
        If levelName = m_logLevelNames Then
            m_logLevel = level
        End If
    Next level
End Property

'Function - ensure level number in bounds
Public Function EnforceLvlLim(level As Integer) As Integer
    Select Case level
        Case LBound(m_logLevelNames) To UBound(m_logLevelNames)
            EnforceLvlLim = level
        Case Is < LBound(m_logLevelNames)
            EnforceLvlLim = LBound(m_logLevelNames)
        Case Is > UBound(m_logLevelNames)
            EnforceLvlLim = UBound(m_logLevelNames)
    End Select
End Function

Public Function EnsureLogPath(Optional ByVal DirName As String = "log") As String
    Dim WorkbookDir As String: WorkbookDir = ActiveWorkbook.Path
    Dim SaveDir As String: SaveDir = WorkbookDir & "\" & DirName & "\"
    If VBA.Right(WorkbookDir, Len(DirName)) = DirName Then
        EnsureLogPath = WorkbookDir & "\"
    ElseIf Dir(SaveDir, vbDirectory) <> "" Then
        EnsureLogPath = SaveDir
    ElseIf Dir(SaveDir, vbDirectory) = "" Then
        MkDir (SaveDir)
        EnsureLogPath = SaveDir
    Else
        EnsureLogPath = Application.DefaultFilePath
    End If
End Function

Public Function TimeStampLog(ByVal FileName As String, _
    Optional ByVal FileExt As String = ".txt") As String
    Dim TimeStamp As String
    TimeStamp = FormatDateTime(Now, vbShortDate) & "_" & _
        FormatDateTime(Now, vbLongTime) & "_"
    TimeStamp = Replace(TimeStamp, "/", "")
    TimeStamp = Replace(TimeStamp, " ", "_")
    TimeStamp = Replace(TimeStamp, ":", "")
    TimeStamp = Replace(TimeStamp, ",", "")
    TimeStamp = Replace(TimeStamp, ".", "")
    TimeStampLog = VBA.Right(TimeStamp & FileName & FileExt, 30)
End Function

