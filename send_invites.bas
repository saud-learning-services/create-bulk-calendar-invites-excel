'Attribute VB_Name = "Send_Invite_Module"
Option Explicit

'Initialize global variables
Dim WB As Workbook
Dim ExamSheet As Worksheet
Dim T1Mails As Worksheet
Dim T1LR As Integer
Dim T1LC As Integer

'Initialize / set values for global variables
Private Sub InitVars()
    Set WB = ThisWorkbook
    Set ExamSheet = WB.Sheets("Exam Sheet")
    Set T1Mails  = WB.Sheets("Tier 1 Mails")
    
    Call MakeRefRange(T1Mails, T1LR, T1LC)
End Sub

'Finds last row and column if applicable
'Can produce a specific reference range also
Private Sub MakeRefRange( _
    RefSheet As Worksheet, _
    NumRow As Integer, _
    NumCol As Integer, _
    Optional RefRange As Range, _
    Optional AnchorRow As Integer = 1, _
    Optional AnchorCol As Integer = 1)

    Call FindLastRC(NumRow, NumCol, RefSheet)
    
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
    LastRow As Integer, _
    LastCol As Integer, _
    FindInSheet As Worksheet)

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
