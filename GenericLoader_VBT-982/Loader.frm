VERSION 5.00
Begin VB.Form Loader 
   Caption         =   "UCLA OCLC Loader"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DEPENDENCIES TO REMEMBER FOR INSTALLATION:
'\windows\system32\voy*.dll
'\windows\system32\mswinsck.ocx

Option Explicit

Private DBUG As Boolean
Private Const ERROR_BAR As String = "*** ERROR ***"
Private Const MAX_RECORD_COUNT As Integer = 3000

Private Sub Form_Load()
    'Main handles everything
    Main
    'Exit from VB if running as program
    End
End Sub

Private Sub Main()
    'This is the controlling procedure for this form
    Set GL = New Globals
    GL.Init Command
    'GL.Init "-f " & App.Path & "\test.mrc -i 0"
    
    DBUG = Not GL.ProductionMode
    
    ProcessRecords
    
    GL.CloseAll
    Set GL = Nothing
End Sub

Private Sub ProcessRecords()
    Dim MarcFile As New Utf8MarcFileClass
    Dim MarcRecord As New Utf8MarcRecordClass
    Dim SourceRecord As OclcRecordType
    Dim RawRecord As String
    Dim RecordNumber As Integer
    Dim F001 As String
    
    RecordNumber = 0
    MarcFile.OpenFile GL.InputFilename
    Do While MarcFile.ReadNextRecord(RawRecord)
        RecordNumber = RecordNumber + 1
        Set MarcRecord = New Utf8MarcRecordClass
        With MarcRecord
            .CharacterSetIn = "U"   'Hooray, UTF-8 from OCLC
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            .MarcRecordIn = RawRecord
            F001 = GetOclcNumberFrom001(MarcRecord)
        End With
        Set SourceRecord.BibRecord = MarcRecord
        SourceRecord.PositionInFile = RecordNumber
        
        'Start of log entry for record
        'Subsequent functions may also write log messages for current record.
        WriteLog GL.Logfile, "Record #" & RecordNumber & ": Incoming record OCLC# " & F001
        
        If RecordIsWanted(SourceRecord) Then
'Tempoarily disable debugging
DBUG = False
            PrepareRecord SourceRecord
            GetOclcNumbers SourceRecord
            SearchVoyager SourceRecord
DBUG = True
            
            'Multiple matches get rejected, so there must be 0 or 1 matches if OK
            If OkToUpdate(SourceRecord) Then
                If SourceRecord.BibMatchCount = 1 Then
                    'Voyager record will be updated, though Voyager records may also be written to a review file.
                    PrepareForUpdate SourceRecord
                    'UpdateVoyager SourceRecord
                Else
                    AddRecord SourceRecord
                End If
            End If 'OkToUpdate
        
        Else
            'Record is not wanted so reject it
            WriteRejectRecord SourceRecord
        End If 'RecordIsWanted
        
        'End of log entry for record
        WriteLog GL.Logfile, ""
    
    Loop 'MarcFile.ReadNextRecord
    
End Sub

Private Sub PrepareRecord(SourceRecord As OclcRecordType)
    'Add/update/delete fields within OCLC record before involving Voyager
    Dim F035 As String
    Dim Ind As String   'Both indicators
    Dim ind2 As String
    Dim Text As String  'Field text
    Dim FldPointer As Integer
    
    With SourceRecord.BibRecord
        'Build 035 from 001/003 (035 $a) and 019 (035 $z)
        .FldFindFirst "001"
        F035 = .SfdMake("a", "(OCoLC)" & LeftPad(GetDigits(.FldText), "0", 8))
        .FldDelete
        .FldFindFirst "003"
        .FldDelete
        '019 is not repeatable, but $a is
        If .FldFindFirst("019") Then
            .SfdFindFirst "a"
            Do While .SfdWasFound
                F035 = F035 & .SfdMake("z", "(OCoLC)" & LeftPad(.SfdText, "0", 8))
                .SfdFindNext
            Loop
            .FldDelete
        End If
        
        'Delete all OCLC-supplied 035 fields before adding our newly-created one
        .FldFindFirst "035"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Finally, add our 035
        .FldAddGeneric "035", "  ", F035, 3
        
        'Remove any 049 and add our own
        .FldFindFirst "049"
        If .FldWasFound Then .FldDelete
        .FldAddGeneric "049", "  ", .SfdMake("a", "CLYY")
        
        'Delete 029 fields
        .FldFindFirst "029"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Delete 7xx with $5
        .FldFindFirst "7"
        Do While .FldWasFound
            .SfdFindFirst "5"
            If .SfdWasFound Then
                .FldDelete
            End If
            .FldFindNext
        Loop
        
        'Delete 891 fields
        .FldFindFirst "891"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Delete all 9XX fields of incoming record except 936 and 948
        .FldFindFirst "9"
        Do While .FldWasFound
            If Not (.FldTag = "936" Or .FldTag = "948") Then
                .FldDelete
            End If
            .FldFindNext
        Loop
        
        'Modify certain 6xx fields with specific indicators
        .FldFindFirst "6"
        Do While .FldWasFound
            ind2 = Right(.FldInd, 1)
            Select Case .FldTag
                Case "600"
                    If ind2 >= "3" And ind2 <= "8" Then
                        Ind = .FldInd
                        Text = .FldText
                        FldPointer = .FldPointer
                        .FldDelete
                        .FldAddGeneric "692", Ind, Text, 3
                        .FldPointer = FldPointer
                    End If
                Case "610"
                    If ind2 >= "3" And ind2 <= "8" Then
                        Ind = .FldInd
                        Text = .FldText
                        FldPointer = .FldPointer
                        .FldDelete
                        .FldAddGeneric "693", Ind, Text, 3
                        .FldPointer = FldPointer
                    End If
                Case "611"
                    If ind2 >= "3" And ind2 <= "8" Then
                        Ind = .FldInd
                        Text = .FldText
                        FldPointer = .FldPointer
                        .FldDelete
                        .FldAddGeneric "694", Ind, Text, 3
                        .FldPointer = FldPointer
                    End If
                Case "630"
                    If ind2 >= "3" And ind2 <= "8" Then
                        Ind = .FldInd
                        Text = .FldText
                        FldPointer = .FldPointer
                        .FldDelete
                        .FldAddGeneric "695", Ind, Text, 3
                        .FldPointer = FldPointer
                    End If
                Case "650"
                    If (ind2 = "3") Or (ind2 >= "5" And ind2 <= "8") Then
                        Ind = .FldInd
                        Text = .FldText
                        FldPointer = .FldPointer
                        .FldDelete
                        .FldAddGeneric "690", Ind, Text, 3
                        .FldPointer = FldPointer
                    End If
                Case "651"
                    If (ind2 = "3") Or (ind2 >= "5" And ind2 <= "8") Then
                        Ind = .FldInd
                        Text = .FldText
                        FldPointer = .FldPointer
                        .FldDelete
                        .FldAddGeneric "691", Ind, Text, 3
                        .FldPointer = FldPointer
                    End If
            End Select
            .FldFindNext
        Loop
        
        'Change LDR/17 from 5 to blank
        If .GetLeaderValue(17, 1) = "5" Then
            .ChangeLeaderValue 17, " "
        End If
        
        'Create 910
        .FldAddGeneric "910", "  ", .SfdMake("a", "BATCHCAT " & Format(Now(), "yymmdd")), 3
        
        'Add date to 948: Assume there's just one
        'Specs say there'll be a $d; put the $c before it
        .FldFindFirst "948"
        If .FldWasFound Then
            .SfdFindFirst "d"
            .SfdInsertBefore "c", Format(Now(), "yyyymmdd")
        End If
        
    End With 'SourceRecord.BibRecord
    
    'Create the UCOCLC 035 fields needed for WorldCat linking
    Set SourceRecord.BibRecord = UpdateUcoclc(SourceRecord.BibRecord)
    
    If DBUG = True Then
        WriteLog GL.Logfile, vbCrLf & "==> Bib record after pre-processing:"
        WriteLog GL.Logfile, SourceRecord.BibRecord.TextRaw
    End If
    
    CreateInternetHoldings SourceRecord
    
End Sub

Private Sub CreateInternetHoldings(SourceRecord As OclcRecordType)
    Dim OclcHR As HoldingsRecordType
    Dim HolRecord As Utf8MarcRecordClass
    Dim current_yymmdd As String
    Dim CallNumber As String
    
    current_yymmdd = Format(Now(), "yymmdd")
    
    Set HolRecord = New Utf8MarcRecordClass
    With HolRecord
        .CharacterSetIn = "U"
        .CharacterSetOut = "U"
        'At present, these are all serials
        .NewRecord "y"
        
        'Add 007: always the same
        .FldAddGeneric "007", "", "cr", 3
        
        'Add 008: always the same
        .FldAddGeneric "008", "", current_yymmdd & "0u    0   0001uu   0" & current_yymmdd, 3
        
        'Add 852, using data from bib record
        CallNumber = GetCallNumber(SourceRecord.BibRecord)
        If CallNumber = "" Then
            WriteLog GL.Logfile, vbTab & "WARNING: No call number found"
            .FldAddGeneric "852", "  ", .SfdMake("b", "in"), 3
        Else
            .FldAddGeneric "852", "0 ", .SfdMake("b", "in") & .SfdMake("h", CallNumber), 3
        End If
        
        'Add 866: always the same
        .FldAddGeneric "866", " 0", .SfdMake("a", "Online access"), 3
    End With 'HolRecord

    If DBUG = True Then
        WriteLog GL.Logfile, vbCrLf & "==> Hol record after pre-processing:"
        WriteLog GL.Logfile, vbCrLf & HolRecord.TextRaw
    End If
    
    Set OclcHR.HolRecord = HolRecord
    SourceRecord.HoldingsRecordCount = 1
    ReDim SourceRecord.HoldingsRecords(1 To 1)
    SourceRecord.HoldingsRecords(1) = OclcHR

End Sub

Private Sub GetOclcNumbers(SourceRecord As OclcRecordType)
    'Populates SourceRecord.OclcNumbers() with values from 035
    '035 $a should always be first subfield, and only one 035 $a
    'May also be any number of $z following the $a
    'Just capture the raw subfield text here, used for logging; normalize for searching in SearchVoyager().
    
    Dim OclcCount As Integer
    ReDim SourceRecord.OclcNumbers(1 To MAX_OCLC_COUNT)
    OclcCount = 0
'Tempoarily enable debugging
DBUG = True
    
    With SourceRecord.BibRecord
        .FldFindFirst "035"
        .SfdMoveTop
        If DBUG Then
            WriteLog GL.Logfile, "==> OCLC numbers for searching:"
        End If
        Do While .SfdMoveNext
            OclcCount = OclcCount + 1
            SourceRecord.OclcNumbers(OclcCount) = .SfdText
            If DBUG Then
                WriteLog GL.Logfile, vbTab & "==> " & .SfdCode & " " & SourceRecord.OclcNumbers(OclcCount)
            End If
        Loop
    End With
    'Remove unused space from array
    ReDim Preserve SourceRecord.OclcNumbers(1 To OclcCount)
End Sub

Private Sub SearchVoyager(SourceRecord As OclcRecordType)
    'Searches Voyager for OCLC number(s) in SourceRecord.OclcNumbers()
    'Populates SourceRecord.BibMatchCount and SourceRecord.BibMatches()
    'Check against UCOCLC values, which are normalized from Voyager 035 $a.
    
    Dim AlreadyExists As Boolean
    Dim BibID As String
    Dim cnt As Integer
    Dim LogMessage As String
    Dim OclcCount As Integer
    Dim SearchNumber As String
    Dim SQL As String
    Dim rs As Integer 'RecordSet
    
    ReDim SourceRecord.BibMatches(1 To MAX_BIB_MATCHES)
    SourceRecord.BibMatchCount = 0
    rs = GL.GetRS
    
    For OclcCount = 1 To UBound(SourceRecord.OclcNumbers())
        SearchNumber = UCase(CalculateUcoclc(SourceRecord.OclcNumbers(OclcCount)))
        
        'Find Voyager bibs with 035 $a OCLC number matching incoming OCLC number
        SQL = _
            "select bi.bib_id " & _
            "from bib_index bi " & _
            "inner join bib_text bt " & _
            "on bi.bib_id = bt.bib_id " & _
            "where bi.index_code = '0350' " & _
            "and bi.normal_heading = '" & SearchNumber & "' " & _
            "order by bi.bib_id"
            
        With GL.Vger
            .ExecuteSQL SQL, rs
            Do While .GetNextRow
                With SourceRecord
                    BibID = GL.Vger.CurrentRow(rs, 1)
                    AlreadyExists = False
                    For cnt = 1 To .BibMatchCount
                        If BibID = .BibMatches(cnt) Then
                            AlreadyExists = True
                        End If
                    Next
                    If AlreadyExists = False Then
                        'Add new unique matches
                        .BibMatchCount = .BibMatchCount + 1
                        .BibMatches(.BibMatchCount) = BibID
                        
                        'Write log messages for matches
                        LogMessage = "Match found: incoming OCLC 035 $" & IIf(OclcCount = 1, "a", "z") & " " & Replace(SearchNumber, "UCOCLC", "")
                        LogMessage = LogMessage & " matches Voyager bib " & BibID
                        WriteLog GL.Logfile, vbTab & LogMessage
                        
                        'Special condition not in specs - reject record per vbross discussion on Jira VBT-217.
                        '$z matches (current OclcCount > 1, which is only true for incoming 035 $z) but incoming $a did not (since current match caused BibMatchCount = 1, meainng no matches before this).
                        'Set BibMatchCount to normally impossible negative value as flag to caller so I don't have to modify the OclcRecordType... yes, rewrite this correctly in another language someday.
                        'TODO: Remove this code if no objection from Melissa on VBT-982
'                        If .BibMatchCount = 1 And OclcCount > 1 Then
'                            .BibMatchCount = -1
'                            'Break out of the OclcCount For..Next block
'                            Exit For
'                        End If
                    End If
                End With 'SourceRecord
            Loop 'GetNextRow
            
        End With 'Vger
        
    Next 'OclcCount
    
    'Remove unused space from array
    With SourceRecord
        If .BibMatchCount > 0 Then
            ReDim Preserve .BibMatches(1 To .BibMatchCount)
        Else
            'Nothing added to the array, but previous results may still exist
            Erase .BibMatches
        End If
    End With
    
    GL.FreeRS rs
    
End Sub

Private Function GetOclcNumberFrom001(MarcRecord As Utf8MarcRecordClass) As String
    Dim F001 As String
    With MarcRecord
        .FldFindFirst "001"
        F001 = GetDigits(.FldText)
    End With
    GetOclcNumberFrom001 = F001
End Function

Private Function OkToUpdate(SourceRecord As OclcRecordType) As Boolean
    'Use incoming/Voyager matching results and other info to determine whether incoming record qualifies to update Voyager.
    'Return TRUE if OK, otherwise return FALSE and write record to REJECT file, along with log messages.
    
    Dim VgerRecord As Utf8MarcRecordClass
    Dim OclcDtSt As String
    Dim VgerDtSt As String
    Dim OclcBlvl As String
    Dim VgerBlvl As String
    
    Dim OK As Boolean
    OK = True   'Any condition below may set to False
    
    With SourceRecord
        'No matches in Voyager: no need to do anything, record will be added
        
        'Multiple matches in Voyager
        If .BibMatchCount > 1 Then
            WriteLog GL.Logfile, vbTab & "REJECTED: Multiple matches in Voyager"
            WriteRejectRecord SourceRecord
            OK = False
        End If
        
        'One match in Voyager, but bib level differs or 008/06 date status differs
        'First get the necessary info for comparison.
        If .BibMatchCount = 1 Then
            With .BibRecord
                OclcBlvl = .GetLeaderValue(7, 1)
                OclcDtSt = .Get008Value(6, 1)
            End With
            
            Set VgerRecord = GetVgerBibRecord(.BibMatches(1))
            With VgerRecord
                VgerBlvl = .GetLeaderValue(7, 1)
                VgerDtSt = .Get008Value(6, 1)
            End With
            
            'Reject if bib levels are different.
            If OclcBlvl <> VgerBlvl Then
                WriteLog GL.Logfile, vbTab & "REJECTED: Bib level mismatch (OCLC " & OclcBlvl & ", Voyager " & VgerBlvl & ")"
                WriteRejectRecord SourceRecord
                OK = False
            End If
            
            'Reject if 008/06 in each record has specific different values.
            If (OclcDtSt = "c" Or OclcDtSt = "u") And VgerDtSt = "d" Then
                WriteLog GL.Logfile, vbTab & "REJECTED: 008/06 mismatch (OCLC " & OclcDtSt & ", Voyager " & VgerDtSt & ")"
                WriteRejectRecord SourceRecord
                OK = False
            End If
            
        End If '.BibMatchCount = 1
        
        'Incoming 035 $z matched but incoming 035 $a did not, detected and flagged in SearchRecords()
        'TODO: Waiting for Kevin Balster/Melissa Beck to clarify if this is needed 20180627.
'        If .BibMatchCount = -1 Then
'            WriteLog GL.Logfile, vbTab & "REJECTED: Incoming record 035 $z matched Voyager, but incoming 035 $a did not"
'            WriteRejectRecord SourceRecord
'            OK = False
'        End If
        
    End With 'SourceRecord
    
    OkToUpdate = OK
End Function

Private Function RecordIsWanted(SourceRecord As OclcRecordType)
    'Some OCLC records are completely unwanted.
    'Log message and set flag for rejection by caller.
    Dim IsWanted As Boolean
    IsWanted = True
    
    'Reject due to various criteria.
    With SourceRecord.BibRecord
        '599 $b Removed from collection...
        .FldFindFirst "599"
        Do While .FldWasFound
            .SfdFindFirst "b"
            Do While .SfdWasFound
                If InStr(1, .SfdText, "Removed from collection", vbTextCompare) = 1 Then
                    WriteLog GL.Logfile, vbTab & "REJECTED: 599 $b Removed from collection..."
                    IsWanted = False
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
    
    RecordIsWanted = IsWanted
End Function

Private Sub PrepareForUpdate(SourceRecord As OclcRecordType)
    'Retrieves the Voyager record which will be updated by the OCLC record.
    'Does various checks for extra info which requires logging.
    Dim OclcDtSt As String
    Dim OclcEntryConv As String
    Dim OclcF1xx As String
    Dim OclcF245anp As String
    Dim VgerDtSt As String
    Dim VgerEntryConv As String
    Dim VgerF1xx As String
    Dim VgerF245anp As String
    Dim VgerHas856xCDL As Boolean
    Dim VgerRecord As Utf8MarcRecordClass
    Dim WriteVger As Boolean
    
    With SourceRecord.BibRecord
        .FldFindFirst "008"
        OclcDtSt = Mid(.FldText, 7, 1)  '008/06, via 1-based array
        OclcEntryConv = Mid(.FldText, 35, 1)    '008/34, via 1-based array
        '1xx field, except for $e
        .FldFindFirst "1"
        If .FldWasFound Then
            OclcF1xx = .FldTag & " " & Replace(.FldInd, " ", "_")
            .SfdMoveTop
            Do While .SfdMoveNext
                If .SfdCode <> "e" Then
                    OclcF1xx = OclcF1xx & " $" & .SfdCode & " " & .SfdText
                End If
            Loop
        End If
        '245 field $anp only
        .FldFindFirst "245"
        If .FldWasFound Then
            OclcF245anp = ""
            .SfdMoveTop
            Do While .SfdMoveNext
                If .SfdCode = "a" Or .SfdCode = "n" Or .SfdCode = "p" Then
                    OclcF245anp = OclcF245anp & "$ " & .SfdCode & " " & .SfdText
                End If
            Loop
        End If
    End With 'OclcRecord

DebugSourceRecord SourceRecord
    
    Set VgerRecord = GetVgerBibRecord(SourceRecord.BibMatches(1))
    With VgerRecord
        .FldFindFirst "008"
        VgerDtSt = Mid(.FldText, 7, 1)  '008/06, via 1-based array
        VgerEntryConv = Mid(.FldText, 35, 1)    '008/34, via 1-based array
        
        '1xx field, except for $e
        .FldFindFirst "1"
        If .FldWasFound Then
            VgerF1xx = .FldTag & " " & Replace(.FldInd, " ", "_")
            .SfdMoveTop
            Do While .SfdMoveNext
                If .SfdCode <> "e" Then
                    VgerF1xx = VgerF1xx & " $" & .SfdCode & " " & .SfdText
                End If
            Loop
        End If
        
        '245 field $anp only
        .FldFindFirst "245"
        If .FldWasFound Then
            VgerF245anp = ""
            .SfdMoveTop
            Do While .SfdMoveNext
                If .SfdCode = "a" Or .SfdCode = "n" Or .SfdCode = "p" Then
                    VgerF245anp = VgerF245anp & "$ " & .SfdCode & " " & .SfdText
                End If
            Loop
        End If
        
    End With 'VgerRecord
    
    'Set WriteVger flag so record gets written only once, regardless of number of log messages written.
    WriteVger = False
    
    'Compare DtSt and log if different
    If OclcDtSt <> VgerDtSt Then
        WriteLog GL.Logfile, vbTab & "INFO: 008/06 mismatch (OCLC " & OclcDtSt & ", Voyager " & VgerDtSt & ")"
        WriteVger = True
    End If
    
    'Log Entry Convention if either record has "latest entry" value (1)
    If OclcEntryConv = "1" Or VgerEntryConv = "1" Then
        WriteLog GL.Logfile, vbTab & "INFO: 008/34 Latest Entry value (OCLC " & OclcEntryConv & ", Voyager " & VgerEntryConv & ")"
        WriteVger = True
    End If
    
    'Compare 1xx fields and log if different
    If NormalizeText(OclcF1xx) <> NormalizeText(VgerF1xx) Then
        WriteLog GL.Logfile, vbTab & "INFO: 1xx mismatch"
        WriteLog GL.Logfile, vbTab & vbTab & "OCLC: " & OclcF1xx
        WriteLog GL.Logfile, vbTab & vbTab & "VGER: " & VgerF1xx
        WriteVger = True
    End If
    
    'Compare 245 $anp and log if different
    If NormalizeText(OclcF245anp) <> NormalizeText(VgerF245anp) Then
        WriteLog GL.Logfile, vbTab & "INFO: 245 $anp mismatch"
        WriteLog GL.Logfile, vbTab & vbTab & "OCLC: " & OclcF245anp
        WriteLog GL.Logfile, vbTab & vbTab & "VGER: " & VgerF245anp
        WriteVger = True
    End If
    
    If WriteVger = True Then WriteVoyagerRecord VgerRecord

End Sub

Private Sub WriteRecordToFile(MarcRecord As String, Description As String)
    'Generic procedure for writing binary MARC to file.
    'Description (rejected, voyager etc.) gets added to file name.
    
    Dim FileHandle As Integer
    Dim Filename As String
    FileHandle = FreeFile
    Filename = GL.BaseFilename & "." & Description & ".mrc"
    Open Filename For Binary Access Write As FileHandle
    'Since opening/closing file for each write, need to position write pointer at end of file using LOF() + 1
    Put FileHandle, LOF(FileHandle) + 1, MarcRecord
    Close FileHandle
End Sub

Private Sub WriteRejectRecord(SourceRecord As OclcRecordType)
    'Convenience method to write binary MARC record to reject file.
    WriteRecordToFile SourceRecord.BibRecord.MarcRecordOut, "reject"
End Sub

Private Sub WriteVoyagerRecord(VgerRecord As Utf8MarcRecordClass)
    'Convenience method to write binary MARC record to voyager (before update) file.
    WriteRecordToFile VgerRecord.MarcRecordOut, "voyager"
End Sub

Private Sub UpdateVoyager(SourceRecord As OclcRecordType)
    'Merges fields from OCLC and Voyager records and updates Voyager.
    'OCLC record is treated as master, with selected fields from Voyager merged into it.
    Dim AddField As Boolean
    Dim OclcBib As Utf8MarcRecordClass
    Dim VgerBib As Utf8MarcRecordClass
    Dim UpdateBibRC As UpdateBibReturnCode
    Dim BibID As String
    
    BibID = SourceRecord.BibMatches(1)
    
    Set OclcBib = SourceRecord.BibRecord
    Set VgerBib = GetVgerBibRecord(BibID)
    
    If DBUG = True Then
        WriteLog GL.Logfile, vbCrLf
        WriteLog GL.Logfile, "***** INCOMING OCLC RECORD:"
        WriteLog GL.Logfile, OclcBib.TextRaw
        WriteLog GL.Logfile, "***************************"
    End If
    
    If DBUG = True Then
        WriteLog GL.Logfile, "***** VOYAGER RECORD BEFORE CHANGES:"
        WriteLog GL.Logfile, VgerBib.TextRaw
        WriteLog GL.Logfile, "***************************"
    End If
    
    'Loop through Vger fields, adding selected ones to Oclc record
    With VgerBib
        .FldMoveTop
        Do While .FldMoveNext
            AddField = False
            Select Case .FldTag
                '001
                Case "001"
                    AddField = True
                '035 $9
                Case "035"
                    If .SfdFindFirst("9") = True Then AddField = True
                '590
                Case "590"
                    AddField = True
                '793
                Case "793"
                    AddField = True
                '856
                Case "856"
                    AddField = True
                '910 handled separately from other 9xx, below
                'Concatenate 910 of incoming record to beginning of existing 910 data.
                'If there are multiple 910 fields in the existing Voyager record, concatenate them to a single field before adding the 910 data from the incoming OCLC record.
                Case "910"
                    With OclcBib
                        .FldFindFirst "910"
                        'Then append current Voyager 910 to OCLC 910
                        .FldText = .FldText & VgerBib.FldText
                        'No AddField for 910 as content is being added directly to the OCLC record
                    End With 'OclcBib
                Case Else
                    'The other 9xx fields, with a few excluded
                    '901-909, 911-935, 937-999
                    If Left(.FldTag, 1) = "9" And (.FldTag <> "910" And .FldTag <> "936") Then
                        AddField = True
                    End If
                    
                    'Any field with $5 CLU
                    If .SfdFindFirst("5") Then  '$5 is not repeatable so FindFirst is right
                        If .SfdText = "CLU" Then
                            AddField = True
                        End If
                    End If
            End Select
            
            If AddField = True Then
                OclcBib.FldAddGeneric .FldTag, .FldInd, .FldText, 3
            End If
        Loop
    End With 'VgerBib
    
    'Rebuild 035 ucoclc fields for WorldCat Local
    Set OclcBib = UpdateUcoclc(OclcBib)
    
    If DBUG = True Then
        WriteLog GL.Logfile, "***** VOYAGER RECORD AFTER CHANGES:"
        'Yes, OclcBib since fields were merged into it...
        WriteLog GL.Logfile, OclcBib.TextRaw
        WriteLog GL.Logfile, "***************************"
    End If
    
    'Finally, update Voyager
    If GL.ProductionMode = True Then
        UpdateBibRC = GL.BatchCat.UpdateBibRecord( _
            CLng(BibID), _
            OclcBib.MarcRecordOut, _
            GL.Vger.BibUpdateDateVB, _
            GL.Vger.BibOwningLibraryNumber, _
            GL.CatLocID, _
            GL.Vger.BibRecordIsSuppressed _
        )
        If UpdateBibRC = ubSuccess Then
            WriteLog GL.Logfile, "Bib #" & BibID & " updated successfully"
        Else
            WriteLog GL.Logfile, "ERROR: Bib #" & BibID & " could not be updated; returncode: " & UpdateBibRC
        End If
    End If

End Sub

Private Function NormalizeText(ByVal str As String) As String
    'QAD normalization function - returns just ASCII 0-9 and letters, uppercased
    'Catalogers want to strip punctuation, capitalization and diacritics...
    'This function is not Unicode-aware but in this context that's OK.
    Dim re As RegExp
    Set re = New RegExp
    str = UCase(str)
    re.Global = True
    re.Pattern = "[^0-9A-Z_]"
    NormalizeText = re.Replace(str, "")
End Function

Private Function GetCallNumber(BibRecord As Utf8MarcRecordClass) As String
    'Returns the first acceptable call number class found in bib record 050/090.
    'Skips fields with unacceptable pseudo-classes used in some LC records.
    Dim callno As String
    Dim temp As String
    With BibRecord
        'Prefer 050 over 090
        .FldFindFirst "050"
        Do While .FldWasFound
            .SfdFindFirst "a"
            Do While .SfdWasFound
                If CallNumberIsOK(.SfdText) Then
                    GetCallNumber = .SfdText
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
        'Made it to here, try the 090
        .FldFindFirst "090"
        Do While .FldWasFound
            .SfdFindFirst "a"
            Do While .SfdWasFound
                If CallNumberIsOK(.SfdText) Then
                    GetCallNumber = .SfdText
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
End Function

Private Function CallNumberIsOK(str As String) As Boolean
    'Supports GetCallNumber by testing text against list of unacceptable pseudo-classes.
    'Reject strings which are used by LC but are not legitimate call#s for 852 $h
    Dim IsOK As Boolean
    IsOK = True 'Assume it passes until it doesn't
    str = UCase(str)
    If str = "CLASSED SEPARATELY" Or str = "CURRENT ISSUES ONLY" Or str = "DISCARD" Or str = "IN PROCESS" Or str = "INTERNET ACCESS" Or str = "ISSN RECORD" _
    Or str = "LAW" Or str = "MICROFICHE" Or str = "MICROFILM" Or str = "NEWSPAPER" Or str = "NOT IN LC" Or str = "PAR" Or str = "REV PAR" _
    Or str = "UNC" Or str = "UNCLASSED" Or str = "WMLC" Then
        IsOK = False
    End If
    
    CallNumberIsOK = IsOK
End Function

Private Sub DebugSourceRecord(SourceRecord As OclcRecordType)
    Dim cnt As Integer
    With SourceRecord
        WriteLog GL.Logfile, "==> Bib match count: " & .BibMatchCount
        'WriteLog GL.Logfile, "Bib record     : " & .BibRecord.TextRaw
        For cnt = 1 To .BibMatchCount
            WriteLog GL.Logfile, "==> Bib match " & cnt & ": " & .BibMatches(cnt)
        Next
    End With
End Sub

Private Sub AddRecord(SourceRecord As OclcRecordType)
    'Add the bib and holdings record associated with SourceRecord to Voyager
    Const LIB_ID As Long = 1
    Dim BibRC As AddBibReturnCode
    Dim HolRC As AddHoldingReturnCode
    Dim BibID As Long
    Dim HolID As Long
    Dim HolLoc As LocationType
    
If DBUG Then
    WriteLog GL.Logfile, "DEBUG: ADD RECORD TO VOYAGER"
Else
    With SourceRecord
        BibID = 0
        BibRC = GL.BatchCat.AddBibRecord( _
            .BibRecord.MarcRecordOut, _
            LIB_ID, _
            GL.CatLocID, _
            False _
        )
        If BibRC = abSuccess Then
            BibID = GL.BatchCat.RecordIDAdded
            WriteLog GL.Logfile, "Added Voyager bib: " & BibID
            
            HolID = 0
            HolLoc = GetLoc("in")   'All of these are Internet holdings
            HolRC = GL.BatchCat.AddHoldingRecord( _
                .HoldingsRecords(1).HolRecord.MarcRecordOut, _
                BibID, _
                GL.CatLocID, _
                False, _
                HolLoc.LocID _
            )
            If HolRC = ahSuccess Then
                HolID = GL.BatchCat.RecordIDAdded
                WriteLog GL.Logfile, vbTab & "Added Voyager hol: " & HolID
            Else
                WriteLog GL.Logfile, "ERROR adding holdings record - return code: " & HolRC
            End If
        Else
            WriteLog GL.Logfile, "ERROR adding bib record - return code: " & BibRC
        End If  'Add bib record
    End With 'SourceRecord
End If
End Sub
