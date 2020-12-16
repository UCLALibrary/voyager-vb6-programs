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
    '''''GL.Init "-t ucladb -f " & App.Path & "\test05.mrc -i 0"
    
    DBUG = Not GL.ProductionMode
    
    ProcessRecords
    
    GL.CloseAll
    Set GL = Nothing
End Sub

Private Sub ProcessRecords()
    Dim MarcFile As New Utf8MarcFileClass
    Dim MarcRecord As New Utf8MarcRecordClass
    Dim WcmRecord As OclcRecordType
    Dim RawRecord As String
    Dim RecordNumber As Long
    Dim F001 As String
    
    RecordNumber = 0
    MarcFile.OpenFile GL.InputFilename
    Do While MarcFile.ReadNextRecord(RawRecord)
        RecordNumber = RecordNumber + 1
        If RecordNumber >= GL.StartRec Then
            Set MarcRecord = New Utf8MarcRecordClass
            With MarcRecord
                .CharacterSetIn = "U"   'Hooray, UTF-8 from OCLC
                .CharacterSetOut = "U"
                .IgnoreSfdOrder = True
                .MarcRecordIn = RawRecord
                F001 = GetOclcNumberFrom001(MarcRecord)
            End With
            Set WcmRecord.BibRecord = MarcRecord
            WcmRecord.PositionInFile = RecordNumber
            
            'Start of log entry for record
            'Subsequent functions may also write log messages for current record.
            WriteLog GL.Logfile, "Record #" & RecordNumber & ": Incoming record OCLC# " & F001
            
            If RecordIsWanted(WcmRecord) Then
                PrepareRecord WcmRecord
                GetOclcNumbers WcmRecord
                SearchVoyager WcmRecord
                
                If OkToUpdate(WcmRecord) Then
                    'Voyager record will be updated, though Voyager records may also be written to a review file.
                    PrepareForUpdate WcmRecord
                    UpdateVoyager WcmRecord
                End If 'OkToUpdate
            
            Else
                'Record is not wanted so reject it
                WriteRejectRecord WcmRecord
            End If 'RecordIsWanted
            
            'End of log entry for record
            WriteLog GL.Logfile, ""
        End If
    
    Loop 'MarcFile.ReadNextRecord
    
End Sub

Private Sub PrepareRecord(WcmRecord As OclcRecordType)
    'Add/update/delete fields within OCLC record before involving Voyager
    Dim F035 As String
    
    With WcmRecord.BibRecord
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
        
        'OCLC puts 049 at end of 0xx block, I prefer it in tag order
        .FldFindFirst "049"
        If .FldWasFound Then .FldDelete
        .FldAddGeneric "049", "  ", .SfdMake("a", "VBIB")
        
        'Early WCM records had incorrect text in 599 $b; change to make consistent with specs and later records
        .FldFindFirst "599"
        Do While .FldWasFound
            .SfdFindFirst "b"
            Do While .SfdWasFound
                If .SfdText = "ucla-coll-mgr" Then .SfdText = "UclaCollMgr"
                .SfdMoveNext
            Loop
            .FldFindNext
        Loop
        
        'Early WCM records had incorrect text in 910 $a; change to make consistent with specs and later records
        .FldFindFirst "910"
        Do While .FldWasFound
            .SfdFindFirst "a"
            Do While .SfdWasFound
                If InStr(1, .SfdText, "ucla-coll-mgr") > 0 Then
                    .SfdText = Replace(.SfdText, "ucla-coll-mgr", "UclaCollMgr")
                End If
                .SfdMoveNext
            Loop
            .FldFindNext
        Loop
        
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
        
        'Delete 856
        .FldFindFirst "856"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Delete 891 fields
        .FldFindFirst "891"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Delete all 9XX fields of incoming record except 910, 936, 948, 949, 987
        .FldFindFirst "9"
        Do While .FldWasFound
            If Not (.FldTag = "910" Or .FldTag = "936" Or .FldTag = "948" Or .FldTag = "949" Or .FldTag = "987") Then
                .FldDelete
            End If
            .FldFindNext
        Loop
        
    End With 'WcmRecord.BibRecord
    
End Sub

Private Sub GetOclcNumbers(WcmRecord As OclcRecordType)
    'Populates WcmRecord.OclcNumbers() with values from 035
    '035 $a should always be first subfield, and only one 035 $a
    'May also be any number of $z following the $a
    'Just capture the raw subfield text here, used for logging; normalize for searching in SearchVoyager().
    
    Dim OclcCount As Integer
    ReDim WcmRecord.OclcNumbers(1 To MAX_OCLC_COUNT)
    OclcCount = 0
    
    With WcmRecord.BibRecord
        .FldFindFirst "035"
        .SfdMoveTop
        Do While .SfdMoveNext
            OclcCount = OclcCount + 1
            WcmRecord.OclcNumbers(OclcCount) = .SfdText
        Loop
    End With
    'Remove unused space from array
    ReDim Preserve WcmRecord.OclcNumbers(1 To OclcCount)
End Sub

Private Sub SearchVoyager(WcmRecord As OclcRecordType)
    'Searches Voyager for OCLC number(s) in WcmRecord.OclcNumbers()
    'Populates WcmRecord.BibMatchCount and WcmRecord.BibMatches()
    'Check against UCOCLC values, which are normalized from Voyager 035 $a.
    'Check for unsuppressed records only.
    
    Dim AlreadyExists As Boolean
    Dim BibID As String
    Dim cnt As Integer
    Dim LogMessage As String
    Dim OclcCount As Integer
    Dim SearchNumber As String
    Dim SQL As String
    Dim rs As Integer 'RecordSet
    
    ReDim WcmRecord.BibMatches(1 To MAX_BIB_MATCHES)
    WcmRecord.BibMatchCount = 0
    rs = GL.GetRS
    
    For OclcCount = 1 To UBound(WcmRecord.OclcNumbers())
        SearchNumber = UCase(CalculateUcoclc(WcmRecord.OclcNumbers(OclcCount)))
        
        'Find unsuppressed Voyager serial bibs with 035 $a OCLC number matching incoming OCLC number
        SQL = _
            "select bi.bib_id " & _
            "from bib_index bi " & _
            "inner join bib_master bm " & _
            "on bi.bib_id = bm.bib_id " & _
            "inner join bib_text bt " & _
            "on bi.bib_id = bt.bib_id " & _
            "where bi.index_code = '0350' " & _
            "and bi.normal_heading = '" & SearchNumber & "' " & _
            "and bm.suppress_in_opac = 'N' " & _
            "and bt.bib_format like '%s' " & _
            "order by bi.bib_id"
            
        With GL.Vger
            .ExecuteSQL SQL, rs
            Do While .GetNextRow
                With WcmRecord
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
                        If .BibMatchCount = 1 And OclcCount > 1 Then
                            .BibMatchCount = -1
                            'Break out of the OclcCount For..Next block
                            Exit For
                        End If
                    End If
                End With 'WcmRecord
            Loop 'GetNextRow
            
        End With 'Vger
        
        'Special case: if no matches after first search (which is always for incoming OCLC 035 $a), no need to continue
        'We don't case if incoming 035 $z matches Voyager if the $a didn't.
        If OclcCount = 1 And WcmRecord.BibMatchCount = 0 Then
            'Logging for this case done later
            Exit For
        End If
    Next 'OclcCount
    
    'Remove unused space from array
    With WcmRecord
        If .BibMatchCount > 0 Then
            ReDim Preserve .BibMatches(1 To .BibMatchCount)
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

Private Function OkToUpdate(WcmRecord As OclcRecordType) As Boolean
    'Use incoming/Voyager matching results and other info to determine whether incoming record qualifies to update Voyager.
    'See specs for conditions being evaluated.
    'Conditions 1-5 are mutually exclusive.
    'Return TRUE if OK, otherwise return FALSE and write record to REJECT file, along with log messages.
    
    Dim OclcDtSt As String
    Dim OclcF005 As String
    Dim VgerDtSt As String
    Dim VgerF005 As String
    Dim VgerRecord As Utf8MarcRecordClass
    
    Dim OK As Boolean
    OK = True   'Any condition below may set to False
    
    With WcmRecord
        'Condition 1: No matches in Voyager
        If .BibMatchCount = 0 Then
            WriteLog GL.Logfile, vbTab & "REJECTED: No matches in Voyager"
            WriteRejectRecord WcmRecord
            OK = False
        End If
        
        'Conditions 2 and 5: multiple matches in Voyager
        If .BibMatchCount > 1 Then
            WriteLog GL.Logfile, vbTab & "REJECTED: Multiple matches in Voyager"
            WriteRejectRecord WcmRecord
            OK = False
        End If
        
        'Conditions 3 and 4: one match in Voyager, but 005/008 date info is not acceptable
        'Get data before evaluating these conditions
        If .BibMatchCount = 1 Then
            With .BibRecord
                .FldFindFirst "005"
                OclcF005 = .FldText
                .FldFindFirst "008"
                OclcDtSt = Mid(.FldText, 7, 1)  '008/06, via 1-based array
            End With
            
            Set VgerRecord = GetVgerBibRecord(.BibMatches(1))
            With VgerRecord
                .FldFindFirst "005"
                VgerF005 = .FldText
                .FldFindFirst "008"
                VgerDtSt = Mid(.FldText, 7, 1)  '008/06, via 1-based array
            End With
        
            'Condition 3
            If OclcF005 <= VgerF005 Then
                WriteLog GL.Logfile, vbTab & "REJECTED: OCLC is older than Voyager"
                WriteRejectRecord WcmRecord
                OK = False
            End If
            
            'Condition 4
            If OclcF005 > VgerF005 And (OclcDtSt = "c" Or OclcDtSt = "u") And VgerDtSt = "d" Then
                WriteLog GL.Logfile, vbTab & "REJECTED: OCLC is newer than Voyager but 008/06 mismatch (OCLC " & OclcDtSt & ", Voyager " & VgerDtSt & ")"
                WriteRejectRecord WcmRecord
                OK = False
            End If
            
        End If '.BibMatchCount = 1
        
        'Condition 7 (incoming 035 $z matched but incoming 035 $a did not, detected and flagged in SearchRecords())
        If .BibMatchCount = -1 Then
            WriteLog GL.Logfile, vbTab & "REJECTED: Incoming record 035 $z matched Voyager, but incoming 035 $a did not"
            WriteRejectRecord WcmRecord
            OK = False
        End If
        
    End With 'WcmRecord
    
    OkToUpdate = OK
End Function

Private Function RecordIsWanted(WcmRecord As OclcRecordType)
    'Some OCLC records are completely unwanted.
    'Log message and set flag for rejection by caller.
    Dim IsWanted As Boolean
    IsWanted = True
    
    'Reject due to various criteria.
    With WcmRecord.BibRecord
        'Non-serials (anything other than "s")
        If Mid(.Leader, 8, 1) <> "s" Then   'LDR/07, via 1-based array
            WriteLog GL.Logfile, vbTab & "REJECTED: not a serial"
            IsWanted = False
        End If
        
        'Last 040 $d = CLU
        .FldFindFirst "040"
        If .FldWasFound Then
            .SfdFindLast "d"
            If .SfdText = "CLU" Then
                WriteLog GL.Logfile, vbTab & "REJECTED: Final 040 $d = CLU"
                IsWanted = False
            End If
        End If
        
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

Private Sub PrepareForUpdate(WcmRecord As OclcRecordType)
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
    
    With WcmRecord.BibRecord
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
    
    Set VgerRecord = GetVgerBibRecord(WcmRecord.BibMatches(1))
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
        
        '856 $x CDL
        VgerHas856xCDL = False
        .FldFindFirst "856"
        Do While .FldWasFound
            .SfdFindFirst "x"
            Do While .SfdWasFound
                If .SfdText = "CDL" Then
                    VgerHas856xCDL = True
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
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
    
    'Log message if Voyager has 856 $x CDL
    If VgerHas856xCDL = True Then
        WriteLog GL.Logfile, vbTab & "INFO: Voyager has 856 $x CDL"
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

Private Sub WriteRejectRecord(WcmRecord As OclcRecordType)
    'Convenience method to write binary MARC record to reject file.
    WriteRecordToFile WcmRecord.BibRecord.MarcRecordOut, "reject"
End Sub

Private Sub WriteVoyagerRecord(VgerRecord As Utf8MarcRecordClass)
    'Convenience method to write binary MARC record to voyager (before update) file.
    WriteRecordToFile VgerRecord.MarcRecordOut, "voyager"
End Sub

Private Sub UpdateVoyager(WcmRecord As OclcRecordType)
    'Merges fields from OCLC and Voyager records and updates Voyager.
    'OCLC record is treated as master, with selected fields from Voyager merged into it.
    Dim AddField As Boolean
    Dim OclcBib As Utf8MarcRecordClass
    Dim VgerBib As Utf8MarcRecordClass
    Dim UpdateBibRC As UpdateBibReturnCode
    Dim BibID As String
    
    BibID = WcmRecord.BibMatches(1)
    
    Set OclcBib = WcmRecord.BibRecord
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
                        If .FldFindFirst("910") = False Then
                            'No OCLC 910 for some reason, so create one first
                            .FldAddGeneric "910", "  ", .SfdMake("a", "UclaCollMgr " & GetDateFromFilename()), 3
                            .FldFindFirst "910"
                        End If
                        'Then append current Voyager 910 to OCLC 910
                        .FldText = .FldText & VgerBib.FldText
                        'No AddField for 910 as content is being added directly to the OCLC record
                    End With 'OclcBib
                Case Else
                    '6xx _4
                    If Left(.FldTag, 1) = "6" And .FldInd2 = "4" Then
                            AddField = True
                    End If
                    '6xx _7 $2 local
                    If Left(.FldTag, 1) = "6" And .FldInd2 = "7" Then
                        '$2 is not repeatable
                        If .SfdFindFirst("2") Then
                            If .SfdText = "local" Then
                                AddField = True
                            End If
                        End If
                    End If
                    '69x (all)
                    If Left(.FldTag, 2) = "69" Then
                        AddField = True
                    End If
                    
                    'The other 9xx fields, with a few excluded
                    '901-909, 911-935, 937-999
                    If Left(.FldTag, 1) = "9" And (.FldTag <> "910" And .FldTag <> "936") Then
                        AddField = True
                    End If
                    
                    'Any field with $5 starting with CLU
                    If .SfdFindFirst("5") Then  '$5 is not repeatable so FindFirst is right
                        If InStr(1, .SfdText, "CLU", vbTextCompare) = 1 Then
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

Private Function GetDateFromFilename() As String
    'QAD function to get date portion (YYYYMMDD) from input filename
    'Input filenames look like: wcmserials_YYYYMMDD.mrc
    Dim YYYYMMDD As String
    YYYYMMDD = GetDigits(GL.BaseFilename)
    GetDateFromFilename = YYYYMMDD
End Function

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
