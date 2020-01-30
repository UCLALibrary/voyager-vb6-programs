Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim BibID As Long
    Dim RawRecord As String
    Dim ReviewFile As String
    Dim ScpFile As Utf8MarcFileClass
    Dim ScpNumber As String
    Dim ScpRecord As Utf8MarcRecordClass
    Dim SearchResult As Long
    Dim VgerRecord As Utf8MarcRecordClass
    Dim WriteToReviewFile As Boolean
    Dim HoldingsType As String

    WriteLog GL.Logfile, "PROD MODE: " & GL.ProductionMode
    
    ReviewFile = GL.BaseFilename & ".rej.mrc"
    Set ScpFile = New Utf8MarcFileClass
    ScpFile.OpenFile GL.InputFilename
    Do While ScpFile.ReadNextRecord(RawRecord)
        DoEvents
        Set ScpRecord = New Utf8MarcRecordClass
        With ScpRecord
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .MarcRecordIn = RawRecord
        End With
        
        ScpNumber = GetRecordNumber(ScpRecord)
        SearchResult = FindBibByOCLC(ScpNumber)
        'SearchResult will be:
        '* Negative (inverse of number of matches (-2 = 2 matches, etc.))
        '* 0 (no matches)
        '* Positive (bib id of 1 matching record)
        If SearchResult > 0 Then
            'Main logic is here
            BibID = SearchResult
            WriteLog GL.Logfile, "SCP " & ScpNumber & " matches 1 record: bib " & BibID
            Set VgerRecord = GetVgerBibRecord(CStr(BibID))
            
            If Scp599cIsNewer(ScpRecord, VgerRecord) Then
                HoldingsType = CharacterizeHoldings(BibID)
                If HasUCLA856(VgerRecord) Then
                    'Scenario 2: Has 856 $x UCLA / UCLA Law
                    If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Has 856 $x UCLA / UCLA Law"
                    'Most complex logic is in, or called from, this procedure
                    ProcessUclaRecord BibID, VgerRecord, ScpRecord, ReviewFile
                Else
                    'Scenario 1: No 856 $x UCLA / UCLA Law
                    If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: No 856 $x UCLA / UCLA Law"
                    If HoldingsType = "INTERNET_ONLY" Then
                        'Scenario 1.b
                        If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 1.b"
                        DeleteInternetHoldings BibID
                        DeleteBibRecord BibID
                    Else
                        'Scenario 1.a
                        If HasCDL856(VgerRecord) Then
                            'Scenario 1.a.i
                            If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 1.a.i"
                            DeleteBibFields VgerRecord
                            DeleteInternetHoldings BibID
                        Else
                            'Scenario 1.a.ii: No action wanted
                            If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 1.a.ii"
                            WriteLog GL.Logfile, vbTab & "No CDL 856: no action taken"
                        End If
                    End If
                End If
            Else    'Scp599cIsNewer() = False
                'Details logged within Scp599cIsNewer()
                WriteToReviewFile = True
            End If  'Scp599cIsNewer
        
        ElseIf SearchResult = 0 Then
            'Logging is enough, no review needed
            WriteLog GL.Logfile, "SCP " & ScpNumber & " matches 0 records"
        Else 'Multiple matches
            WriteLog GL.Logfile, "SCP " & ScpNumber & " matches " & (0 - SearchResult) & " records"
            WriteToReviewFile = True
        End If
        
        
        If WriteToReviewFile = True Then
            WriteRawRecord ReviewFile, RawRecord
            WriteLog GL.Logfile, vbTab & "Record written to review file"
        End If
            
        If GL.ProductionMode = True Then
            WriteLog GL.Logfile, ""
        Else
            WriteLog GL.Logfile, "***********************************************************************************************"
        End If
        NiceSleep GL.Interval
    Loop
    ScpFile.CloseFile
    SkeletonForm.lblStatus.Caption = "Done!"
End Sub

Private Function GetRecordNumber(BibRecord As Utf8MarcRecordClass) As String
    'Returns contents of 001 field - usually an OCLC number but sometimes a vendor number.
    'These both get treated the same way later on.
    Dim RecordNumber As String
    With BibRecord
        .FldFindFirst "001"
        If .FldWasFound Then
            RecordNumber = Trim(.FldText)
        Else
            WriteLog GL.Logfile, "ERROR: no 001 field"
            WriteLog GL.Logfile, .TextRaw
            RecordNumber = ""
        End If
    End With
    GetRecordNumber = RecordNumber
End Function

Private Function FindBibByOCLC(Oclc As String) As Long
    'Searches Voyager for (presumed) OCLC number.
    'Per VBT-826, also checks for vendor-specific numbers from 001.
    'If 0 matches, returns 0
    'If 1 match, returns bib id for the match
    'If multiple matches, returns negative of # of matches (e.g., 2 matches returns -2)
    
    Dim result As Long
    Dim BibID As Long
    Dim SQL As String
    Dim rs As Integer
    Dim cnt As Integer
    
    rs = GL.GetRS
    cnt = 0
    BibID = 0
    result = 0
    SQL = _
        "SELECT Bib_ID " & _
        "FROM bib_index " & _
        "WHERE index_code = '0350' " & _
        "AND ( normal_heading = 'UCOCLC" & Oclc & "' OR normal_heading = 'SCP " & UCase(Oclc) & "' ) " & _
        "ORDER BY Bib_ID"

    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While True
            If Not .GetNextRow Then
                Exit Do
            End If
            BibID = .CurrentRow(rs, 1)
            cnt = cnt + 1
        Loop
    End With 'Vger
    
    GL.FreeRS rs
    
    'If multiple matches, return negative number of matches
    Select Case cnt
        Case 0
            result = 0
        Case 1
            result = BibID
        Case Else
            result = 0 - cnt
    End Select
    
    FindBibByOCLC = result
End Function

Private Function All856AreCDL(BibRecord As Utf8MarcRecordClass) As Boolean
    'Each 856 must have an $x beginning with CDL or UC_ (space - not UCLA, just UC)
    'If no 856 field, or 856 has no $x, or any $x doesn't match CDL-only values, return false.
    'VBT-825: Check for CDL, not CDL_ (space); and verified logic in subfield comparisons.
    Dim result As Boolean
    
    result = True
    With BibRecord
        .FldFindFirst "856"
        If Not .FldWasFound Then
            result = False
        End If
        Do While .FldWasFound
            .SfdFindFirst "x"
            If Not .SfdWasFound Then
                result = False
            End If
            Do While .SfdWasFound
                '856 $x does not start with CDL or UC_ (space), case-insensitive.
                'If it does, then the test is true, and the function will fail (return false).
                If InStr(1, .SfdText, "CDL", vbTextCompare) <> 1 And InStr(1, .SfdText, "UC ", vbTextCompare) <> 1 Then
                    result = False
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
    
    All856AreCDL = result
End Function

Private Function HasCDL856(BibRecord As Utf8MarcRecordClass) As Boolean
    'Returns True if any 856 has an $x beginning with CDL or UC_ (space - not UCLA, just UC)
    Dim result As Boolean
    
    result = False
    With BibRecord
        .FldFindFirst "856"
        Do While .FldWasFound
            .SfdFindFirst "x"
            Do While .SfdWasFound
                If InStr(1, .SfdText, "CDL", vbTextCompare) = 1 Or InStr(1, .SfdText, "UC ", vbTextCompare) = 1 Then
                    result = True
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
    
    HasCDL856 = result
End Function

Private Function HasUCLA856(BibRecord As Utf8MarcRecordClass) As Boolean
    'Returns True if any 856 has an $x beginning with UCLA or UCLA Law
    Dim result As Boolean
    
    result = False
    With BibRecord
        .FldFindFirst "856"
        Do While .FldWasFound
            .SfdFindFirst "x"
            Do While .SfdWasFound
                'Data is dirty, but checking "start with UCLA" seems to cover only appropriate cases
                If InStr(1, .SfdText, "UCLA", vbTextCompare) = 1 Then
                    result = True
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
    
    HasUCLA856 = result
End Function

Private Function CharacterizeHoldings(BibID As Long) As String
    'Return strings describing the set of holdings attached to the bib.
    'PRINT means "non-internet" here.
    '* INTERNET_ONLY
    '* SUPPRESSED_PRINT_ONLY
    '* UNSUPPRESSED_PRINT
    Dim result As String
    Dim SQL As String
    Dim rs As Integer
    Dim InternetHoldings As Integer
    Dim SuppressedPrint As Integer
    Dim UnsuppressedPrint As Integer
    
    result = ""
    
    SQL = _
        "with d as ( " & vbCrLf & _
        "  select " & vbCrLf & _
        "    bm.* " & vbCrLf & _
        "  , l.location_code " & vbCrLf & _
        "  , mm.suppress_in_opac as mfhd_suppress " & vbCrLf & _
        "  from bib_mfhd bm " & vbCrLf & _
        "  inner join mfhd_master mm on bm.mfhd_id = mm.mfhd_id " & vbCrLf & _
        "  inner join location l on mm.location_id = l.location_id " & vbCrLf & _
        "  where bm.bib_id = " & BibID & vbCrLf & _
        ") " & vbCrLf & _
        "select " & vbCrLf & _
        "  (select count(*) from d where location_code = 'in') as internet_holdings " & vbCrLf & _
        ", (select count(*) from d where location_code != 'in' and mfhd_suppress = 'Y') as suppressed_print " & vbCrLf & _
        ", (select count(*) from d where location_code != 'in' and mfhd_suppress = 'N') as unsuppressed_print " & vbCrLf & _
        "from dual " & vbCrLf
    
    rs = GL.GetRS
    With GL.Vger
        .ExecuteSQL SQL, rs
        If .GetNextRow Then
            InternetHoldings = .CurrentRow(rs, 1)
            SuppressedPrint = .CurrentRow(rs, 2)
            UnsuppressedPrint = .CurrentRow(rs, 3)
        End If
    End With
    
    GL.FreeRS rs
    
    'Assume they all have internet holdings, but verify for INTERNET_ONLY
    If InternetHoldings > 0 And SuppressedPrint = 0 And UnsuppressedPrint = 0 Then
        result = "INTERNET_ONLY"
    ElseIf SuppressedPrint > 0 And UnsuppressedPrint = 0 Then
        result = "SUPPRESSED_PRINT_ONLY"
    Else
        result = "UNSUPPRESSED_PRINT"
    End If
    
    If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: CharacterizeHoldings = " & result
    
    CharacterizeHoldings = result
End Function

Private Sub DeleteInternetHoldings(BibID As Long)
    Dim rs As Integer
    Dim rc As DeleteHoldingReturnCode
    Dim HolID As Long

If GL.ProductionMode Then
    rs = GL.GetRS
    With GL.Vger
        .SearchHoldNumbersForBib CStr(BibID), rs
        Do While .GetNextRow(rs)
            .RetrieveHoldRecord .CurrentRow(1)
            If .HoldLocationCode = "in" Then
                HolID = .HoldRecordNumber
                rc = GL.BatchCat.DeleteHoldingRecord(HolID)
                If rc = dhSuccess Then
                    WriteLog GL.Logfile, vbTab & "Deleted hol " & HolID
                Else
                    WriteLog GL.Logfile, vbTab & "ERROR deleting hol " & HolID & " - return code: " & rc
                End If
            End If
        Loop
    End With
    GL.FreeRS rs
Else
    WriteLog GL.Logfile, vbTab & "DEBUG: Internet holdings would be deleted from bib: " & BibID
End If

End Sub

Private Sub DeleteBibRecord(BibID As Long)
    Dim rc As DeleteBibReturnCode
If GL.ProductionMode Then
    rc = GL.BatchCat.DeleteBibRecord(BibID)
    If rc = dbSuccess Then
        WriteLog GL.Logfile, vbTab & "Deleted bib " & BibID
    Else
        WriteLog GL.Logfile, vbTab & "ERROR deleting bib " & BibID & " - return code: " & rc
    End If
Else
    WriteLog GL.Logfile, vbTab & "DEBUG: Bib record would be deleted: " & BibID
End If

End Sub

Private Sub DeleteBibFields(BibRecord As Utf8MarcRecordClass)
    'i.  Delete all 856 fields containing $x "CDL" and/or $x "UC open access"
    'ii. Delete all 793 fields
    'iii.Delete 590 field containing $a "UCLA Library - CDL shared resource"
    'iv. Delete 599 field where $a equals "UPD," "DEL," or "NEW", and $c is present
    '''If GL.ProductionMode = False Then WriteLog GL.Logfile, "*** BEFORE *** " & vbCrLf & BibRecord.TextFormatted & vbCrLf
    Dim BibID As Long
    
    With BibRecord
        .FldFindFirst "856"
        Do While .FldWasFound
            .SfdFindFirst "x"
            Do While .SfdWasFound
                If InStr(1, .SfdText, "CDL", vbTextCompare) = 1 Or InStr(1, .SfdText, "UC open access", vbTextCompare) = 1 Then
                    .FldDelete
                    Exit Do
                Else
                    .SfdFindNext
                End If
            Loop
            .FldFindNext
        Loop
        
        .FldFindFirst "793"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop

        .FldFindFirst "590"
        Do While .FldWasFound
            .SfdFindFirst "a"
            Do While .SfdWasFound
                If InStr(1, .SfdText, "UCLA Library - CDL shared resource", vbTextCompare) = 1 Then
                    .FldDelete
                    Exit Do
                Else
                    .SfdFindNext
                End If
            Loop
            .FldFindNext
        Loop

        .FldFindFirst "599"
        Do While .FldWasFound
            .SfdFindFirst "a"
            Do While .SfdWasFound
                If .SfdText = "DEL" Or .SfdText = "NEW" Or .SfdText = "UPD" Then
                    .SfdFindFirst "c"
                    If .SfdWasFound Then
                        .FldDelete
                        Exit Do
                    End If
                Else
                    .SfdFindNext
                End If
            Loop
            .FldFindNext
        Loop
    End With
    '''If GL.ProductionMode = False Then WriteLog GL.Logfile, "*** AFTER *** " & vbCrLf & BibRecord.TextFormatted & vbCrLf
    
    BibID = CLng(GetRecordNumber(BibRecord))
    If GL.ProductionMode = True Then
        'Add code to save record here, or in caller?
    Else
        WriteLog GL.Logfile, vbTab & "DEBUG: Fields deleted, bib record would be updated: " & BibID
    End If
    
End Sub

Private Function Scp599cIsNewer(ScpRecord As Utf8MarcRecordClass, VgerRecord As Utf8MarcRecordClass) As Boolean
    Dim Scp599c As String
    Dim Vgr599c As String
    
    Dim result As Boolean
    result = False
    
    Scp599c = Get599c(ScpRecord)
    Vgr599c = Get599c(VgerRecord)
    If (Scp599c <> "") And (Vgr599c <> "") And (Scp599c > Vgr599c) Then
        result = True
    '20181009: Also treat records with no Voyager 599 $c as older than SCP deletion candidates
    ElseIf (Scp599c <> "") And (Vgr599c = "") Then
        result = True
    End If

If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: SCP 599: " & Scp599c & " *** VGR 599: " & Vgr599c
    
    'Log here since we have the specific 599 $c values
    If result = False Then
        WriteLog GL.Logfile, vbTab & "SKIPPING: SCP 599 (" & Scp599c & ") is not newer than VGR 599 (" & Vgr599c & ")"
    End If
    
    Scp599cIsNewer = result
End Function

Private Function Get599c(BibRecord As Utf8MarcRecordClass) As String
    'Returns first 599 $c found in record
    Dim f599c As String
    f599c = ""
    With BibRecord
        .FldFindFirst "599"
        Do While .FldWasFound And f599c = ""
            .SfdFindFirst "c"
            If .SfdWasFound Then
                f599c = .SfdText
            End If
            .FldFindNext
        Loop
    End With
    Get599c = f599c
End Function

Private Function IsRcRecord(BibRecord As Utf8MarcRecordClass) As Boolean
    'Determines whether record has 599 $b rc
    Dim result As Boolean
    result = False
    With BibRecord
        .FldFindFirst "599"
        Do While .FldWasFound
            .SfdFindFirst "b"
            Do While .SfdWasFound
                If .SfdText = "rc" Then
                    result = True
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
    If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: RC record: " & result
    IsRcRecord = result
End Function

Private Sub ProcessUclaRecord(BibID As Long, VgerRecord As Utf8MarcRecordClass, ScpRecord As Utf8MarcRecordClass, ReviewFile As String)
    'Implements complex logic for processing SCP deletes where the Voyager record has UCLA 856 fields.
    'Most steps are the same across all cases, but there are variations.
    'BibID and VgerRecord point to same Voyager record matching the SCP record on OCLC#, both passed for convenience.
    'ScpRecord needed for 599 $b
    'TargetRecord (declared and searched for below) is the "Online version" of the record, in Voyager - not the same as VgerRecord, which represents print.
    
    Dim HoldingsType As String
    Dim TargetRecord As Utf8MarcRecordClass
    Dim TargetBibID As Long
        
    'The rest depend on holdings associated with the source Voyager record.
    HoldingsType = CharacterizeHoldings(BibID)
    Select Case HoldingsType
        'Scenario 2.a
        Case "UNSUPPRESSED_PRINT"
            If IsRcRecord(ScpRecord) Then
            'Scenario 2.a.i
                If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 2.a.i"
                Set TargetRecord = FindTargetRecord(ScpRecord)
                If Not (TargetRecord Is Nothing) Then
                    TargetBibID = CLng(GetRecordNumber(TargetRecord))
                    Move856Fields VgerRecord, TargetRecord
                    DeleteInternetHoldings TargetBibID
                    MoveInternetHoldings BibID, TargetBibID
                    DeleteBibFields VgerRecord
                    '''''TODO: Report on online OCLC# and UCLA holdings info (2.a.i.9-10)
                Else
                    WriteLog GL.Logfile, vbTab & "WARNING: No online version found in Voyager - see review file"
                    WriteRawRecord ReviewFile, VgerRecord.MarcRecordOut
                End If
            Else
                'Scenario 2.a.ii
                If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 2.a.ii"
                DeleteBibFields VgerRecord
            End If
        'Scenario 2.b
        Case "SUPPRESSED_PRINT_ONLY"
            If IsRcRecord(ScpRecord) Then
            'Scenario 2.b.i
                If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 2.b.i"
                Set TargetRecord = FindTargetRecord(ScpRecord)
                If Not (TargetRecord Is Nothing) Then
                    TargetBibID = CLng(GetRecordNumber(TargetRecord))
                    Move856Fields VgerRecord, TargetRecord
                    DeleteInternetHoldings TargetBibID
                    MoveInternetHoldings BibID, TargetBibID
                    DeleteBibFields VgerRecord
                    '''''TODO: Report on online OCLC# and print OCLC# (2.b.i.9-10)
                    SuppressBibRecord BibID, VgerRecord
                Else
                    WriteLog GL.Logfile, vbTab & "WARNING: No online version found in Voyager - see review file"
                    WriteRawRecord ReviewFile, VgerRecord.MarcRecordOut
                End If
            Else
                'Scenario 2.b.ii
                If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 2.b.ii"
                DeleteBibFields VgerRecord
            End If
        'Scenario 2.c
        Case "INTERNET_ONLY"
            If IsRcRecord(ScpRecord) Then
            'Scenario 2.c.i
                If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 2.c.i"
                Set TargetRecord = FindTargetRecord(ScpRecord)
                If Not (TargetRecord Is Nothing) Then
                    TargetBibID = CLng(GetRecordNumber(TargetRecord))
                    Move856Fields VgerRecord, TargetRecord
                    DeleteInternetHoldings TargetBibID
                    MoveInternetHoldings BibID, TargetBibID
                    '''''TODO: Report on online OCLC# and print OCLC# (2.c.i.5-6)
                    DeleteBibRecord BibID
                Else
                    WriteLog GL.Logfile, vbTab & "WARNING: No online version found in Voyager - see review file"
                    WriteRawRecord ReviewFile, VgerRecord.MarcRecordOut
                End If
            Else
                'Scenario 2.c.ii
                If GL.ProductionMode = False Then WriteLog GL.Logfile, vbTab & "DEBUG: Scenario 2.c.i"
                DeleteBibFields VgerRecord
            End If
    End Select
    
End Sub

Private Function FindTargetRecord(BibRecord As Utf8MarcRecordClass) As Utf8MarcRecordClass
    'Searches Voyager for the new SCP online record we should have, which basically supersedes the current record.
    
    Dim Oclc As String
    Dim BibID As Long
    Dim Success As Boolean
    Dim TargetRecord As Utf8MarcRecordClass
    
    Success = False
    With BibRecord
        .FldFindFirst "776"
        Do While .FldWasFound And Success = False
            .SfdFindFirst "w"
            Do While .SfdWasFound And Success = False
                If InStr(1, .SfdText, "(OCoLC)", vbTextCompare) = 1 Then
                    Oclc = Replace(.SfdText, "(OCoLC)", "", vbTextCompare)
                    WriteLog GL.Logfile, vbTab & "Searching 776 $w " & Oclc
                    BibID = FindBibByOCLC(Oclc)
                    If BibID > 0 Then
                        Set TargetRecord = GetVgerBibRecord(CStr(BibID))
                        If HasCDL856(TargetRecord) Then
                            Success = True
                        End If
                    End If
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
    
    If Success = True Then
        WriteLog GL.Logfile, vbCrLf & "FOUND VIA 776: " & Oclc & vbCrLf & TargetRecord.TextFormatted & vbCrLf
        Set FindTargetRecord = TargetRecord
    Else
        WriteLog GL.Logfile, vbCrLf & "FOUND VIA 776: NO QUALIFYING RECORD"
        Set FindTargetRecord = Nothing
    End If
End Function

Private Sub Move856Fields(SourceRecord As Utf8MarcRecordClass, TargetRecord As Utf8MarcRecordClass)
    'Moves selected 856 fields from one Voyager record to another.
    'Also changes indicators for 856 fields in the target record.
    'Saves the source and target records after the updates.
    
    Dim Changed As Boolean
    
    'If GL.ProductionMode = False Then WriteLog GL.Logfile, "*** SOURCE BEFORE *** " & vbCrLf & SourceRecord.TextFormatted & vbCrLf
    'If GL.ProductionMode = False Then WriteLog GL.Logfile, "*** TARGET BEFORE *** " & vbCrLf & TargetRecord.TextFormatted & vbCrLf

    'Move selected 856 fields from one Voyager record to another.
    With SourceRecord
        Changed = False
        .FldFindFirst "856"
        Do While .FldWasFound
            .SfdFindFirst "x"
            Do While .SfdWasFound
                If .SfdText = "UCLA" Or .SfdText = "UCLA Law" Then
                    TargetRecord.FldAddGeneric .FldTag, .FldInd, .FldText, 3
                    .FldDelete
                    Changed = True
                    Exit Do 'Break out of subfield loop
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
    End With
    
    If Changed = True Then
        'Save source and target
    End If
    
    'If GL.ProductionMode = False Then WriteLog GL.Logfile, "*** SOURCE AFTER *** " & vbCrLf & SourceRecord.TextFormatted & vbCrLf
    
    'Change indicators for 856 fields in the target record.
    With TargetRecord
        Changed = False
        .FldFindFirst "856"
        Do While .FldWasFound
            If .FldInd <> "40" Then
                .FldInd = "40"
                Changed = True
            End If
            .FldFindNext
        Loop
    End With
    
    If Changed = True Then
        'Save target
        'If GL.ProductionMode = False Then WriteLog GL.Logfile, "*** TARGET AFTER *** " & vbCrLf & TargetRecord.TextFormatted & vbCrLf
    End If
    
End Sub

Private Sub MoveInternetHoldings(SourceBibID As Long, TargetBibID As Long)
    'Moves all internet holdings records from the source bib to the target bib
    Dim rs As Integer
    Dim HolID As Long
    Dim rc As UpdateHoldingReturnCode

    rs = GL.GetRS
    With GL.Vger
        .SearchHoldNumbersForBib CStr(SourceBibID), rs
        Do While .GetNextRow(rs)
            .RetrieveHoldRecord .CurrentRow(1)
            If .HoldLocationCode = "in" Then
                HolID = .HoldRecordNumber
                If GL.ProductionMode = True Then
                    rc = GL.BatchCat.UpdateHoldingRecord(HolID, .HoldRecord, .HoldUpdateDateVB, GL.CatLocID, TargetBibID, .HoldLocationID, .HoldRecordIsSuppressed)
                    If rc = uhSuccess Then
                        WriteLog GL.Logfile, vbTab & "Moved hol " & HolID & " from bib " & SourceBibID & " to " & TargetBibID
                    Else
                        WriteLog GL.Logfile, vbTab & "ERROR moving hol " & HolID & " from bib " & SourceBibID & " to " & TargetBibID & " - return code: " & rc
                    End If
                Else
                    WriteLog GL.Logfile, vbTab & "DEBUG: Hol record " & HolID & " would be moved from bib " & SourceBibID & " to " & TargetBibID
                End If
            End If
        Loop
    End With
    GL.FreeRS rs
End Sub

Private Sub SuppressBibRecord(BibID As Long, BibRecord As Utf8MarcRecordClass)
    'Supresses the given bib record
    Dim rc As UpdateBibReturnCode
If GL.ProductionMode Then
    With GL.Vger
        rc = GL.BatchCat.UpdateBibRecord(BibID, BibRecord.MarcRecordOut, .BibUpdateDateVB, .BibOwningLibraryNumber, GL.CatLocID, True)
        If rc = ubSuccess Then
            WriteLog GL.Logfile, vbTab & "Suppressed bib " & BibID
        Else
            WriteLog GL.Logfile, vbTab & "ERROR suppressing bib " & BibID & " - return code: " & rc
        End If
    End With
Else
    WriteLog GL.Logfile, vbTab & "DEBUG: Bib record would be suppressed: " & BibID
End If
End Sub
