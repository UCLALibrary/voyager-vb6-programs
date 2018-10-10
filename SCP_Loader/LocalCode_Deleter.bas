Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim BibID As Long
    Dim OCLC As String
    Dim BibRC As UpdateBibReturnCode
    Dim SourceFile As Utf8MarcFileClass
    Dim BibRecord As Utf8MarcRecordClass
    Dim RawRecord As String
    Dim SearchResult As Long
    Dim VgerRecord As Utf8MarcRecordClass
    Dim WriteToReviewFile As Boolean
    Dim Scp599 As String
    Dim Vgr599 As String
    Dim RejectFilename As String

    RejectFilename = GL.BaseFilename & ".rej.mrc"
    
    Set SourceFile = New Utf8MarcFileClass
    SourceFile.OpenFile GL.InputFilename

    Do While SourceFile.ReadNextRecord(RawRecord)
        DoEvents
        
        ' Initialise some variables for each record
        WriteToReviewFile = False
        Scp599 = ""
        Vgr599 = ""
        Set BibRecord = New Utf8MarcRecordClass
        
        With BibRecord
            'SCP records are unicode now
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .MarcRecordIn = RawRecord

            .FldFindFirst "001"
            If .FldWasFound Then
                OCLC = Trim(.FldText)
            Else
                WriteLog GL.Logfile, "ERROR: no 001 field"
                WriteLog GL.Logfile, .TextRaw
                OCLC = ""
            End If
            
            'For date comparison with Voyager
            .FldFindFirst "599"
            If .FldWasFound Then
                .SfdFindFirst "c"
                If .SfdWasFound Then
                    Scp599 = .SfdText
                End If
            End If

        End With

        SearchResult = FindBibForOCLC(OCLC)
        'SkeletonForm.lblStatus.Caption = "Processing " & OCLC & " *** " & SearchResult
       
        ' SearchResult will be:
        ' * Negative (inverse of number of matches (-2 = 2 matches, etc.))
        ' * 0 (no matches)
        ' * Positive (bib id of 1 matching record)
        If SearchResult > 0 Then
            BibID = SearchResult
            WriteLog GL.Logfile, "OCLC " & OCLC & " matches 1 record: bib " & BibID
            Set VgerRecord = GetVgerBibRecord(CStr(BibID))
            
            'For date comparison with incoming SCP record
            With VgerRecord
                .FldFindFirst "599"
                If .FldWasFound Then
                    .SfdFindFirst "c"
                    If .SfdWasFound Then
                        Vgr599 = .SfdText
                    End If
                End If
            End With
            
            If (All856AreCDL(VgerRecord) = True) And (HasOnlyInternetHoldings(BibID) = True) And (Scp599cIsNewer(Scp599, Vgr599) = True) Then
                ' Log messages are in deletion subroutines
                DeleteInternetHoldings BibID
                DeleteBibRecord BibID
            Else
                If All856AreCDL(VgerRecord) = False Then
                    WriteLog GL.Logfile, vbTab & "Not deleted: non-matching 856 $x"
                    WriteToReviewFile = True
                End If
                
                If HasOnlyInternetHoldings(BibID) = False Then
                    WriteLog GL.Logfile, vbTab & "Not deleted: non-internet holdings"
                    WriteToReviewFile = True
                End If
                
                If Scp599cIsNewer(Scp599, Vgr599) = False Then
                    WriteLog GL.Logfile, vbTab & "Not deleted: SCP 599 (" & Scp599 & ") is not newer than VGR 599 (" & Vgr599 & ")"
                    WriteToReviewFile = True
                End If
            End If
        ElseIf SearchResult = 0 Then
            ' Logging is enough, no review needed
            WriteLog GL.Logfile, "OCLC " & OCLC & " matches 0 records"
        Else ' Multiple matches
            WriteLog GL.Logfile, "OCLC " & OCLC & " matches " & (0 - SearchResult) & " records"
            WriteToReviewFile = True
        End If
        
        If WriteToReviewFile = True Then
            WriteRawRecord RejectFilename, RawRecord
            WriteLog GL.Logfile, vbTab & "Record written to review file"
        End If

        WriteLog GL.Logfile, ""
        
        NiceSleep GL.Interval
    Loop

    SourceFile.CloseFile
    SkeletonForm.lblStatus.Caption = "Done!"
End Sub

Private Function FindBibForOCLC(OCLC As String) As Long
    ' Searches Voyager for OCLC number.
    ' 2017-05-16: Per VBT-826, also check for vendor-specific numbers from 001.
    ' If 0 matches, returns 0
    ' If 1 match, returns bib id for the match
    ' If multiple matches, returns negative of # of matches (e.g., 2 matches returns -2)
    
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
        "AND ( normal_heading = 'UCOCLC" & OCLC & "' OR normal_heading = 'SCP " & UCase(OCLC) & "' ) " & _
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
    
    ' If multiple matches, return negative number of matches
    Select Case cnt
        Case 0
            result = 0
        Case 1
            result = BibID
        Case Else
            result = 0 - cnt
    End Select
    
    FindBibForOCLC = result
End Function

Private Function All856AreCDL(BibRecord As Utf8MarcRecordClass) As Boolean
    ' Each 856 must have an $x beginning with CDL or UC_ (space - not UCLA, just UC)
    ' If no 856 field, or 856 has no $x, or any $x doesn't match CDL-only values, return false.
    ' VBT-825: Check for CDL, not CDL_ (space); and verified logic in subfield comparisons.
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

Private Function HasOnlyInternetHoldings(BibID As Long) As Boolean
    ' Return true if bib record has internet holdings, and no others
    ' Else returns false (no internet, or has non-internet)
    Dim result As Boolean
    Dim SQL As String
    Dim rs As Integer
    Dim Internet_cnt As Integer
    Dim NonInternet_cnt As Integer
    
    result = False
    Internet_cnt = 0
    NonInternet_cnt = 0
    
    SQL = _
        "WITH mfhds AS ( " & vbCrLf & _
            "SELECT mm.mfhd_id, l.location_code " & vbCrLf & _
            "FROM bib_mfhd bm " & vbCrLf & _
            "INNER JOIN mfhd_master mm ON bm.mfhd_id = mm.mfhd_id " & vbCrLf & _
            "INNER JOIN location l ON mm.location_id = l.location_id " & vbCrLf & _
            "WHERE bm.bib_id = " & BibID & vbCrLf & _
        ") " & vbCrLf & _
        "SELECT " & vbCrLf & _
            "(SELECT Count(*) AS internet FROM mfhds WHERE location_code = 'in') AS internet " & vbCrLf & _
        ",  (SELECT Count(*) AS internet FROM mfhds WHERE location_code <> 'in') AS non_internet " & vbCrLf & _
        "FROM dual"

    rs = GL.GetRS
    With GL.Vger
        .ExecuteSQL SQL, rs
        If .GetNextRow Then
            Internet_cnt = .CurrentRow(rs, 1)
            NonInternet_cnt = .CurrentRow(rs, 2)
        End If
    End With
    
    GL.FreeRS rs
    
    If Internet_cnt > 0 And NonInternet_cnt = 0 Then
        result = True
    End If
    
    HasOnlyInternetHoldings = result
End Function

Private Function Scp599cIsNewer(Scp599 As String, Vgr599 As String) As Boolean
    Dim result As Boolean
    result = False
    
    If (Scp599 <> "") And (Vgr599 <> "") And (Scp599 > Vgr599) Then
        result = True
    '20181009: Also treat records with no Voyager 599 $c as older than SCP deletion candidates
    ElseIf (Scp599 <> "") And (Vgr599 = "") Then
        result = True
    End If
    
    Scp599cIsNewer = result
End Function

Private Sub DeleteInternetHoldings(BibID As Long)
    Dim rs As Integer
    Dim rc As DeleteHoldingReturnCode
    Dim HolID As Long
    
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
End Sub

Private Sub DeleteBibRecord(BibID As Long)
    Dim rc As DeleteBibReturnCode
    rc = GL.BatchCat.DeleteBibRecord(BibID)
    If rc = dbSuccess Then
        WriteLog GL.Logfile, vbTab & "Deleted bib " & BibID
    Else
        WriteLog GL.Logfile, vbTab & "ERROR deleting bib " & BibID & " - return code: " & rc
    End If
    
End Sub

