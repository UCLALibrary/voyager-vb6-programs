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
            If DBUG = True Then
                WriteLog GL.Logfile, vbCrLf & "================================================================================"
            End If
            
            WriteLog GL.Logfile, "Record #" & RecordNumber & ": Incoming record OCLC# " & F001
            If RecordIsWanted(WcmRecord) Then
                PrepareRecord WcmRecord
                GetOclcNumbers WcmRecord
                SearchVoyager WcmRecord
            
                If DBUG = True Then
                    WriteLog GL.Logfile, vbCrLf
                    WriteLog GL.Logfile, "***** INCOMING OCLC RECORD:"
                    WriteLog GL.Logfile, WcmRecord.BibRecord.TextRaw
                    WriteLog GL.Logfile, "***************************"
                End If
                
                If OkToUpdate(WcmRecord) Then
                    UpdateVoyager WcmRecord
                End If 'OkToUpdate
            
            Else
                'Record is not wanted so reject it; no need to write to file, or already logged when evaluated
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
        
        'Delete 029 fields
        .FldFindFirst "029"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Delete 049
        .FldFindFirst "049"
        If .FldWasFound Then .FldDelete
        
        'Delete 583 (retention info from other institutions)
        .FldFindFirst "583"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Delete 5xx with $5 other than CLU
        .FldFindFirst "5"
        Do While .FldWasFound
            '$5 not repeatable
            If .SfdFindFirst("5") Then
                If .SfdText <> "CLU" Then
                    .FldDelete
                End If
            End If
            .FldFindNext
        Loop
        
        'Delete 655 with 2nd indicator 4
        .FldFindFirst "655"
        Do While .FldWasFound
            If .FldInd2 = "4" Then
                .FldDelete
            End If
            .FldFindNext
        Loop
        
        'Delete 7xx with $5 other than CLU
        .FldFindFirst "7"
        Do While .FldWasFound
            '$5 not repeatable
            If .SfdFindFirst("5") Then
                If .SfdText <> "CLU" Then
                    .FldDelete
                End If
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
        
        'Delete 938
        .FldFindFirst "938"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        'Delete 994
        .FldFindFirst "994"
        Do While .FldWasFound
            .FldDelete
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
    Dim SkipLoc As String
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
        SQL = GetUcoclcSql(SearchNumber)
            
        With GL.Vger
            .ExecuteSQL SQL, rs
            Do While .GetNextRow
                With WcmRecord
                    BibID = GL.Vger.CurrentRow(rs, 1)
                    SkipLoc = GL.Vger.CurrentRow(rs, 2)
                    AlreadyExists = False
                    For cnt = 1 To .BibMatchCount
                        If BibID = .BibMatches(cnt) Then
                            AlreadyExists = True
                        End If
                    Next
                    LogMessage = ""
                    If AlreadyExists = False Then
                        'Add new unique matches, or skip if bib has locs requring rejection
                        If SkipLoc = "Y" Then
                            WriteLocationRecord WcmRecord
                            LogMessage = "REJECTED DUE TO LOCATIONS: "
                            'Set BibMatchCount to normally impossible negative value as flag to caller so I don't have to modify the OclcRecordType
                            .BibMatchCount = -1
                        Else
                            .BibMatchCount = .BibMatchCount + 1
                            .BibMatches(.BibMatchCount) = BibID
                        End If
                        'Write log messages for matches
                        LogMessage = LogMessage & "Match found: incoming OCLC 035 $" & IIf(OclcCount = 1, "a", "z") & " " & Replace(SearchNumber, "UCOCLC", "")
                        LogMessage = LogMessage & " matches Voyager bib " & BibID
                        WriteLog GL.Logfile, vbTab & LogMessage
                        'Break out of the OclcCount For..Next block if skipped due to location
                        If .BibMatchCount = -1 Then
                            Exit For
                        End If
                    End If
                End With 'WcmRecord
            Loop 'GetNextRow
            
        End With 'Vger
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
    'Return TRUE if OK, otherwise return FALSE and write record to REJECT file, along with log messages.
    
    Dim VgerRecord As Utf8MarcRecordClass
    Dim OclcSize As Long
    Dim VgerSize As Long
    Dim BibID As String
    Dim OK As Boolean
    OK = True   'Any condition below may set to False
    
    With WcmRecord
        'Skipped due to special locations in Voyager:
        If .BibMatchCount = -1 Then
            OK = False
        End If
        'No matches in Voyager: Log and discard
        If .BibMatchCount = 0 Then
            WriteLog GL.Logfile, vbTab & "DISCARDED: No matches in Voyager"
            OK = False
        End If
        
        'Multiple matches in Voyager: Log, save for review
        If .BibMatchCount > 1 Then
            WriteLog GL.Logfile, vbTab & "REVIEW NEEDED: Multiple matches in Voyager on " & .OclcNumbers(1)
            WriteReviewRecord WcmRecord
            OK = False
        End If
        
        'Conditions 3 and 4: one match in Voyager, may or may not be suppressed
        'If suppressed: Log and discard
        'Otherwise, OK to update
        'Get data before evaluating these conditions
        If .BibMatchCount = 1 Then
            BibID = WcmRecord.BibMatches(1)
            Set VgerRecord = GetVgerBibRecord(BibID)
            If GL.Vger.BibRecordIsSuppressed Then
                WriteLog GL.Logfile, vbTab & "DISCARDED: Voyager record is suppressed"
                OK = False
            
            End If
            
            'Reject incoming OCLC records which are smaller than the matching Voyager record.
            'Log and discard, write to review file
            OclcSize = Len(WcmRecord.BibRecord.MarcRecordOut)
            VgerSize = Len(VgerRecord.MarcRecordOut)
            If OclcSize < VgerSize Then
                WriteLog GL.Logfile, vbTab & "REVIEW NEEDED: Possible data loss for Voyager bib " & BibID & ": OCLC " & OclcSize & ", VGER " & VgerSize
                WriteReviewRecord WcmRecord
                OK = False
            End If
        
        End If '.BibMatchCount = 1
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
        'Non-monos (anything other than "m")
        If Mid(.Leader, 8, 1) <> "m" Then   'LDR/07, via 1-based array
            WriteLog GL.Logfile, vbTab & "DISCARDED: not a monograph"
            IsWanted = False
        End If
        '008/06 (date status) of m - multi-date mono, not sure why not wanted
        If .Get008Value(6, 1) = "m" Then
            WriteLog GL.Logfile, vbTab & "REJECTED: DtSt m"
            WriteDtstRecord WcmRecord
            IsWanted = False
        End If
    End With
    
    RecordIsWanted = IsWanted
End Function

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

Private Sub WriteReviewRecord(WcmRecord As OclcRecordType)
    'Convenience method to write binary MARC record to review file.
    WriteRecordToFile WcmRecord.BibRecord.MarcRecordOut, "review"
End Sub

Private Sub WriteLocationRecord(WcmRecord As OclcRecordType)
    'Convenience method to write binary MARC record when excluded due to Voyager locations.
    WriteRecordToFile WcmRecord.BibRecord.MarcRecordOut, "location"
End Sub

Private Sub WriteDtstRecord(WcmRecord As OclcRecordType)
    'Convenience method to write binary MARC record when excluded due to OCLC 008/06 DtSt
    WriteRecordToFile WcmRecord.BibRecord.MarcRecordOut, "dtst"
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
                '590
                Case "590"
                    AddField = True
                '599
                Case "599"
                    AddField = True
                '793
                Case "793"
                    AddField = True
                '856
                Case "856"
                    .SfdFindFirst "x"
                    Do While .SfdWasFound
                        If InStr(1, .SfdText, "CDL") = 1 Or InStr(1, .SfdText, "UCLA") = 1 Then
                            AddField = True
                        End If
                        .SfdFindNext
                    Loop
'                '910 handled separately from other 9xx, below
'                'Concatenate 910 of incoming record to beginning of existing 910 data.
'                'If there are multiple 910 fields in the existing Voyager record, concatenate them to a single field before adding the 910 data from the incoming OCLC record.
'                Case "910"
'                    With OclcBib
'                        If .FldFindFirst("910") = False Then
'                            'No OCLC 910 for some reason, so create one first
'                            .FldAddGeneric "910", "  ", .SfdMake("a", "UclaCollMgr " & GetDateFromFilename()), 3
'                            .FldFindFirst "910"
'                        End If
'                        'Then append current Voyager 910 to OCLC 910
'                        .FldText = .FldText & VgerBib.FldText
'                        'No AddField for 910 as content is being added directly to the OCLC record
'                    End With 'OclcBib
                Case Else
                    '035 $9, or 035 $a (SFXObjID)
                    If .FldTag = "035" Then
                        If .SfdFindFirst("9") = True Then
                            AddField = True
                        End If
                        If .SfdFindFirst("a") Then
                            If InStr(1, .SfdText, "(SFXObjID)") Then AddField = True
                        End If
                    End If
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
                    
                    'The other 9xx fields, with a few excluded *936/985/987
                    If Left(.FldTag, 1) = "9" And (.FldTag <> "936" And .FldTag <> "985" And .FldTag <> "987") Then
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
'        UpdateBibRC = GL.BatchCat.UpdateBibRecord( _
'            CLng(BibID), _
'            OclcBib.MarcRecordOut, _
'            GL.Vger.BibUpdateDateVB, _
'            GL.Vger.BibOwningLibraryNumber, _
'            GL.CatLocID, _
'            GL.Vger.BibRecordIsSuppressed _
'        )
'        If UpdateBibRC = ubSuccess Then
'            WriteLog GL.Logfile, "Bib #" & BibID & " updated successfully"
'        Else
'            WriteLog GL.Logfile, "ERROR: Bib #" & BibID & " could not be updated; returncode: " & UpdateBibRC
'        End If
        
        'Save merged record to file for load via bulkimport, avoiding keyword indexing performance problems
        WriteLog GL.Logfile, "Bib #" & BibID & " will be updated via bulkimport"
        WriteRecordToFile OclcBib.MarcRecordOut, "vgerload"
    Else
        WriteLog GL.Logfile, "Bib #" & BibID & " WOULD BE UPDATED"
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

Private Function GetUcoclcSql(ucoclc As String) As String
    'Given an OCLC number (formatted as UCOCLC), returns SQL this project needs
    'to search voyager for that number.
    Dim SQL As String
    SQL = GetTextFromFile("wcm_oclc_query.sql")
    SQL = Replace(SQL, "UCOCLC_PLACEHOLDER", ucoclc)
    GetUcoclcSql = SQL
End Function
