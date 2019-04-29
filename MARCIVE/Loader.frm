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

#Const DBUG = False
Private Const UBERLOGMODE As Boolean = False

Private Const ERROR_BAR As String = "*** ERROR ***"
Private Const MAX_RECORD_COUNT As Integer = 4000

'*****
'Form-level globals - let's keep this list short!
'MARCIVE: different processing rules based on types of records, pre-segmented by vendor
'These are set by SetModes() based on filename
'New or Changed?
Private g_IsNew As Boolean
'Mono or Serial? (according to vendor, not MARC)
Private g_IsSerial As Boolean
'Mixed or Online-only? (only applies to serials)
Private g_IsOnlineOnly As Boolean
'*****

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

    ReDim OclcRecords(1 To MAX_RECORD_COUNT) As OclcRecordType  '2000 should be plenty
    
    SetModes
    GetLoadableRecords GL.InputFilename, OclcRecords()
    LoadRecords OclcRecords()
    
    GL.CloseAll
    Set GL = Nothing
End Sub

Private Sub SetModes()
    'Filename encodes info needed to process file correctly
    'Char 1=M (Marcive); 2=C/N (changed/new); 3=M/S (mono/serial); 4=M/O/X (mixed-format/online-only/null)
    'Char 5-10 = YYMMDD
    Dim AllOK As Boolean
    Dim Filename As String
    Dim Modes As String

    AllOK = False
    Filename = GetFileFromPath(GL.InputFilename)
    Modes = ""
    
    Select Case Mid(Filename, 2, 1)
        Case "C"
            g_IsNew = False
            Modes = "Changed"
            AllOK = True
        Case "N"
            g_IsNew = True
            Modes = "New"
            AllOK = True
        Case Else
            WriteLog GL.Logfile, "ERROR: Can't determine new/changed from filename: " & Filename & " - exiting"
            AllOK = False
    End Select
    
    Select Case Mid(Filename, 3, 1)
        Case "M"
            g_IsSerial = False
            Modes = Modes & ", Mono"
            AllOK = True
        Case "S"
            g_IsSerial = True
            Modes = Modes & ", Serial"
            AllOK = True
        Case Else
            WriteLog GL.Logfile, "ERROR: Can't determine mono/serial from filename: " & Filename & " - exiting"
            AllOK = False
    End Select

    Select Case Mid(Filename, 4, 1)
        Case "M"
            g_IsOnlineOnly = False
            Modes = Modes & ", Mixed"
            AllOK = True
        Case "O"
            g_IsOnlineOnly = True
            Modes = Modes & ", Online only"
            AllOK = True
        Case "X"
            g_IsOnlineOnly = False
            Modes = Modes 'no change: neither mixed nor online only
            AllOK = True
        Case Else
            WriteLog GL.Logfile, "ERROR: Can't determine mixed/online-only from filename: " & Filename & " - exiting"
            AllOK = False
    End Select
    
    If AllOK = True Then
        WriteLog GL.Logfile, "These records are: " & Modes
        WriteLog GL.Logfile, ""
    Else
        GL.CloseAll
        End
    End If
End Sub

Private Sub GetLoadableRecords(InputFilename As String, OclcRecords() As OclcRecordType)
    'Fills OclcRecords() with records eligible for loading:
    '   - dedups input file on 001, keeping final version of record
    '   - writes earlier versions of dup records to DupFile
    '   - keeps original position of each record and occurrence count of 001 for reporting
    
    Dim DupFile As Integer  'file handle
    Dim MarcFile As New Utf8MarcFileClass
    Dim MarcRecord As New Utf8MarcRecordClass
    Dim OclcRecord As OclcRecordType
    Dim RecordsKept As Integer
    Dim RecordsRead As Integer
    Dim InternalDupFound As Boolean
    Dim RawRecord As String
    Dim TxtFile As Integer  'File handle
    Dim F001 As String
    Dim cnt As Integer
    
    DupFile = FreeFile
    Open GL.BaseFilename + ".dups" For Binary As DupFile

    'For now, at least, write text version of all all records in input (pre-deduped) file
    TxtFile = OpenLog(InputFilename + ".txt") 'Yes, inputfilename (*.ucla -> *.ucla.txt)
    WriteLog TxtFile, "UCLA records before deduping"
    WriteLog TxtFile, ""
    
    WriteLog GL.Logfile, "Deduping " & GetFileFromPath(InputFilename) & " on 001 field (OCLC#), keeping latest occurrence only:"
    
    RecordsKept = 0
    RecordsRead = 0
    MarcFile.OpenFile InputFilename
    Do While MarcFile.ReadNextRecord(RawRecord)
        RecordsRead = RecordsRead + 1
        Set MarcRecord = New Utf8MarcRecordClass    'Treat each record as a separate instance
        With MarcRecord
            .CharacterSetIn = "O"   'OCLC records
            .IgnoreSfdOrder = True
            .MarcRecordIn = RawRecord
            .FldFindFirst ("001")
            'OCLC's 001 fields have 1 trailing space, which we don't want; if not using GetDigits be sure to Trim()
            F001 = GetDigits(.FldText)
            
            'Write it to the text file
            '.CharacterSetOut = "U" '*** DOES CONVERSION TO UNICODE HERE CAUSE PROBLEMS LATER ??? ***
            WriteLog TxtFile, "*** Record number " & RecordsRead & " ***" '& .CharacterSetOut
            WriteLog TxtFile, .TextFormatted(latin1)
            WriteLog TxtFile, ""
        End With
        
        'Check array to see if we already have a record to load with this OCLC#
        'Replace earlier record with the current one, noting # of occurrences of this oclc#
        InternalDupFound = False
        For cnt = 1 To RecordsKept
            OclcRecord = OclcRecords(cnt)
            If OclcRecord.OclcNumbers(1) = F001 Then
                With OclcRecord
                    'Write the earlier bib record to dup file
                    Put DupFile, , RawRecord    'Write the record unchanged
                    'Replace the earlier with the new one
                    Set .BibRecord = MarcRecord
                    .OccurrenceCount = .OccurrenceCount + 1
                    InternalDupFound = True
                    WriteLog GL.Logfile, "Duplicate (OCLC# " & F001 & "): replaced record #" & .PositionInFile & _
                        " with #" & RecordsRead 'equal to the current record#
                End With
                OclcRecords(cnt) = OclcRecord
                Exit For
            End If
        Next

        'Not a dup, so add this record to the array
        If InternalDupFound = False Then
            RecordsKept = RecordsKept + 1
            With OclcRecord
                Set .BibRecord = MarcRecord
                'This record's OCLC number(s)
                ReDim .OclcNumbers(1 To MAX_OCLC_COUNT)
                .OclcNumbers(1) = F001
                '.OclcNumber = f001
                .OccurrenceCount = 1
                .PositionInFile = RecordsRead
            End With
            OclcRecords(RecordsKept) = OclcRecord
        End If
    Loop 'MarcFile.ReadNextRecord

    CloseLog TxtFile
    Close DupFile
    MarcFile.CloseFile

    WriteLog GL.Logfile, "Records read: " & RecordsRead
    WriteLog GL.Logfile, "Records kept: " & RecordsKept
    WriteLog GL.Logfile, ""
    'Remove unused part of the array
    If RecordsKept > 0 Then
        ReDim Preserve OclcRecords(1 To RecordsKept)
    Else
        'No records to process, so quit
'**** Make this more graceful.....
        GL.CloseAll
        End
    End If
End Sub

Private Sub LoadRecords(OclcRecords() As OclcRecordType)
    Const LIB_ID As Long = 1
    
    Dim ExistingHols(1 To 1000) As HoldingsLocType '1000 is several times current max
    Dim HolLoc As HoldingsLocType
    
    Dim LocCode As String
    Dim OclcCnt As Integer
    Dim BibID As Long
    Dim HolID As Long
    Dim NewHolID As Long
    Dim itemID As Long
    Dim NewItemID As Long
    Dim OclcRecord As OclcRecordType
    Dim MarcRecord As Utf8MarcRecordClass
    Dim OclcHolRecord As HoldingsRecordType
    Dim Message As String
    Dim HolCnt As Integer
    Dim HolLocCnt As Integer
    Dim HolMatch As Boolean
    Dim ItemCnt As Integer
    Dim ReviewFile As Integer       'FileHandle for binary file
    Dim ReviewTextFile As Integer   'FileHandle for text file
    Dim pos As Integer
    Dim cnt As Integer
    Dim rs As Integer
    Dim OkToProcess As Boolean
    Dim BibReplaced As Boolean
    
    rs = GL.GetRS
    
    ReviewFile = FreeFile
    Open GL.BaseFilename & ".review" For Binary As ReviewFile
    
    ReviewTextFile = FreeFile
    Open GL.BaseFilename & ".review.txt" For Output As ReviewTextFile
    
    '*** Start of main loop through records ***
    For OclcCnt = GL.StartRec To UBound(OclcRecords)
        OclcRecord = OclcRecords(OclcCnt)
If UBERLOGMODE Then
    WriteLog GL.Logfile, "**********"
    WriteLog GL.Logfile, "*** INCOMING OCLC RECORD (BEFORE PREPROCESSING) ***"
    WriteLog GL.Logfile, OclcRecord.BibRecord.TextRaw
    WriteLog GL.Logfile, ""
End If
        
        PreprocessRecord OclcRecord
        Parse049 OclcRecord

If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** INCOMING OCLC RECORD (AFTER PREPROCESSING) ***"
    WriteLog GL.Logfile, OclcRecord.BibRecord.TextRaw
    WriteLog GL.Logfile, ""
End If
        
        BuildHoldings OclcRecord
        SearchDB OclcRecord
        With OclcRecord
            Message = "Record #" & OclcCnt & ": Incoming record OCLC# " & .OclcNumbers(1)
            'BibMatchCount is based on 035 $a $z.  Before loading MARCIVE serials, must check to see:
            '- does record also have matches based on 776 $w (OCoLC)?
            '- is record "new" or "changed" per MARCIVE?
            If .BibMatchCount = 0 Then
                'Load new-to-file
                WriteLog GL.Logfile, Message & " : no match found based on incoming 035 $a $z"
                OkToProcess = True  'prove me wrong
                If g_IsSerial = True Then
                    If SearchDB776w(OclcRecord) = True Then
                        WriteLog GL.Logfile, "*** ERROR: record matches Voyager only on 776 $w (OCoLC) - rejected, see review file"
                        Put ReviewFile, , .BibRecord.MarcRecordOut
                        Print #ReviewTextFile, "### Incoming record matches Voyager only via 776 $w (OCoLC) ###"
                        Print #ReviewTextFile, .BibRecord.TextFormatted
                        Print #ReviewTextFile, ""
                        OkToProcess = False
                    Else 'serial, no 035 matches, no 776 matches
                        If g_IsNew = False Then
                            'Record is "changed" serial per MARCIVE, not already in Voyager, reject to review file
                            WriteLog GL.Logfile, "*** ERROR: 'changed' serial MARCIVE record does not match Voyager - rejected, see review file"
                            Put ReviewFile, , .BibRecord.MarcRecordOut
                            Print #ReviewTextFile, "### 'Changed' MARCIVE serial with no match in Voyager ###"
                            Print #ReviewTextFile, .BibRecord.TextFormatted
                            Print #ReviewTextFile, ""
                            OkToProcess = False
                        End If 'g_IsNew
                    End If 'SearchDB776w
                Else
                    'mono, no match in Voyager, no restrictions on adding
                End If 'g_IsSerial
                If OkToProcess = True Then
                    'AddBibRecord writes success/failure messages to log
                    BibID = AddBibRecord(OclcRecord, LIB_ID)
                    If BibID <> 0 Then
                        'Now add holdings
                        For HolCnt = 1 To .HoldingsRecordCount
                            OclcHolRecord = .HoldingsRecords(HolCnt)
                            NewHolID = AddHolRecord(OclcHolRecord, BibID)
                        Next
                    End If 'BibID <> 0
                End If 'OkToProcess 0 matches
            ElseIf .BibMatchCount = 1 Then 'incoming matches exactly one Voyager 035 $a or $z
                'Attempt overlay of existing Voyager record
                'For MARCIVE serials, must also check to see whether 776 $w matches
                '2006-08-25: per vbross, for now at least, reject MNSO (new, serial, onlineonly) records that match Voyager; cataloging differences
                '  require human review
                WriteLog GL.Logfile, Message & " : found 1 match based on incoming 035 $a $z"
                OkToProcess = True  'prove me wrong
                If g_IsSerial = True Then
                    If g_IsNew = False Or g_IsOnlineOnly = False Then
                        If SearchDB776w(OclcRecord) = True Then
                            WriteLog GL.Logfile, "*** ERROR: record matches Voyager on 035 and on 776 $w (OCoLC) - rejected, see review file"
                            Put ReviewFile, , .BibRecord.MarcRecordOut
                            Print #ReviewTextFile, "### Multiple Matches on 035 $a $z and on 776 $w (OCoLC) ###"
                            Print #ReviewTextFile, .BibRecord.TextFormatted
                            Print #ReviewTextFile, ""
                            OkToProcess = False
                        End If 'SearchDB776w
                    Else
                        'Serial, new and online-only: reject uncategorically
                        WriteLog GL.Logfile, "*** ERROR: MNSO record matches Voyager - rejected, see review file"
                        Put ReviewFile, , .BibRecord.MarcRecordOut
                        Print #ReviewTextFile, "### MNSO record matches Voyager ###"
                        Print #ReviewTextFile, .BibRecord.TextFormatted
                        Print #ReviewTextFile, ""
                        OkToProcess = False
                    End If 'not new, not online-only
                End If 'g_IsSerial
                If OkToProcess = True Then
                    BibID = .BibMatches(1)
                    'ReplaceBibRecord writes success/failure messages to log
                    BibReplaced = ReplaceBibRecord(OclcRecord, BibID)
                    If BibReplaced = True Then
                        'Get existing Voyager holdings records for this bib
                        With GL.Vger
                            .SearchHoldNumbersForBib CStr(BibID), rs
                            'Save HolID & LocCode for each in an array, since Vger rs can only be iterated through once
                            HolLocCnt = 0
                            Do While .GetNextRow(rs)
                                HolLocCnt = HolLocCnt + 1
                                HolLoc.HolID = .CurrentRow(rs, 1)
                                HolLoc.LocCode = GetHolLocationCode(HolLoc.HolID)
                                ExistingHols(HolLocCnt) = HolLoc
                            Loop
                        End With
                        'For replaced bibs, add (or replace) *ONLY* the Internet holdings; discard/ignore others
                        For HolCnt = 1 To OclcRecord.HoldingsRecordCount
                            OclcHolRecord = OclcRecord.HoldingsRecords(HolCnt)
                            With OclcHolRecord
                                If .ClCode = "CLYY" Then
                                    HolMatch = False
                                    'Replace record if loc matches
                                    For cnt = 1 To HolLocCnt
                                        If .MatchLoc = ExistingHols(cnt).LocCode Then
                                            HolMatch = True
                                            HolID = ExistingHols(cnt).HolID
                                            ReplaceHolRecord OclcHolRecord, HolID
                                        End If 'MatchLoc
                                    Next 'HolLocCnt
        
                                    'No existing holdings, or none match on loc, so add
                                    If HolMatch = False Then
                                        NewHolID = AddHolRecord(OclcHolRecord, BibID)
                                    End If 'HolMatch false
                                End If '.ClCode = CLYY
                            End With 'OclcHolRecord
                        Next 'HolCnt
                    Else 'BibReplaced = false
                        WriteLog GL.Logfile, "*** ERROR: Bib record was not replaced due to above errors - see review file"
                        Put ReviewFile, , .BibRecord.MarcRecordOut
                        Print #ReviewTextFile, "### Rejected due to data mismatch - see log file ###"
                        Print #ReviewTextFile, .BibRecord.TextFormatted
                        Print #ReviewTextFile, ""
                    End If 'BibReplaced
                End If 'OkToProcess one match
            Else 'incoming matches more than 1 Voyager record on 035 $a $z
                WriteLog GL.Logfile, Message & " : found multiple matches based on incoming 035 $a $z"
                WriteLog GL.Logfile, "*** ERROR: record matches " & .BibMatchCount & " records on 035 $a or $z - rejected, see review file"
                Put ReviewFile, , .BibRecord.MarcRecordOut
                Print #ReviewTextFile, "### Multiple Matches on 035 $a $z ###"
                Print #ReviewTextFile, .BibRecord.TextFormatted
                Print #ReviewTextFile, ""
            End If 'BibMatchCount
        End With 'OclcRecord
        
'VB has no "continue" to exit current loop iteration - must use evil goto
EndOfLoop:
        'Blank line in log after each full record is handled
        WriteLog GL.Logfile, ""
        'Let some time go by so we don't flood the server
        NiceSleep GL.Interval
    Next 'OclcCnt
    Close ReviewFile
    Close ReviewTextFile
    GL.FreeRS rs
End Sub

Private Sub PreprocessRecord(RecordIn As OclcRecordType)
    'Makes several changes to bib record:
    '- Create new 035 $a from 003 & 001: 035 $a(003)[001 digits only]; remove 001/003
    '- Move each 019 $a to a $z in the above 035
    '- Remove most 9xx fields
    '- Create 910
    '- Fix leader if needed
    '
    'While handling 001 & 019 OCLC numbers, add to OclcRecord's OclcNumbers()
    '20071108: per slayne, added 65x -> 69x routine from daily OCLC loader
    
    Dim Ind As String
    Dim ind2 As String
    Dim Text As String
    Dim FldPointer As Integer
    
    Dim F001 As String
    Dim f003 As String
    Dim f019a As String
    Dim f035 As String
    Dim EncLevel As String
    Dim OclcCnt As Integer
    Dim Add856x_UCLA As Boolean
    Dim MarcRecord As Utf8MarcRecordClass

    Set MarcRecord = RecordIn.BibRecord
    With MarcRecord
        If .FldFindFirst("001") Then
            F001 = LeftPad(GetDigits(.FldText), "0", 8)
            .FldDelete
        End If
        If .FldFindFirst("003") Then
            f003 = .FldText
            .FldDelete
        End If
        
        f035 = .SfdMake("a", "(" & f003 & ")" & F001)
        
        'Convert 019 $a to 035 $z, using same 035 created from 001/003
        '019 is not repeatable, but $a is
        OclcCnt = 1     '001 was added when deduping file in GetLoadableRecords
        If .FldFindFirst("019") Then
            .SfdFindFirst "a"
            Do While .SfdWasFound
                f019a = LeftPad(.SfdText, "0", 8)
                f035 = f035 & .SfdMake("z", "(" & f003 & ")" & f019a)
                OclcCnt = OclcCnt + 1
                RecordIn.OclcNumbers(OclcCnt) = f019a
                .SfdFindNext
            Loop
            .FldDelete
        End If
        .FldAddGeneric "035", "  ", f035, 3
        
        'Remove unused space from array
        ReDim Preserve RecordIn.OclcNumbers(1 To OclcCnt)
        
        'Add $x UCLA to all 856 fields
        'For serials only: If no $3, add $3 Available issues:
        .FldFindFirst ("856")
        Do While .FldWasFound
            If .GetLeaderValue(7, 1) = "s" Then
                .SfdFindFirst "3"
                If Not .SfdWasFound Then
                    .FldText = .SfdMake("3", "Available issues:") & .FldText
                End If
            End If
            .FldText = .FldText & .SfdMake("x", "UCLA")
            .FldFindNext
        Loop '856
        
        'Delete existing 9XX fields
        .FldFindFirst "9"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop '9XX
        
        'Add 910 $a govdocs with date
        .FldAddGeneric "910", "  ", .SfdMake("a", "marcive " & DateToYYMMDD(Now())), 3
        
        'Make sure LDR/20-23 = 4500
        .ChangeLeaderValue 20, "4500"
        
        'For serials, change LDR/17 (encoding level) as needed
        If .GetLeaderValue(7, 1) = "s" Then
            Select Case .GetLeaderValue(17, 1)
                Case "3"
                    .ChangeLeaderValue 17, "K"
                Case "5"
                    .ChangeLeaderValue 17, " "
                Case Else
                    'no change
            End Select
        Else 'non-serials have different rules
            Select Case .GetLeaderValue(17, 1)
                Case "3"
                    .ChangeLeaderValue 17, "K"
                Case "5"
                    .ChangeLeaderValue 17, "K"
                Case Else
                    'no change
            End Select
        End If
        
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

    End With 'MarcRecord
    
    Set RecordIn.BibRecord = MarcRecord
End Sub

Private Sub Parse049(RecordIn As OclcRecordType)
    'Parse bib records 049 field to create holdings record(s) and item(s)
    ReDim RecordIn.HoldingsRecords(1 To 10) As HoldingsRecordType   '10 should be plenty
    Dim HolRecord As HoldingsRecordType
    Dim HolRecordCnt As Integer
    
    Dim AddInternetHoldings As Boolean
    Dim F049 As String
    Dim CluArray() As String
    Dim CluCnt As Integer
    
    Dim item As ItemRecordType
    Dim BibLevel As String

    Dim cnt As Integer
    Dim fldptr As Integer
    Dim sfdptr As Integer
        
    ReDim Parsed049Chunks(1 To 3) As String

    Dim BibRecord As Utf8MarcRecordClass
    Set BibRecord = RecordIn.BibRecord
   
    With BibRecord
        BibLevel = .GetLeaderValue(7, 1)
        'If 856 meets certain criteria, we'll need to make an Internet holdings record later
        'Assume all 856 fields fail to meet criteria, until we find otherwise
        AddInternetHoldings = False
        .FldFindFirst "856"
        Do While (.FldWasFound = True) And (AddInternetHoldings = False)
            'Check mono records for undesired descriptions in $3
            'Leaving references to serials, in case we process gov doc serials in future
            If BibLevel = "b" Or BibLevel = "i" Or BibLevel = "s" Then
                'Serials - we don't care about $3
                AddInternetHoldings = True
            Else
                'Monos = not serials
                .SfdFindFirst "3"
                '$3 is not repeatable so no need to check beyond the 1st
                If .SfdWasFound Then
                    If Is8563TextOK(.SfdText) Then
                        'Phrases not found so this URL is good
                        AddInternetHoldings = True
                    End If
                Else
                    'No $3, so no bad phrases, so this URL is good
                    AddInternetHoldings = True
                End If '$3
            
                'Check indicator 2 for undesirable values - this trumps $3 evaluation above
                'Applies to monos only per vbross 2006-08-25
                If .FldInd2 = "2" Then
                    AddInternetHoldings = False
                End If
            
            End If 'BibLevel
            
            .FldFindNext
        Loop '856
        
        'Marcive records should come with no 049; if one exists, log it and delete it
        .FldFindFirst "049"
        Do While .FldWasFound
            WriteLog GL.Logfile, "*** Warning: Marcive record contained 049 field: " & .FldText
            .FldDelete
            .FldFindNext
        Loop
        
        If AddInternetHoldings = True Then
            .FldAddGeneric "049", "  ", .SfdMake("a", "CLYY"), 3
        Else
            WriteLog GL.Logfile, "*** Warning: No valid 856 fields, so no 049 created"
        End If

        If .FldFindFirst("049") Then
            HolRecordCnt = 0
            .SfdMoveTop
            Do While .SfdMoveNext
                Select Case .SfdCode
                    'a: Call number prefix, CL code, call number suffix
                    Case "a"
                        'Finish the current HolRecord record (if appropriate)
                        If HolRecordCnt > 0 Then
                            'Free up unused allocated space
                            With HolRecord
                                If .ItemCount > 0 Then ReDim Preserve .Items(1 To .ItemCount)
                                If .NoteCount > 0 Then ReDim Preserve .Notes(1 To .NoteCount)
                            End With
                            'Save it if it's good
                            If GetClInfo(HolRecord.ClCode).DefaultLoc <> "INVALID" Then
                                RecordIn.HoldingsRecords(HolRecordCnt) = HolRecord
                            Else
                                'No good, so decrement HolRecordCnt - next 049 HolRecord will overwrite this one
                                WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - BAD CL CODE: " & HolRecord.ClCode & " - holdings record not created"
                                HolRecordCnt = HolRecordCnt - 1
                            End If
                        End If

                        'Initialize new record
                        HolRecordCnt = HolRecordCnt + 1
                        RecordIn.HoldingsRecords(HolRecordCnt) = HolRecord
                        With HolRecord
                            ReDim .Items(1 To MAX_ITEM_COUNT)
                            ReDim .Notes(1 To MAX_NOTE_COUNT)
                            .CallNumPrefix = ""
                            .CallNumSuffix = ""
                            .ClCode = ""
                            .CopyNum = ""
                            .ItemCount = 0
                            .MatchLoc = ""
                            .NewLoc = ""
                            .NoteCount = 0
                            .Summary = ""
                        End With
                        
                        'Parse the $a
                        Parse049Chunk .SfdText, Parsed049Chunks
                        With HolRecord
                            .CallNumPrefix = Parsed049Chunks(1)
                            .ClCode = UCase(Parsed049Chunks(2)) 'force to upper for later matching on table
                            .CallNumSuffix = Parsed049Chunks(3)
                        End With
                    
                    'c: Copy "number"
                    Case "c"
                        HolRecord.CopyNum = .SfdText
                    
                    'l: Item enumeration, barcode, item type
                    Case "l"
                        With HolRecord
                            Parse049Chunk BibRecord.SfdText, Parsed049Chunks
                            .ItemCount = .ItemCount + 1
                            With item
                                .Enum = Parsed049Chunks(1)
                                .Barcode = UCase(Parsed049Chunks(2))    'barcode should always be upper case
                                .ItemCode = LCase(Parsed049Chunks(3))   'item code should always be lower case
                            End With
                            .Items(.ItemCount) = item
                        End With
                    
                    'n: Notes
                    Case "n"
                        With HolRecord
                            .NoteCount = .NoteCount + 1
                            .Notes(.NoteCount) = BibRecord.SfdText
                        End With
                    
                    'o: Location override & match
                    Case "o"
                        With HolRecord
                            'Fudge a bit to allow reuse of Parse049Chunk
                            Parse049Chunk "[]" & BibRecord.SfdText, Parsed049Chunks
                            '(1) is "", not of interest
                            'If these are not set here, they'll be set in BuildHoldings, based on 049 $a (ClCode)
                            .MatchLoc = LCase(Parsed049Chunks(2))   'locations are always lower case
                            .NewLoc = LCase(Parsed049Chunks(3))     'locations are always lower case
                        End With
                    
                    
                    'v: Summary holdings
                    Case "v"
                        HolRecord.Summary = .SfdText
                    
                    'None of the above
                    Case Else
                        WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - Invalid 049 subfield " & .SfdCode & ":" & .SfdText
                End Select
            Loop '049 .SfdMoveNext
            
            'Clean up & store the final HolRecord in the array
            'Free up unused allocated space
            With HolRecord
                If .ItemCount > 0 Then ReDim Preserve .Items(1 To .ItemCount)
                If .NoteCount > 0 Then ReDim Preserve .Notes(1 To .NoteCount)
            End With
            'Save it if it's good
            If GetClInfo(HolRecord.ClCode).DefaultLoc <> "INVALID" Then
                RecordIn.HoldingsRecords(HolRecordCnt) = HolRecord
            Else
                'No good, so decrement HolRecordCnt - next 049 HolRecord will overwrite this one
                WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - BAD CL CODE: " & HolRecord.ClCode & " - holdings record not created"
                HolRecordCnt = HolRecordCnt - 1
            End If
        Else 'No 049
            WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - No 049 found"
        End If '049
        
    End With 'BibRecord
    
    'If 049 contained an $a code which means "bib-only", drop any holdings created (via 049 or 856)
    With RecordIn
        For cnt = 1 To HolRecordCnt
            If GetClInfo(.HoldingsRecords(cnt).ClCode).DefaultLoc = "SKIP" Then
                HolRecordCnt = 0
                Exit For
            End If
        Next
        'Remove unused allocated space
        .HoldingsRecordCount = HolRecordCnt
        If .HoldingsRecordCount > 0 Then
            ReDim Preserve .HoldingsRecords(1 To .HoldingsRecordCount)
        Else
            ReDim .HoldingsRecords(1 To 1)  'can't do 1 to 0
        End If
    End With
End Sub

Private Sub Parse049Chunk(Chunk As String, arr() As String)
    'Input: Text from 049, with up to 3 parts:
    '- [text1] in [] (optional)
    '- text2 (mandatory)
    '- [text3] in [] (optional)
    'Output: arr(1 to 3), with all values in order; optional values not supplied returned as ""
    Dim temp As String
    Dim pos1 As Integer
    Dim pos2 As Integer
    ReDim arr(1 To 3)

    temp = Trim(Chunk)
    
    '1st part: (optional) text in brackets
    If Left(temp, 1) = "[" Then
        pos1 = InStr(1, temp, "]")
        If pos1 > 1 Then
            arr(1) = GetBracketedText(Mid(temp, 1, pos1))
            'Chop off the prefix
            temp = Trim(Right(temp, Len(temp) - pos1))
        Else
            arr(1) = ""
        End If
    End If
    '2nd part: (mandatory) text
    pos1 = InStr(1, temp, "[")
    If pos1 = 0 Then
        'No (optional) 3rd part, so 2nd part is everything remaining
        arr(2) = temp
        arr(3) = ""
    Else
        '2nd part is what's in front of 3rd part...
        arr(2) = Trim(Left(temp, pos1 - 1))
        arr(3) = GetBracketedText(Right(temp, Len(temp) - pos1 + 1))
    End If
End Sub

Private Sub BuildHoldings(RecordIn As OclcRecordType)
    'Populates RecordIn.HoldingsRecords()
    Dim OclcHoldingsRecord As HoldingsRecordType
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolCnt As Integer
    
    Dim ValidHoldingsRecord As Boolean
    
    Dim InternetHoldings As Boolean
    Dim BibLDR_06 As String
    Dim BibLDR_07 As String
    Dim Bib008_06 As String     'for internet holdings records
    Dim CallNum_H As String
    Dim CallNum_I As String
    Dim CallNumInd As String
    Dim ClTag As String
    Dim ClTagArray() As String
    Dim F007 As String
    Dim ItemCnt As Integer
    
    Dim cnt As Integer
    
    With RecordIn
        'First get bib info, which will apply to all holdings records for this bib
        With .BibRecord
            'LDR
            BibLDR_06 = .GetLeaderValue(6, 1)
            BibLDR_07 = .GetLeaderValue(7, 1)
            '007
            If .FldFindFirst("007") Then
                F007 = .FldText
            Else
                F007 = "t"  'default, if there's no bib 007
            End If
            '008
            Bib008_06 = .Get008Value(6, 1)
        End With 'BibRecord
        
        'Get info necessary to build each specific holdings record
        For HolCnt = 1 To .HoldingsRecordCount
            OclcHoldingsRecord = .HoldingsRecords(HolCnt)
            'Assume it's good until proven otherwise
            ValidHoldingsRecord = True
            With OclcHoldingsRecord
                Select Case .ClCode
                    ' Should be only CLYY for MARCIVE records
                    Case "CLYY"
                        ClTagArray() = Split("050,090", ",")
                        .NewLoc = "in"
                End Select
                
                CallNum_H = ""
                CallNum_I = ""
                CallNumInd = "  "
                
                For cnt = 0 To UBound(ClTagArray)
                    With RecordIn.BibRecord
                        ClTag = ClTagArray(cnt)
                        Select Case ClTag
                            Case "050"
                                .FldFindLast ClTag
                                CallNumInd = "0 "
                            Case "090"
                                .FldFindFirst ClTag
                                CallNumInd = "0 "
                        End Select
                        If .FldWasFound Then
                            .SfdFindFirst "a"
                            If .SfdWasFound Then
                                'If not an LC call number, keep looking
                                If IsLcClass(.SfdText) Then
                                    CallNum_H = .SfdText
                                    'Stop looking - we have our match
                                    Exit For
                                End If
                            End If
                        End If
                    End With 'RecordIn.BibRecord
                Next 'cnt CallNumberTags
                    
                'Store call# info in OclcHoldingsRecord
                .CallNum_H = CallNum_H
                .CallNum_I = CallNum_I
                .CallNumInd = CallNumInd
                    
                If .MatchLoc = "" Then
                    'Setting MatchLoc allows replacement of holdings
                    .MatchLoc = .NewLoc
                End If
            End With 'OclcHoldingsRecord
            
            ' For govdocs, all holdings records are constructed like internet holdings,
            ' with a few location-specific tweaks for print holdings (loc already done, above; rest below)
            InternetHoldings = True
            CreateNewHoldingsRecord OclcHoldingsRecord, BibLDR_07, F007, Bib008_06, InternetHoldings
            With OclcHoldingsRecord
                Select Case .ClCode
                    Case "CLUD", "CLUL"
                        With .HolRecord
                            .FldFindFirst "007"
                            If .FldWasFound Then
                                .FldText = "t"
                            End If
                            .FldFindFirst "866"
                            If .FldWasFound Then
                                .FldDelete
                            End If
                        End With
                End Select
            End With

If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** NEW HOLDINGS RECORD ***"
    WriteLog GL.Logfile, OclcHoldingsRecord.HolRecord.TextRaw
    WriteLog GL.Logfile, ""
End If

            'Store the updated OclcHoldingsRecord back in RecordIn.HoldingsRecords()
            .HoldingsRecords(HolCnt) = OclcHoldingsRecord

        Next 'HolCnt
    End With 'RecordIn

#If DBUG Then
'    Dump_OHR RecordIn
#End If

End Sub

Private Sub CreateNewHoldingsRecord(OclcHR As HoldingsRecordType, BibLDR_07 As String, F007 As String, Bib008_06 As String, InternetHoldings As Boolean)
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolLDR_06 As String
    Dim F852 As String
    Dim F901 As String
    Dim cnt As Integer
    
    'Serial
    If IsSerial(BibLDR_07) Then
        HolLDR_06 = "y"
    'Mono (with/without summary holdings)
    ElseIf OclcHR.Summary = "" Then
        HolLDR_06 = "x"
    Else
        HolLDR_06 = "v"
    End If
        
    Set HolRecord = New Utf8MarcRecordClass
    With HolRecord
        .CharacterSetIn = "O"   'Not Unicode yet - some data comes from bib record, which is OCLC
        .CharacterSetOut = "O"  'Not Unicode yet - some data comes from bib record, which is OCLC
        .NewRecord HolLDR_06

'Kludge to allow conversion from Oclc to Unicode of .NewRecord - if .MarcRecordIn = "" conversion not possible 01 Aug 2004 ak
.MarcRecordIn = .MarcRecordOut

        Select Case HolLDR_06
            Case "x"
                .ChangeLeaderValue 17, "2"
                .Change008Value 6, "0"
                .Change008Value 7, "u"
                .Change008Value 8, "    "   '08-11: 4 blanks
                .Change008Value 12, "8"
                .Change008Value 13, "   "   '13-15: 3 blanks
                .Change008Value 16, "4"
                .Change008Value 17, "001"   '17-19
                .Change008Value 20, "a"
                .Change008Value 21, "u"
                .Change008Value 22, "eng"   '22-24
                .Change008Value 25, "0"
            Case "v"
                .ChangeLeaderValue 17, "3"
                .Change008Value 6, "0"
                .Change008Value 7, "u"
                .Change008Value 8, "    "   '08-11: 4 blanks
                .Change008Value 12, "8"
                .Change008Value 13, "   "   '13-15: 3 blanks
                .Change008Value 16, "0"
                .Change008Value 17, "001"   '17-19
                .Change008Value 20, "a"
                .Change008Value 21, "u"
                .Change008Value 22, "eng"   '22-24
                .Change008Value 25, "0"
            Case "y"
                .ChangeLeaderValue 17, "3"
                .Change008Value 6, "5"
                .Change008Value 7, "u"
                .Change008Value 8, "uuuu"   '08-11: 4 u
                .Change008Value 12, "8"
                .Change008Value 13, "   "   '13-15: 3 blanks
                .Change008Value 16, "0"
                .Change008Value 17, "001"   '17-19
                .Change008Value 20, "u"
                .Change008Value 21, "u"
                .Change008Value 22, "eng"   '22-24
                .Change008Value 25, "0"
        End Select
        
        'Override some LDR & 008 values for internet records
        'MARCIVE - some values differ from standard OCLC load
        If InternetHoldings Then
            'LDR/06: for internet monos, reset this based on bib 008/06
            If HolLDR_06 <> "y" Then
                If Bib008_06 = "s" Then
                    .ChangeLeaderValue 6, "x"
                Else
                    .ChangeLeaderValue 6, "v"
                End If
            End If
            'LDR/17
            .ChangeLeaderValue 17, "2"      'for all records
            .ChangeLeaderValue 18, "n"      'for all records: no item information
            '008
            'Internet serial (y)
            If HolLDR_06 = "y" Then
                .Change008Value 6, "0"
                .Change008Value 7, "d"      'MARCIVE: method of acq: depository
                .Change008Value 8, "    "   '08-11: 4 blanks
                .Change008Value 12, "0"     'MARCIVE: General retention: not applicable
                .Change008Value 22, "   "   'MARCIVE: 22-24: (language) 3 blanks
            Else
            'Internet mono (x, v)
                .Change008Value 6, "0"
                .Change008Value 7, "d"      'MARCIVE: method of acq: depository
                .Change008Value 12, "0"     'MARCIVE: General retention: not applicable
                .Change008Value 16, "4"     'MARCIVE: completeness = other
                .Change008Value 20, "u"
                .Change008Value 22, "eng"   'MARCIVE: language = english
            End If
        End If
        
        'Override some data if internet holdings
        'Changed phrase from "Access available online" to "Online access" per PSC/slayne
        'Also no longer restricted to serials - applies internet holdings for all record types 31 Aug 2004 ak
        If InternetHoldings Then
            F007 = "cr"
            With OclcHR
                .CallNum_I = ""
                .CallNumPrefix = ""
                .CallNumSuffix = ""
                .CopyNum = ""
                .Summary = "Online access"
            End With
        End If
        
        'Add 007
        .FldAddGeneric "007", "", F007, 3
        
        'Build 852 from data in OclcHR
        F852 = .SfdMake("b", OclcHR.NewLoc)
        If OclcHR.CallNumPrefix <> "" Then
            F852 = F852 & .SfdMake("k", OclcHR.CallNumPrefix)
        End If
        If OclcHR.CallNum_H <> "" Then
            F852 = F852 & .SfdMake("h", OclcHR.CallNum_H)
        Else
            'No $h: undefined call number type, override 852 indicators
            OclcHR.CallNumInd = "  "
        End If
        If OclcHR.CallNum_I <> "" Then
            F852 = F852 & .SfdMake("i", OclcHR.CallNum_I)
        End If
        If OclcHR.CallNumSuffix <> "" Then
            F852 = F852 & .SfdMake("m", OclcHR.CallNumSuffix)
        End If
        If OclcHR.CopyNum <> "" Then
            F852 = F852 & .SfdMake("t", OclcHR.CopyNum)
        End If
        If OclcHR.NoteCount > 0 Then
            For cnt = 1 To OclcHR.NoteCount
                F852 = F852 & .SfdMake("z", OclcHR.Notes(cnt))
            Next
        End If

        .FldAddGeneric "852", OclcHR.CallNumInd, F852, 3
        
        If OclcHR.Summary <> "" Then
            .FldAddGeneric "866", " 0", .SfdMake("a", OclcHR.Summary), 3
        End If
        
        'This should convert to Unicode but doesn't (email to GS 01 Aug 2004)
        'See .MarcRecordIn = .MarcRecordOut above for workaround
        .CharacterSetOut = "U"  'Now set it to Unicode
    End With
        
    Set OclcHR.HolRecord = HolRecord
End Sub

Private Sub SearchDB(RecordIn As OclcRecordType)
    'Searches Voyager for all OCLC numbers in RecordIn.OclcNumbers()
    'Modifies RecordIn: places all matching Voyager BibIDs in RecordIn.BibMatches, with total count in .BibMatchCount for convenience
    'MARCIVE: For serials, check incoming 776 $w(OCoLC) as well, and handle differently if match if via that number

    Dim SearchNumber As String
    Dim SQL As String
    Dim BibID As String
    Dim rs As Integer
    Dim cnt As Integer
    Dim AlreadyExists As Boolean
    Dim OclcCnt As Integer
    Dim F776w As String

    rs = GL.GetRS

    RecordIn.BibMatchCount = 0
    ReDim RecordIn.BibMatches(1 To MAX_BIB_MATCHES)

    For OclcCnt = 1 To UBound(RecordIn.OclcNumbers)
        SearchNumber = "OCOLC " & Normalize0350(RecordIn.OclcNumbers(OclcCnt))
    
        SQL = _
            "SELECT Bib_ID " & _
            "FROM Bib_Index " & _
            "WHERE Index_Code = '0350' " & _
            "AND Normal_Heading = '" & SearchNumber & "' " & _
            "ORDER BY Bib_ID"
        
        With GL.Vger
#If DBUG Then
'    WriteLog GL.Logfile, "Searchnumber: " & SearchNumber & " - OclcCnt = " & OclcCnt
#End If
            .ExecuteSQL SQL, rs
            Do While True
                If Not .GetNextRow Then
                    Exit Do
                End If
                With RecordIn
                    BibID = GL.Vger.CurrentRow(rs, 1)
                    AlreadyExists = False
                    For cnt = 1 To .BibMatchCount
                        If BibID = .BibMatches(cnt) Then
                            AlreadyExists = True
                        End If
                    Next
                    If AlreadyExists = False Then
                        .BibMatchCount = .BibMatchCount + 1
                        .BibMatches(.BibMatchCount) = BibID
#If DBUG Then
'    WriteLog GL.Logfile, "Found bibid: " & BibID
#End If
                    End If
                End With
            Loop
        End With 'Vger
    Next 'OclcCnt
        
    'Shrink the array
    With RecordIn
        If .BibMatchCount > 0 Then
            ReDim Preserve .BibMatches(1 To .BibMatchCount)
        End If
    End With

    GL.FreeRS rs
End Sub

Private Function SearchDB776w(RecordIn As OclcRecordType) As Boolean
    'MARCIVE: Searches Voyager 0350 OCLC# index using 776 $w(OCoLC) of incoming record
    'Return true if match found

    Dim SearchNumber As String
    Dim SQL As String
    Dim BibID As String
    Dim rs As Integer
    Dim cnt As Integer
    Dim MatchFound As Boolean
    Dim F776w As String

    rs = GL.GetRS

    MatchFound = False
    With RecordIn.BibRecord
        .FldFindFirst "776"
        Do While .FldWasFound
            .SfdFindFirst "w"
            Do While .SfdWasFound
                F776w = ""
                If Left(.SfdText, 7) = "(OCoLC)" Then
                    F776w = LeftPad(GetDigits(.SfdText), "0", 8)
                    SearchNumber = "OCOLC " & Normalize0350(F776w)
                
                    SQL = _
                        "SELECT Bib_ID " & _
                        "FROM Bib_Index " & _
                        "WHERE Index_Code = '0350' " & _
                        "AND Normal_Heading = '" & SearchNumber & "' " & _
                        "ORDER BY Bib_ID"
                    
                    With GL.Vger
                        #If DBUG Then
                            'WriteLog GL.Logfile, "OCLC number: " & SearchNumber & " - from 776 $w"
                        #End If
                        .ExecuteSQL SQL, rs
                        Do While True
                            If Not .GetNextRow Then
                                Exit Do
                            End If
                            MatchFound = True
                            BibID = GL.Vger.CurrentRow(rs, 1)
                            #If DBUG Then
                                'WriteLog GL.Logfile, vbTab & "Found bibid: " & BibID & " - from 776 $w"
                            #End If
                        Loop 'True
                    End With 'Vger
                End If 'OCoLC
                .SfdFindNext
            Loop '.SfdWasFound
            .FldFindNext
        Loop 'FldWasFound
    End With 'RecordIn.BibRecord
    
    GL.FreeRS rs
    SearchDB776w = MatchFound
End Function

Private Function AddBibRecord(RecordIn As OclcRecordType, LibraryID As Long) As Long
    'Adds new Voyager bib record; returns new record's ID
    
    Dim ReturnCode As AddBibReturnCode
    Dim BibID As Long
    
    Dim OclcBib As New Utf8MarcRecordClass
    Set OclcBib = RecordIn.BibRecord
    OclcBib.CharacterSetOut = "U"
    
    ' 20080416: Add/update ucoclc 035 fields needed for WorldCat Local before updating Voyager
    Set OclcBib = UpdateUcoclc(OclcBib)
    
    If GL.ProductionMode Then
        BibID = 0
        ReturnCode = GL.BatchCat.AddBibRecord(OclcBib.MarcRecordOut, LibraryID, GL.CatLocID, False)
        
        If ReturnCode = abSuccess Then
            BibID = GL.BatchCat.RecordIDAdded
            '20090615: added OCLC number to "Added Voyager" log entry for later reporting
            WriteLog GL.Logfile, "Added Voyager bib#" & BibID & " : OCLC " & RecordIn.OclcNumbers(1)
        Else
            WriteLog GL.Logfile, ERROR_BAR
            WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - AddBibRecord failed with returncode: " & ReturnCode
            WriteLog GL.Logfile, OclcBib.TextFormatted
            WriteLog GL.Logfile, ERROR_BAR
        End If
        
        AddBibRecord = BibID
    Else
        AddBibRecord = -1
        WriteLog GL.Logfile, "TESTMODE ONLY: new Voyager bib# NOT added"
    End If
End Function

Private Function AddHolRecord(RecordIn As HoldingsRecordType, BibID As Long) As Long
    'Adds new Voyager holdings record, linked to BibID; returns new record's ID
    Dim ReturnCode As AddHoldingReturnCode
    Dim HolID As Long
    Dim HolLoc As LocationType
    Dim OclcHol As New Utf8MarcRecordClass
    Set OclcHol = RecordIn.HolRecord
    OclcHol.CharacterSetOut = "U"
    
    HolLoc = GetLoc(RecordIn.NewLoc)
    'If loc's no good, reject holdings record to log file and return 0 as the new ID
    If HolLoc.LocID = 0 Then
        WriteLog GL.Logfile, ERROR_BAR
        WriteLog GL.Logfile, "ERROR: invalid location " & RecordIn.NewLoc & " - holdings record not added"
        WriteLog GL.Logfile, OclcHol.TextFormatted
        WriteLog GL.Logfile, ERROR_BAR
        AddHolRecord = 0
        Exit Function
    End If

    If GL.ProductionMode Then
        HolID = 0
        ReturnCode = GL.BatchCat.AddHoldingRecord(OclcHol.MarcRecordOut, BibID, GL.CatLocID, HolLoc.Suppressed, HolLoc.LocID)
        If ReturnCode = ahSuccess Then
            HolID = GL.BatchCat.RecordIDAdded
            WriteLog GL.Logfile, vbTab & "Added Voyager hol#" & HolID
        Else
            WriteLog GL.Logfile, ERROR_BAR
            WriteLog GL.Logfile, "ERROR - AddHolRecord failed with returncode: " & ReturnCode
            WriteLog GL.Logfile, OclcHol.TextRaw
            WriteLog GL.Logfile, ERROR_BAR
        End If
        
        AddHolRecord = HolID
    Else
        AddHolRecord = -1
        WriteLog GL.Logfile, "TESTMODE ONLY: new Voyager hol# NOT added"
    End If
End Function

Private Function ReplaceBibRecord(RecordIn As OclcRecordType, BibID As Long)
    'Replaces Voyager record# identified by BibID with the RecordIn.BibRecord
    'Incoming record is treated as the "master" - only selected fields from the Voyager record are preserved
    
    Dim BibReturnCode As UpdateBibReturnCode
    Dim OclcBib As New Utf8MarcRecordClass
    Dim VgerBib As New Utf8MarcRecordClass
    
    Dim AddField As Boolean
    Dim F910 As String
    
    Dim OclcBlvl As String
    Dim VgerBlvl As String
    Dim OclcElvl As String
    Dim VgerElvl As String
    Dim Oclc00806 As String
    Dim Vger00806 As String
    Dim Oclc00834 As String
    Dim Vger00834 As String
    Dim OkToReplace As Boolean
    
    Set OclcBib = RecordIn.BibRecord
    
    Set VgerBib = GetVgerBibRecord(CStr(BibID))     'this method requires String, not Long
    
If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** VOYAGER RECORD (BEFORE OVERLAY) ***"
    WriteLog GL.Logfile, VgerBib.TextRaw
    WriteLog GL.Logfile, ""
    
'    WriteLog GL.Logfile, "*** OCLC RECORD ***"
'    WriteLog GL.Logfile, OclcBib.TextFormatted
'    WriteLog GL.Logfile, ""
End If
    
    OclcBlvl = OclcBib.GetLeaderValue(7, 1)
    VgerBlvl = VgerBib.GetLeaderValue(7, 1)
    OclcElvl = OclcBib.GetLeaderValue(17, 1)
    VgerElvl = VgerBib.GetLeaderValue(17, 1)
    ' 008/34, used for serials only
    Oclc00834 = OclcBib.Get008Value(34, 1)
    Vger00834 = VgerBib.Get008Value(34, 1)
    ' 2006-08-24: per vbross, some Voyager records have 008/34 = blank, which is invalid
    ' Treat those as 008/34 = 0 (zero)
    If Vger00834 = " " Then
        Vger00834 = "0"
    End If

    'Only replace if LDR/07 matches, and (for serials only) 008/34 also matches
    OkToReplace = False
    If OclcBlvl = VgerBlvl Then
        If OclcBlvl = "s" Then
            If Oclc00834 = Vger00834 Then
                OkToReplace = True
            Else
                WriteLog GL.Logfile, "Bib #" & BibID & " - 008/34 mismatch (Voyager, Marcive): " & _
                    TranslateBlank(Vger00834) & " - " & TranslateBlank(Oclc00834) & " - see review file"
            End If
        Else
            OkToReplace = True
        End If
    Else
        WriteLog GL.Logfile, "Bib #" & BibID & " - bib level mismatch (Voyager, Marcive): " & _
            TranslateBlank(VgerBlvl) & " - " & TranslateBlank(OclcBlvl) & " - see review file"
    End If
    
    'Reject incoming record if encoding level is lower than existing record's
    If (VgerElvl = " " Or VgerElvl = "I") Then
        If ElvlScore(VgerElvl) > ElvlScore(OclcElvl) Then
            WriteLog GL.Logfile, "Bib #" & BibID & " - ELvl mismatch (Voyager, Marcive): " & _
                TranslateBlank(VgerElvl) & " - " & TranslateBlank(OclcElvl) & " - see review file"
            OkToReplace = False
        End If
    End If
    
    'For serials only, compare 008/06 DtSt
    'MARCIVE: reject if fatal mismatch (Voyager=d, Marcive=c); allow but report if non-fatal
    Oclc00806 = OclcBib.Get008Value(6, 1)
    Vger00806 = VgerBib.Get008Value(6, 1)
    If OclcBlvl = "s" And Oclc00806 <> Vger00806 Then
        If Oclc00806 = "c" And Vger00806 = "d" Then
            WriteLog GL.Logfile, "Bib #" & BibID & " - FATAL DtSt mismatch (Voyager, Marcive): " & Vger00806 & " - " & Oclc00806 & " - see review file"
            OkToReplace = False
        Else
            WriteLog GL.Logfile, "Bib #" & BibID & " - nonfatal DtSt mismatch (Voyager, Marcive): " & Vger00806 & " - " & Oclc00806
            'Don't change OkToReplace - might be false from earlier checks
        End If
    End If
    
    'OK, so merge the records and update Voyager
    If OkToReplace Then
        With VgerBib
            .FldMoveTop
            Do While .FldMoveNext
                AddField = False
                Select Case .FldTag
                    '001
                    Case "001"
                        AddField = True
                    '035
                    Case "035"
                        If .SfdFindFirst("9") Then
                            AddField = True
                        End If
                        '20090615: added 035 with $z (OCoLC) to preserved fields
                        If .SfdFindFirst("a") = True Or .SfdFindFirst("z") = True Then
                            If InStr(1, .SfdText, "(OCoLC)", vbTextCompare) = 1 Then
                                AddField = True
                            End If
                        End If
                    '590
                    Case "590"
                        AddField = True
                    '655 (other 6xx handled separately)
                    Case "655"
                        AddField = True
                    '793
                    Case "793"
                        AddField = True
                    '856 (keep those with $xCDL or $xUCLA Law)
                    Case "856"
                        .SfdFindFirst "x"
                        Do While .SfdWasFound
                            If .SfdText = "CDL" Or .SfdText = "UCLA Law" Then
                                AddField = True
                            End If
                            .SfdFindNext
                        Loop
                    '910 will be handled separately from other 9XX
                    Case "910"
                        With OclcBib
                            'Append Voyager 910 to OCLC marcive 910, filtering out any existing "marcive" or "MARS" $a
                            .FldFindFirst "910"
                            VgerBib.SfdMoveTop
                            Do While VgerBib.SfdMoveNext
                                If (InStr(1, VgerBib.SfdText, "marcive", vbTextCompare) = 0) And _
                                   (InStr(1, VgerBib.SfdText, "MARS", vbBinaryCompare) = 0) Then
                                    .FldText = .FldText & .SfdMake(VgerBib.SfdCode, VgerBib.SfdText)
                                End If
                            Loop
                        End With
                    'the rest
                    Case Else
                        '6XX _2
                        If Left(.FldTag, 1) = "6" And .FldInd = " 2" Then
                            AddField = True
                        End If
                        '9XX, except 936 not kept for some reason
                        If Left(.FldTag, 1) = "9" And .FldTag <> "936" Then
                            AddField = True
                        End If
                        'XXX $5 CLU
                        If .SfdFindFirst("5") Then  '$5 is not repeatable so FindFirst is right
                            If .SfdText = "CLU" Then
                                AddField = True
                            End If
                        End If
                End Select
                If AddField Then
                    'Make sure we don't add duplicate field content
                    OclcBib.FldFindFirst .FldTag
                    Do While OclcBib.FldWasFound
                        'Compare normalized forms & reject duplicates
                        'For some reason, at this point VgerBib.NormString (or VgerBib.FldNorm) always returns empty string,
                        '   so must use OclcBib.NormString for both fields (can't use .FldNorm for both, obviously)
                        If OclcBib.NormString(OclcBib.FldText) = OclcBib.NormString(VgerBib.FldText) Then
                            AddField = False
                            Exit Do
                        End If
                        OclcBib.FldFindNext
                    Loop
                End If
                'Still OK to add the field?
                If AddField Then
                    OclcBib.FldAddGeneric .FldTag, .FldInd, .FldText, 3
                End If
    
            Loop 'FldMoveNext
        End With 'VgerBib
        
        If UBERLOGMODE Then
            WriteLog GL.Logfile, "*** COMBINED RECORD (AFTER OVERLAY) ***"
            WriteLog GL.Logfile, OclcBib.TextRaw
            WriteLog GL.Logfile, ""
        End If
        
        ' 20080416: Now that records are merged, add/update ucoclc 035 fields needed for WorldCat Local before updating Voyager
        Set OclcBib = UpdateUcoclc(OclcBib)
        
        If GL.ProductionMode Then
            With GL.BatchCat
                'Convert incoming record to Unicode
                OclcBib.CharacterSetOut = "U"

                BibReturnCode = .UpdateBibRecord(BibID, OclcBib.MarcRecordOut, GL.Vger.BibUpdateDateVB, GL.Vger.BibOwningLibraryNumber, GL.CatLocID, False)
                If BibReturnCode = ubSuccess Then
                    WriteLog GL.Logfile, "Updated Voyager bib#" & BibID
                Else
                    WriteLog GL.Logfile, ERROR_BAR
                    WriteLog GL.Logfile, "ERROR - ReplaceBibRecord failed with returncode: " & BibReturnCode
                    WriteLog GL.Logfile, OclcBib.TextRaw
                    WriteLog GL.Logfile, ERROR_BAR
                End If
            End With
        Else
            WriteLog GL.Logfile, "TESTMODE ONLY: Voyager bib#" & BibID & " NOT updated"
        End If 'ProductionMode
    Else 'not OK to replace
        'Calling routine (LoadRecords) will write review file based on ReplaceBibRecord's return code
    End If 'OkToReplace

    If OkToReplace = True And BibReturnCode = ubSuccess Then
        ReplaceBibRecord = True
    Else
        ReplaceBibRecord = False
    End If

End Function

Private Sub ReplaceHolRecord(OclcHolRecord As HoldingsRecordType, HolID As Long)
    'Replaces Voyager record# identified by HolID with OclcHolRecord.HolRecord
    'Voyager record is treated as "master" - only selected fields are replaced from the incoming record
    
    Dim HolReturnCode As UpdateHoldingReturnCode
    
    Dim OclcHol As New Utf8MarcRecordClass
    Dim VgerHol As New Utf8MarcRecordClass
    
    Dim F007 As String
    Dim F852_New As String
    Dim F852_Old As String
    Dim F901 As String
    Dim NewLoc As LocationType
    Dim OldLoc As LocationType
    Dim Suppress As Boolean
    Dim sfd As String
    Dim cnt As Integer
    
    Set OclcHol = OclcHolRecord.HolRecord
    Set VgerHol = GetVgerHolRecord(CStr(HolID))     'this method requires String, not Long

If UBERLOGMODE Then
    WriteLog GL.Logfile, ""
    WriteLog GL.Logfile, "*** VOYAGER HOLDINGS RECORD (BEFORE OVERLAY) ***"
    WriteLog GL.Logfile, VgerHol.TextRaw
    WriteLog GL.Logfile, ""
    
    WriteLog GL.Logfile, "*** INCOMING OCLC HOLDINGS RECORD ***"
    WriteLog GL.Logfile, OclcHol.TextRaw
    WriteLog GL.Logfile, ""
End If

    If OclcHol.FldFindFirst("007") Then
        F007 = OclcHol.FldText
    End If

    'Find elements in incoming record and use them to modify Voyager record
    With VgerHol
        'Update leader
        .ChangeLeaderValue 6, OclcHol.GetLeaderValue(6, 1)
        .ChangeLeaderValue 17, OclcHol.GetLeaderValue(17, 1)
        
        'Replace 007 (or add if none already)
        If .FldFindFirst("007") Then
            .FldText = F007
        Else
            If F007 <> "" Then
                .FldAddGeneric "007", "", F007, 3
            End If
        End If
        
        'Update 852 indicators & selected subfields
        .FldFindFirst "852"
        .FldInd = OclcHolRecord.CallNumInd
        
        'Remove subfields from Voyager record - they'll be replaced by incoming data
        .SfdMoveTop
        Do While .SfdMoveNext
            Select Case .SfdCode
                Case "b", "h", "i", "k", "m", "t"
                    .SfdDelete
            End Select
        Loop
        'Save what's left - we'll want this later
        F852_Old = .FldText
        With OclcHol
            .FldFindFirst "852"
            .SfdMoveTop
            'End up with: (Incoming subfields except $z)(existing subfields [after above deletion])(incoming $z)
            Do While .SfdMoveNext
                If .SfdCode = "b" Then
                    NewLoc = GetLoc(.SfdText)  'needed for BatchCat.UpdateHoldingRecord
                End If
                If .SfdCode <> "z" Then
                    F852_New = F852_New & .SfdMake(.SfdCode, .SfdText)
                Else
                    F852_Old = F852_Old & .SfdMake(.SfdCode, .SfdText)
                End If
            Loop
        End With
        'Now replace Voyager 852 with our new field
        .FldText = F852_New & F852_Old
        
        'Add 866, if none already (and only if a monograph)
        If OclcHolRecord.Summary <> "" Then
            'If incoming holdings LDR/06 = 'y' it's a serial
            If (.FldFindFirst("866") = False) And (OclcHolRecord.HolRecord.GetLeaderValue(6, 1) <> "y") Then
                .FldAddGeneric "866", " 0", OclcHolRecord.Summary, 3
            End If
        End If
        
    End With 'VgerHol

If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** COMBINED HOLDINGS RECORD (AFTER OVERLAY) ***"
    WriteLog GL.Logfile, VgerHol.TextRaw
    WriteLog GL.Logfile, ""
End If

    'Should the updated record be suppressed?
    OldLoc = GetLoc(GL.Vger.HoldLocationCode)
    If NewLoc.LocID = OldLoc.LocID Then
        'Keep current suppression value (could be manually suppressed record in non-suppressed loc)
        Suppress = GL.Vger.HoldRecordIsSuppressed
    Else
        'Go with whatever the new loc's rules are
        Suppress = NewLoc.Suppressed
    End If

If GL.ProductionMode Then
'********** QUICK HACK - CHECK THIS MORE CAREFULLY 02 Oct 2004
    If NewLoc.LocID > 0 Then
        With GL.BatchCat
            'Can't use Vger.HoldLocationID since 852 $b may have changed, causing mismatch
            HolReturnCode = .UpdateHoldingRecord _
                (HolID, VgerHol.MarcRecordOut, GL.Vger.HoldUpdateDateVB, GL.CatLocID, GL.Vger.HoldBibRecordNumber, NewLoc.LocID, Suppress)
            If HolReturnCode = uhSuccess Then
                WriteLog GL.Logfile, vbTab & "Updated Voyager hol#" & HolID
            Else
                WriteLog GL.Logfile, ERROR_BAR
                WriteLog GL.Logfile, "ERROR - ReplaceHolRecord failed with returncode: " & HolReturnCode
                WriteLog GL.Logfile, OclcHol.TextRaw
                WriteLog GL.Logfile, ERROR_BAR
            End If
        End With
    Else
        WriteLog GL.Logfile, "ERROR - Invalid location: " & NewLoc.Code & " - hol#" & HolID & " not updated"
    End If
Else
    WriteLog GL.Logfile, "TESTMODE ONLY: Voyager hol#" & HolID & " NOT updated"
End If

End Sub

Private Function RecordIsOnlineOnlySerial(ByRef OclcRecord As OclcRecordType) As Boolean
    'MARCIVE supposedly separates online-only serials from others, based on filenames
    'However, test file was inconsistent so double-check
    Dim AllOK As Boolean
    AllOK = True    'prove me wrong
    
    With OclcRecord.BibRecord
        'Serial?
        If .GetLeaderValue(7, 1) <> "s" Then
            AllOK = False
        End If
        
        'Has 245 $h[electronic resource]?
        .FldFindFirst "245"
        If .FldWasFound Then
            .SfdFindFirst "h"
            If .SfdWasFound Then
                If InStr(1, .SfdText, "[electronic resource]", vbTextCompare) = 0 Then
                    AllOK = False
                End If
            End If
        End If
        
        'Lacks 300?
        .FldFindFirst "300"
        If .FldWasFound Then
            AllOK = False
        End If
        
        'Lacks 776?
        .FldFindFirst "776"
        If .FldWasFound Then
            AllOK = False
        End If
    End With
    RecordIsOnlineOnlySerial = AllOK
End Function
