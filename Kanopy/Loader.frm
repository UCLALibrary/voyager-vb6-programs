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
Option Explicit

#Const DBUG = True
Private Const UBERLOGMODE As Boolean = False

Private Const CATLOC As String = "lissystem"
Private Const ERROR_BAR As String = "*** ERROR ***"
Private Const MAX_RECORD_COUNT As Integer = 2000

'Form-level globals - let's keep this list short!
Private SpacMap As Collection       'Maps SpacCode to SpacText
Private CatLocID As Long            'Cataloging Location ID

'DEPENDENCIES TO REMEMBER FOR INSTALLATION:
'\windows\system32\voy*.dll
'\windows\system32\mswinsck.ocx
'.\spac_map.txt

Private Sub Form_Load()
    'Main handles everything
    Main
'   Test
    'Exit from VB if running as program
    End
End Sub

Private Sub Main()
    'This is the controlling procedure for this form
    Set GL = New Globals
    GL.Init Command
'GL.Init "-t ucla_testdb -f " & App.Path & "\foo.002.oclc"

    CatLocID = GetLoc(CATLOC).LocID
    ReDim OclcRecords(1 To MAX_RECORD_COUNT) As OclcRecordType  '2000 should be plenty
    
    BuildSpacMap
    
    KeepOrReject GL.InputFilename
DoEvents
    SplitByDatabase GL.BaseFilename & ".keep"
DoEvents
    GetLoadableRecords GL.BaseFilename & ".ucla", OclcRecords()
DoEvents
    LoadRecords OclcRecords()
    GL.CloseAll
    Set GL = Nothing
End Sub

Private Sub KeepOrReject(InputFilename As String)
    'Opens input file of MARC records, reads through & splits into 2 output files:
    '   1) .reject (unwanted records)
    '       LDR/22 [994 $a, starting 20060926] - this check removed 2016-03-15 akohler per VBT-528.
    '       049 $a certain values
    '       079 field (20080923)
    '   2) .keep (everything else)
    '
    '20060925: Per OCLC Technical Bulletin 253 (http://www.oclc.org/support/documentation/worldcat/tb/253/default.htm),
    '   as of Nov 12 2006 transaction codes from LDR/22 will be only in 994 $a; LDR/22 will always be '0'.
    '   994 $a is not documented here http://www.oclc.org/bibformats/en/9xx/default.shtm
    '   but values are listed here: http://www.oclc.org/support/documentation/worldcat/records/subscription/2/2.pdf
    '   Generally contains hexadecimal representation of the LDR/22 value: 01, 02 11, etc.
    '   but since some values are not valid hex ("X0", "Z0") or are exceptions ("E0" instead of the old "e"),
    '   handle all as strings from now on.
    '20160314 akohler: 994 now documented with values: http://www.oclc.org/bibformats/en/9xx/994.html
    '
    '20080923 akohler: Per Sara Layne, reject records with 079 fields, which are updates to RLIN-derived institutional records
    
    Dim SourceFile As New Utf8MarcFileClass
    Dim KeepFile As Integer     'File handle#
    Dim RejectFile As Integer   'File handle#
    Dim KeepCount As Integer
    Dim RejectCount As Integer
    
    Dim MarcRecord As New Utf8MarcRecordClass
    Dim RawRecord As String
'    Dim OclcTransactionCode As String
    Dim KeepRecord As Boolean
    
    Dim CharacterSet As String
    
    KeepFile = FreeFile
    Open GL.BaseFilename + ".keep" For Binary As KeepFile
    
    RejectFile = FreeFile
    Open GL.BaseFilename + ".reject" For Binary As RejectFile
    
    SourceFile.OpenFile InputFilename
    Do While SourceFile.ReadNextRecord(RawRecord)
        
        'Assume it's good, prove it's bad
        KeepRecord = True
        
        '20160815: Per VBT-662, records exported from OCLC might be structurally corrupt but this
        'code library doesn't catch them cleanly.
        '20170327: Try other techniques to find and log bad records.
        If Not RecordIsValid(RawRecord) Then
            WriteLog GL.Logfile, ERROR_BAR
            WriteLog GL.Logfile, "ERROR: RECORD " & SourceFile.RecordIndex & " IS INVALID MARC - SEE REJECT FILE AND COPY OF RECORD BELOW:"
            WriteLog GL.Logfile, RawRecord
            WriteLog GL.Logfile, ERROR_BAR
            KeepRecord = False
            
            'Bad record, so skip the rest of this iteration of the DO loop and jump down to the keep/reject code
            'No "continue" option in VB6 so must use GoTo
            GoTo UpdateFiles
        End If
        
        Set MarcRecord = New Utf8MarcRecordClass    'Treat each record as a separate instance
        'MarcRecordOut automatically changes LDR/22-23 to "00"
        'It also appears to trim leading/trailing space from non-control fields (010 & higher)
        With MarcRecord
            .MarcRecordIn = RawRecord
            'Records *should* be Unicode, but try to handle others
            'Make best guess about ambiguous values: X -> O (OCLC), Y -> M (MARC-8)
            CharacterSet = .CharacterizeCharacterSet
            Select Case CharacterSet
                Case "L", "M", "O", "U", "V"
                    'These values are all OK - no changes
                    
                Case "X"
                    CharacterSet = "O"  'OCLC
                Case "Y"
                    CharacterSet = "M"  'MARC-8
                Case Else
                    WriteLog GL.Logfile, ERROR_BAR
                    WriteLog GL.Logfile, "ERROR: RECORD " & SourceFile.RecordIndex & " HAS INVALID CHARACTER SET - SEE REJECT FILE AND COPY OF RECORD BELOW:"
                    WriteLog GL.Logfile, "Character set: " & CharacterSet
                    WriteLog GL.Logfile, RawRecord
                    WriteLog GL.Logfile, ERROR_BAR
                    KeepRecord = False
                    
                    'Bad record, so skip the rest of this iteration of the DO loop and jump down to the keep/reject code
                    'No "continue" option in VB6 so must use GoTo
                    GoTo UpdateFiles
                
            End Select
            .CharacterSetIn = CharacterSet
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            
            'Reject records with 049 $aUPDT or 049 $aCLU& (regardless of other content)
            .FldFindFirst "049"
            If InStr(1, .FldText, .MarcDelimiter & "aUPDT", vbTextCompare) > 0 _
                Or InStr(1, .FldText, .MarcDelimiter & "aCLU&", vbTextCompare) > 0 _
            Then
                KeepRecord = False
            End If
            
            'Reject records with any 079 field
            .FldFindFirst "079"
            If .FldWasFound Then
                KeepRecord = False
            End If
            
            'Reject records lacking certain OCLC transaction codes in 994 $a
            '2016-03-15 akohler: 994 check disabled per VBT-528
'            OclcTransactionCode = ""
'            .FldFindFirst "994"
'            If .FldWasFound Then
'                .SfdFindFirst "a"
'                If .SfdWasFound Then
'                    OclcTransactionCode = .SfdText
'                End If
'            End If
'
'            Select Case OclcTransactionCode
'                Case OCLC_PRODUCE, OCLC_PRODUCE_ALL, OCLC_UPDATE, OCLC_CONNEXION
'                    'Do nothing, we'll keep these (unless already rejected above)
'                Case Else
'                    'Otherwise, reject it
'                    KeepRecord = False
'            End Select
            
            'Put record in appropriate file and increment count
UpdateFiles:
            If KeepRecord = True Then
                Put #KeepFile, , RawRecord
                KeepCount = KeepCount + 1
            Else
                Put #RejectFile, , RawRecord
                RejectCount = RejectCount + 1
            End If
        End With 'MarcRecord
    Loop 'SourceFile
    
    WriteLog GL.Logfile, GetFileFromPath(InputFilename) & " split based on 049/079 and fatal errors: Kept " & KeepCount & ", rejected " & RejectCount
    
    SourceFile.CloseFile
    Close #KeepFile
    Close #RejectFile
End Sub

Private Sub SplitByDatabase(InputFilename As String)
    'Opens input file of MARC records, reads through & splits into 2 output files:
    '   1) .ethno (049 $aCLYM)
    '   2) .ucla (everything else)
        
    Dim SourceFile As New Utf8MarcFileClass
    Dim EthnoFile As Integer    'File handle#
    Dim UclaFile As Integer     'File handle#
    Dim EthnoCount As Integer
    Dim UclaCount As Integer
    
    Dim MarcRecord As New Utf8MarcRecordClass
    Dim RawRecord As String
    
    EthnoFile = FreeFile
    Open GL.BaseFilename + ".ethno" For Binary As EthnoFile
    
    UclaFile = FreeFile
    Open GL.BaseFilename + ".ucla" For Binary As UclaFile
    
    SourceFile.OpenFile InputFilename
    Do While SourceFile.ReadNextRecord(RawRecord)
        Set MarcRecord = New Utf8MarcRecordClass    'Treat each record as a separate instance
        With MarcRecord
            'All records should be Unicode at this point
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            .MarcRecordIn = RawRecord
            .FldFindFirst "049"
            If .FldWasFound Then
                If InStr(1, .FldText, "CLYM", vbBinaryCompare) > 0 Then
                    Put #EthnoFile, , RawRecord    'Write the record unchanged
                    EthnoCount = EthnoCount + 1
                Else
                    Put #UclaFile, , RawRecord    'Write the record unchanged
                    UclaCount = UclaCount + 1
                End If
            End If
        End With
    Loop
    
    WriteLog GL.Logfile, GetFileFromPath(InputFilename) & " split based on 049 CLYM: UCLA " & UclaCount & ", Ethno " & EthnoCount
    
    SourceFile.CloseFile
    Close #EthnoFile
    Close #UclaFile
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
    Dim f001 As String
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
            'All records should be Unicode at this point
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            .MarcRecordIn = RawRecord '2011-02-01: now doing MarcRecordIn *after* CharacterSetOut, apparently conversion problem no longer exists
            .FldFindFirst ("001")
            'OCLC's 001 fields have 1 trailing space, which we don't want; if not using GetDigits be sure to Trim()
            'Change for Kanopy records
            f001 = Trim(.FldText)
            
            'Write it to the text file
            WriteLog TxtFile, "*** Record number " & RecordsRead & " ***" '& .CharacterSetOut
            WriteLog TxtFile, .TextFormatted(latin1)
            WriteLog TxtFile, ""
        End With
        
        'Check array to see if we already have a record to load with this OCLC#
        'Replace earlier record with the current one, noting # of occurrences of this oclc#
        InternalDupFound = False
        For cnt = 1 To RecordsKept
            OclcRecord = OclcRecords(cnt)
            If OclcRecord.OclcNumbers(1) = f001 Then
                With OclcRecord
                    'Write the earlier bib record to dup file
                    Put DupFile, , RawRecord    'Write the record unchanged
                    'Replace the earlier with the new one
                    Set .BibRecord = MarcRecord
                    .OccurrenceCount = .OccurrenceCount + 1
                    InternalDupFound = True
                    WriteLog GL.Logfile, "Duplicate (OCLC# " & f001 & "): replaced record #" & .PositionInFile & _
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
                .OclcNumbers(1) = f001
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
    Dim BibMatchNum As Integer
    Dim BibID As Long
    Dim HolID As Long
    Dim NewHolID As Long
    Dim itemID As Long
    Dim NewItemID As Long
    Dim OclcRecord As OclcRecordType
    'Dim MarcRecord As Utf8MarcRecordClass 'UNUSED?
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
    
    rs = GL.GetRS
    
    ReviewFile = FreeFile
    Open GL.BaseFilename & ".review" For Binary As ReviewFile
    
    ReviewTextFile = FreeFile
    Open GL.BaseFilename & ".review.txt" For Output As ReviewTextFile
    
    For OclcCnt = GL.StartRec To UBound(OclcRecords)
        OclcRecord = OclcRecords(OclcCnt)
        PreprocessRecord OclcRecord
        Parse049 OclcRecord
        BuildHoldings OclcRecord
        SearchDB OclcRecord
        With OclcRecord
            Message = "Record #" & OclcCnt & ": Incoming record OCLC# " & .OclcNumbers(1)
            'Load new-to-file
            If .BibMatchCount = 0 Then
                BibID = AddBibRecord(OclcRecord, LIB_ID)
                If BibID <> 0 Then
                    WriteLog GL.Logfile, Message & " : no match found"
                    WriteLog GL.Logfile, "Added Voyager bib#" & BibID
                    'Log possible error for manual review
                    If .F949a <> "" Then
                        WriteLog GL.Logfile, "*** Warning: record with 949 (" & .F949a & ") added as new - failed overlay?"
                    End If
                    'Now add holdings
                    For HolCnt = 1 To .HoldingsRecordCount
                        OclcHolRecord = .HoldingsRecords(HolCnt)
                        NewHolID = AddHolRecord(OclcHolRecord, BibID)
                        If NewHolID <> 0 Then
                            With OclcHolRecord
                                For ItemCnt = 1 To .ItemCount
                                    'Item copy numbers must be numbers (as opposed to holdings 852 $t, which is a string)
                                    'For now, if OclcHolRecord.CopyNum is not purely numeric, take leading digit(s)
                                    'e.g., "2-4" becomes 2
                                    'Could improve later by parsing full CopyNum
                                    itemID = AddItemRecord(.Items(ItemCnt), GetCopyNumber(.CopyNum), NewHolID)
                                    If itemID > 0 And OclcRecord.NeedsInProcess = True Then
                                        AddItemStatus itemID, IN_PROCESS
                                    End If
                                Next
                            End With
                        End If 'NewHolID <> 0
                    Next
                End If 'BibID <> 0
            'Attempt overlay of existing Voyager record
            Else
                WriteLog GL.Logfile, Message & " : found " & .BibMatchCount & IIf(.BibMatchCount = 1, " match", " matches") & " on all OCLC numbers"
                If .F949a <> "" Then
                    For BibMatchNum = 1 To .BibMatchCount
                        BibID = .BibMatches(BibMatchNum)
                        If CStr(BibID) = .F949a Then    'Cast (long) BibID to string for comparison with 949 $a
                            'Writelog gl.logfile, vbTab & "Replacing bib# " & BibID & " with incoming record"
                            'ReplaceBibRecord writes success/failure messages to log
                            ReplaceBibRecord OclcRecord, BibID
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
                            'Compare each incoming HR to the existing set of Voyager HR, replacing if loc matches, else adding as new
                            For HolCnt = 1 To OclcRecord.HoldingsRecordCount
                                OclcHolRecord = OclcRecord.HoldingsRecords(HolCnt)
                                With OclcHolRecord
                                    HolMatch = False
                                    'Replace record if loc matches
                                    For cnt = 1 To HolLocCnt
                                        If .MatchLoc = ExistingHols(cnt).LocCode Then
                                            HolMatch = True
                                            HolID = ExistingHols(cnt).HolID
                                            ReplaceHolRecord OclcHolRecord, HolID
                                            For ItemCnt = 1 To .ItemCount
                                                'If BarcodeExists(.Items(ItemCnt).Barcode) Then
                                                If BarcodeMatchesHol(.Items(ItemCnt).Barcode, HolID) Then
                                                    ReplaceItemRecord .Items(ItemCnt), HolID
                                                Else
                                                    NewItemID = AddItemRecord(.Items(ItemCnt), GetCopyNumber(.CopyNum), HolID)
                                                    If NewItemID > 0 And OclcRecord.NeedsInProcess = True Then
                                                        AddItemStatus NewItemID, IN_PROCESS
                                                    End If
                                                End If
                                            Next 'ItemCnt
                                        End If 'MatchLoc
                                    Next 'HolLocCnt

                                    'No existing holdings, or none match on loc, so add
                                    If HolMatch = False Then
                                        NewHolID = AddHolRecord(OclcHolRecord, BibID)
                                        If NewHolID <> 0 Then
                                            For ItemCnt = 1 To .ItemCount
                                                NewItemID = AddItemRecord(.Items(ItemCnt), GetCopyNumber(.CopyNum), NewHolID)
                                                If NewItemID > 0 And OclcRecord.NeedsInProcess = True Then
                                                    AddItemStatus NewItemID, IN_PROCESS
                                                End If
                                            Next
                                        End If
                                    End If 'HolMatch false
                                End With 'OclcHolRecord
                            Next 'HolCnt
                        Else
                            WriteLog GL.Logfile, "*** Warning: Bib# " & BibID & " not replaced - incoming 949 $a does not match: check review file"
                            Put ReviewFile, , .BibRecord.MarcRecordOut
                            Print #ReviewTextFile, "### Partial Match ###"
                            Print #ReviewTextFile, .BibRecord.TextFormatted
                            Print #ReviewTextFile, ""
                        End If
                    Next
                Else 'no 949 $a
                    For BibMatchNum = 1 To .BibMatchCount
                        BibID = .BibMatches(BibMatchNum)
                        WriteLog GL.Logfile, "*** Warning: Bib# " & BibID & " not replaced - no 949 $a in incoming record: check review file"
                    Next
                    Put ReviewFile, , .BibRecord.MarcRecordOut
                    Print #ReviewTextFile, "### Partial Match ###"
                    Print #ReviewTextFile, .BibRecord.TextFormatted
                    Print #ReviewTextFile, ""
                End If
            End If 'BibMatchCount
        End With 'OclcRecord
        'Blank line in log after each full record is handled
        WriteLog GL.Logfile, ""
        'Let some time go by so we don't flood the server
        NiceSleep GL.Interval
    Next
    Close ReviewFile
    Close ReviewTextFile
    GL.FreeRS rs
End Sub

Private Sub PreprocessRecord(RecordIn As OclcRecordType)
    'Makes several changes to bib record:
    '- Create new 035 $a from 003 & 001: 035 $a(003)[001 digits only]; remove 001/003
    '- Move each 019 $a to a $z in the above 035
'    '- Add 655 to records with 856 (if this 655 doesn't already exist)
    '- Remove most 9xx fields
    '
    'While handling 001 & 019 OCLC numbers, add to OclcRecord's OclcNumbers()
    
    Dim ONLINE_655 As String ' can't use CONST, some data comes from function
    
    Dim Ind As String
    Dim ind2 As String
    Dim Text As String
    Dim FldPointer As Integer
    
    Dim f001 As String
    Dim f003 As String
    Dim f005 As String
    Dim f019a As String
    Dim f035 As String
    Dim OclcCnt As Integer
'    Dim Has655_Online As Boolean
'    Dim Has856 As Boolean
    Dim Add856x_UCLA As Boolean
    Dim MarcRecord As Utf8MarcRecordClass

    Set MarcRecord = RecordIn.BibRecord
    With MarcRecord
        If .FldFindFirst("001") Then
            f001 = .FldText
            'RecordIn.OtherOclcNumbers(OclcCnt) = f001
            .FldDelete
        End If
        If .FldFindFirst("003") Then
            f003 = .FldText
            .FldDelete
        End If
        
        f035 = .SfdMake("a", "(" & f003 & ")" & f001)
        
        'Convert 019 $a to 035 $z, using same 035 created from 001/003
        '019 is not repeatable, but $a is
        OclcCnt = 1     '001 added when deduping file in GetLoadableRecords
        If .FldFindFirst("019") Then
            .SfdFindFirst "a"
            Do While .SfdWasFound
                f019a = .SfdText
                f035 = f035 & .SfdMake("z", "(" & f003 & ")" & f019a)
                OclcCnt = OclcCnt + 1
                RecordIn.OclcNumbers(OclcCnt) = f019a
                .SfdFindNext
            Loop
            .FldDelete
        End If
        
        '2006-11-16: OCLC is now providing 035 fields with $a $z, but not 0-padded
        '   Remove OCLC-provided 035 since ours are better
        '2008-09-03: Naively assumed OCLC-provided 035 fields were (OCoLC), leading to infinite loop when encountered (CStRLIN);
        '   Now delete *all* OCLC-supplied 035 fields
        Do While .FldFindFirst("035")
            'If InStr(1, .FldText, "(OCoLC)", vbBinaryCompare) > 0 Then
                .FldDelete
            'End If
        Loop
        
        'Now add our new 035 field
        .FldAddGeneric "035", "  ", f035, 3
        
        ReDim Preserve RecordIn.OclcNumbers(1 To OclcCnt)
        
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
        
        .FldFindFirst ("856")
        Do While .FldWasFound
'            Has856 = True   'for use with 655 check later on
            'Add $xUCLA if 856 has neither $xUCLA nor $xCDL nor $xUCLA Law
            Add856x_UCLA = True
            .SfdFindFirst "x"
            Do While .SfdWasFound
                If .SfdText = "CDL" Or .SfdText = "UCLA" Or .SfdText = "UCLA Law" Then
                    Add856x_UCLA = False
                End If
                .SfdFindNext
            Loop
            If Add856x_UCLA Then
                .SfdAdd "x", "UCLA"
            End If
            .FldFindNext
        Loop '856
        
' Disabled the below block - we will no longer add 655 fields to records with 856 fields, per slayne 28 Jul 2004 ak
' Leaving code for now, just in case....
'        If Has856 Then
'            ONLINE_655 = .MarcDelimiter & "aOnline resources." & .MarcDelimiter & "2local"
'            Has655_Online = False
'            .FldFindFirst "655"
'            Do While .FldWasFound
'                If .FldInd = " 7" And .FldText = ONLINE_655 Then
'                    Has655_Online = True
'                End If
'                .FldFindNext
'            Loop '655
'            If Has655_Online = False Then
'                .FldAddGeneric "655", " 7", ONLINE_655, 3
'            End If
'        End If
        
        .FldFindFirst "9"
        Do While .FldWasFound
            Select Case .FldTag
                '2007-06-27: per ACCM, keep OCLC 948; will be used for cataloging workload statistics
                Case "910", "936", "948", "987"
                    'Do nothing - we're keeping these
                Case "949"
                    If .SfdFindFirst("a") Then
                        RecordIn.F949a = .SfdText
                    End If
                Case Else
                    .FldDelete
            End Select
            .FldFindNext
        Loop '9XX
        
        '20070627: get YYYYMMDD from 005 if present, else from OCLC filename
        .FldFindFirst "005"
        If .FldWasFound Then
            f005 = Mid(.FldText, 1, 8) 'YYYYMMDD
        Else
            f005 = Mid(Format(Date, "yyyy"), 1, 2) & Mid(GL.BaseFilename, 2, 6)  'CC + YYMMDD; assumes files are named in OCLC's normal way, with D followed by YYMMDD
        End If
        
        '20070627: add 948 $c with YYYYMMDD
        '20090921: check $a for 'pacq' and set RecordIn.NeedsInProcess if found
        RecordIn.NeedsInProcess = False
        .FldFindFirst "948"
        Do While .FldWasFound
            .SfdFindFirst "a"
            Do While .SfdWasFound
                If .SfdText = "pacq" Then
                    RecordIn.NeedsInProcess = True
                End If
                .SfdFindNext
            Loop
            'Remove any existing $c
            .SfdFindFirst "c"
            Do While .SfdWasFound
                .SfdDelete
                .SfdFindNext
            Loop
            'Add new $c in appropriate place
            .SfdMoveFirst
            Do While .SfdCode <> "" And .SfdCode < "c"
                .SfdMoveNext
            Loop
            If .SfdPointer >= 0 Then
                .SfdInsertBefore "c", f005
            Else
                .SfdMoveLast
                .SfdInsertAfter "c", f005
            End If
            .FldFindNext
        Loop
        
        '20130819: Temporary workaround for records with 6xx ind2=7 but no $2
        'These cause catsvr crash in Voyager 8.2.0; supposedly fixed in 8.2.1+
        'See Jira VBT-64
        'Does not affect ALL 6xx fields, but those it doesn't (653, 654, 658 and 662) are not valid with ind2=7
        'so this broad check is OK.
        .FldFindFirst "6"
        Do While .FldWasFound
            If .FldInd2 = "7" Then
                .SfdFindFirst "2"
                If .SfdWasFound = False Then
                    WriteLog GL.Logfile, "ERROR: REMOVING UNSUPPORTED 6XX FIELD WITH 2ND IND=7 BUT NO $2, IN OCLC# " & RecordIn.OclcNumbers(1)
                    WriteLog GL.Logfile, "You will need to fix this manually in Voyager."
                    WriteLog GL.Logfile, vbTab & .FldText
                    .FldDelete
                End If
            End If
            .FldFindNext
        Loop
        
    End With
    ' 20080415: Rebuild 035 ucoclc fields for WorldCat Local
    Set RecordIn.BibRecord = UpdateUcoclc(MarcRecord)
End Sub

Private Sub Parse049(RecordIn As OclcRecordType)
    'Parse bib records 049 field to create holdings record(s) and item(s)
    ReDim RecordIn.HoldingsRecords(1 To 10) As HoldingsRecordType   '10 should be plenty
    Dim HolRecord As HoldingsRecordType
    Dim HolRecordCnt As Integer
    
    Dim item As ItemRecordType
    Dim SpacCode As String
    Dim SpacText As String
    Dim AddSpac As Boolean
    Dim AddInternetHoldings As Boolean
    
    Dim cnt As Integer
    Dim fldptr As Integer
    Dim sfdptr As Integer
    
    ReDim Parsed049Chunks(1 To 3) As String
    
    Dim BibRecord As Utf8MarcRecordClass
    Set BibRecord = RecordIn.BibRecord
   
    With BibRecord
        'If 856 meets certain criteria, we'll need to make an Internet holdings record later
        'Assume all 856 fields fail to meet criteria, until we find otherwise
        AddInternetHoldings = False
        .FldFindFirst "856"
        Do While (.FldWasFound = True) And (AddInternetHoldings = False)
            'Check mono records for undesired descriptions in $3
'***** CONSIDER MAKING FUNCTION: IsSerial() *****
            If .GetLeaderValue(7, 1) <> "b" And .GetLeaderValue(7, 1) <> "s" Then
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
            Else
                'Serials - we don't care about $3
                AddInternetHoldings = True
            End If 'BibLevel <> b,s
            
            'Check indicator 2 for undesirable values - this trumps $3 and biblevel evaluation above
            If .FldInd2 = "2" Then
                AddInternetHoldings = False
            End If
            
            .FldFindNext
        Loop '856
        
        'Per 856 fields we should add internet holdings (CLYY)
        'If 049 doesn't already have one, add a barebones $aCLYY
        If AddInternetHoldings = True Then
            .FldFindFirst "049"
            If .FldWasFound Then
                .SfdFindFirst "a"
                Do While .SfdWasFound
                    If InStr(1, .SfdText, "CLYY", vbTextCompare) = 0 Then
                        .SfdAdd "a", "CLYY"
                    End If
                    .SfdFindNext
                Loop
            End If
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
                                If .ShelvingLocCount > 0 Then ReDim Preserve .ShelvingLocs(1 To .ShelvingLocCount)
                                If .SpacCount > 0 Then ReDim Preserve .SpacCodes(1 To .SpacCount)
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
                            ReDim .ShelvingLocs(1 To MAX_SHEVLOC_COUNT)
                            ReDim .SpacCodes(1 To MAX_SPAC_COUNT)
                            .CallNumPrefix = ""
                            .CallNumSuffix = ""
                            .ClCode = ""
                            .CopyNum = ""
                            .ItemCount = 0
                            .MatchLoc = ""
                            .NewLoc = ""
                            .NoteCount = 0
                            .SpacCount = 0
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
                    
                    'p: SPAC code
                    Case "p"
                        With HolRecord
                            SpacCode = UCase(BibRecord.SfdText)     'SPAC codes are always upper case
                            'Check against collection - have to trap error in case code not found
                            SpacText = GetSpacText(SpacCode)
                            If SpacText <> "" Then
                                .SpacCount = .SpacCount + 1
                                .SpacCodes(.SpacCount) = SpacCode
                                With BibRecord
                                    'Keep track of where we are in the record, so we can get back when done with the 901
                                    fldptr = .FldPointer
                                    sfdptr = .SfdPointer
                                    AddSpac = True
                                    .FldFindFirst "901"
                                    Do While .FldWasFound
                                        If .SfdFindFirst("a") Then
                                            If .SfdText = SpacCode Then
                                                'Incoming SPAC matches one already in the record
                                                AddSpac = False
                                            End If
                                        End If
                                        .FldFindNext
                                    Loop '901 .FldWasFound
                                    If AddSpac Then
                                        .FldAddGeneric "901", "  ", .SfdMake("a", SpacCode) & .SfdMake("b", SpacText), 3
                                    End If
                                    'Done with the 901, so back to the 049
                                    .FldPointer = fldptr
                                    .SfdPointer = sfdptr
                                End With 'BibRecord
                            Else
                                WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - Invalid SPAC code not added: " & SpacCode
                            End If
                            'Turn off this error trap
                            On Error GoTo 0
                        End With
                    
                    '20080918: Added support for 049 $q, which will create holdings 852 $c
                    'q: 852 $c
                    Case "q"
                        With HolRecord
                            .ShelvingLocCount = .ShelvingLocCount + 1
                            .ShelvingLocs(.ShelvingLocCount) = BibRecord.SfdText
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
                If .ShelvingLocCount > 0 Then ReDim Preserve .ShelvingLocs(1 To .ShelvingLocCount)
                If .SpacCount > 0 Then ReDim Preserve .SpacCodes(1 To .SpacCount)
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
    Dim ClInfo As ClInfoType
    Dim ValidHoldingsRecord As Boolean
    
    Dim InternetHoldings As Boolean
    Dim BibLDR_06 As String
    Dim BibLDR_07 As String
    Dim Bib008_06 As String     'for internet holdings records
    Dim CallNum_H As String
    Dim CallNum_I As String
    Dim CallNumInd As String
    Dim ClTag As String
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
                ClInfo = GetClInfo(.ClCode)
''''' THIS ClInfo test needs reworking - relevant parts now in Parse049()
                If ClInfo.DefaultLoc = "INVALID" Then
                    WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - BAD CL CODE: " & .ClCode
                    ValidHoldingsRecord = False
                ElseIf ClInfo.DefaultLoc = "SKIP" Then
                    ValidHoldingsRecord = False
                Else
                    CallNum_H = ""
                    CallNum_I = ""
                    For cnt = 0 To UBound(ClInfo.CallNumberTags)
                        With RecordIn.BibRecord
                            ClTag = ClInfo.CallNumberTags(cnt)
                            Select Case ClTag
                                Case "050"
                                    .FldFindLast ClTag
                                    CallNumInd = "0 "
                                Case "060"
                                    .FldFindLast ClTag
                                    CallNumInd = "2 "
                                Case "082" 'for UES
                                    .FldFindFirst ClTag
                                    CallNumInd = "1 "
                                Case "090"
                                    .FldFindFirst ClTag
                                    CallNumInd = "0 "
                                Case "092" 'for UES
                                    .FldFindFirst ClTag
                                    CallNumInd = "1 "
                                Case "096"
                                    .FldFindFirst ClTag
                                    CallNumInd = "2 "
                                Case "099"
                                    .FldFindFirst ClTag
                                    CallNumInd = "8 "
                            End Select
                            If .FldWasFound Then
                                'Debug.Print .FldTag, Replace(.FldText, .MarcDelimiter, " |")
                                .SfdFindFirst "a"
                                CallNum_H = .SfdText
                                .SfdFindFirst "b"
                                CallNum_I = .SfdText
                                'Debug.Print OclcHoldingsRecord.ClCode, CallNum_H, CallNum_I
                                'Stop looking - we have our match
                                Exit For
                            End If
                        End With 'RecordIn.BibRecord
                    Next 'cnt CallNumberTags
                    
                    'Store call# info in OclcHoldingsRecord
                    .CallNum_H = CallNum_H
                    .CallNum_I = CallNum_I
                    .CallNumInd = CallNumInd
' *** This block commented out 21 Jul 2004 - logic changed per slayne
'                    'If no loc code(s) specified via 049 $o, use default for this ClCode
'                    '049 $o abc [def] : abc = MatchLoc, def = NewLoc
'                    If .MatchLoc = "" Then
'                        'Can't set MatchLoc to default - should be used only when present
'                        '.MatchLoc = ClInfo.DefaultLoc
'                    End If
'                    'NewLoc must have a value - it goes into 852 $b of new holdings record
'                    If .NewLoc = "" Then
'                        If .MatchLoc = "" Then
'                            .NewLoc = ClInfo.DefaultLoc
'                        Else
'                            .NewLoc = .MatchLoc
'                        End If
'                    End If
' *** This block added 21 Jul 2004 - logic changed per slayne
                    '049 $o abc [def] : abc = MatchLoc, def = NewLoc
                    If .MatchLoc = "" Then
                        'Setting MatchLoc even if no $o allows update of holdings based on default loc from 049 $a
                        .MatchLoc = ClInfo.DefaultLoc
                    End If
                    'NewLoc must have a value - it goes into 852 $b of new holdings record
                    If .NewLoc = "" Then
                        .NewLoc = .MatchLoc
                    End If
' *** End of changes 21 Jul 2004

                    'If no item type code specified, use default based on bib LDR/06-07
                    For ItemCnt = 1 To .ItemCount
                        If .Items(ItemCnt).ItemCode = "" Then
                            .Items(ItemCnt).ItemCode = GetDefaultItemCode(BibLDR_06, BibLDR_07)
                        End If
                    Next
                    
                    'If CLYY, we'll create an internet holdings record
                    If .ClCode = "CLYY" Then
                        InternetHoldings = True
                    Else
                        InternetHoldings = False
                    End If
                End If 'ClInfo
            End With 'OclcHoldingsRecord
            
            CreateNewHoldingsRecord OclcHoldingsRecord, BibLDR_07, F007, Bib008_06, InternetHoldings
            
            'Store the updated OclcHoldingsRecord back in RecordIn.HoldingsRecords()
            .HoldingsRecords(HolCnt) = OclcHoldingsRecord

            If UBERLOGMODE Then
                WriteLog GL.Logfile, "*** NEW HOLDINGS RECORD ***"
                WriteLog GL.Logfile, OclcHoldingsRecord.HolRecord.TextFormatted
                WriteLog GL.Logfile, ""
            End If
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
    Dim SpacCode As String
    Dim SpacText As String
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
        'All records should be Unicode at this point
        .CharacterSetIn = "U"
        .CharacterSetOut = "U"
        .NewRecord HolLDR_06

'Kludge to allow conversion from Oclc to Unicode of .NewRecord - if .MarcRecordIn = "" conversion not possible 01 Aug 2004 ak
.MarcRecordIn = .MarcRecordOut

        Select Case HolLDR_06
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
        If InternetHoldings Then
            'LDR/06: for internet monos, reset this based on bib 008/06
            If .GetLeaderValue(6, 1) <> "y" Then
                If Bib008_06 = "s" Then
                    .ChangeLeaderValue 6, "x"
                Else
                    .ChangeLeaderValue 6, "v"
                End If
            End If
            'LDR/17
            .ChangeLeaderValue 17, "2"      'for all records
            '008
            If HolLDR_06 = "y" Then
                .Change008Value 6, "0"
                .Change008Value 8, "    "   '08-11: 4 blanks
                .Change008Value 22, "   "   '22-24: 3 blanks
            Else
                .Change008Value 16, "4"
                .Change008Value 20, "u"
                .Change008Value 22, "   "   '22-24: 3 blanks
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
        '20080918: Add 852 $c if provided
        If OclcHR.ShelvingLocCount > 0 Then
            For cnt = 1 To OclcHR.ShelvingLocCount
                F852 = F852 & .SfdMake("c", OclcHR.ShelvingLocs(cnt))
            Next
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
        
        If OclcHR.SpacCount > 0 Then
            For cnt = 1 To OclcHR.SpacCount
                'build F901
                SpacCode = OclcHR.SpacCodes(cnt)
                On Error Resume Next    'to trap error if item not found in collection
                'Error shouldn't happen since SpacCode was validated in Parse049()...
                SpacText = SpacMap.item(SpacCode)
                F901 = .SfdMake("a", SpacCode) & .SfdMake("b", SpacText)
                .FldAddGeneric "901", "  ", F901, 3
            Next
        End If
    End With
        
    Set OclcHR.HolRecord = HolRecord
End Sub

Private Sub SearchDB(RecordIn As OclcRecordType)
    'Searches Voyager for all OCLC numbers in RecordIn.OclcNumbers()
    'Modifies RecordIn: places all matching Voyager BibIDs in RecordIn.BibMatches, with total count in .BibMatchCount for convenience
    
    Dim SearchNumber As String
    Dim OclcCnt As Integer
    Dim BibID As String
    Dim rs As Integer
    Dim cnt As Integer
    Dim AlreadyExists As Boolean
    
    rs = GL.GetRS
    
    RecordIn.BibMatchCount = 0
    ReDim RecordIn.BibMatches(1 To MAX_BIB_MATCHES)
    
    For OclcCnt = 1 To UBound(RecordIn.OclcNumbers)
        With GL.Vger
            SearchNumber = RecordIn.OclcNumbers(OclcCnt)
            
            'Search both 035 $a and $z using existing method
            'Consider writing custom SQL to search both at once
            
            'Search Voyager 035 $a
            .SearchStandardNumber "B", "035", "a", SearchNumber, rs, False, False
            Do While True
                If Not .GetNextRow Then
                    Exit Do
                End If
                With RecordIn
                    BibID = GL.Vger.CurrentRow(1)
                    AlreadyExists = False
                    For cnt = 1 To .BibMatchCount
                        If BibID = .BibMatches(cnt) Then
                            AlreadyExists = True
                        End If
                    Next
                    If AlreadyExists = False Then
                        .BibMatchCount = .BibMatchCount + 1
                        .BibMatches(.BibMatchCount) = BibID
                    End If
                End With
            Loop
            
            ' Search Voyager 035 $z
            .SearchStandardNumber "B", "035", "z", SearchNumber, rs, False, False
            Do While True
                If Not .GetNextRow Then
                    Exit Do
                End If
                With RecordIn
'Debug.Print "Found match on 035 $z: " & SearchNumber
                    BibID = GL.Vger.CurrentRow(1)
                    AlreadyExists = False
                    For cnt = 1 To .BibMatchCount
                        If BibID = .BibMatches(cnt) Then
                            AlreadyExists = True
                        End If
                    Next
                    If AlreadyExists = False Then
                        .BibMatchCount = .BibMatchCount + 1
                        .BibMatches(.BibMatchCount) = BibID
                    End If
                End With
            Loop
        
        End With 'Vger
    Next
    
    With RecordIn
        If .BibMatchCount > 0 Then
            ReDim Preserve .BibMatches(1 To .BibMatchCount)
        End If
    End With
    
    GL.FreeRS rs
End Sub

Private Function AddBibRecord(RecordIn As OclcRecordType, LibraryID As Long) As Long
    'Adds new Voyager bib record; returns new record's ID
    Dim ReturnCode As AddBibReturnCode
    Dim BibID As Long
    
    Dim OclcBib As New Utf8MarcRecordClass
    Set OclcBib = RecordIn.BibRecord
    
    BibID = 0
    ReturnCode = GL.BatchCat.AddBibRecord(OclcBib.MarcRecordOut, LibraryID, CatLocID, False)
    
    If ReturnCode = abSuccess Then
        BibID = GL.BatchCat.RecordIDAdded
        'Log written in LoadRecords
        'Writelog gl.logfile, "OCLC#" & RecordIn.OclcNumbers(1) & " added as Voyager bib#" & BibID
    Else
        WriteLog GL.Logfile, ERROR_BAR
        WriteLog GL.Logfile, "ERROR - OCLC#" & RecordIn.OclcNumbers(1) & " - AddBibRecord failed with returncode: " & ReturnCode
        WriteLog GL.Logfile, OclcBib.TextFormatted
        WriteLog GL.Logfile, ERROR_BAR
    End If
    
    AddBibRecord = BibID
End Function

Private Function AddHolRecord(RecordIn As HoldingsRecordType, BibID As Long) As Long
    'Adds new Voyager holdings record, linked to BibID; returns new record's ID
    Dim ReturnCode As AddHoldingReturnCode
    Dim HolID As Long
    Dim HolLoc As LocationType
    Dim OclcHol As New Utf8MarcRecordClass
    Set OclcHol = RecordIn.HolRecord
    
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

    HolID = 0
    ReturnCode = GL.BatchCat.AddHoldingRecord(OclcHol.MarcRecordOut, BibID, CatLocID, HolLoc.Suppressed, HolLoc.LocID)
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
End Function

Private Function AddItemRecord(item As ItemRecordType, CopyNum As Integer, HolID As Long) As Long
    'Adds new Voyager item record, linked to HolID; returns new record's ID
    Dim itemID As Long
    Dim ItemReturnCode As AddItemReturnCode
    Dim BCReturnCode As AddItemBarCodeReturnCode
    
    With GL.BatchCat.cItem
        .AddItemToTop = False
        .CopyNumber = CopyNum
        .Enumeration = item.Enum
        .HoldingID = HolID
        .ItemTypeID = GetItemTypeID(item.ItemCode)
        .Sensitize = "Y"
        'If type isn't valid, use a default; control here instead of in GetItemTypeID
        If .ItemTypeID = 0 Then
            .ItemTypeID = GetItemTypeID("book")
            WriteLog GL.Logfile, ERROR_BAR
            WriteLog GL.Logfile, "ERROR - Barcode " & item.Barcode & " - Invalid item code [" & item.ItemCode & "] replaced by default (book)"
            WriteLog GL.Logfile, ERROR_BAR
        End If
        .PermLocationID = 0 'BatchCat will use holding record's location
    End With
    
'Debug.Print Item.ItemCode, GL.BatchCat.cItem.ItemTypeID
    
    itemID = 0
    ItemReturnCode = GL.BatchCat.AddItemData(CatLocID)
    If ItemReturnCode = aiSuccess Then
        itemID = GL.BatchCat.RecordIDAdded
        WriteLog GL.Logfile, vbTab & vbTab & "Added Voyager item#" & itemID
        'Now add the barcode (if any) to the new item
        If item.Barcode <> "" Then
            'BatchCat doesn't allow addition of duplicate barcodes
            If BarcodeExists(item.Barcode) Then
                WriteLog GL.Logfile, ERROR_BAR
                WriteLog GL.Logfile, "ERROR - Barcode " & item.Barcode & " already exists - cannot add duplicate barcode to item #" & itemID
                WriteLog GL.Logfile, ERROR_BAR
            Else
                BCReturnCode = GL.BatchCat.AddItemBarCode(itemID, item.Barcode)
                If BCReturnCode = aibSuccess Then
                    WriteLog GL.Logfile, vbTab & vbTab & vbTab & "Added Voyager barcode " & item.Barcode
                Else
                    WriteLog GL.Logfile, ERROR_BAR
                    WriteLog GL.Logfile, "ERROR - Barcode " & item.Barcode & " - AddItemRecord failed with returncode: " & BCReturnCode
                    WriteLog GL.Logfile, ERROR_BAR
                End If 'BCReturnCode
            End If 'Add barcode
        End If 'Barcode <> ""
    Else
        WriteLog GL.Logfile, ERROR_BAR
        WriteLog GL.Logfile, "ERROR - Barcode " & item.Barcode & " - AddItemRecord failed with returncode: " & BCReturnCode
        WriteLog GL.Logfile, ERROR_BAR
    End If 'Add item
    
    AddItemRecord = itemID
End Function

Private Sub AddItemStatus(itemID As Long, ItemStatus As Long)
    Dim ItemRC As AddItemStatusReturnCode
    ItemRC = GL.BatchCat.AddItemStatus(itemID, IN_PROCESS)
    'Return value could be aisNotAValidItemStatus if this status was already set
    If (ItemRC <> aisSuccess And ItemRC <> aisNotAValidItemStatus) Then
        WriteLog GL.Logfile, "*** ERROR: could not set IN PROCESS item status for item_id " & _
            itemID & " : " & TranslateAddItemStatusCode(ItemRC)
    End If
End Sub

Private Sub ReplaceBibRecord(RecordIn As OclcRecordType, BibID As Long)
    'Replaces Voyager record# identified by BibID with the RecordIn.BibRecord
    'Incoming record is treated as the "master" - only selected fields from the Voyager record are preserved
    
    Dim BibReturnCode As UpdateBibReturnCode
    Dim OclcBib As New Utf8MarcRecordClass
    Dim VgerBib As New Utf8MarcRecordClass
    Dim SpacCode As String
    
    Dim AddField As Boolean
    
    Set OclcBib = RecordIn.BibRecord
    
    Set VgerBib = GetVgerBibRecord(CStr(BibID))     'this method requires String, not Long
    
If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** VOYAGER RECORD ***"
    WriteLog GL.Logfile, VgerBib.TextFormatted
'Debug.Print VgerBib.TextRaw
    WriteLog GL.Logfile, ""
    
    WriteLog GL.Logfile, "*** OCLC RECORD ***"
    WriteLog GL.Logfile, OclcBib.TextFormatted
'Debug.Print ""
'Debug.Print OclcBib.TextRaw
    WriteLog GL.Logfile, ""
End If
    
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
                    If .SfdFindFirst("9") Then
                        AddField = True
                    End If
                    '2012-12-20: Also preserve 035 with $a starting with (SFXObjID)
                    .SfdFindFirst "a"   'Non-repeatable
                    If .SfdWasFound Then
                        If InStr(1, .SfdText, "(SFXObjID)", vbTextCompare) = 1 Then
                            AddField = True
                        End If
                    End If
                    '2012-12-20: Also preserve (OCoLC) (in $a or $z) 035 for continuing resources (LDR/07 = b/i/s)
                    If .GetLeaderValue(7, 1) = "b" Or .GetLeaderValue(7, 1) = "i" Or .GetLeaderValue(7, 1) = "s" Then
                        .SfdFindFirst "a"   'Non-repeatable
                        If .SfdWasFound Then
                            If InStr(1, .SfdText, "(OCoLC)", vbTextCompare) = 1 Then
                                AddField = True
                            End If
                        End If
                        .SfdFindFirst "z"   'Repeatable
                        Do While .SfdWasFound And AddField = False
                            If InStr(1, .SfdText, "(OCoLC)", vbTextCompare) = 1 Then
                                AddField = True
                            End If
                            .SfdFindNext
                        Loop
                    End If
                '590
                Case "590"
                    AddField = True
                '599 (added 2012-12-20)
                Case "599"
                    AddField = True
                '793 (added 2012-12-20)
                Case "793"
                    AddField = True
                '856
                Case "856"
                    'All continuing resources (LDR/07 = b/i/s)
                    If .GetLeaderValue(7, 1) = "b" Or .GetLeaderValue(7, 1) = "i" Or .GetLeaderValue(7, 1) = "s" Then
                        AddField = True
                    End If
                    'All formats, for Law or CDL
                    .SfdFindFirst "x"
                    Do While .SfdWasFound
                        If UCase(.SfdText) = "UCLA LAW" Or UCase(.SfdText) = "CDL" Then
                            AddField = True
                        End If
                        .SfdFindNext
                    Loop
                '901 will be handled separately from other 9XX
                Case "901"
                    'Get the existing Voyager 901 $a
                    .SfdFindFirst "a"
                    SpacCode = .SfdText
                    With OclcBib
                        'Compare this Voyager 901 $a to all OCLC 901 $a
                        'We'll add (preserve) the Voyager 901 $a, unless we find it in the OCLC record
                        AddField = True
                        .FldFindFirst "901"
                        Do While .FldWasFound
                            If .SfdFindFirst("a") Then
                                If .SfdText = SpacCode Then
                                    AddField = False
                                End If
                            End If
                            .FldFindNext
                        Loop
                    End With
                '910 will be handled separately from other 9XX
                Case "910"
                    With OclcBib
                        If .FldFindFirst("910") Then
                            'Append old Voyager 910 to incoming OCLC 910
                            .FldText = .FldText & VgerBib.FldText
                        Else
                            'no OCLC 910 for some reason, so add old Voyager 910
                            AddField = True
                        End If
                    End With
                'the rest
                Case Else
                    '6XX _4
                    If Left(.FldTag, 1) = "6" And .FldInd = " 4" Then
                        AddField = True
                    End If
                    '9XX (other than those handled above, and exceptions here)
                    '2012-12-20: 936 and 985 no longer retained
                    If Left(.FldTag, 1) = "9" And (.FldTag <> "936" And .FldTag <> "985") Then
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
                OclcBib.FldAddGeneric .FldTag, .FldInd, .FldText, 3
            End If
        Loop 'FldMoveNext
    End With 'VgerBib
    
If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** COMBINED RECORD ***"
    WriteLog GL.Logfile, OclcBib.TextFormatted
'Debug.Print ""
'Debug.Print OclcBib.TextRaw
    WriteLog GL.Logfile, ""
    WriteLog GL.Logfile, "********************"
End If

    With GL.BatchCat
        BibReturnCode = .UpdateBibRecord(BibID, OclcBib.MarcRecordOut, GL.Vger.BibUpdateDateVB, GL.Vger.BibOwningLibraryNumber, CatLocID, False)
        If BibReturnCode = ubSuccess Then
            WriteLog GL.Logfile, "Updated Voyager bib#" & BibID
        Else
            WriteLog GL.Logfile, ERROR_BAR
            WriteLog GL.Logfile, "ERROR - ReplaceBibRecord failed with returncode: " & BibReturnCode
            WriteLog GL.Logfile, OclcBib.TextRaw
            WriteLog GL.Logfile, ERROR_BAR
        End If
    End With

End Sub

Private Sub ReplaceHolRecord(OclcHolRecord As HoldingsRecordType, HolID As Long)
    'Replaces Voyager record# identified by HolID with OclcHolRecord.HolRecord
    'Voyager record is treated as "master" - only selected fields are replaced from the incoming record
    
    Dim HolReturnCode As UpdateHoldingReturnCode
    
    Dim OclcHol As New Utf8MarcRecordClass
    Dim VgerHol As New Utf8MarcRecordClass
    
    Dim F007 As String
    Dim F852_New As String
    Dim F852_Vger As String
    Dim F852c As String
    Dim F901 As String
    Dim AddSpac As Boolean
    Dim NewLoc As LocationType
    Dim OldLoc As LocationType
    Dim Suppress As Boolean
    Dim sfd As String
    Dim cnt As Integer
    
    Set OclcHol = OclcHolRecord.HolRecord
    Set VgerHol = GetVgerHolRecord(CStr(HolID))     'this method requires String, not Long

If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** VOYAGER RECORD ***"
    WriteLog GL.Logfile, VgerHol.TextFormatted
    WriteLog GL.Logfile, ""
    
    WriteLog GL.Logfile, "*** OCLC RECORD ***"
    WriteLog GL.Logfile, OclcHol.TextFormatted
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
        
        'Remove subfields from Voyager record - they'll be replaced or augmented (852 $c) by incoming data
        F852c = ""
        .SfdMoveTop
        Do While .SfdMoveNext
            Select Case .SfdCode
                Case "b", "h", "i", "k", "m", "t"
                    .SfdDelete
                Case "c"
                    F852c = F852c & .SfdMake(.SfdCode, .SfdText)
                    .SfdDelete
            End Select
        Loop
        'Save what's left - we'll want this later
        F852_Vger = .FldText
        
        'Loop thru OCLC 852, collecting subfields in chunks to reassemble later into one merged field
        F852_New = ""
        With OclcHol
            .FldFindFirst "852"
            .SfdMoveTop
            Do While .SfdMoveNext
                Select Case .SfdCode
                    Case "b"
                        'Capture OCLC $b for BatchCat.UpdateHoldingRecord
                        NewLoc = GetLoc(.SfdText)
                        F852_New = F852_New & .SfdMake(.SfdCode, .SfdText)
                        'Add Voyager 852 $c after OCLC $b
                        F852_New = F852_New & F852c
                    Case "z"
                        'Append OCLC $z after Voyager 852 remnants to add later
                        F852_Vger = F852_Vger & .SfdMake(.SfdCode, .SfdText)
                    Case Else
                        'Other OCLC subfields
                        F852_New = F852_New & .SfdMake(.SfdCode, .SfdText)
                End Select
            Loop
        End With 'OclcHol
        'Now replace Voyager 852 with our new field
        'Goal: (OCLC 852 $b) (Vger 852 $c) (OCLC 852 $c) (OCLC other 852 subfields except $z) (Vger other 852 subfields) (OCLC 852 $z)
        .FldText = F852_New & F852_Vger
        
        'Add 866, if none already (and only if a monograph)
        If OclcHolRecord.Summary <> "" Then
            'If incoming holdings LDR/06 = 'y' it's a serial
            If (.FldFindFirst("866") = False) And (OclcHolRecord.HolRecord.GetLeaderValue(6, 1) <> "y") Then
                .FldAddGeneric "866", " 0", OclcHolRecord.Summary, 3
            End If
        End If
        
        'Add SPAC info
        For cnt = 1 To OclcHolRecord.SpacCount
            AddSpac = True
            .FldFindFirst "901"
            Do While .FldWasFound
                If .SfdFindFirst("a") Then
                    If .SfdText = OclcHolRecord.SpacCodes(cnt) Then
                        'Incoming SPAC matches one already in the record
                        AddSpac = False
                    End If
                End If
                .FldFindNext
            Loop '901 .FldWasFound
            If AddSpac Then
                F901 = .SfdMake("a", OclcHolRecord.SpacCodes(cnt)) & .SfdMake("b", GetSpacText(OclcHolRecord.SpacCodes(cnt)))
                .FldAddGeneric "901", "  ", F901
            End If
        Next
    End With 'VgerHol

If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** COMBINED RECORD ***"
    WriteLog GL.Logfile, VgerHol.TextFormatted
    WriteLog GL.Logfile, ""
    WriteLog GL.Logfile, "********************"
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

'********** QUICK HACK - CHECK THIS MORE CAREFULLY 02 Oct 2004
    If NewLoc.LocID > 0 Then
        With GL.BatchCat
            'Can't use Vger.HoldLocationID since 852 $b may have changed, causing mismatch
            HolReturnCode = .UpdateHoldingRecord _
                (HolID, VgerHol.MarcRecordOut, GL.Vger.HoldUpdateDateVB, CatLocID, GL.Vger.HoldBibRecordNumber, NewLoc.LocID, Suppress)
            If HolReturnCode = uhSuccess Then
                WriteLog GL.Logfile, vbTab & "Updated Voyager hol#" & HolID
            
                'If holdings loc changed, item locs may need changing too
                If NewLoc.LocID <> OldLoc.LocID Then
                    UpdateItemLocs HolID, OldLoc.LocID, NewLoc.LocID
                End If
            
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
End Sub

Private Sub ReplaceItemRecord(item As ItemRecordType, HolID As Long)
    Dim itemID As Long
    Dim ItemReturnCode As AddItemReturnCode
    Dim rs As Integer
    rs = GL.GetRS
    
    With GL.Vger
        If item.Barcode <> "" Then
            .ItemBarcodeIsNumeric = False
            .SearchItemBarcode item.Barcode, rs
            Do While .GetNextRow
                itemID = CLng(.CurrentRow(rs, 1))
                .RetrieveItemRecord CStr(itemID)    'requires string
                'Confirm that this item's barcode is on this holdings record
                If .ItemHoldRecordNumber = HolID Then
                    .CopyItemObject GL.BatchCat.cItem
                    With GL.BatchCat
                        'Replace Enum with whatever cataloger just supplied
                        .cItem.Enumeration = item.Enum
                        'Only replace ItemCode if cataloger supplied new value
                        If item.ItemCode <> "" Then
                            .cItem.ItemTypeID = GetItemTypeID(item.ItemCode)
                        End If
                        'Need Permanent location
                        .cItem.PermLocationID = GetHolLocationID(HolID)
                        ItemReturnCode = .UpdateItemData(CatLocID)
                        If ItemReturnCode = aiSuccess Then
                            WriteLog GL.Logfile, vbTab & vbTab & "Updated Voyager item# " & itemID & ", barcode " & item.Barcode
                        Else
                            WriteLog GL.Logfile, ERROR_BAR
                            WriteLog GL.Logfile, "ERROR - Barcode " & item.Barcode & " - ReplaceItemRecord failed with returncode: " & ItemReturnCode
                            WriteLog GL.Logfile, ERROR_BAR
                        End If
                    End With 'BatchCat
                Else
                    'This item barcode is NOT on the provided HolID
                    WriteLog GL.Logfile, ERROR_BAR
                    WriteLog GL.Logfile, "ERROR - Barcode " & item.Barcode & " already exists on another MFHD (hol#" & .ItemHoldRecordNumber _
                        & ") - item not updated"
                    WriteLog GL.Logfile, ERROR_BAR
                End If 'HolID match
            Loop 'rs items
        End If 'Barcode <> ""
    End With 'Vger
    GL.FreeRS rs
End Sub

Private Sub UpdateItemLocs(HolID As Long, OldLocID As Long, NewLocID As Long)
    Dim ItemReturnCode As UpdateItemReturnCode
    Dim rs As Integer
    rs = GL.GetRS
    
    'Global VGER object already should have this holdings record, but get it if necessary
    If GL.Vger.HoldRecordNumber <> CStr(HolID) Then
        GL.Vger.RetrieveHoldRecord CStr(HolID)
    End If
    
    'Validation of locs would be nice....

    With GL.Vger
        .SearchItemNumbersForHold .HoldRecordNumber, rs
        Do While .GetNextRow(rs)
            .RetrieveItemRecord .CurrentRow(rs, 1)
            .CopyItemObject GL.BatchCat.cItem
            With GL.BatchCat
                'It is possible (and valid) for holdings to have items with different permanent locations
                'Change the loc ONLY for items which have OldLocID
                If .cItem.PermLocationID = OldLocID Then
                    .cItem.PermLocationID = NewLocID
                    ItemReturnCode = .UpdateItemData(CatLocID)
                    If ItemReturnCode <> uiSuccess Then
                        WriteLog GL.Logfile, ERROR_BAR
                        WriteLog GL.Logfile, "ERROR - UpdateItemLocs failed for item# " & .cItem.itemID & " with returncode: " & ItemReturnCode
                        WriteLog GL.Logfile, ERROR_BAR
                    End If
                End If
            End With
        Loop
    End With
    GL.FreeRS rs
End Sub

Private Sub BuildSpacMap()
    Dim SpacFile As Integer
    Dim Line As String
    Dim SpacCode As String
    Dim SpacText As String
    Dim TabPos As Integer
    
    Set SpacMap = New Collection
    
    SpacFile = FreeFile
    
    Open App.Path & "\spac_map.txt" For Input As SpacFile
    
    Do While Not EOF(SpacFile)
        Line Input #SpacFile, Line      'requires # symbol
        '2006-07-26: changed spac file format, now not position based: just tab-delimited
        TabPos = InStr(1, Line, Chr(9))
        If TabPos > 0 Then
            SpacCode = Left(Line, TabPos - 1)
            SpacText = Right(Line, Len(Line) - TabPos)
            SpacMap.Add SpacText, SpacCode
        End If
    Loop
    Close SpacFile
End Sub

Private Function GetSpacText(SpacCode As String) As String
    Dim SpacText As String
    
    'Check against collection - have to trap error in case code not found
    SpacText = ""   'otherwise could have old value if error occurs
    On Error Resume Next
    SpacText = SpacMap.item(UCase(SpacCode))    'SPAC codes are always upper case
    GetSpacText = SpacText
End Function

Private Sub Test()
    Set GL = New Globals
    GL.Init Command
   
'    OpenVger "UCLA_TESTDB"
'    Dim Loc As LocationType
'    Loc = GetLoc("arbtsra")
'    With Loc
'        Debug.Print .Code
'        Debug.Print .LocID
'        Debug.Print .OpacDisplay
'        Debug.Print .SpineLabel
'        Debug.Print .StaffName
'        Debug.Print .Suppressed
'    End With
'    CloseVger
'
'    ReDim Foo(1 To 1000) As HoldingsLocType

End Sub

Private Function RecordIsValid(RawRecord As String) As Boolean
    Dim Leader As String
    Dim Directory As String
    Dim allegedBaseAddress As Long
    Dim realBaseAddress As Long
    Dim allegedSize As Long
    Dim realSize As Long
    Dim minSize As Long
    
    Dim valid As Boolean
    
    'Assume it's good, prove otherwise
    valid = True
    
    minSize = 40 'semi-arbitrary minimum size of marc record
    
    realSize = Len(RawRecord)
    If realSize > minSize Then
        allegedSize = Mid(RawRecord, 1, 5)  'LDR/00-04
        allegedBaseAddress = Mid(RawRecord, 13, 5) 'LDR/12-16
    Else
        allegedSize = 0
        allegedBaseAddress = 0
        valid = False
    End If
    
    'InStr is 1-based, MARC is 0-based, so don't add 1 to get real base address
    realBaseAddress = InStr(1, RawRecord, Chr(30), vbBinaryCompare) 'position of first field separator, which should follow directory
    
    If realBaseAddress < realSize Then
        Leader = Mid(RawRecord, 1, 24)
        Directory = Mid(RawRecord, 25, realBaseAddress - 25) 'Start after leader, with length of real base address - leader
    End If
    
    If Len(Directory) Mod 12 <> 0 Then
        valid = False
    End If
    
    If valid = False Then
        WriteLog GL.Logfile, "Length of record (alleged): " & allegedSize
        WriteLog GL.Logfile, "Length of record (real): " & realSize
        WriteLog GL.Logfile, "Base address (alleged): " & allegedBaseAddress
        WriteLog GL.Logfile, "Base address (real): " & realBaseAddress
        WriteLog GL.Logfile, "Leader: " & Leader
        WriteLog GL.Logfile, "Directory: " & Directory
        WriteLog GL.Logfile, "Directory size: " & Len(Directory)
    End If
    
    RecordIsValid = valid
End Function
