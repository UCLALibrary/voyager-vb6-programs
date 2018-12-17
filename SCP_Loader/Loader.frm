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

#Const DBUG = True
Private Const UBERLOGMODE As Boolean = False

Private Const ERROR_BAR As String = "*** ERROR ***"
Private Const RECORD_ARRAY_INCREMENT As Integer = 2000

'Form-level globals - let's keep this list short!
Private NewMonos() As OclcRecordType
Private NewSerials() As OclcRecordType
Private UpdMonos() As OclcRecordType
Private UpdSerials() As OclcRecordType
'

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
'    GL.Init "-t ucla_testdb -f " & App.Path & "\bad_880.mrc"
    
    ReDim NewMonos(1 To RECORD_ARRAY_INCREMENT) As OclcRecordType
    ReDim NewSerials(1 To RECORD_ARRAY_INCREMENT) As OclcRecordType
    ReDim UpdMonos(1 To RECORD_ARRAY_INCREMENT) As OclcRecordType
    ReDim UpdSerials(1 To RECORD_ARRAY_INCREMENT) As OclcRecordType
    
    KeepOrReject GL.InputFilename
    'Now process the "keep" records, in each of the arrays
    WriteLog GL.Logfile, "==================="
    WriteLog GL.Logfile, "Loading UPD serials"
    WriteLog GL.Logfile, "==================="
    LoadRecords UpdSerials

    WriteLog GL.Logfile, "==================="
    WriteLog GL.Logfile, "Loading NEW serials"
    WriteLog GL.Logfile, "==================="
    LoadRecords NewSerials

    WriteLog GL.Logfile, "================="
    WriteLog GL.Logfile, "Loading UPD monos"
    WriteLog GL.Logfile, "================="
    LoadRecords UpdMonos

    WriteLog GL.Logfile, "================="
    WriteLog GL.Logfile, "Loading NEW monos"
    WriteLog GL.Logfile, "================="
    LoadRecords NewMonos
    
    GL.CloseAll
    Set GL = Nothing
End Sub

Private Sub KeepOrReject(InputFilename As String)
    'Opens input file of bib MARC records, reads through & splits into:
    '   1) .del (records with 599 $aDEL)
    '   2) .reject (records with inappropriate bib level, no 599 $a, & other problems)
    '   3) Everything else goes into 1 of 4 arrays (serial NEW, serial UPD, mono NEW, mono UPD)

    Dim SourceFile As Utf8MarcFileClass
    Dim MarcRecord As Utf8MarcRecordClass
    Dim RawRecord As String
    Dim F001 As String
    Dim F599a As String
    Dim BibLevel As String
    Dim ScpRecord As OclcRecordType
    
    'Scads of separate files & counts for various problems
    Dim DelMonoFilename As String
    Dim DelSerialFilename As String
    Dim RejBlvlFilename As String
    Dim RejMonoFilename As String
    Dim RejSerialFilename As String
    
    Dim RecordCount As Long
    Dim DelMonoCount As Long
    Dim DelSerialCount As Long
    Dim NewMonoCount As Long
    Dim NewSerialCount As Long
    Dim RejBlvlCount As Long
    Dim RejMonoCount As Long
    Dim RejSerialCount As Long
    Dim UpdMonoCount As Long
    Dim UpdSerialCount As Long

'**** TEMP TO SUPPORT RELOAD OF REJECTS ****
'Dim F035 As String
'**** TEMP TO SUPPORT RELOAD OF REJECTS ****
    
    
    DelMonoFilename = GL.BaseFilename & ".mono.del"
    DelSerialFilename = GL.BaseFilename & ".serial.del"
    RejBlvlFilename = GL.BaseFilename & ".blvl.rej"
    RejMonoFilename = GL.BaseFilename & ".mono.rej"
    RejSerialFilename = GL.BaseFilename & ".serial.rej"
    
    RecordCount = 0
    DelMonoCount = 0
    DelSerialCount = 0
    NewMonoCount = 0
    NewSerialCount = 0
    RejBlvlCount = 0
    RejMonoCount = 0
    RejSerialCount = 0
    UpdMonoCount = 0
    UpdSerialCount = 0
    
    Set SourceFile = New Utf8MarcFileClass
    SourceFile.OpenFile InputFilename

    WriteLog GL.Logfile, ""
    WriteLog GL.Logfile, "Rejecting invalid & unwanted records..."
    
'Loops through all records in sourcefile
'Records can be rejected for several reasons, not mutually exclusive; using evil GoTo since no break/continue in VB
MainLoop:
    Do While SourceFile.ReadNextRecord(RawRecord)
        DoEvents
        RecordCount = RecordCount + 1
'Debug.Print RecordCount
        Set MarcRecord = New Utf8MarcRecordClass    'Treat each record as a separate instance
        'MarcRecordOut automatically changes LDR/22-23 to "00"
        'It also appears to trim leading/trailing space from non-control fields (010 & higher)
        With MarcRecord
            '2012-04-24: SCP now supplies Unicode records
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            .MarcRecordIn = RawRecord
            
            'Get 001 field, for reporting errors
            F001 = ""
            .FldFindFirst "001"
            If .FldWasFound Then
                F001 = .FldText
            End If
            
            'Reject records with bib level other than 'i', 'm' or 's'; use biblevel later for mono/serial distinctions
            BibLevel = .GetLeaderValue(7, 1)
            If BibLevel <> "i" And BibLevel <> "m" And BibLevel <> "s" Then
                WriteRawRecord RejBlvlFilename, RawRecord, RejBlvlCount
                WriteLog GL.Logfile, "Invalid bib level (" & BibLevel & "): " & F001
                GoTo MainLoop
            End If
            
            'Some quality issues... check 001 & 245
            If F001 = "" Then
                If IsSerial(BibLevel) Then
                    WriteRawRecord RejSerialFilename, RawRecord, RejSerialCount
                Else
                    WriteRawRecord RejMonoFilename, RawRecord, RejMonoCount
                End If
                WriteLog GL.Logfile, "No 001 field " & TranslateBibLevel(BibLevel) & " : "
                WriteLog GL.Logfile, .TextFormatted
                WriteLog GL.Logfile, ""
                GoTo MainLoop
            End If
            
            .FldFindFirst "245"
            If .FldWasFound = False Then
                If IsSerial(BibLevel) Then
                    WriteRawRecord RejSerialFilename, RawRecord, RejSerialCount
                Else
                    WriteRawRecord RejMonoFilename, RawRecord, RejMonoCount
                End If
                WriteLog GL.Logfile, "No 245 field " & TranslateBibLevel(BibLevel) & " : " & F001
                GoTo MainLoop
            End If
            
            F599a = Get599a(MarcRecord)
            If F599a = "" Then
                If IsSerial(BibLevel) Then
                    WriteRawRecord RejSerialFilename, RawRecord, RejSerialCount
                Else
                    WriteRawRecord RejMonoFilename, RawRecord, RejMonoCount
                End If
                WriteLog GL.Logfile, "No 599 $a " & TranslateBibLevel(BibLevel) & " : " & F001
                GoTo MainLoop
            End If

            If IsSerial(BibLevel) = False And IsBadEMono(MarcRecord) Then
                'Monos only
                WriteRawRecord RejMonoFilename, RawRecord, RejMonoCount
                WriteLog GL.Logfile, "Badly coded e-mono (check 245 $h / 338 $a and 599 $b) " & TranslateBibLevel(BibLevel) & " : " & F001
                GoTo MainLoop
            End If

            If LacksGood856(MarcRecord) Then
                If IsSerial(BibLevel) Then
                    WriteRawRecord RejSerialFilename, RawRecord, RejSerialCount
                Else
                    WriteRawRecord RejMonoFilename, RawRecord, RejMonoCount
                End If
                WriteLog GL.Logfile, "No usable 856 field " & TranslateBibLevel(BibLevel) & " : " & F001
                GoTo MainLoop
            End If

            'Record passed the tests, so we'll keep it
            With ScpRecord
                Set .BibRecord = MarcRecord
                ReDim .OclcNumbers(1 To 1)
                '.OclcNumbers(1) = F001
            End With
            
            '*** Consider separate subroutine to handle all these similar array assignments
            Select Case F599a
                Case "DEL"
                    If IsSerial(BibLevel) Then
                        WriteRawRecord DelSerialFilename, RawRecord, DelSerialCount
                    Else
                        WriteRawRecord DelMonoFilename, RawRecord, DelMonoCount
                    End If
                    WriteLog GL.Logfile, "DEL record per 599 $a " & TranslateBibLevel(BibLevel) & " : " & F001
                Case "NEW"
                    If IsSerial(BibLevel) Then
                        NewSerialCount = NewSerialCount + 1
                        If NewSerialCount > UBound(NewSerials) Then
                            GrowArray NewSerials, RECORD_ARRAY_INCREMENT
                        End If
                        NewSerials(NewSerialCount) = ScpRecord
                    Else
                        NewMonoCount = NewMonoCount + 1
                        If NewMonoCount > UBound(NewMonos) Then
                            GrowArray NewMonos, RECORD_ARRAY_INCREMENT
                        End If
                        NewMonos(NewMonoCount) = ScpRecord
                    End If
                Case "UPD"
                    If IsSerial(BibLevel) Then
                        UpdSerialCount = UpdSerialCount + 1
                        If UpdSerialCount > UBound(UpdSerials) Then
                            GrowArray UpdSerials, RECORD_ARRAY_INCREMENT
                        End If
                        UpdSerials(UpdSerialCount) = ScpRecord
                    Else
                        UpdMonoCount = UpdMonoCount + 1
                        If UpdMonoCount > UBound(UpdMonos) Then
                            GrowArray UpdMonos, RECORD_ARRAY_INCREMENT
                        End If
                        UpdMonos(UpdMonoCount) = ScpRecord
                    End If
                Case Else
                    'Not DEL/NEW/UPD, so reject it
                    If IsSerial(BibLevel) Then
                        WriteRawRecord RejSerialFilename, RawRecord, RejSerialCount
                    Else
                        WriteRawRecord RejMonoFilename, RawRecord, RejMonoCount
                    End If
                    WriteLog GL.Logfile, "Unknown 599 $a (" & F599a & ") " & TranslateBibLevel(BibLevel) & " : " & F001
                    GoTo MainLoop
            End Select
        End With 'MarcRecord
    Loop 'SourceFile

    'Clear out the unused space
    TruncateArray NewMonos, NewMonoCount
    TruncateArray NewSerials, NewSerialCount
    TruncateArray UpdMonos, UpdMonoCount
    TruncateArray UpdSerials, UpdSerialCount
    
    WriteLog GL.Logfile, ""
    WriteLog GL.Logfile, GetFileFromPath(InputFilename) & " split - processed " & RecordCount & " records..."
    WriteLog GL.Logfile, "Rejected: " & LeftPad(CStr(RejBlvlCount), " ", 5) & " bad biblevel"
    WriteLog GL.Logfile, "          " & LeftPad(CStr(RejMonoCount), " ", 5) & " rejected mono(s)"
    WriteLog GL.Logfile, "          " & LeftPad(CStr(RejSerialCount), " ", 5) & " rejected serial(s)"
    WriteLog GL.Logfile, "          " & LeftPad(CStr(DelMonoCount), " ", 5) & " DEL mono(s)"
    WriteLog GL.Logfile, "          " & LeftPad(CStr(DelSerialCount), " ", 5) & " DEL serial(s)"
    WriteLog GL.Logfile, ""
    WriteLog GL.Logfile, "Kept:     " & LeftPad(CStr(NewMonoCount), " ", 5) & " NEW mono(s)"
    WriteLog GL.Logfile, "          " & LeftPad(CStr(NewSerialCount), " ", 5) & " NEW serial(s)"
    WriteLog GL.Logfile, "          " & LeftPad(CStr(UpdMonoCount), " ", 5) & " UPD mono(s)"
    WriteLog GL.Logfile, "          " & LeftPad(CStr(UpdSerialCount), " ", 5) & " UPD serial(s)"
    
    SourceFile.CloseFile
End Sub

Private Function IsBadEMono(ByVal BibRecord As Utf8MarcRecordClass) As Boolean
    Dim Lacks245h As Boolean
    Dim Lacks338a As Boolean
    Dim Lacks599b As Boolean

    Lacks245h = True
    Lacks338a = True
    Lacks599b = True
    
    With BibRecord
        .FldFindFirst "245"
        .SfdFindFirst "h"
        If .SfdWasFound And InStr(1, .SfdText, "[electronic resource]", vbTextCompare) > 0 Then
            Lacks245h = False
        End If
        
        'Check for 338 $a "online resource" added 2013-06-12 for RDA, per Footprints ticket #33053
        '338 is repeatable, and $a is repeatable
        .FldFindFirst "338"
        Do While .FldWasFound
            .SfdFindFirst "a"
            Do While .SfdWasFound
                If InStr(1, .SfdText, "online resource", vbTextCompare) > 0 Then
                    Lacks338a = False
                End If
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
        
        .FldFindFirst "599"
        .SfdFindFirst "b"
        If .SfdWasFound And _
            (InStr(1, .SfdText, "cadocs", vbTextCompare) > 0 Or InStr(1, .SfdText, "db", vbTextCompare) > 0) _
        Then
            Lacks599b = False
        End If
    End With
    
    IsBadEMono = Lacks599b And (Lacks245h And Lacks338a)
End Function

Private Function LacksGood856(ByVal BibRecord As Utf8MarcRecordClass) As Boolean
    'Check 856 $3 and indicators for bad values
    Dim BibLevel As String
    
    With BibRecord
        BibLevel = .GetLeaderValue(7, 1)
        'Prove me wrong...
        LacksGood856 = True
        .FldFindFirst "856"
        Do While (.FldWasFound = True) And (LacksGood856 = True)
            'Check mono records for undesired descriptions in $3
            If IsSerial(BibLevel) = False Then
                .SfdFindFirst "3"
                '$3 is not repeatable so no need to check beyond the 1st
                If .SfdWasFound Then
                    If Is8563TextOK(.SfdText) Then
                        'Phrases not found so this URL is good
                        LacksGood856 = False
                    End If
                Else
                    'No $3, so no bad phrases, so this URL is good
                    LacksGood856 = False
                End If '$3
            Else
                'Serials - we don't care about $3
                LacksGood856 = False
            End If 'monos
        
            'Check indicator 2 for undesirable values - this trumps $3 and biblevel evaluation above, for this URL
            If .FldInd2 = "2" Then
                LacksGood856 = True
            End If
        
            .FldFindNext
        Loop '856
    End With 'BibRecord
    'Retval already set
End Function

Private Sub PrepRecord(ByRef ScpRecord As OclcRecordType)
    'Add/change/delete fields before loading
    '20071108: per slayne, added 65x -> 69x routine from daily OCLC loader
    
    Dim DelFields() As Variant
    Dim DelField As String
    Dim DelFieldCnt As Integer
    
    Dim Biblvl As String
    Dim LDR17 As String
    Dim F001 As String
    Dim F003 As String
    Dim F856 As String
    Dim F916 As String
    Dim F035 As String
    Dim F001_Digits As String
    Dim F001_Suffix As String
    Dim Ind As String
    Dim ind2 As String
    Dim OldInd As String
    Dim Text As String
    Dim FldPointer As Integer

'WriteLog GL.Logfile, "BEFORE:"
'WriteLog GL.Logfile, BibRecord.TextFormatted

    With ScpRecord.BibRecord
        'Get BibLvl (LDR/07): some actions vary based on this
        Biblvl = .GetLeaderValue(7, 1)
        
        '2010-10-25: move "mobile" 956 to 856 - must do before 9xx fields are deleted
        .FldFindFirst "956"
        Do While .FldWasFound
            OldInd = .FldInd
            F856 = .FldText
            .FldDelete
            .FldAddGeneric "856", OldInd, F856, 3
'            .FldFindFirst "856"
'            If .FldWasFound Then
'                .FldInsertBefore "856", OldInd, F856
'            Else
'                .FldAddGeneric "856", OldInd, F856, 3
'            End If
            .FldFindFirst "956" 'not .FldFindNext since deleting/adding affects .FldPointer
        Loop
        
        'Delete fields
        'CDL doesn't usually include 003, but when they do, it's wrong, so remove it
        'No longer delete 655 fields (regardless of 2nd indicator), per VBT-577 20160603
        If IsSerial(Biblvl) = True Then
            DelFields() = Array("003", "049", "099", "510", "590", "690", "9XX")
        Else
            DelFields() = Array("003", "049", "099", "690", "9XX")
        End If
        'Arrays created by Array() are always 0-based
        For DelFieldCnt = 0 To UBound(DelFields)
            DelField = DelFields(DelFieldCnt)
            .FldFindFirst DelField
            Do While .FldWasFound
                .FldDelete
                .FldFindNext
            Loop
        Next
        
        'Change LDR/17
        LDR17 = .GetLeaderValue(17, 1)
        Select Case LDR17
            Case "3"
                .ChangeLeaderValue 17, "K"
            Case "5"
                .ChangeLeaderValue 17, " "
        End Select
        
        'Change LDR/20-23
        .ChangeLeaderValue 20, "4500"
        
        '20070716: Remove 035 OCLC fields supplied by SCP (field starts with (OCoLC) )
        ' The 035 will be rebuilt more accurately using 001 and 019 below
        '20071130: Changed to remove all 035 fields except for those starting with $a (SFXObjID)
        '20090706: Changed to preserve (Sc-P) as well as (SFXObjID)
'20080414 special LION record reload only: do not remove OCLC numbers, commented out this block:
'20080506 special EAI record reload only: do not remove OCLC numbers, commented out this block:
        .FldFindFirst "035"
        Do While .FldWasFound
            'If neither (SFXObjID) nor (Sc-P) is found, delete the field
            If InStr(1, .FldText, "a(SFXObjID)", vbTextCompare) = 0 And InStr(1, .FldText, "a(Sc-P)", vbTextCompare) = 0 Then
                .FldDelete
            End If
            .FldFindNext
        Loop
        
        'Add 035; convert any 019 $a to 035 $z in same 035
        .FldFindFirst "001"
        If .FldWasFound Then
            'Records with no 001 were rejected already in KeepOrReject
            '2005-03-09: SCP sometimes allows 001 to have trailing spaces... trim them
            F001 = Trim(.FldText)
            .FldDelete
        End If

        F001_Digits = GetDigits(F001)
        F001_Suffix = GetTerminalNondigits(F001)
        If F001 = F001_Digits Then
            F003 = "(OCoLC)" & LeftPad(F001, "0", 8)
        Else
            If F001_Suffix = "" Then
                F003 = "(SCP)"
            Else
                F003 = "(SCP-" & F001_Suffix & ")"
            End If
            If F001_Suffix = "eo" Then
                F001 = LeftPad(F001_Digits, "0", 8) & F001_Suffix
            End If
            F003 = F003 & F001
        End If
        'Use this new 003 + 001 as the "OCLC" number for matching
        ScpRecord.OclcNumbers(1) = F003
        
        'Build the 035
        F035 = .SfdMake("a", F003)
        .FldFindFirst "019"
        If .FldWasFound Then
            .SfdFindFirst "a"
            Do While .SfdWasFound
                F035 = F035 & .SfdMake("z", LeftPad(.SfdText, "0", 8))
                .SfdFindNext
            Loop
        End If
        
        .FldAddGeneric "035", "  ", F035, 3
        
        'Add 049 to create internet holdings
        .FldAddGeneric "049", "  ", .SfdMake("a", "CLYY"), 3
        
        'Add $xCDL to each 856
        .FldFindFirst "856"
        Do While .FldWasFound
            .SfdAdd "x", "CDL"
            .FldFindNext
        Loop
        
        '2005-03-09: Added check for certain 856 data
        'Remove $xSCP UCSD
        'Replace "Restricted to your local campus" with "Restricted to UCLA" in $z
        .FldFindFirst "856"
        Do While .FldWasFound
            .SfdFindFirst "x"
            Do While .SfdWasFound
                If .SfdText = "SCP UCSD" Then
                    .SfdDelete
                End If
                .SfdFindNext
            Loop
            .SfdFindFirst "z"
            Do While .SfdWasFound
                .SfdText = Replace(.SfdText, "Restricted to your local campus", "Restricted to UCLA", 1, , vbTextCompare)
                .SfdFindNext
            Loop
            '20071130: Correct certain problems in the URL, if they exist
            .SfdFindFirst "u"
            Do While .SfdWasFound
                .SfdText = FixUrl(.SfdText)
                .SfdFindNext
            Loop
            .FldFindNext
        Loop
        
        'Add 910
        .FldAddGeneric "910", "  ", .SfdMake("a", "cdl " & DateToYYMMDD(Now())), 3
    
        'MONOS ONLY: Move SCP-supplied 590s to 916
        If IsSerial(Biblvl) = False Then
            .FldFindFirst "590"
            Do While .FldWasFound
                F916 = .FldText
                .FldDelete
                AddFieldInOrder "916", "  ", F916, ScpRecord.BibRecord
                .FldFindFirst "590" 'NOT FindNext, since FldDelete & FldAdd can interfere with fld pointers
            Loop
        End If
        
        'Add generic 590 for all bib levels
        AddFieldInOrder "590", "  ", .SfdMake("a", "UCLA Library - CDL shared resource."), ScpRecord.BibRecord
        
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
        
    End With 'SCP Bibrecord

'WriteLog GL.Logfile, "AFTER:"
'WriteLog GL.Logfile, BibRecord.TextFormatted

End Sub

Private Sub BuildHoldings(RecordIn As OclcRecordType)
    'Populates RecordIn.HoldingsRecords()
    '*** NOT the same as in UCLA_Loader
    Dim OclcHoldingsRecord As HoldingsRecordType
    Dim HolRecord As Utf8MarcRecordClass
    Dim ClInfo As ClInfoType

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

        ReDim .HoldingsRecords(1 To 1)
        .HoldingsRecordCount = 1
        OclcHoldingsRecord = .HoldingsRecords(.HoldingsRecordCount)
        With OclcHoldingsRecord
            .ClCode = "CLYY"
            ClInfo = GetClInfo(.ClCode)
            CallNum_H = ""
            CallNum_I = ""
            For cnt = 0 To UBound(ClInfo.CallNumberTags)
                With RecordIn.BibRecord
                    ClTag = ClInfo.CallNumberTags(cnt)
                    Select Case ClTag
                        Case "050"
                            'FIRST 050, per specs
                            .FldFindFirst ClTag
                            CallNumInd = "0 "
                        Case "090"
                            .FldFindFirst ClTag
                            CallNumInd = "0 "
                        Case Else
                            'CLYY allows 099, but don't want to consider that field for SCP records so look for completely nonexistent field
                            .FldFindFirst "000"
                    End Select
                    If .FldWasFound Then
                        .SfdFindFirst "a"
                        If .SfdWasFound Then
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

            .MatchLoc = ClInfo.DefaultLoc
            .NewLoc = .MatchLoc

            'If CLYY, we'll create an internet holdings record
            If .ClCode = "CLYY" Then
                InternetHoldings = True
            Else
                InternetHoldings = False
            End If
        End With 'OclcHoldingsRecord

        CreateNewHoldingsRecord OclcHoldingsRecord, BibLDR_07, F007, Bib008_06, InternetHoldings

        'Store the updated OclcHoldingsRecord back in RecordIn.HoldingsRecords()
        .HoldingsRecords(1) = OclcHoldingsRecord

'        If UBERLOGMODE Then
'            WriteLog GL.Logfile, ""
'            WriteLog GL.Logfile, "*** NEW HOLDINGS RECORD ***"
'            WriteLog GL.Logfile, OclcHoldingsRecord.HolRecord.TextFormatted
'            WriteLog GL.Logfile, ""
'        End If
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
        '2012-04-24: SCP now supplies Unicode records
        .CharacterSetIn = "U"
        .CharacterSetOut = "U"
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
    'Searches Voyager for primary OCLC number in RecordIn.OclcNumbers(1)
    'Modifies RecordIn: places all matching Voyager BibIDs in RecordIn.BibMatches, with total count in .BibMatchCount for convenience

    Dim SearchNumber As String
    Dim SQL As String
    Dim BibID As String
    Dim rs As Integer
    Dim cnt As Integer
    Dim AlreadyExists As Boolean

    rs = GL.GetRS

    RecordIn.BibMatchCount = 0
    ReDim RecordIn.BibMatches(1 To MAX_BIB_MATCHES)

    SearchNumber = Normalize0350(RecordIn.OclcNumbers(1))
    
    SQL = _
        "SELECT Bib_ID " & _
        "FROM Bib_Index " & _
        "WHERE Index_Code = '0350' " & _
        "AND Normal_Heading = '" & SearchNumber & "' " & _
        "ORDER BY Bib_ID"
    
    With GL.Vger
'Debug.Print "Searchnumber: " & SearchNumber
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
'Debug.Print "Found bibid: " & BibID
                End If
            End With
        Loop
    End With 'Vger

    With RecordIn
        If .BibMatchCount > 0 Then
            ReDim Preserve .BibMatches(1 To .BibMatchCount)
        End If
    End With

    GL.FreeRS rs
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
    Dim OclcRecord As OclcRecordType
    Dim MarcRecord As Utf8MarcRecordClass
    Dim OclcHolRecord As HoldingsRecordType
    Dim Message As String
    Dim HolCnt As Integer
    Dim HolLocCnt As Integer
    Dim HolMatch As Boolean
    Dim ReviewFilename As String
    Dim pos As Integer
    Dim cnt As Integer
    Dim rs As Integer
    Dim BibReplaced As Boolean

    rs = GL.GetRS

    For OclcCnt = GL.StartRec To UBound(OclcRecords)
        OclcRecord = OclcRecords(OclcCnt)
        PrepRecord OclcRecord
        BuildHoldings OclcRecord
        SearchDB OclcRecord
        With OclcRecord
            Message = "Record #" & OclcCnt & ": Incoming record OCLC# " & .OclcNumbers(1)
            Select Case .BibMatchCount
                Case 0
                    'Load new-to-file
                    BibID = AddBibRecord(OclcRecord, LIB_ID)
                    If BibID <> 0 Then
                        WriteLog GL.Logfile, Message & " : no match found"
                        WriteLog GL.Logfile, "Added Voyager bib#" & BibID
                        'Monos & serials: write log message if UPD record added as new to Voyager
                        If Get599a(.BibRecord) = "UPD" Then
                            WriteLog GL.Logfile, "*** This UPD loaded as new."
                        End If
                        'Monos only: write log message if 008/06 = 'm', for records added as new to Voyager
                        If IsMono00806m(.BibRecord) = True Then
                            WriteLog GL.Logfile, "*** DtSt m record loaded as new."
                        End If
                        'Now add holdings
                        For HolCnt = 1 To .HoldingsRecordCount
                            OclcHolRecord = .HoldingsRecords(HolCnt)
                            NewHolID = AddHolRecord(OclcHolRecord, BibID)
                        Next
                    End If 'BibID <> 0
                Case 1
                    'Attempt overlay of existing Voyager record
                    WriteLog GL.Logfile, Message & " : found " & .BibMatchCount & IIf(.BibMatchCount = 1, " match", " matches") & " on all OCLC numbers"
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
                                    End If 'MatchLoc
                                Next 'HolLocCnt

                                'No existing holdings, or none match on loc, so add
                                If HolMatch = False Then
                                    NewHolID = AddHolRecord(OclcHolRecord, BibID)
                                End If 'HolMatch false
                            End With 'OclcHolRecord
                        Next 'HolCnt
                    Else
                        'Bib not replaced
                        'Already wrote to review file in ReplaceBibRecord -nothing needed here?
'                        WriteLog GL.Logfile, "*** Warning: Bib# " & BibID & " not replaced - LDR or 008 mismatch: check review file"
'                        Put ReviewFile, , .BibRecord.MarcRecordOut
'                        Print #ReviewTextFile, "### Partial Match ###"
'                        Print #ReviewTextFile, .BibRecord.TextFormatted
'                        Print #ReviewTextFile, ""
                    End If 'BibReplaced
                Case Else 'not 0 or 1
                    WriteLog GL.Logfile, Message & " : found " & .BibMatchCount & " matches on all OCLC numbers: check review file"
                    If IsSerial(.BibRecord.GetLeaderValue(7, 1)) = False Then
                        ReviewFilename = GL.BaseFilename & ".mono.review"
                    Else
                        ReviewFilename = GL.BaseFilename & ".serial.review"
                    End If
                    WriteRawRecord ReviewFilename, .BibRecord.MarcRecordOut
            End Select
        End With 'OclcRecord
        'Blank line in log after each full record is handled
        WriteLog GL.Logfile, ""
        'Let some time go by so we don't flood the server
        Sleep GL.Interval
    Next
    GL.FreeRS rs
End Sub

Private Function AddBibRecord(RecordIn As OclcRecordType, LibraryID As Long) As Long
    'Adds new Voyager bib record; returns new record's ID
    Dim ReturnCode As AddBibReturnCode
    Dim BibID As Long

'*** This is NOT a safe way to create new reference from old: Unicode conversion can have errors,
'    especially with superscripts (& probably subscripts and greek symbols).
    'Dim OclcBib As New Utf8MarcRecordClass
    'Set OclcBib = RecordIn.BibRecord
    'OclcBib.CharacterSetOut = "U"

'*** This appears to be safe: works correctly for superscripts, needs more testing
    Dim OclcBib As Utf8MarcRecordClass
    Set OclcBib = New Utf8MarcRecordClass
    With OclcBib
        '2012-04-24: SCP now supplies Unicode records
        .CharacterSetIn = "U"
        .IgnoreSfdOrder = True
        .MarcRecordIn = RecordIn.BibRecord.MarcRecordOut
        .CharacterSetOut = "U"
    End With
 
'Test856 OclcBib

    ' 20080416: Add/update ucoclc 035 fields needed for WorldCat Local before updating Voyager
    Set OclcBib = UpdateUcoclc(OclcBib)

    BibID = 0
    ReturnCode = GL.BatchCat.AddBibRecord(OclcBib.MarcRecordOut, LibraryID, GL.CatLocID, False)

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
End Function

Private Function ReplaceBibRecord(RecordIn As OclcRecordType, BibID As Long) As Boolean
    'Replaces Voyager record# identified by BibID with the RecordIn.BibRecord
    'Incoming record is treated as the "master" - only selected fields from the Voyager record are preserved

    Dim BibReturnCode As UpdateBibReturnCode
    'Dim OclcBib As New Utf8MarcRecordClass
    Dim VgerBib As New Utf8MarcRecordClass
    Dim ReviewFilename As String
    
    Dim OclcBlvl As String
    Dim VgerBlvl As String
    Dim OclcElvl As String
    Dim VgerElvl As String
    Dim Oclc00806 As String
    Dim Vger00806 As String
    Dim Oclc00834 As String
    Dim Vger00834 As String
    Dim OkToReplace As Boolean
    Dim WriteReview As Boolean

    Dim AddField As Boolean

    'Set OclcBib = RecordIn.BibRecord
    'OclcBib.CharacterSetOut = "U"

'*** This appears to be safe: works correctly for superscripts, needs more testing
    Dim OclcBib As Utf8MarcRecordClass
    Set OclcBib = New Utf8MarcRecordClass
    With OclcBib
        '2012-04-24: SCP now supplies Unicode records
        .CharacterSetIn = "U"
        .IgnoreSfdOrder = True
        .CharacterSetOut = "U"
        .MarcRecordIn = RecordIn.BibRecord.MarcRecordOut
    End With

    Set VgerBib = GetVgerBibRecord(CStr(BibID))     'this method requires String, not Long

If UBERLOGMODE Then
    WriteLog GL.Logfile, "*** VOYAGER RECORD (BEFORE) ***"
    WriteLog GL.Logfile, VgerBib.TextFormatted
    WriteLog GL.Logfile, ""
End If

    OclcBlvl = OclcBib.GetLeaderValue(7, 1)
    VgerBlvl = VgerBib.GetLeaderValue(7, 1)
    OclcElvl = OclcBib.GetLeaderValue(17, 1)
    VgerElvl = VgerBib.GetLeaderValue(17, 1)
    Oclc00834 = OclcBib.Get008Value(34, 1)
    Vger00834 = VgerBib.Get008Value(34, 1)
    
    If IsSerial(OclcBlvl) Then
        ReviewFilename = GL.BaseFilename & ".serial.review"
    Else
        ReviewFilename = GL.BaseFilename & ".mono.review"
    End If
    
    WriteReview = False
    
    'Only replace if LDR/07 matches; for serials, reject incoming if Voyager 008/34 = 1
    OkToReplace = False
    If OclcBlvl = VgerBlvl Then
        If IsSerial(OclcBlvl) Then
            If Vger00834 <> "1" Then
                OkToReplace = True
            Else
                WriteLog GL.Logfile, "Bib #" & BibID & " - 008/34 = 1: " & _
                    "(SCP 00/834) = " & TranslateBlank(Oclc00834) & " - see review file"
            End If
        Else
            OkToReplace = True
        End If
    Else
        WriteLog GL.Logfile, "Bib #" & BibID & " - bib level mismatch (Voyager, SCP): " & _
            TranslateBlank(VgerBlvl) & " - " & TranslateBlank(OclcBlvl) & " - see review file"
    End If
    
    '2006-11-17: no longer rejecting elvl mismatches per vbross, but write log message and put copy of SCP record in review file
'    'Reject incoming record if encoding level is lower than existing record's
'    If (VgerElvl = " " Or VgerElvl = "I") Then
'        If ElvlScore(VgerElvl) > ElvlScore(OclcElvl) Then
'            WriteLog GL.Logfile, "Bib #" & BibID & " - ELvl mismatch (Voyager, SCP): " & _
'                TranslateBlank(VgerElvl) & " - " & TranslateBlank(OclcElvl) & " - see review file"
'            OkToReplace = False
'        End If
'    End If
    If VgerElvl <> OclcElvl Then
        WriteLog GL.Logfile, "Bib #" & BibID & " - ELvl mismatch (Voyager, SCP): " & _
            TranslateBlank(VgerElvl) & " - " & TranslateBlank(OclcElvl) & " - Voyager updated AND SCP record in review file"
        WriteReview = True
    End If
    
    '2006-11-17: no longer rejecting DtSt mismatches per vbross, but write log message and put copy of SCP record in review file
'    'Compare 008/06 DtSt (serials only) and reject if no match
'    Oclc00806 = OclcBib.Get008Value(6, 1)
'    Vger00806 = VgerBib.Get008Value(6, 1)
'    If OclcBlvl = "s" And Oclc00806 <> Vger00806 Then
'        WriteLog GL.Logfile, "Bib #" & BibID & " - DtSt mismatch (Voyager, SCP): " & Vger00806 & " - " & Oclc00806 & " - see review file"
'        OkToReplace = False
'    End If

    Oclc00806 = OclcBib.Get008Value(6, 1)
    Vger00806 = VgerBib.Get008Value(6, 1)
    If IsSerial(OclcBlvl) = True And Oclc00806 <> Vger00806 Then
        WriteLog GL.Logfile, "Bib #" & BibID & " - DtSt mismatch (Voyager, SCP): " & _
            TranslateBlank(Vger00806) & " - " & TranslateBlank(Oclc00806) & " - Voyager updated AND SCP record in review file"
        WriteReview = True
    End If

    'OK, so merge the records and update Voyager
    If OkToReplace Then
        'Check each Voyager field & set flag if it should be merged into the incoming record
        With VgerBib
            .FldMoveTop
            Do While .FldMoveNext
                AddField = False
                Select Case .FldTag
                    '001
                    Case "001"
                        AddField = True
                    '035
                    '20090615: formerly kept all *except* those with $a(OCoLC); now keeping all
                    Case "035"
                        AddField = True
                    '590
                    Case "590"
                        AddField = True
                    '650 _4
                    Case "650"
                        If .FldInd2 = "4" Then
                            AddField = True
                        End If
                    'No longer retain 655 fields (unless preserved below by $5 CLU), per VBT-577 20160603
                    '655
                    '856 (keep those with $xUCLA or $xUCLA Law)
                    Case "856"
                        .SfdFindFirst "x"
                        Do While .SfdWasFound
                            If .SfdText = "UCLA" Or .SfdText = "UCLA Law" Then
                                AddField = True
                            End If
                            .SfdFindNext
                        Loop
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
                        '6XX _2
                        If Left(.FldTag, 1) = "6" And .FldInd2 = "2" Then
                            AddField = True
                        End If
                        '9XX, except 936 for some reason
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
'Debug.Print "OCLC: " & vbTab & OclcBib.FldText & vbTab & "Norm: " & OclcBib.NormString(OclcBib.FldText)
'Debug.Print "VGER: " & vbTab & VgerBib.FldText & vbTab & "Norm: " & OclcBib.NormString(VgerBib.FldText)
                        'Compare normalized forms & reject duplicates
                        'For some reason, at this point VgerBib.NormString (or VgerBib.FldNorm) always returns empty string,
                        '   so must use OclcBib.NormString for both fields (can't use .FldNorm for both, obviously)
                        If OclcBib.NormString(OclcBib.FldText) = OclcBib.NormString(VgerBib.FldText) Then
'Debug.Print "Rejected dup field: ", OclcBib.FldText, VgerBib.FldText
                            AddField = False
                            Exit Do
                        End If
                        OclcBib.FldFindNext
                    Loop
                End If
                'Still OK to add the field?
                If AddField Then
                    AddFieldInOrder .FldTag, .FldInd, .FldText, OclcBib
                End If
            Loop 'FldMoveNext
        End With 'VgerBib

        If UBERLOGMODE Then
            WriteLog GL.Logfile, "*** COMBINED RECORD (AFTER) ***"
            WriteLog GL.Logfile, OclcBib.TextFormatted
            'WriteLog GL.Logfile, "********************"
            WriteLog GL.Logfile, ""
        End If
    
        ' 20080416: Now that records are merged, add/update ucoclc 035 fields needed for WorldCat Local before updating Voyager
        Set OclcBib = UpdateUcoclc(OclcBib)
        
        With GL.BatchCat
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
    End If
    
    If OkToReplace = False Or WriteReview = True Then
        'Not a complete match so write to review file
        'Write it as OCLC format, not Unicode, for consistency with records rejected elsewhere
        OclcBib.CharacterSetOut = "O"
        WriteRawRecord ReviewFilename, OclcBib.MarcRecordOut
    End If 'OK to replace
    
    If OkToReplace = True And BibReturnCode = ubSuccess Then
        ReplaceBibRecord = True
    Else
        ReplaceBibRecord = False
    End If

End Function

Private Sub ReplaceHolRecord(OclcHolRecord As HoldingsRecordType, HolID As Long)
    'Replaces Voyager record# identified by HolID with OclcHolRecord.HolRecord
    'Total replacement - nothing from the Voyager record is retained (except 001, and possibly call#)
    '22 Nov 2004 change: Voyager record is *not* replaced if it has a call# (852 $h)

    Dim HolReturnCode As UpdateHoldingReturnCode
    Dim OkToReplace As Boolean

    Dim OclcHol As New Utf8MarcRecordClass
    Dim VgerHol As New Utf8MarcRecordClass

    Set OclcHol = OclcHolRecord.HolRecord
    Set VgerHol = GetVgerHolRecord(CStr(HolID))     'this method requires String, not Long
    
    'Assume we'll replace the Voyager record
    OkToReplace = True
    
    With VgerHol
        'Reject Oclc MFHD if Voyager record has call#
        .FldFindFirst "852" 'always has 852 with at least $b
        .SfdFindFirst "h"
        If .SfdWasFound Then
            WriteLog GL.Logfile, vbTab & "Voyager hol#" & HolID & " not updated - record has call number"
            OkToReplace = False
        End If
        '20091118 Reject if Voyager record has 852 $c starting with 'lw' (Law Internet holdings)
        .SfdFindFirst "c"
        If .SfdWasFound Then
            If Left(.SfdText, 2) = "lw" Then
                WriteLog GL.Logfile, vbTab & "Voyager hol#" & HolID & " not updated - Law Internet holdings: " & .SfdText
                OkToReplace = False
            End If
        End If
    End With
    
'    'If Voyager record has call# but Oclc record doesn't, retain Voyager call#
'    If OclcHolRecord.CallNum_H = "" Then
'        With VgerHol
'            .FldFindFirst "852" 'always has 852 at this point
'            .SfdFindFirst "h"
'            If .SfdWasFound Then
'                With OclcHol
'                    .FldFindFirst "852"
'                    .SfdFindFirst "b"
'                    'put the $h after the $b
'                    .SfdInsertAfter "h", VgerHol.SfdText
'                    'Since OclcHol didn't have call#, and now does, must set 852 1st indicator to 0
'                    .FldInd1 = "0"
'                End With
'            End If
'        End With
'    End If
    
    'Add Voyager 001 & 004 to Oclc record
    OclcHol.FldAddGeneric "001", "", HolID, 3
    OclcHol.FldAddGeneric "004", "", GL.Vger.HoldBibRecordNumber, 3
    
    If OkToReplace = True Then
        With GL.BatchCat
            HolReturnCode = .UpdateHoldingRecord _
                (HolID, OclcHol.MarcRecordOut, GL.Vger.HoldUpdateDateVB, GL.CatLocID, GL.Vger.HoldBibRecordNumber, GL.Vger.HoldLocationID, False)
            If HolReturnCode = uhSuccess Then
                WriteLog GL.Logfile, vbTab & "Updated Voyager hol#" & HolID
            Else
                WriteLog GL.Logfile, ERROR_BAR
                WriteLog GL.Logfile, "ERROR - ReplaceHolRecord failed with returncode: " & HolReturnCode
                WriteLog GL.Logfile, OclcHol.TextRaw
                WriteLog GL.Logfile, ERROR_BAR
            End If
        End With 'Batchcat
    End If 'OK to replace
End Sub

Private Function Get599a(ByVal BibRecord As Utf8MarcRecordClass) As String
    '2016-05-06 akohler: Per VBT-538
    'Iterate through all 599 $a and use this hierarchy: DEL -> UPD -> NEW
    'Some SCP records have multiple 599 fields, causing problems.
    'Assume each 599 has only one $a, take the first.
    
    Dim F599a As String
    
    F599a = ""
    With BibRecord
        .FldFindFirst "599"
        Do While .FldWasFound
            .SfdFindFirst "a"
                If .SfdWasFound Then
                    Select Case .SfdText
                        'Take the first value in this hierarchy
                        'Override lower values if a higher one is found later.
                        Case "DEL"
                            'Always take DEL
                            F599a = "DEL"
                        Case "UPD"
                            'Change (or reset) only if not DEL
                            F599a = IIf(F599a <> "DEL", "UPD", F599a)
                        Case "NEW"
                            'Change only if empty
                            F599a = IIf(F599a = "", "NEW", F599a)
                        Case Else
                            'Unknown 599 $a, ignore.
                    End Select
                End If
            .FldFindNext
        Loop
    End With
    Get599a = F599a
End Function

Private Function IsMono00806m(ByVal BibRecord As Utf8MarcRecordClass) As Boolean
    Dim retval As Boolean
    retval = False
    If IsSerial(BibRecord.GetLeaderValue(7, 1)) = False And BibRecord.Get008Value(6, 1) = "m" Then
        retval = True
    End If
    IsMono00806m = retval
End Function

