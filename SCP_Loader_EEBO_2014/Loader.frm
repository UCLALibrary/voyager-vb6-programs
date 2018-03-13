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
    'GL.Init "-t ucla_testdb -f " & App.Path & "\eebo2.mrc"
    
    'EEBO: 129K records to load, will not do via this/SCP load but
    'create interleaved file for more efficient bulkimport via server.
    MakeInterleavedFile
    
    GL.CloseAll
    Set GL = Nothing
End Sub

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
            .FldFindFirst "956" 'not .FldFindNext since deleting/adding affects .FldPointer
        Loop
        
        'Delete fields
        'CDL doesn't usually include 003, but when they do, it's wrong, so remove it
        If IsSerial(Biblvl) = True Then
            DelFields() = Array("003", "049", "099", "510", "590", "655", "690", "9XX")
        Else
            DelFields() = Array("003", "049", "099", "655", "690", "9XX")
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
        
        'EEBO: Non-OCLC, so add prefix and preserve 035
        'EEBO: All 035s are clean, one $a, consistent format
        'EEBO: Make 035 $a look like: (SCP-eo)99854629eo
        .FldFindFirst "035"
        Do While .FldWasFound
            .SfdFindFirst "a"
            If .SfdWasFound Then
                .SfdText = "(SCP-eo)" & .SfdText
            End If
            .FldFindNext
        Loop
        
        'EEBO: Remove incoming 049
        .FldFindFirst "049"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
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

Private Sub MakeInterleavedFile()
    Const LIB_ID As Long = 1

    Dim OclcRecord As OclcRecordType
    Dim SourceFile As Utf8MarcFileClass
    Dim MarcRecord As Utf8MarcRecordClass
    Dim RawRecord As String
    Dim OutFilename As String
    
    OutFilename = GL.BaseFilename & ".interleaved.mrc"

    Set SourceFile = New Utf8MarcFileClass
    SourceFile.OpenFile GL.InputFilename

    Do While SourceFile.ReadNextRecord(RawRecord)
        Set MarcRecord = New Utf8MarcRecordClass
        With MarcRecord
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            .MarcRecordIn = RawRecord
        End With 'MarcRecord
        
        Set OclcRecord.BibRecord = MarcRecord
        PrepRecord OclcRecord
        BuildHoldings OclcRecord
        
        WriteRawRecord OutFilename, OclcRecord.BibRecord.MarcRecordOut
        WriteRawRecord OutFilename, OclcRecord.HoldingsRecords(1).HolRecord.MarcRecordOut
        
    Loop 'SourceFile
    
End Sub

