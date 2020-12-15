Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim rs As Integer
    
    Dim BibID As Long
    Dim BibRC As UpdateBibReturnCode
    Dim SourceFile As Utf8MarcFileClass
    Dim OclcRecord As Utf8MarcRecordClass
    Dim VgerRecord As Utf8MarcRecordClass
    Dim RawRecord As String
    Dim OclcNumber As String
    Dim OclcFile As Integer 'File handle
    Dim F948 As String

    Set SourceFile = New Utf8MarcFileClass
    SourceFile.OpenFile GL.InputFilename

    OclcFile = FreeFile
    Open GL.BaseFilename + ".worldcat" For Binary As OclcFile

    'Process each OCLC record in file
    Do While SourceFile.ReadNextRecord(RawRecord)
        DoEvents
        Set OclcRecord = New Utf8MarcRecordClass
        With OclcRecord
            .CharacterSetIn = "U"   'OCLC records
            .CharacterSetOut = "U"
            .MarcRecordIn = RawRecord

            'Get OCLC number from 001
            .FldFindFirst "001"
            OclcNumber = GetDigits(.FldText)
            
            'Get Voyager bib id from lookup table, based on OCLC number
            'Verfied 1-1 match already
            SQL = "select bib_id from vger_report.tmp_vbt_1690_oclc where oclc_number = " & OclcNumber
            rs = GL.GetRS
            With GL.Vger
                DoEvents
                .ExecuteSQL SQL, rs
                Do While .GetNextRow(rs)
                    BibID = .CurrentRow(rs, 1)
                Loop
            End With
            GL.FreeRS rs
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
        End With

        'Get Voyager record to be merged with this OCLC record
        Set VgerRecord = GetVgerBibRecord(CStr(BibID))
        
        'WriteLog GL.Logfile, vbCrLf & "************************************************************************"
        WriteLog GL.Logfile, "VERSION 2: Merging OCLC " & OclcNumber & " and Voyager bib " & BibID
        'WriteLog GL.Logfile, "VGER BEFORE: " & vbCrLf & VgerRecord.TextRaw
        'WriteLog GL.Logfile, ""
        'WriteLog GL.Logfile, "OCLC BEFORE: " & vbCrLf & OclcRecord.TextRaw
        'WriteLog GL.Logfile, ""
        'Set OclcRecord = MergeRecords(OclcRecord, VgerRecord)
        'WriteLog GL.Logfile, "OCLC AFTER: " & vbCrLf & OclcRecord.TextRaw
        
        'Fix some problems in the OCLC record from version 1
        With OclcRecord
            'Change 008/06 to s
            .Change008Value 6, "s"
            'Add 040 $b - assume 040 exists
            .FldFindFirst "040"
            .SfdFindFirst "b"
            If .SfdWasFound = False Then
                .SfdFindFirst "a"
                .SfdInsertAfter "b", "eng"
            End If
            'Fix black diamond Unicode character - and possibly other problematic data - in 520 field
            .FldFindFirst "520"
            If .FldWasFound Then
                .FldText = DeleteNonAscii(.FldText)
            End If
            
        End With 'OclcRecord
        'WriteLog GL.Logfile, "OCLC AFTER: " & vbCrLf & OclcRecord.TextRaw
        
        'Write the merged record to file, for updating later in WorldCat
        Put #OclcFile, , OclcRecord.MarcRecordOut
        
        'After saving the WorldCat version of the record, make some Voyager-specific changes and update Voyager
        With OclcRecord
            'Add a Voyager-only 793 field
            .FldAddGeneric "793", "  ", .SfdMake("a", "UCLA online theses and dissertations."), 3
            'Remove 035s, then copy all 035 from Voyager record
            .FldFindFirst "035"
            Do While .FldWasFound
                .FldDelete
                .FldFindNext
            Loop
            CopyField VgerRecord, OclcRecord, "035"
            'Replace the OCLC 001/003 with Voyager 001
            .FldFindFirst "001"
            .FldDelete
            .FldFindFirst "003"
            .FldDelete
            CopyField VgerRecord, OclcRecord, "001"
            'Add 910
            .FldAddGeneric "910", "  ", .SfdMake("a", "oclcetdupdate"), 3
            'Add 948 |a cmc |b oclcetdupdate |c [date] |d 2rev |k batchcat |h e
            F948 = _
                .SfdMake("a", "cmc") & _
                .SfdMake("b", "oclcetdupdate") & _
                .SfdMake("c", Format(Now(), "yyyymmdd")) & _
                .SfdMake("d", "2rev") & _
                .SfdMake("k", "batchcat") & _
                .SfdMake("h", "e")
            .FldAddGeneric "948", "  ", F948, 3
            
        End With 'OclcRecord, which will replace the Voyager record
        'WriteLog GL.Logfile, ""
        'WriteLog GL.Logfile, "VGER FINAL: " & vbCrLf & OclcRecord.TextRaw

        'Update the Voyager record, replacing it with the newly-merged data in OclcRecord
        BibRC = GL.BatchCat.UpdateBibRecord( _
            BibID, _
            OclcRecord.MarcRecordOut, _
            GL.Vger.BibUpdateDateVB, _
            GL.Vger.BibOwningLibraryNumber, _
            GL.CatLocID, _
            GL.Vger.BibRecordIsSuppressed _
            )
        If BibRC = ubSuccess Then
            WriteLog GL.Logfile, "Updated bib " & BibID
        Else
            WriteLog GL.Logfile, "ERROR updating bib " & BibID & " : return code " & BibRC
        End If

        NiceSleep GL.Interval
    Loop 'ReadNextRecord

    SourceFile.CloseFile
    Close OclcFile
    SkeletonForm.lblStatus.Caption = "Done!"
End Sub

Private Function MergeRecords(OclcRecord As Utf8MarcRecordClass, VgerRecord As Utf8MarcRecordClass) As Utf8MarcRecordClass
    'OCLC record is the (target) primary, with some data copied in from the (source) Voyager record.
    Dim AlwaysAddFields() As String
    Dim DeleteFields() As String
    Dim PreferOclcFields() As String
    Dim cnt As Integer
    Dim Tag As String
    Dim F506 As String
    
    With OclcRecord
        'LDR: Absolute changes
        .ChangeLeaderValue 6, "a"
        .ChangeLeaderValue 7, "m"
        .ChangeLeaderValue 17, "K"
        .ChangeLeaderValue 18, "i"
        
        '008: Absolute changes
        .Change008Value 6, "t"
        .Change008Value 24, "b"
        .Change008Value 25, "m"
        .Change008Value 29, "0"
        .Change008Value 30, "0"
        .Change008Value 31, "0"
        .Change008Value 33, "0"
        
        'Always add these Voyager fields to the OCLC record
        AlwaysAddFields = Split("264,500,655", ",")
        For cnt = 0 To UBound(AlwaysAddFields)
            Tag = AlwaysAddFields(cnt)
            CopyField VgerRecord, OclcRecord, Tag
        Next
        
        'Delete these fields from the OCLC record (some will be copied from Voyager later, or created whole, later)
        DeleteFields = Split("245,260,506,520", ",")
        For cnt = 0 To UBound(DeleteFields)
            Tag = DeleteFields(cnt)
            .FldFindFirst Tag
            Do While .FldWasFound
                .FldDelete
                .FldFindNext
            Loop
        Next
        
        'Use OCLC if present, else Voyager
        'Some of these were deleted from the OCLC record earlier
        PreferOclcFields = Split("006,007,100,245,260,300,336,337,338,502,520", ",")
        For cnt = 0 To UBound(PreferOclcFields)
            Tag = PreferOclcFields(cnt)
            If .FldFindFirst(Tag) = False Then
                CopyField VgerRecord, OclcRecord, Tag
            End If
        Next
        
        'Clean up some OCLC fields, whether original or copied from Voyager
        '* 100 $e
        If .FldFindFirst("100") Then
            If .SfdFindFirst("e") = False Then
                .SfdFindFirst "a"
                .SfdText = .SfdText & ","
                .SfdInsertAfter "e", "author."
            End If
        End If
        '* 245 $h, including basic punctuation handling
        If .FldFindFirst("245") Then
            If .SfdFindFirst("h") Then
                .SfdDelete
                .SfdFindFirst "a"
                If Right(.SfdText, 1) <> "/" Then
                    .SfdText = .SfdText & " /"
                End If
            End If
        End If
        '* 502 $g
        If .FldFindFirst("502") Then
            If .SfdFindFirst("g") = False Then
                .SfdAdd "g", "Thesis"
            End If
        End If
        '* 506: Removed earlier, now add a constant one
        F506 = _
            .SfdMake("a", "Open access") & _
            .SfdMake("f", "Unrestricted online access") & _
            .SfdMake("2", "star")
        .FldAddGeneric "506", "0 ", F506, 3
        '* 856 $7
        .FldFindFirst "856"
        Do While .FldWasFound
            If .FldInd2 = "0" Then
                'Loop over subfields, removing all except $u
                .SfdMoveTop
                Do While .SfdMoveNext
                    If .SfdCode <> "u" Then
                        .SfdDelete
                    End If
                Loop
                .SfdAdd "7", "0"
            End If
            If .FldInd2 = "2" Then
                .FldDelete
            End If
            .FldFindNext
        Loop
        
    End With
    Set MergeRecords = OclcRecord
End Function

Private Sub CopyField(Source As Utf8MarcRecordClass, ByRef Target As Utf8MarcRecordClass, Tag As String)
    'Copies all instances of Tag from Source record to Target record.
    'Modifies Target record passed by reference.
    With Source
        .FldFindFirst Tag
        Do While .FldWasFound
            Target.FldAddGeneric .FldTag, .FldInd, .FldText, 3
            .FldFindNext
        Loop
    End With 'Source
End Sub

Private Function DeleteNonAscii(ByRef Text As String)
    Dim i As Long
    Dim J As Long
    Dim Char As String

    i = 1
    For J = 1 To Len(Text)
        Char = Mid$(Text, J, 1)
        'WriteLog GL.Logfile, Char & " === " & AscW(Char) & " === " & (AscW(Char) And &HFFFF&)
        If (AscW(Char) And &HFFFF&) <= &H7F& Then
            Mid$(Text, i, 1) = Char
            i = i + 1
        Else
            Mid$(Text, i, 1) = " "
            i = i + 1
        End If
    Next
    Text = Left$(Text, i - 1)
    Text = Replace(Text, "   ", " ")
    DeleteNonAscii = Text
End Function

