Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm

    Dim BibID As String
    Dim RecordID As String
    Dim ISSN As String
    Dim MatchFound As Boolean
    
    Dim SourceFile As Utf8MarcFileClass
    Dim BibRecord As Utf8MarcRecordClass
    Dim RawRecord As String
    
    Dim Match035FileName As String
    Dim MatchISSNFileName As String
    Dim NoMatchSerialFileName As String
    Dim LoadableRecordsFileName As String

    Match035FileName = GL.BaseFilename & ".match_035.mrc"
    MatchISSNFileName = GL.BaseFilename & ".match_issn.mrc"
    NoMatchSerialFileName = GL.BaseFilename & ".nomatch_serial.mrc"
    LoadableRecordsFileName = GL.BaseFilename & ".loadable.mrc"
    
    Set SourceFile = New Utf8MarcFileClass
    SourceFile.OpenFile GL.InputFilename

    Do While SourceFile.ReadNextRecord(RawRecord)
        'Fake inner loop, just so we can exit to real outer loop as needed
        Do While 1 = 1
            Set BibRecord = New Utf8MarcRecordClass
            With BibRecord
                'convert to Unicode for Voyager
                .CharacterSetIn = "O"   'OCLC records
                .CharacterSetOut = "U"
                .MarcRecordIn = RawRecord
    
                RecordID = ""
                
                .FldFindFirst "001"
                If .FldWasFound Then
                    RecordID = Replace(.FldText, "wln", "ccn")
                Else
                    WriteLog GL.Logfile, "ERROR: no 001 field"
                    WriteLog GL.Logfile, .TextRaw
                    Exit Do
                End If
                
                SkeletonForm.lblStatus.Caption = "Processing " & RecordID
                DoEvents
                
                'Filter out dups based on CCN number (incoming 001, Voyager 035)
                BibID = SearchDBFor035(RecordID)
                'Found a record based on CCN: log it and skip to next record
                If BibID <> "" Then
                    WriteLog GL.Logfile, RecordID & " 035 matches Voyager bib " & BibID & " - rejected"
                    WriteRawRecord Match035FileName, RawRecord
                    Exit Do
                End If
                
                'Filter out dups based on ISSN
                MatchFound = False
                BibID = ""
                .FldFindFirst "022"
                Do While .FldWasFound = True And MatchFound = False
                    .SfdMoveFirst
                    Do While .SfdWasFound = True And MatchFound = False
                        If (.SfdCode = "a" Or .SfdCode = "y") Then
                            ISSN = NormalizeISSN(.SfdText)
                            'WriteLog GL.Logfile, vbTab & ISSN & " *** " & .SfdText
                            BibID = SearchDBForISSN(ISSN)
                            If BibID <> "" Then
                                MatchFound = True
                            End If
                        End If
                        .SfdMoveNext
                    Loop
                    .FldFindNext
                Loop
                'Found a record based on ISSN: log it and skip to next record
                If MatchFound = True Then
                    WriteLog GL.Logfile, RecordID & " ISSN " & ISSN & " matches Voyager bib " & BibID & " - rejected"
                    WriteRawRecord MatchISSNFileName, RawRecord
                    Exit Do
                End If

                'Filter out any remaining serials (LDR/07 = 's', only)
                'Log and skip to next record
                If .GetLeaderValue(7, 1) = "s" Then
                    WriteLog GL.Logfile, RecordID & " is serial - rejected"
                    WriteRawRecord NoMatchSerialFileName, RawRecord
                    Exit Do
                End If
                
                'Whatever's left is loadable
                WriteLog GL.Logfile, RecordID & " is loadable - kept"
                WriteRawRecord LoadableRecordsFileName, RawRecord
                
                'Bail out of fake inner loop to go on to next record
                Exit Do
    
            End With
    
        Loop 'Fake inner loop
    Loop 'Real outer loop, processing each record

    SourceFile.CloseFile
    SkeletonForm.lblStatus.Caption = "Done!"
End Sub

Public Function SearchDBFor035(RecordID As String) As String
    Dim BibID As String
    Dim SQL As String
    Dim rs As Integer
    
    BibID = ""
    '035A index entries like this: CCN12345678; 0350 entries like this: CASSIDY CCN12345678
    SQL = _
        "SELECT Bib_ID " & _
        "FROM Bib_Index " & _
        "WHERE Index_Code = '035A' " & _
        "AND Normal_Heading = '" & Normalize0350(RecordID) & "' " & _
        "ORDER BY Bib_ID"
    rs = GL.GetRS
    With GL.Vger
        .ExecuteSQL SQL, rs
        If .GetNextRow(rs) Then
            BibID = .CurrentRow(rs, 1)
        End If
    End With
    GL.FreeRS rs
    SearchDBFor035 = BibID
End Function

Public Function SearchDBForISSN(ISSN As String) As String
    'For this project, only search 022 $a and $y, not $z
    Dim BibID As String
    Dim SQL As String
    Dim rs As Integer
    
    BibID = ""
    SQL = _
        "SELECT Bib_ID " & _
        "FROM Bib_Index " & _
        "WHERE Index_Code IN ('ISSA', 'ISSY') " & _
        "AND Normal_Heading = '" & ISSN & "' " & _
        "ORDER BY Bib_ID"
    rs = GL.GetRS
    With GL.Vger
        .ExecuteSQL SQL, rs
        If .GetNextRow(rs) Then
            BibID = .CurrentRow(rs, 1)
        End If
    End With
    GL.FreeRS rs
    
    SearchDBForISSN = BibID
End Function
