Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim rs As Integer
    Dim BibID As String
    Dim oclc As String
    Dim BibRecord As Utf8MarcRecordClass
    Dim BibRC As UpdateBibReturnCode
    Dim Message As String
    
    'ZAS reclamation only
    SQL = "select bib_id, oclc_number from vger_report.zas_matches_voyager order by bib_id"
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            BibID = CStr(.CurrentRow(rs, 1))
            oclc = CStr(.CurrentRow(rs, 2))
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            Set BibRecord = GetVgerBibRecord(BibID)
            With BibRecord
                Message = ""
                .FldFindFirst "035"
                Do While .FldWasFound
                    'Remove bib id in 035 from old project
                    If .FldText = .SfdDelim("a") & BibID Then
                        .FldDelete
                    'For SRLF (ZAS), change existing $a(OCoLC) to $z per slayne 20091216
                    Else
                        'Change all $a(OCoLC) to $z
                        .SfdFindFirst "a"
                        Do While .SfdWasFound
                            If Left(.SfdText, 7) = "(OCoLC)" Then
                                .SfdCode = "z"
                                Message = "changed $a->$z; "
                            End If
                            .SfdFindNext
                        Loop
                    End If
                    .FldFindNext
                Loop
                'Add new field with OCLC-supplied number
                .FldAddGeneric "035", "  ", .SfdMake("a", "(OCoLC)" & LeftPad(oclc, "0", 8)), 3
                Message = Message & "added oclc " & oclc
            End With
            
            Set BibRecord = UpdateUcoclc(BibRecord)
            
            With GL.Vger
                BibRC = GL.BatchCat.UpdateBibRecord( _
                    CLng(BibID), _
                    BibRecord.MarcRecordOut, _
                    .BibUpdateDateVB, _
                    .BibOwningLibraryNumber, _
                    GL.CatLocID, _
                    .BibRecordIsSuppressed _
                )
                If BibRC = ubSuccess Then
                    WriteLog GL.Logfile, BibID & vbTab & Message
                Else
                    WriteLog GL.Logfile, BibID & vbTab & "ERROR updating record: " & BibRC
                End If
            End With
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
End Sub
