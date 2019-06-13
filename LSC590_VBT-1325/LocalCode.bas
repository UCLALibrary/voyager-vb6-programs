Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim BibID As Long
    Dim Treatment As Integer
    Dim rs As Integer
    Dim BibRC As UpdateBibReturnCode
    Dim BibRecord As Utf8MarcRecordClass
    Dim F590_combined As String
    
    SQL = GetTextFromFile(GL.InputFilename)
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            BibID = .CurrentRow(rs, 1)
            Treatment = .CurrentRow(rs, 2)
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            
            Set BibRecord = GetVgerBibRecord(CStr(BibID))
            With BibRecord
                'Initialize new 590 field
                F590_combined = "Spec. Coll. copy:"
                'Loop through 590 fields, appending text from each into one field
                'Pre-checked that all 590s have only $a, and have only 1 $a
                'Also each record is present only once in the record set.
                .FldFindFirst "590"
                Do While .FldWasFound
                    .SfdFindFirst "a"
                    F590_combined = F590_combined & " " & .SfdText
                    .FldDelete
                    .FldFindNext
                Loop
                'Finally, add the new combined 590 field
                .FldAddGeneric "590", "  ", .SfdMake("a", F590_combined)
                'WriteLog GL.Logfile, .TextFormatted
            End With
            
            'Update Voyager
            BibRC = GL.BatchCat.UpdateBibRecord( _
                BibID, _
                BibRecord.MarcRecordOut, _
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
        Loop
    End With
    
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
End Sub
