Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim BibID As Long
    Dim BibRC As UpdateBibReturnCode
    Dim BibRecord As Utf8MarcRecordClass
    Dim rs As Integer
    
    SQL = GetTextFromFile(GL.InputFilename)
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            BibID = .CurrentRow(rs, 1)
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            DoEvents
            Set BibRecord = GetVgerBibRecord(CLng(BibID))
            With BibRecord
                'TODO: Get field info from external source so recompiling isn't necessary
                .FldAddGeneric "590", "  ", .SfdMake("a", "Spec. Coll. Belt Copy: Gift, 1961.")
                'WriteLog GL.Logfile, .TextFormatted
            End With
            
            BibRC = GL.BatchCat.UpdateBibRecord( _
                CLng(BibID), _
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
