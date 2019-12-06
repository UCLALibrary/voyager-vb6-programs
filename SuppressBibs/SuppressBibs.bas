Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    Dim BibID As String
    Dim SQL As String
    Dim rs As Integer
    Dim BibRC As UpdateBibReturnCode
    Dim CatLocID As Long
    
    CatLocID = GetLocID("lissystem")

    ' 20090619 akohler: new SQL, finds bibs where there's at least 1 suppressed mfhd, no unsuppressed mfhds, and bib is unsuppressed
    SQL = GetTextFromFile(GL.InputFilename)
 
    rs = GL.GetRS
    
    With GL.Vger
        WriteLog GL.Logfile, "Executing SQL to retrieve records... " & Now()
        SkeletonForm.lblStatus.Caption = "Executing SQL..."
        DoEvents
        .ExecuteSQL SQL, rs
        WriteLog GL.Logfile, "Finished executing SQL: " & Now()
        Do While .GetNextRow(rs)
            BibID = .CurrentRow(rs, 1)
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            .RetrieveBibRecord BibID
            BibRC = GL.BatchCat.UpdateBibRecord( _
                CLng(BibID), _
                .BibRecord, _
                .BibUpdateDateVB, _
                .BibOwningLibraryNumber, _
                CatLocID, _
                True _
            )
            If BibRC = ubSuccess Then
                WriteLog GL.Logfile, "Suppressed " & BibID
            Else
                WriteLog GL.Logfile, "Error suppressing " & BibID & " - returncode: " & BibRC
            End If
            NiceSleep GL.Interval
        Loop
    End With
    
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
End Sub
