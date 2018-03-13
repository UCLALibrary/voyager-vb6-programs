Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    Dim SQL As String
    Dim rs As Integer
    Dim BibID As Long
    Dim BibRC As DeleteBibReturnCode
    
    SQL = _
        "SELECT B.Bib_ID " & _
        "FROM Bib_Master B " & _
        "WHERE B.Suppress_In_Opac = 'N' " & _
        "AND NOT EXISTS (SELECT * FROM Bib_MFHD WHERE Bib_ID = B.Bib_ID) " & _
        "AND NOT EXISTS (SELECT * FROM Bib_History WHERE Bib_ID = B.Bib_ID AND Action_Date >= (SYSDATE - 7)) " & _
        "ORDER BY B.Bib_ID "

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
            BibRC = GL.BatchCat.DeleteBibRecord(BibID)
            If BibRC = dbSuccess Then
                WriteLog GL.Logfile, "Deleted bib #" & BibID
            Else
                WriteLog GL.Logfile, "Error deleting BibID " & BibID & " : " & TranslateBibDeleteCode(BibRC)
            End If
            NiceSleep GL.Interval
        Loop
    End With
    
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
End Sub
