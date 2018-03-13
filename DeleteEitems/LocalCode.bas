Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim InputFileNum As Integer
    Dim Line As String
    Dim BibID As Long
    Dim MfhdID As Long
    Dim IdArray() As String
    Dim DelBibRc As DeleteBibReturnCode
    Dim DelMfhdRc As DeleteHoldingReturnCode
    
    InputFileNum = FreeFile
    Open GL.InputFilename For Input As InputFileNum
    Do While Not EOF(InputFileNum)
        Line Input #InputFileNum, Line
        IdArray = Split(Line, vbTab)
        BibID = CLng(IdArray(0))
        MfhdID = CLng(IdArray(1))
        SkeletonForm.lblStatus.Caption = "Processing " & MfhdID
        DelMfhdRc = GL.BatchCat.DeleteHoldingRecord(MfhdID)
        If DelMfhdRc = dhSuccess Then
            WriteLog GL.Logfile, "Deleted mfhd #" & MfhdID
            ' Delete mfhd - try to delete bib
            DelBibRc = GL.BatchCat.DeleteBibRecord(BibID)
            If DelBibRc = dbSuccess Then
                WriteLog GL.Logfile, "Deleted bib #" & BibID
            Else
                WriteLog GL.Logfile, "*** Error deleting bib #" & BibID & " - returncode: " & TranslateBibDeleteCode(DelBibRc)
            End If
        Else
            WriteLog GL.Logfile, "*** Error deleting mfhd #" & MfhdID & " - returncode: " & TranslateHoldingsDeleteCode(DelMfhdRc)
        End If
        WriteLog GL.Logfile, ""
        NiceSleep GL.Interval
    Loop
    SkeletonForm.lblStatus.Caption = "Done!"
    Close InputFileNum
End Sub
