Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim HolID As Long
    Dim HolRS As Integer
    Dim HolRC As DeleteHoldingReturnCode
    Dim HolDelCnt As Long
    
    SQL = GetTextFromFile(GL.InputFilename)
    HolRS = GL.GetRS
    HolDelCnt = 0
    
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    ' Looks only at holdings records found via HolRS SQL; does not look at items, orders, bibs
    With GL.Vger
        .ExecuteSQL SQL, HolRS
        Do While .GetNextRow(HolRS)
            HolID = CLng(.CurrentRow(HolRS, 1))
            SkeletonForm.lblStatus.Caption = "Processing " & HolID
            
            HolRC = GL.BatchCat.DeleteHoldingRecord(HolID)
            If HolRC = dhSuccess Then
                WriteLog GL.Logfile, "Deleted HolID " & HolID
                HolDelCnt = HolDelCnt + 1
            Else
                WriteLog GL.Logfile, "Error deleting HolID " & HolID & " : " & TranslateHoldingsDeleteCode(HolRC)
            End If
            
            DoEvents
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    
    WriteLog GL.Logfile, ""
    WriteLog GL.Logfile, "Deleted: " & HolDelCnt & " hols"
    WriteLog GL.Logfile, ""
    
    GL.FreeRS HolRS
End Sub

