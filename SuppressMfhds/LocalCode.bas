Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim InputFileNum As Integer
    Dim Line As String
    Dim HolID As String
    Dim HolRC As UpdateHoldingReturnCode
    Dim rs As Integer
    
    InputFileNum = FreeFile
    Open GL.InputFilename For Input As InputFileNum
    Do While Not EOF(InputFileNum)
        Line Input #InputFileNum, HolID
        With GL.Vger
            'Get the record and suppress it
            .RetrieveHoldRecord HolID
            HolRC = GL.BatchCat.UpdateHoldingRecord( _
                CLng(HolID), _
                .HoldRecord, _
                .HoldUpdateDateVB, _
                GL.CatLocID, _
                .HoldBibRecordNumber, _
                .HoldLocationID, _
                True _
            )
            If HolRC = uhSuccess Then
                WriteLog GL.Logfile, "Suppressed " & HolID
            Else
                WriteLog GL.Logfile, "Error suppressing " & HolID & " - returncode: " & HolRC
            End If
        End With
        NiceSleep GL.Interval
    Loop
    Close InputFileNum
End Sub
