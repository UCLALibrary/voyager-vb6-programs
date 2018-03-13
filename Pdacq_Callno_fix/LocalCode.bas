Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim HolID As Long
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    Dim rs As Integer
    
    
    SQL = GetTextFromFile(GL.InputFilename)
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
        HolID = .CurrentRow(rs, 1)
        'do something with results
            SkeletonForm.lblStatus.Caption = "Processing " & HolID
            Set HolRecord = GetVgerHolRecord(CStr(HolID))
            With HolRecord
                ' Records were pre-checked so QAD processing...
                .FldFindFirst "852"
                .SfdFindFirst "h"
                If .SfdWasFound Then
                    ' UCLA faculty and students ONLY: Click on "Request an Item" and select "Buy this item-Purchase Request" to ask Library to order this
                    .SfdText = "UCLA faculty and students ONLY: Click on ""Request an Item"" and select ""Buy this item-Purchase Request"" to ask Library to order this"
                End If
            End With
            
            HolRC = GL.BatchCat.UpdateHoldingRecord( _
                HolID, _
                HolRecord.MarcRecordOut, _
                GL.Vger.HoldUpdateDateVB, _
                GL.CatLocID, _
                GL.Vger.HoldBibRecordNumber, _
                GL.Vger.HoldLocationID, _
                GL.Vger.HoldRecordIsSuppressed _
            )
            If HolRC = uhSuccess Then
                WriteLog GL.Logfile, "Updated hol " & HolID
            Else
                WriteLog GL.Logfile, "ERROR updating hol " & HolID & " : return code " & HolRC
            End If
            
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
End Sub
