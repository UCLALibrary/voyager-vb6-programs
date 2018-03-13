Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim InputFileNum As Integer
    Dim Line As String
    Dim SQL As String
    Dim HolID As Long
    Dim Barcode As String
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    Dim rs As Integer
    Dim changed As Boolean
    
    SQL = GetTextFromFile(GL.InputFilename)
    
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            HolID = .CurrentRow(rs, 1)
            Barcode = .CurrentRow(rs, 2)
            SkeletonForm.lblStatus.Caption = "Processing " & HolID
            Set HolRecord = GetVgerHolRecord(CStr(HolID))
            ' Record might have been deleted since SQL executed
            If HolRecord.MarcRecordIn <> "" Then
                With HolRecord
                    changed = False
                    .FldFindFirst "852"
                    .SfdFindFirst "j"
                    If .SfdWasFound Then
                        If .SfdText <> Barcode Then
                            .SfdText = Barcode
                            changed = True
                        End If
                    Else
                        .SfdAdd "j", Barcode
                        changed = True
                    End If
                End With
                If changed = True Then
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
                        WriteLog GL.Logfile, "HolID " & HolID & " updated"
                    Else
                        WriteLog GL.Logfile, "HolID " & HolID & " error - return code " & HolRC
                    End If
                End If
            End If
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
    
End Sub
