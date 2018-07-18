Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim HolID As Long
    Dim HolRC As UpdateHoldingReturnCode
    Dim HolRecord As Utf8MarcRecordClass
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
            DoEvents
            Set HolRecord = GetVgerHolRecord(CStr(HolID))
            With HolRecord
                .FldFindFirst "583"
                'Records have multiple 583 fields, get the right one
                Do While .FldWasFound
                    .SfdFindFirst "a"
                    If .SfdWasFound Then
                        If .SfdText = "committed to retain" Then
                            'Find the first $f, add a new one after it
                            .SfdFindFirst "f"
                            .SfdInsertAfter "f", "Licensed Content"
                        End If
                    End If
                    .FldFindNext
                Loop
            End With 'HolRecord
            
            HolRC = GL.BatchCat.UpdateHoldingRecord( _
                HolID, _
                HolRecord.MarcRecordOut, _
                .HoldUpdateDateVB, _
                GL.CatLocID, _
                .HoldBibRecordNumber, _
                .HoldLocationID, _
                .HoldRecordIsSuppressed _
            )
            If HolRC = uhSuccess Then
                WriteLog GL.Logfile, "Updated mfhd " & HolID
            Else
                WriteLog GL.Logfile, "ERROR updating mfhd " & HolID & " : return code " & HolRC
            End If
            NiceSleep GL.Interval
        Loop
    End With 'Vger
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
End Sub
