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
    Dim Changed As Boolean
    
    SQL = GetTextFromFile(GL.InputFilename)
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
        HolID = .CurrentRow(rs, 1)
            Changed = False
            SkeletonForm.lblStatus.Caption = "Processing " & HolID
            Set HolRecord = GetVgerHolRecord(CStr(HolID))
            With HolRecord
                WriteLog GL.Logfile, "Processing mfhd " & HolID
                .FldFindFirst "852"
                WriteLog GL.Logfile, vbTab & "Before: " & .FldTextFormatted
                .SfdFindFirst "k"
                If .SfdText = "Sadleir" Then
                    .SfdDelete
                    .SfdFindFirst "h"
                    If .SfdWasFound Then
                        .SfdText = "Sadleir " & .SfdText
                    End If
                    WriteLog GL.Logfile, vbTab & "After : " & .FldTextFormatted
                    Changed = True
                Else
                    WriteLog GL.Logfile, "WARNING: NO CHANGE"
                End If
            End With
            
            If Changed = True Then
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
                    WriteLog GL.Logfile, "Updated mfhd " & HolID
                Else
                    WriteLog GL.Logfile, "ERROR updating mfhd " & HolID & " : return code " & HolRC
                End If
            End If
            WriteLog GL.Logfile, ""
            
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
End Sub
    

