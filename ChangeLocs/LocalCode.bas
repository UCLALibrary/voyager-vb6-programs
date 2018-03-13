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
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    Dim HolRS As Integer
    Dim OldLoc As String
    Dim NewLoc As String
    
    SQL = GetTextFromFile(GL.InputFilename)
    
    HolRS = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    'SQL provides mfhd id, current loc code and new loc code
    With GL.Vger
        .ExecuteSQL SQL, HolRS
        Do While .GetNextRow(HolRS)
        HolID = .CurrentRow(HolRS, 1)
        OldLoc = .CurrentRow(HolRS, 2)
        NewLoc = .CurrentRow(HolRS, 3)
        'do something with results
            SkeletonForm.lblStatus.Caption = "Processing " & HolID
            Set HolRecord = GetVgerHolRecord(CStr(HolID))
            With HolRecord
                'Update the holdings record
                .FldFindFirst "852"
                .SfdFindFirst "b"
                If .SfdWasFound = True And .SfdText = OldLoc Then
                    .SfdText = NewLoc
                    HolRC = GL.BatchCat.UpdateHoldingRecord( _
                        HolID, _
                        .MarcRecordOut, _
                        GL.Vger.HoldUpdateDateVB, _
                        GL.CatLocID, _
                        GL.Vger.HoldBibRecordNumber, _
                        GetLocID(NewLoc), _
                        GL.Vger.HoldRecordIsSuppressed _
                    )
                    If HolRC = uhSuccess Then
                        WriteLog GL.Logfile, "UPDATED" & vbTab & HolID & vbTab & OldLoc & vbTab & NewLoc
                    Else
                        WriteLog GL.Logfile, "ERROR updating hol " & HolID & " : return code " & HolRC
                    End If
                Else
                    WriteLog GL.Logfile, "ERROR: wrong loc in " & HolID & " - expected " & OldLoc & " , found " & .SfdText
                End If
            End With
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS HolRS
    

'        If BibRC = ubSuccess Then
'            WriteLog GL.Logfile, "Updated bib " & BibID
'        Else
'            WriteLog GL.Logfile, "ERROR updating bib " & BibID & " : return code " & BibRC
'        End If
'
'        NiceSleep GL.Interval
'    Loop
'
'    SourceFile.CloseFile
'    SkeletonForm.lblStatus.Caption = "Done!"

End Sub
