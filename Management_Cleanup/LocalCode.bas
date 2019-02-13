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
    
    Dim f852x As String
    
    SQL = "select distinct mfhd_id from vger_report.tmp_vbt_975_mfhds order by mfhd_id"
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            HolID = .CurrentRow(rs, 1)
            SkeletonForm.lblStatus.Caption = "Processing " & HolID
            DoEvents
            
            Set HolRecord = GetVgerHolRecord(CStr(HolID))
            With HolRecord
                'Reset 852 $x target value
                f852x = ""
                .FldFindFirst "852"
                WriteLog GL.Logfile, "Checking MFHD " & HolID
                WriteLog GL.Logfile, vbTab & "Before: " & .FldTextFormatted
                'Set 852 1st indicator to blank, since we're removing the call number
                .FldInd1 = " "
                .SfdFindFirst "h"
                If .SfdWasFound Then
                    If InStr(1, .SfdText, "SUPPRESSED", vbTextCompare) = 1 Then
                        'No action
                    Else
                        'Move the old $h to $x
                        f852x = .SfdText
                        'Change the 852 $h
                        .SfdText = "SUPPRESSED"
                        'Now combine the $i with the $h in the new $x and remove the $i
                        .SfdFindFirst "i"
                        If .SfdWasFound Then
                            f852x = f852x & " " & .SfdText
                            .SfdDelete
                        End If
                        .SfdAdd "x", f852x
                    End If
                Else
                    'No 852 $h found, so add one
                    .SfdFindFirst "b"
                    .SfdInsertAfter "h", "SUPPRESSED"
                End If
                
                WriteLog GL.Logfile, vbTab & "After : " & .FldTextFormatted
                
                'Save the record, suppressing it if not already
                HolRC = GL.BatchCat.UpdateHoldingRecord( _
                    HolID, _
                    HolRecord.MarcRecordOut, _
                    GL.Vger.HoldUpdateDateVB, _
                    GL.CatLocID, _
                    GL.Vger.HoldBibRecordNumber, _
                    GL.Vger.HoldLocationID, _
                    True _
                )
                If HolRC = uhSuccess Then
                    WriteLog GL.Logfile, "Updated MFHD " & HolID
                Else
                    WriteLog GL.Logfile, "ERROR updating MFHD " & HolID & " : return code " & HolRC
                End If
                WriteLog GL.Logfile, ""
            End With
            
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
    

'        SkeletonForm.lblStatus.Caption = "Processing " & BibID
'
'        'To populate GL.Vger's convenience fields - retrieved record not actually used
'        GetVgerBibRecord BibID
'
'        BibRC = GL.BatchCat.UpdateBibRecord( _
'            CLng(BibID), _
'            BibRecord.MarcRecordOut, _
'            GL.Vger.BibUpdateDateVB, _
'            GL.Vger.BibOwningLibraryNumber, _
'            GL.CatLocID, _
'            GL.Vger.BibRecordIsSuppressed _
'            )
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
