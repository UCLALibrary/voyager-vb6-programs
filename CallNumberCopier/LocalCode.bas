Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim rs As Integer
    
    Dim SourceHolID As Long
    Dim TargetHolID As Long

    SQL = GetTextFromFile(GL.InputFilename)
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            SourceHolID = .CurrentRow(rs, 1)
            TargetHolID = .CurrentRow(rs, 2)
            SkeletonForm.lblStatus.Caption = "Processing " & SourceHolID
            ' Tweak for special indicator-only subset of records needed on VBT-817
            If SourceHolID = -1 Then
                UpdateIndicatorVbt817 TargetHolID
            Else
                CopyCallNumber SourceHolID, TargetHolID
            End If
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs

End Sub

Private Sub CopyCallNumber(SourceHolID As Long, TargetHolID As Long)
    Dim SourceHolRecord As Utf8MarcRecordClass
    Dim TargetHolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    Dim CallNumber As String
    Dim Indicators As String
    Dim Changed As Boolean
    
    Set SourceHolRecord = GetVgerHolRecord(CStr(SourceHolID))
    Set TargetHolRecord = GetVgerHolRecord(CStr(TargetHolID))
    ' Testing for Nothing doesn't work, even if record wasn't found; test for data
    If SourceHolRecord.MarcRecordIn = "" Or TargetHolRecord.MarcRecordIn = "" Then
        WriteLog GL.Logfile, "ERROR: Source record " & SourceHolID & " and/or Target record " & TargetHolID & " not found in Voyager - skipping"
        Exit Sub
    End If
    
'    WriteLog GL.Logfile, "SOURCE BEFORE*****************************"
'    WriteLog GL.Logfile, SourceHolRecord.TextFormatted
'    WriteLog GL.Logfile, ""
'    WriteLog GL.Logfile, "TARGET BEFORE*****************************"
'    WriteLog GL.Logfile, TargetHolRecord.TextFormatted
'    WriteLog GL.Logfile, ""
    
    With SourceHolRecord
        .FldFindFirst "852"
        If .FldWasFound Then
            Indicators = .FldInd
            .SfdFindFirst "k"
            If .SfdWasFound Then
                CallNumber = CallNumber & .SfdMake("k", .SfdText)
            End If
            .SfdFindFirst "h"
            If .SfdWasFound Then
                CallNumber = CallNumber & .SfdMake("h", .SfdText)
            End If
            .SfdFindFirst "i"
            If .SfdWasFound Then
                CallNumber = CallNumber & .SfdMake("i", .SfdText)
            End If
        Else
            WriteLog GL.Logfile, "ERROR: No 852 found in source record " & SourceHolID
            Exit Sub
        End If
    End With
    
    Changed = False
    With TargetHolRecord
        .FldFindFirst "852"
        If .FldWasFound Then
            ' Remove any existing call number subfields: $k, $h, $i, but *not* $j to support SRLF single-volume monos.
            ' Might be OK to do that as the SRLF852j program probably would recreate them but I'm not ready to make this
            ' program that generic yet.
            ' Does *not* consider possibility of $m (call number suffix) or $l (ell - shelving form of title) - rarely used but may be present.
            
            ' Should be only one (max) each of $k $h $i but loop to be sure
            .SfdFindFirst "k"
            Do While .SfdWasFound
                .SfdDelete
                .SfdFindNext
            Loop
            .SfdFindFirst "h"
            Do While .SfdWasFound
                .SfdDelete
                .SfdFindNext
            Loop
            .SfdFindFirst "i"
            Do While .SfdWasFound
                .SfdDelete
                .SfdFindNext
            Loop
            
            ' Now add call number from source record to what's left of the 852 and set the indicators from source as well.
            .FldText = .FldText & CallNumber
            .FldInd = Indicators
            Changed = True
        Else
            WriteLog GL.Logfile, "ERROR: No 852 found in target record " & TargetHolID
            Exit Sub
        End If
    End With
    
'    If Changed = True Then
'        WriteLog GL.Logfile, "TARGET AFTER*****************************"
'        WriteLog GL.Logfile, TargetHolRecord.TextFormatted
'        WriteLog GL.Logfile, ""
'    End If

    If Changed = True Then
        HolRC = GL.BatchCat.UpdateHoldingRecord( _
            TargetHolID, _
            TargetHolRecord.MarcRecordOut, _
            GL.Vger.HoldUpdateDateVB, _
            GL.CatLocID, _
            GL.Vger.HoldBibRecordNumber, _
            GL.Vger.HoldLocationID, _
            GL.Vger.HoldRecordIsSuppressed _
        )
        If HolRC = uhSuccess Then
            WriteLog GL.Logfile, "Copied call number from source mfhd " & SourceHolID & " to mfhd " & TargetHolID & " : " & Replace(CallNumber, Chr(31), "$")
        Else
            WriteLog GL.Logfile, "ERROR updating target record " & TargetHolID & " : return code: " & HolRC
        End If
    End If

End Sub

Private Sub UpdateIndicatorVbt817(TargetHolID As Long)
    ' Set 852 1st indicator to 0
    ' Special subset of VBT-817, not currently intended for general use
    
    Dim TargetHolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    Dim Ind1 As String
    Dim Changed As Boolean
    
    Set TargetHolRecord = GetVgerHolRecord(CStr(TargetHolID))
    ' Testing for Nothing doesn't work, even if record wasn't found; test for data
    If TargetHolRecord.MarcRecordIn = "" Then
        WriteLog GL.Logfile, "ERROR: Target record " & TargetHolID & " not found in Voyager - skipping"
        Exit Sub
    End If
    
    Changed = False
    With TargetHolRecord
        .FldFindFirst "852"
        If .FldWasFound Then
            .FldInd1 = "0"
            Changed = True
        Else
            WriteLog GL.Logfile, "ERROR: No 852 found in target record " & TargetHolID
            Exit Sub
        End If
    End With

    If Changed = True Then
        HolRC = GL.BatchCat.UpdateHoldingRecord( _
            TargetHolID, _
            TargetHolRecord.MarcRecordOut, _
            GL.Vger.HoldUpdateDateVB, _
            GL.CatLocID, _
            GL.Vger.HoldBibRecordNumber, _
            GL.Vger.HoldLocationID, _
            GL.Vger.HoldRecordIsSuppressed _
        )
        If HolRC = uhSuccess Then
            WriteLog GL.Logfile, "Updated indicator1 to 0 on mfhd " & TargetHolID
        Else
            WriteLog GL.Logfile, "ERROR updating target record " & TargetHolID & " : return code: " & HolRC
        End If
    End If
    
End Sub
