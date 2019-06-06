Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim rs As Integer 'recordset
    
    Dim BibID As Long
    Dim HolID As Long
    Dim Record As Utf8MarcRecordClass
    
    Dim SpacCode As String
    Dim SpacText As String
    
    SQL = GetTextFromFile(GL.InputFilename)
    rs = GL.GetRS
    DoEvents
    
    With GL.Vger
        SkeletonForm.lblStatus.Caption = "Executing SQL..."
        DoEvents
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            BibID = .CurrentRow(rs, 1)
            HolID = .CurrentRow(rs, 2)
            SpacCode = .CurrentRow(rs, 3)
            SpacText = .CurrentRow(rs, 4)
            SkeletonForm.lblStatus.Caption = "Processing " & SpacCode
            DoEvents
            
            WriteLog GL.Logfile, "Checking bib " & BibID & ", hol " & HolID
            
            'Process bib record
            Set Record = GetVgerBibRecord(CStr(BibID))
            AddSpacToRecord Record, "bib", BibID, SpacCode, SpacText
            
            'Process holdings record
            Set Record = GetVgerHolRecord(CStr(HolID))
            AddSpacToRecord Record, "hol", HolID, SpacCode, SpacText
            
            DoEvents
            NiceSleep GL.Interval
        Loop 'GetNextRow
    End With
    
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs

End Sub

'Sub signature is awkward, I'm not proud of it.....
Private Sub AddSpacToRecord(ByVal Record As Utf8MarcRecordClass, RecordType As String, RecordID As Long, SpacCode As String, SpacText As String)
    Dim AddSpac As Boolean
    Dim NewSpacField As String
    Dim BibRC As UpdateBibReturnCode
    Dim HolRC As UpdateHoldingReturnCode
    
    With Record
'                WriteLog GL.Logfile, "BEFORE:"
'                WriteLog GL.Logfile, .TextFormatted
        AddSpac = True
        'Loop through 901 fields and add SPAC if not already present
        .FldFindFirst "901"
        Do While .FldWasFound
            .SfdFindFirst "a"
            If .SfdWasFound Then
                If .SfdText = SpacCode Then
                    AddSpac = False
                End If
            End If
            .FldFindNext
        Loop
        If AddSpac = True Then
            NewSpacField = .SfdMake("a", SpacCode) & .SfdMake("b", SpacText)
            .FldAddGeneric "901", "  ", NewSpacField, 3
'                    WriteLog GL.Logfile, " "
'                    WriteLog GL.Logfile, "AFTER:"
'                    WriteLog GL.Logfile, .TextFormatted

            Select Case RecordType
                Case "bib"
                    BibRC = GL.BatchCat.UpdateBibRecord( _
                        CLng(RecordID), _
                        Record.MarcRecordOut, _
                        GL.Vger.BibUpdateDateVB, _
                        GL.Vger.BibOwningLibraryNumber, _
                        GL.CatLocID, _
                        GL.Vger.BibRecordIsSuppressed _
                        )
                    If BibRC = ubSuccess Then
                        WriteLog GL.Logfile, vbTab & "Added " & SpacCode & " (" & SpacText & ")" & " to " & RecordType & " " & RecordID
                    Else
                        WriteLog GL.Logfile, "ERROR updating " & RecordType & " " & RecordID & " : return code " & BibRC
                    End If
                Case "hol"
                    HolRC = GL.BatchCat.UpdateHoldingRecord( _
                        RecordID, _
                        Record.MarcRecordOut, _
                        GL.Vger.HoldUpdateDateVB, _
                        GL.CatLocID, _
                        GL.Vger.HoldBibRecordNumber, _
                        GL.Vger.HoldLocationID, _
                        GL.Vger.HoldRecordIsSuppressed _
                        )
                    'Error 42 is spurious "could not update record" almost always due to call number / 852 indicator mismatch; ignore in this program
                    If HolRC = uhSuccess Or HolRC = uhCouldNotUpdateRecord Then
                        WriteLog GL.Logfile, vbTab & "Added " & SpacCode & " (" & SpacText & ")" & " to " & RecordType & " " & RecordID
                    Else
                        WriteLog GL.Logfile, "ERROR updating " & RecordType & " " & RecordID & " : return code " & HolRC
                    End If
                Case Else
                    'Serious problem, exit immediately
                    End
            End Select
        Else
            WriteLog GL.Logfile, vbTab & "Found " & SpacCode & ", no change made" & " to " & RecordType & " " & RecordID
        End If
    End With
End Sub
