Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim rs As Integer 'recordset
    
    Dim RecordID As Long
    Dim Record As Utf8MarcRecordClass
    Dim BibRC As UpdateBibReturnCode
    Dim HolRC As UpdateHoldingReturnCode
    
    Dim SpacCode As String
    Dim NewSpacText As String
    Dim OldSpacText As String
    Dim Changed As Boolean
    
    Dim RecordType As String
    Dim RecordTypes(1 To 2) As String
    RecordTypes(1) = "bib"
    RecordTypes(2) = "mfhd"
    
    Dim i As Integer
    
    'Not really a location, but currently best way to pass parameter to this program unchanged
    SpacCode = GL.Location
    
    For i = LBound(RecordTypes) To UBound(RecordTypes)
        RecordType = RecordTypes(i)
        SQL = "select s.record_id, sm.code, sm.name " _
        & "from vger_support.spac_map sm " _
        & "inner join vger_subfields." & GL.TableSpace & "_" & RecordType & "_subfield s " _
        & "on s.tag = '901a' and sm.code = s.subfield " _
        & "where sm.code = '" & SpacCode & "' " _
        & "order by record_id"
        
        rs = GL.GetRS
        SkeletonForm.lblStatus.Caption = "Executing SQL for " & RecordType & " records..."
        DoEvents
        
        With GL.Vger
            .ExecuteSQL SQL, rs
            Do While .GetNextRow(rs)
                RecordID = .CurrentRow(rs, 1)
                SpacCode = .CurrentRow(rs, 2)
                NewSpacText = .CurrentRow(rs, 3)
                SkeletonForm.lblStatus.Caption = "Processing " & RecordType & " " & RecordID
                Select Case RecordType
                    Case "bib"
                        Set Record = GetVgerBibRecord(CStr(RecordID))
                    Case "mfhd"
                        Set Record = GetVgerHolRecord(CStr(RecordID))
                    Case Else
                        'Serious problem, exit immediately
                        End
                End Select
                        
                With Record
                    Changed = False
                    OldSpacText = ""
                    .FldFindFirst "901"
                    Do While .FldWasFound
                        .SfdFindFirst "a"
                        If .SfdText = SpacCode Then
                            .SfdFindFirst "b"
                            If .SfdWasFound Then
                                'IS .SFDTEXT AUTOMATICALLY TRIMMED?  SEEMS IT IS.... NOT FIXED 2011-10-17
                                If .SfdText <> NewSpacText Then
                                    OldSpacText = .SfdText
                                    .SfdText = NewSpacText
                                    Changed = True
                                End If
                            Else
                                .SfdAdd "b", NewSpacText
                                Changed = True
                            End If
                        End If
                        .FldFindNext
                    Loop 'FldWasFound
                End With 'Record
                
                If Changed Then
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
                                WriteLog GL.Logfile, "Updated SPAC text for " & SpacCode & " on " & RecordType & " " & RecordID
                                WriteLog GL.Logfile, vbTab & "From: " & OldSpacText
                                WriteLog GL.Logfile, vbTab & "To  : " & NewSpacText
                            Else
                                WriteLog GL.Logfile, "ERROR updating " & RecordType & " " & RecordID & " : return code " & BibRC
                            End If
                        Case "mfhd"
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
                                WriteLog GL.Logfile, "Updated SPAC text for " & SpacCode & " on " & RecordType & " " & RecordID
                                WriteLog GL.Logfile, vbTab & "From: " & OldSpacText
                                WriteLog GL.Logfile, vbTab & "To  : " & NewSpacText
                            Else
                                WriteLog GL.Logfile, "ERROR updating " & RecordType & " " & RecordID & " : return code " & HolRC
                            End If
                        Case Else
                            'Serious problem, exit immediately
                            End
                    End Select
                End If 'Changed
            Loop '.GetNextRow
            NiceSleep GL.Interval
        End With 'GL.Vger
    Next 'RecordType

    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs

End Sub
