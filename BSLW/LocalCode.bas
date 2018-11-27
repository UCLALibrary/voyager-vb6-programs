Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    
'GL.Init "-t ucladb -f " & App.Path & "\4934310.mrc"
    SplitFile
    LoadRecords GL.InputFilename
End Sub

Sub SplitFile()
    Dim SourceFile As New Utf8MarcFileClass
    Dim ChangedFile As Integer      'file handle
    Dim UnchangedFile As Integer    'file handle
    
    Dim ChangedCount As Long
    Dim UnchangedCount As Long
    
    Dim MarcRecord As New Utf8MarcRecordClass
    Dim RawRecord As String
    Dim Changed As Boolean
    
    ChangedFile = FreeFile
    Open GL.BaseFilename + ".changed" For Binary As ChangedFile
    UnchangedFile = FreeFile
    Open GL.BaseFilename + ".unchanged" For Binary As UnchangedFile
    
    ChangedCount = 0
    UnchangedCount = 0
    
    SourceFile.OpenFile GL.InputFilename
    Do While SourceFile.ReadNextRecord(RawRecord)
        Set MarcRecord = New Utf8MarcRecordClass    'Treat each record as a separate instance
        'MarcRecordOut automatically changes LDR/22-23 to "00"
        'It also appears to trim leading/trailing space from non-control fields (010 & higher)
        With MarcRecord
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            .MarcRecordIn = RawRecord
            
            Changed = False
            .FldFindFirst "040"
            If .FldWasFound Then
                .SfdFindFirst "d"
                Do While .SfdWasFound
                    If .SfdText = "UtOrBLW" Then
                        Changed = True
                        Exit Do
                    End If
                    .SfdFindNext
                Loop
            End If
            
            If Changed = True Then
                Put #ChangedFile, , RawRecord
                ChangedCount = ChangedCount + 1
            Else
                Put #UnchangedFile, , RawRecord
                UnchangedCount = UnchangedCount + 1
            End If
        End With
    Loop
    
    WriteLog GL.Logfile, "Read " & (ChangedCount + UnchangedCount) & " records: " & ChangedCount & " changed, " & UnchangedCount & " unchanged"
    WriteLog GL.Logfile, ""
    SourceFile.CloseFile
    Close #ChangedFile
    Close #UnchangedFile
    
End Sub

Sub LoadRecords(Filename As String)
    Dim SourceFile As New Utf8MarcFileClass
    Dim DeleteFile As Integer   'file handle
    Dim ReviewFile As Integer   'file handle
    
    Dim BslwRecord As New Utf8MarcRecordClass
    Dim VgerRecord As New Utf8MarcRecordClass
    
    Dim RawRecord As String
    Dim RecordNumber As Long
    Dim Changed As Boolean
    Dim F001 As String
    Dim BslwF005 As String
    Dim VgerF005 As String
    
    Dim UpdateBibRC As UpdateBibReturnCode
    
    DeleteFile = FreeFile
    Open GL.BaseFilename + ".deleted" For Binary As DeleteFile
    ReviewFile = FreeFile
    Open GL.BaseFilename + ".review" For Binary As ReviewFile
    
    WriteLog GL.Logfile, "Processing " & GL.InputFilename & ", starting with record #" & GL.StartRec & ":"
    
    RecordNumber = 0
    SourceFile.OpenFile Filename
    Do While SourceFile.ReadNextRecord(RawRecord)
        RecordNumber = RecordNumber + 1
        'Allow restart at specified record
        If RecordNumber >= GL.StartRec Then
            Set BslwRecord = New Utf8MarcRecordClass    'Treat each record as a separate instance
            With BslwRecord
                '20080429: records now are in UTF-8, no conversion needed
                '.CharacterSetIn = "M"
                '.MarcRecordIn = RawRecord 'MarcRecordIn *must* be done before CharacterSetOut, else conversion to UTF-8 can be incorrect
                .CharacterSetIn = "U"
                .CharacterSetOut = "U"
                .IgnoreSfdOrder = True
                .MarcRecordIn = RawRecord
                
                .FldFindFirst "001"
                If .FldWasFound Then
                    F001 = .FldText
                Else
                    F001 = "ERROR: NO BSLW 001"
                End If
                
                If GL.Use_GUI = True Then
                    SkeletonForm.lblStatus.Caption = "Processing bib #" & F001
                    DoEvents
                End If
    
                .FldFindFirst "005"
                If .FldWasFound Then
                    BslwF005 = .FldText
                Else
                    BslwF005 = "ERROR: NO BSLW 005"
                End If
'Debug.Print BslwRecord.TextFormatted
                Set VgerRecord = GetVgerBibRecord(F001)
                'If nothing in VgerRecord for this 001, the record has been deleted since extract
                If VgerRecord.MarcRecordIn = "" Then
                    WriteLog GL.Logfile, "Bib #" & F001 & " not found in Voyager: BSLW record written to deleted file"
                    Put #DeleteFile, , RawRecord    'the record as received from BSLW
                Else
                    'Check the 005; update Voyager if it matches, reject BSLW record to review file if it doesn't
                    With VgerRecord
                        .FldFindFirst "005"
                        If .FldWasFound Then
                            VgerF005 = .FldText
                        Else
                            VgerF005 = "ERROR: NO VGER 005"
                        End If
                    End With
    
' Hack for reload of October 2007 file
'                    If (BslwF005 = VgerF005) Or (VgerF005 >= "20071106183052" And VgerF005 <= "20071106194705") Then
' Hack for reload of December 2007 file (005 mismatches due to 035 ucoclc updates)
'   Already ran with standard match on 005, so re-running with special 005 check
'                    If (VgerF005 >= "20071230200000" And VgerF005 <= "20080102070000") Then
' Hack for reload of Feb 2008 records with 880 problems
'                    If (BslwF005 = VgerF005) _
'                        Or (VgerF005 >= "20080403192000" And VgerF005 <= "20080407100259") _
'                        Or (VgerF005 >= "20080410113000" And VgerF005 <= "20080410213659") Then
                    If BslwF005 = VgerF005 Then
                        If GL.ProductionMode = True Then
                            UpdateBibRC = GL.BatchCat.UpdateBibRecord( _
                                CLng(F001), _
                                BslwRecord.MarcRecordOut, _
                                GL.Vger.BibUpdateDateVB, _
                                GL.Vger.BibOwningLibraryNumber, _
                                GL.CatLocID, _
                                GL.Vger.BibRecordIsSuppressed _
                            )
                            If UpdateBibRC = ubSuccess Then
                                WriteLog GL.Logfile, "Bib #" & F001 & " match on 005: Voyager updated"
                            Else
                                WriteLog GL.Logfile, "ERROR: Bib #" & F001 & " match on 005, but could not update Voyager; returncode: " & UpdateBibRC
                            End If
                        Else
                            WriteLog GL.Logfile, "Bib #" & F001 & " match on 005: Voyager WOULD BE updated"
                        End If
                    Else
                        WriteLog GL.Logfile, "Bib #" & F001 & " no match on 005: BSLW record written to review file"
                        Put #ReviewFile, , RawRecord    'the record as received from BSLW, still in MARC-8
                    End If
                End If
            End With 'BslwRecord
            NiceSleep GL.Interval
        End If 'RecordNumber
    Loop 'ReadNextRecord
    
    If GL.Use_GUI = True Then
        SkeletonForm.lblStatus.Caption = "Done!"
        DoEvents
    End If
    SourceFile.CloseFile
    Close #DeleteFile
    Close #ReviewFile
End Sub

