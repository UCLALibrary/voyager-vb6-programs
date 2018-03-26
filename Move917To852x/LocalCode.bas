Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim InputFileNum As Integer
    Dim HolID As String
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    Dim rs As Integer
    Dim s917a As String
    
    InputFileNum = FreeFile
    Open GL.InputFilename For Input As InputFileNum
    Do While Not EOF(InputFileNum)
        Line Input #InputFileNum, HolID
        SkeletonForm.lblStatus.Caption = "Processing " & HolID
        DoEvents
        
        Set HolRecord = GetVgerHolRecord(HolID)
        ' Some records may no longer exist
        'If HolRecord Is Nothing Then 'NOPE
        'If IsEmpty(HolRecord) Then 'NOPE
        If Len(HolRecord.MarcRecordIn) = 0 Then
            WriteLog GL.Logfile, "HolID " & HolID & " does not exist"
        Else
            With HolRecord
                'Special purpose program
                'Verified separately that there's at most 1 917 per record, and those have only $a
                'However, some records may not have a 917 field.
                .FldFindFirst "917"
                If .FldWasFound Then
                    WriteLog GL.Logfile, "Changing HolID " & HolID
                    'Grab 917 $a
                    .SfdFindFirst "a"
                    s917a = .SfdText
                    .FldDelete
                    
                    'Add text of 917 $a to 852 as new $x
                    .FldFindFirst "852"
                    'All holdings have 852...
                    WriteLog GL.Logfile, vbTab & "From: " & .FldText
                    .SfdAdd "x", s917a
                    WriteLog GL.Logfile, vbTab & "To  : " & .FldText
                    
                    HolRC = GL.BatchCat.UpdateHoldingRecord( _
                        CLng(HolID), _
                        .MarcRecordOut, _
                        GL.Vger.HoldUpdateDateVB, _
                        GL.CatLocID, _
                        GL.Vger.HoldBibRecordNumber, _
                        GL.Vger.HoldLocationID, _
                        GL.Vger.HoldRecordIsSuppressed _
                    )
                    If HolRC = uhSuccess Then
                        WriteLog GL.Logfile, "Updated HolID " & HolID
                    Else
                        WriteLog GL.Logfile, "ERROR updating HolID " & HolID & " : return code " & HolRC
                    End If
                Else
                    WriteLog GL.Logfile, "No 917 field found in HolID " & HolID
                End If
            End With
        End If
        
        NiceSleep GL.Interval
    Loop
    Close InputFileNum
    SkeletonForm.lblStatus.Caption = "Done!"

'Or MARC-based:
'    Dim BibID As String
'    Dim BibRC As UpdateBibReturnCode
'    Dim SourceFile As Utf8MarcFileClass
'    Dim BibRecord As Utf8MarcRecordClass
'    Dim RawRecord As String
'
'    Set SourceFile = New Utf8MarcFileClass
'    SourceFile.OpenFile GL.InputFilename
'
'    Do While SourceFile.ReadNextRecord(RawRecord)
'        DoEvents
'        Set BibRecord = New Utf8MarcRecordClass
'        With BibRecord
'            'convert to Unicode for Voyager
'            .CharacterSetIn = "O"   'OCLC records
'            .CharacterSetOut = "U"
'            .MarcRecordIn = RawRecord
'
'            .FldFindFirst "001"
'            If .FldWasFound Then
'                BibID = Trim(.FldText)
'            Else
'                WriteLog GL.Logfile, "ERROR: no 001 field"
'                WriteLog GL.Logfile, .TextRaw
'                BibID = ""
'            End If
'
'        End With
'
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
