Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim InputFileNum As Integer
    Dim Line As String
    Dim SQL As String
    Dim BibID As Long
    Dim rs As Integer
    
    Dim i As Long
    For i = 5000001 To 5000050
    'WriteLog LogFile, GetVgerHolRecord("6909939").TextFormatted
        WriteLog GL.Logfile, GetVgerHolRecord(CStr(i)).TextFormatted
        SkeletonForm.lblStatus.Caption = "Processing " & i
        NiceSleep GL.Interval
    Next
    
'Or SQL-based:
'    SQL = "SELECT..."
'    rs = GL.GetRS
'    SkeletonForm.lblStatus.Caption = "Executing SQL..."
'    DoEvents
'    With GL.Vger
'        .ExecuteSQL SQL, rs
'        Do While .GetNextRow(rs)
'        BibID = .CurrentRow(rs, 1)
'        'do something with results
'            SkeletonForm.lblStatus.Caption = "Processing " & somevar
'            NiceSleep GL.Interval
'        Loop
'    End With
'    SkeletonForm.lblStatus.Caption = "Done!"
'    GL.FreeRS rs
    
'Or file-based:
'    InputFileNum = FreeFile
'    Open GL.InputFilename For Input As InputFileNum
'    Do While Not EOF(InputFileNum)
'        Line Input #InputFileNum, Line
'        'do something
'        NiceSleep GL.Interval
'    Loop
'    Close InputFileNum

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

End Sub
