Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim rs As Integer
    Dim BibID As Long
    Dim BibRecord As Utf8MarcRecordClass
    Dim BibRc As UpdateBibReturnCode
    
    '20090429 akohler: New logic: records with OCLC numbers but no UCOCLC entry at all - catch a few which might have fallen through the cracks
    '20091201 akohler: Modified SQL: check for 'OCOLC %' (space after OCOLC) to avoid FATA's (OCoLCIR) numbers
    SQL = _
        "select bib_id " & vbCrLf & _
        "from bib_index bi " & vbCrLf & _
        "where index_code = '0350' " & vbCrLf & _
        "and normal_heading like 'OCOLC %' " & vbCrLf & _
        "and not exists ( " & vbCrLf & _
          "select *  " & vbCrLf & _
          "from bib_index  " & vbCrLf & _
          "where bib_id = bi.bib_id  " & vbCrLf & _
          "and index_code = '0350'  " & vbCrLf & _
          "and normal_heading like 'UCOCLC%' " & vbCrLf & _
        ") " & vbCrLf & _
        "order by bib_id "

    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            BibID = CLng(.CurrentRow(rs, 1))
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            Set BibRecord = UpdateUcoclc(GetVgerBibRecord(CStr(BibID)))
            BibRc = GL.BatchCat.UpdateBibRecord(BibID, BibRecord.MarcRecordOut, GL.Vger.BibUpdateDateVB, GL.Vger.BibOwningLibraryNumber, GL.CatLocID, GL.Vger.BibRecordIsSuppressed)
            If BibRc = ubSuccess Then
                WriteLog GL.Logfile, "Updated bib# " & BibID
            Else
                WriteLog GL.Logfile, "Error updating bib# " & BibID & " - returncode: " & BibRc
            End If
    
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
    
End Sub

Private Sub TestRE()
    Dim re As RegExp
    Dim str As String
    Dim Match As Match
    Dim Matches As MatchCollection
    Dim matchno As Integer
    Dim submatch As Variant
    Dim tests As Variant
    Dim test As Variant
    
    
    Set re = New RegExp
    'RE.pattern = "(^\(OCoLC\)|)(oc[mn]|ocl7|)(0*)(\d+)" 'from perl; can't get the | to handle the optional parts correctly
    re.Pattern = "(^\(OCoLC\)|\(OCoLC\)ocl7|\(OCoLC\)ocm|\(OCoLC\)ocn|oc[mn]|ocl7)(0*)(\d+)"
    Debug.Print "Pattern: " & re.Pattern & vbCrLf
    tests = Array("(OCoLC)123456789", "ocm12345678", "(OCoLC)ocm87654321", "(OCoLC)ocl7123456", "ocl7654321", "(OCoLC)00012345", "(SfxObjID)222222222222")
   
    For Each test In tests
        Debug.Print "Testing: " & test
        Set Matches = re.Execute(test)
        Debug.Print "Found matches: " & Matches.Count
        For Each Match In Matches
            With Match
                Debug.Print vbTab & "FirstIndex: " & .FirstIndex
                Debug.Print vbTab & "Length: " & .length
                Debug.Print vbTab & "Value: " & .Value
                Debug.Print vbTab & "SubMatches: " & .SubMatches.Count
                For Each submatch In .SubMatches
                    Debug.Print vbTab & vbTab & submatch
                Next
            End With
        Next
        Debug.Print vbCrLf
    Next
    
End Sub
