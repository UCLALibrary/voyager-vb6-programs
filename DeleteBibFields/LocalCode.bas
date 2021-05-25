Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    Dim BibID As Long
    Dim PrevBibID As Long
    Dim Tag As String
    Dim Indicators As String
    Dim Field As String
    Dim rs As Integer
    Dim BibRecord As Utf8MarcRecordClass
    Dim BibRC As UpdateBibReturnCode
    Dim DeleteHoldings As Boolean
    
    'Case-Insensitive (CI) which could trigger delete
    Dim DeleteStringsCI_Vals As String
    Dim DeleteStringsCI() As String
    'Case-Sensitive (CS) which could trigger delete
    Dim DeleteStringsCS_Vals As String
    Dim DeleteStringsCS() As String
    'Case-Insensitive (CI) which could prevent delete
    'No need for CS keeps
    Dim KeepStringsCI_Vals As String
    Dim KeepStringsCI() As String
    
    DeleteStringsCI_Vals = "Publisher description,Publisher's description,Publication information"
    DeleteStringsCI() = Split(DeleteStringsCI_Vals, ",")
    DeleteStringsCS_Vals = ""
    DeleteStringsCS() = Split(DeleteStringsCS_Vals, ",")
    KeepStringsCI_Vals = ""
    KeepStringsCI() = Split(KeepStringsCI_Vals, ", ")
    
    SQL = GetTextFromFile(GL.InputFilename)
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            BibID = .CurrentRow(rs, 1)
            Tag = .CurrentRow(rs, 2)
            'Indicators = .CurrentRow(rs, 3)
            
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            
            Set BibRecord = GetVgerBibRecord(CStr(BibID))
            With BibRecord
                Field = ""
                .FldFindFirst Tag
                Do While .FldWasFound
                    Field = Trim(Replace(.FldText, Chr(31), " $"))
                    If (FieldContainsStrings(Field, DeleteStringsCI, False) = True Or FieldContainsStrings(Field, DeleteStringsCS, True) = True) _
                    And FieldContainsStrings(Field, KeepStringsCI, False) = False Then
                        WriteLog GL.Logfile, "Deleting: " & vbTab & BibID & vbTab & Field
                        .FldDelete
                    Else
                        'For testing
                        'WriteLog GL.Logfile, "KEEPING: " & vbTab & BibID & vbTab & Field
                    End If
                    .FldFindNext
                Loop
            End With
            
            BibRC = GL.BatchCat.UpdateBibRecord( _
                CLng(BibID), _
                BibRecord.MarcRecordOut, _
                GL.Vger.BibUpdateDateVB, _
                GL.Vger.BibOwningLibraryNumber, _
                GL.CatLocID, _
                GL.Vger.BibRecordIsSuppressed _
                )
            If BibRC = ubSuccess Then
                WriteLog GL.Logfile, "Updated bib " & BibID
            Else
                WriteLog GL.Logfile, "ERROR updating bib " & BibID & " : return code " & BibRC
            End If

            'Check just 856 fields - in case we do use generic tag logic above for field deletion
            With BibRecord
                .FldFindFirst "856"
                If .FldWasFound = False Then
                    'No 856 fields, delete internet holdings
                    DeleteHoldings = True
                Else
                    'Start true, change to false if warranted
                    DeleteHoldings = True
                    Do While .FldWasFound
                        If .FldInd <> "42" Then
                            WriteLog GL.Logfile, "DEBUG NO DELETE: " & .FldInd & vbTab & .FldTextFormatted
                            DeleteHoldings = False
                            Exit Do
                        End If
                        .FldFindNext
                    Loop
                End If
                
                If DeleteHoldings = True Then
                    DeleteInternetHoldings BibID
                End If
                End With

            NiceSleep GL.Interval
        Loop
    End With
    GL.FreeRS rs
    SkeletonForm.lblStatus.Caption = "Done!"
    
End Sub

Private Function FieldContainsStrings(Field As String, StringArray() As String, CaseSensitive As Boolean) As Boolean
    Dim ArrayPos As Integer
    Dim Term As String
    Dim Found As Boolean
    Dim Compare As Integer
    
    If CaseSensitive = True Then
        Compare = vbBinaryCompare
    Else
        Compare = vbTextCompare
    End If
    
    Found = False
    For ArrayPos = 0 To UBound(StringArray)
        Term = StringArray(ArrayPos)
        If InStr(1, Field, Term, Compare) Then
            'Found it, no need to continue checking terms
            Found = True
            Exit For
        End If
    Next
    FieldContainsStrings = Found
End Function

Private Sub DeleteInternetHoldings(BibID As Long)
    'Depends on caller to decide whether holdings *should* be deleted....
    
    Dim rs As Integer
    Dim rc As DeleteHoldingReturnCode
    Dim HolID As Long
    
    rs = GL.GetRS
    With GL.Vger
        .SearchHoldNumbersForBib CStr(BibID), rs
        Do While .GetNextRow(rs)
            .RetrieveHoldRecord .CurrentRow(1)
            If .HoldLocationCode = "in" Then
                HolID = .HoldRecordNumber
                rc = GL.BatchCat.DeleteHoldingRecord(HolID)
                If rc = dhSuccess Then
                    WriteLog GL.Logfile, vbTab & "Deleted internet hol " & HolID
                Else
                    WriteLog GL.Logfile, vbTab & "ERROR deleting internet hol " & HolID & " : " & TranslateHoldingsDeleteCode(rc)
                End If
            End If
        Loop
    End With
    GL.FreeRS rs

End Sub
