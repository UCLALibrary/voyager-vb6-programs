Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim OuterSQL As String
    Dim OuterRS As Integer
    
    Dim Code As String
    Dim Name As String
    Dim Url As String
    
    OuterSQL = _
        "SELECT code, name, url " & _
        "FROM vger_support.spac_map " & _
        "WHERE url is not null " & _
        "ORDER BY code"
    
    OuterRS = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL OuterSQL, OuterRS
        Do While .GetNextRow(OuterRS)
            Code = .CurrentRow(OuterRS, 1)
            Name = .CurrentRow(OuterRS, 2)
            Url = .CurrentRow(OuterRS, 3)
            
            SkeletonForm.lblStatus.Caption = "Processing " & Code
            ProcessHoldings Code, Name, Url
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS OuterRS
End Sub

Private Sub ProcessHoldings(Code As String, Name As String, Url As String)
    Dim InnerSQL As String
    Dim InnerRS As Integer
    Dim HolID As Long
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    Dim NeedsUpdate As Boolean
    Dim AcquiredNote As String
    Dim Fld856 As String
    
    InnerSQL = GetTextFromFile(GL.BaseFilename & "_TEMPLATE.sql")
    InnerSQL = Replace(InnerSQL, "_CODE_", Code, 1, -1, vbBinaryCompare)
    InnerSQL = Replace(InnerSQL, "_URL_", Url, 1, -1, vbBinaryCompare)
    
    InnerRS = GL.GetRS
    DoEvents
    With GL.Vger
        WriteLog GL.Logfile, "Processing code " & Code
        
        .ExecuteSQL InnerSQL, InnerRS
        
        Do While .GetNextRow(InnerRS)
            HolID = .CurrentRow(InnerRS, 1)
            Set HolRecord = GetVgerHolRecord(CStr(HolID))
            With HolRecord
                'Double-check existing 856 fields to reject records which already have this URL
                'Initial check done by InnerSQL above, but that depends on subfield database which may be out of date.
                NeedsUpdate = True
                .FldFindFirst "856"
                Do While .FldWasFound
                    .SfdFindFirst "u"
                    Do While .SfdWasFound
                        If .SfdText = Url Then
                            NeedsUpdate = False
                        End If
                        .SfdFindNext
                    Loop
                    .FldFindNext
                Loop
            
                If NeedsUpdate = True Then
                    ' 856 42 $3 Bookplate: $u [raw URL] $z Acquired as part of the [name of fund, as cited in the NAME column of the attached table]
                    AcquiredNote = "Acquired as part of the "
                    'Strip leading "The " from SPAC name, for consistent wording in AcquiredNote
                    If Left(Name, 4) = "The " Then
                        Name = Replace(Name, "The ", "", 1, 1, vbBinaryCompare)
                    End If
                    AcquiredNote = AcquiredNote & Name
                    Fld856 = _
                        .SfdMake("3", "Bookplate:") & _
                        .SfdMake("u", Url) & _
                        .SfdMake("z", AcquiredNote)
                    .FldAddGeneric "856", "42", Fld856
                
                    HolRC = GL.BatchCat.UpdateHoldingRecord( _
                        HolID, _
                        HolRecord.MarcRecordOut, _
                        GL.Vger.HoldUpdateDateVB, _
                        GL.CatLocID, _
                        GL.Vger.HoldBibRecordNumber, _
                        GL.Vger.HoldLocationID, _
                        GL.Vger.HoldRecordIsSuppressed _
                    )
                    If HolRC = uhSuccess Then
                        WriteLog GL.Logfile, vbTab & "Updated hol " & HolID
                    Else
                        WriteLog GL.Logfile, vbTab & "ERROR updating hol " & HolID & " : return code: " & HolRC
                    End If
                
                End If
            End With
            
            NiceSleep GL.Interval
        Loop
    End With

End Sub
