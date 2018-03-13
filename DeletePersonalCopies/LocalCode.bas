Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    
    Dim BibID As Long
    Dim HolID As Long
    Dim itemID As Long
    
    Dim HolRS As Integer
    Dim ItemRS As Integer
    
    Dim BibRC As DeleteBibReturnCode
    Dim HolRC As DeleteHoldingReturnCode
    Dim ItemRC As DeleteItemReturnCode
    
' ('arrsprscp', 'birsprscp', 'clrsprscp', 'mgrsprscp', 'mursprscp', 'scrsprscp', 'sgrsprscp', 'smrsprscp', 'yrrsprscp')
'    SQL = _
'        "SELECT " & vbCrLf & _
'            "bm.bib_id " & vbCrLf & _
'        ",  bm.mfhd_id " & vbCrLf & _
'        "FROM mfhd_master mm " & vbCrLf & _
'        "INNER JOIN bib_mfhd bm ON mm.mfhd_id = bm.mfhd_id " & vbCrLf & _
'        "INNER JOIN location l ON mm.location_id = l.location_id " & vbCrLf & _
'        "WHERE l.location_code LIKE '%prscp' " & vbCrLf & _
'        "AND NOT EXISTS (SELECT * FROM mfhd_history WHERE mfhd_id = mm.mfhd_id AND action_date > To_Date('20090715', 'YYYYMMDD')) " & vbCrLf & _
'        "ORDER BY bm.mfhd_id " & vbCrLf
    
    SQL = GetTextFromFile(GL.InputFilename)
    
    HolRS = GL.GetRS
    ItemRS = GL.GetRS
    
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, HolRS
        'Iterate through the holdings, deleting each record's items, then the holdings, then the bib (if possible)
        Do While .GetNextRow(HolRS)
            BibID = CLng(.CurrentRow(HolRS, 1))
            HolID = CLng(.CurrentRow(HolRS, 2))
            SkeletonForm.lblStatus.Caption = "Processing " & HolID
            
            WriteLog GL.Logfile, "BibID: " & BibID
            WriteLog GL.Logfile, vbTab & "HolID: " & HolID
            
            'Get and delete the items
            .SearchItemNumbersForHold CStr(HolID), ItemRS
            Do While .GetNextRow(ItemRS)
                itemID = CLng(.CurrentRow(ItemRS, 1))
                ItemRC = GL.BatchCat.DeleteItemRecord(itemID)
                If ItemRC = diSuccess Then
                    WriteLog GL.Logfile, vbTab & vbTab & "Deleted ItemID " & itemID
                Else
                    .RetrieveItemRecord CStr(itemID)
                    WriteLog GL.Logfile, vbTab & vbTab & "Error deleting ItemID " & itemID & " (" & .ItemBarcode & ") : " & _
                        TranslateItemDeleteCode(ItemRC)
                End If
            Loop
            
            'Delete the holdings record
            HolRC = GL.BatchCat.DeleteHoldingRecord(HolID)
            If HolRC = dhSuccess Then
                WriteLog GL.Logfile, vbTab & "Deleted HolID " & HolID
            Else
                WriteLog GL.Logfile, vbTab & "Error deleting HolID " & HolID & " : " & TranslateHoldingsDeleteCode(HolRC)
            End If
    
            'Delete the bib record
            BibRC = GL.BatchCat.DeleteBibRecord(BibID)
            If BibRC = dbSuccess Then
                WriteLog GL.Logfile, "Deleted BibID " & BibID
            Else
                WriteLog GL.Logfile, "Error deleting BibID " & BibID & " : " & TranslateBibDeleteCode(BibRC)
            End If
        
            'Get ready for the next one
            WriteLog GL.Logfile, ""
            DoEvents
            NiceSleep GL.Interval
        
        Loop 'next holdings record
    End With
    
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS ItemRS
    GL.FreeRS HolRS
End Sub
