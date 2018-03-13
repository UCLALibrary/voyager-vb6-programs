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
    
    Dim BibDelCnt As Long
    Dim HolDelCnt As Long
    Dim ItemDelCnt As Long
    
    SQL = GetTextFromFile(GL.InputFilename)

    HolRS = GL.GetRS
    ItemRS = GL.GetRS
    
    BibDelCnt = 0
    HolDelCnt = 0
    ItemDelCnt = 0
    
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    ' For each holdings record in HolRS, try to retrieve & delete items (if any), then the holdings record, then the bib record
    ' Looks only at holdings records found via HolRS SQL, so if other (non-qualifying) holdings exist, bib cannot be deleted
    With GL.Vger
        .ExecuteSQL SQL, HolRS
        Do While .GetNextRow(HolRS)
            BibID = CLng(.CurrentRow(HolRS, 1))
            HolID = CLng(.CurrentRow(HolRS, 2))
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            .SearchItemNumbersForHold CStr(HolID), ItemRS
            Do While .GetNextRow(ItemRS)
                itemID = CLng(.CurrentRow(ItemRS, 1))
                ItemRC = GL.BatchCat.DeleteItemRecord(itemID)
                If ItemRC = diSuccess Then
                    WriteLog GL.Logfile, "Deleted ItemID " & itemID
                    ItemDelCnt = ItemDelCnt + 1
                Else
                    WriteLog GL.Logfile, "Error deleting ItemID " & itemID & " : " & TranslateItemDeleteCode(ItemRC)
                End If
            Loop
            
            HolRC = GL.BatchCat.DeleteHoldingRecord(HolID)
            If HolRC = dhSuccess Then
                WriteLog GL.Logfile, vbTab & "Deleted HolID " & HolID
                HolDelCnt = HolDelCnt + 1
            Else
                WriteLog GL.Logfile, vbTab & "Error deleting HolID " & HolID & " : " & TranslateHoldingsDeleteCode(HolRC)
                SuppressHoldings HolID
            End If
            
            BibRC = GL.BatchCat.DeleteBibRecord(BibID)
            If BibRC = dbSuccess Then
                WriteLog GL.Logfile, vbTab & vbTab & "Deleted BibID " & BibID
                BibDelCnt = BibDelCnt + 1
            Else
                WriteLog GL.Logfile, vbTab & vbTab & "Error deleting BibID " & BibID & " : " & TranslateBibDeleteCode(BibRC)
            End If
            
            WriteLog GL.Logfile, ""
            DoEvents
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    
    WriteLog GL.Logfile, "Deleted: " & BibDelCnt & " bibs, " & HolDelCnt & " hols, " & ItemDelCnt & " items"
    WriteLog GL.Logfile, ""
    
    GL.FreeRS ItemRS
    GL.FreeRS HolRS
    
End Sub

