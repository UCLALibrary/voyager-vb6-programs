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
    
    Dim BibRS As Integer
    Dim HolRS As Integer
    Dim ItemRS As Integer
    
    Dim BibRC As DeleteBibReturnCode
    Dim HolRC As DeleteHoldingReturnCode
    Dim ItemRC As DeleteItemReturnCode
    
    SQL = GetTextFromFile(GL.InputFilename)

    BibRS = GL.GetRS
    HolRS = GL.GetRS
    ItemRS = GL.GetRS
    
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, BibRS
        Do While .GetNextRow(BibRS)
            BibID = CLng(.CurrentRow(BibRS, 1))
            WriteLog GL.Logfile, "BibID: " & BibID
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            .SearchHoldNumbersForBib CStr(BibID), HolRS
            Do While .GetNextRow(HolRS)
                HolID = CLng(.CurrentRow(HolRS, 1))
                WriteLog GL.Logfile, vbTab & "HolID: " & HolID
                .SearchItemNumbersForHold CStr(HolID), ItemRS
                Do While .GetNextRow(ItemRS)
                    itemID = CLng(.CurrentRow(ItemRS, 1))
                    'WriteLog GL.Logfile, vbTab & vbTab & "ItemID: " & ItemID
                    ItemRC = GL.BatchCat.DeleteItemRecord(itemID)
                    If ItemRC = diSuccess Then
                        WriteLog GL.Logfile, vbTab & vbTab & "Deleted ItemID " & itemID
                    Else
                        WriteLog GL.Logfile, vbTab & vbTab & "Error deleting ItemID " & itemID & " : " & TranslateItemDeleteCode(ItemRC)
                    End If
                Loop
                HolRC = GL.BatchCat.DeleteHoldingRecord(HolID)
                If HolRC = dhSuccess Then
                    WriteLog GL.Logfile, vbTab & "Deleted HolID " & HolID
                Else
                    WriteLog GL.Logfile, vbTab & "Error deleting HolID " & HolID & " : " & TranslateHoldingsDeleteCode(HolRC)
                End If
            Loop
            BibRC = GL.BatchCat.DeleteBibRecord(BibID)
            If BibRC = dbSuccess Then
                WriteLog GL.Logfile, "Deleted BibID " & BibID
            Else
                WriteLog GL.Logfile, "Error deleting BibID " & BibID & " : " & TranslateBibDeleteCode(BibRC)
            End If
            WriteLog GL.Logfile, ""
            DoEvents
            NiceSleep GL.Interval
        Loop
    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    
    GL.FreeRS ItemRS
    GL.FreeRS HolRS
    GL.FreeRS BibRS
    
End Sub
