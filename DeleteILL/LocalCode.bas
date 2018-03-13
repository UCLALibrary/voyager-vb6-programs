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
    
    Dim ItemRS As Integer
    
    Dim BibRC As DeleteBibReturnCode
    Dim HolRC As DeleteHoldingReturnCode
    Dim ItemRC As DeleteItemReturnCode
    
    Dim BibDelCnt As Long
    Dim HolDelCnt As Long
    Dim ItemDelCnt As Long
    
    SQL = GetTextFromFile(GL.InputFilename)

    ItemRS = GL.GetRS
    
    BibDelCnt = 0
    HolDelCnt = 0
    ItemDelCnt = 0
    
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    'ItemRS contains item ids (and their bib & mfhd ids) to be deleted.
    'Try to delete item, then mfhd, then bib
    With GL.Vger
        .ExecuteSQL SQL, ItemRS
        Do While .GetNextRow(ItemRS)
            BibID = CLng(.CurrentRow(ItemRS, 1))
            HolID = CLng(.CurrentRow(ItemRS, 2))
            itemID = CLng(.CurrentRow(ItemRS, 3))
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            WriteLog GL.Logfile, "BibID: " & BibID
            WriteLog GL.Logfile, vbTab & "HolID: " & HolID
            ItemRC = GL.BatchCat.DeleteItemRecord(itemID)
            If ItemRC = diSuccess Then
                WriteLog GL.Logfile, vbTab & vbTab & "Deleted ItemID " & itemID
                ItemDelCnt = ItemDelCnt + 1
            Else
                WriteLog GL.Logfile, vbTab & vbTab & "Error deleting ItemID " & itemID & " : " & TranslateItemDeleteCode(ItemRC)
            End If
            
            HolRC = GL.BatchCat.DeleteHoldingRecord(HolID)
            If HolRC = dhSuccess Then
                WriteLog GL.Logfile, vbTab & "Deleted HolID " & HolID
                HolDelCnt = HolDelCnt + 1
            Else
                WriteLog GL.Logfile, vbTab & "Error deleting HolID " & HolID & " : " & TranslateHoldingsDeleteCode(HolRC)
            End If
            
            BibRC = GL.BatchCat.DeleteBibRecord(BibID)
            If BibRC = dbSuccess Then
                WriteLog GL.Logfile, "Deleted BibID " & BibID
                BibDelCnt = BibDelCnt + 1
            Else
                WriteLog GL.Logfile, "Error deleting BibID " & BibID & " : " & TranslateBibDeleteCode(BibRC)
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
    
End Sub


'Public Sub RunLocalCode()
'    'This public procedure is called from SkeletonForm
'    'It controls what happens for most projects
'    'Global (GL) init/termination handled on SkeletonForm
'    Dim SQL As String
'
'    Dim BibID As Long
'    Dim HolID As Long
'    Dim itemID As Long
'
'    Dim HolRS As Integer
'    Dim ItemRS As Integer
'
'    Dim BibRC As DeleteBibReturnCode
'    Dim HolRC As DeleteHoldingReturnCode
'    Dim ItemRC As DeleteItemReturnCode
'
'    Dim BibDelCnt As Long
'    Dim HolDelCnt As Long
'    Dim ItemDelCnt As Long
'
''*** BE SURE TO USE THE RIGHT SQL ***'
'    SQL = _
'        "SELECT  " & vbCrLf & _
'            "BT.Bib_ID " & vbCrLf & _
'        ",  MM.MFHD_ID " & vbCrLf & _
'        "--,    L.Location_Code " & vbCrLf & _
'        "--,    BT.Title " & vbCrLf & _
'        "FROM Bib_Text BT " & vbCrLf & _
'        "INNER JOIN Bib_MFHD BM " & vbCrLf & _
'            "ON BT.Bib_ID = BM.Bib_ID " & vbCrLf & _
'        "INNER JOIN MFHD_Master MM " & vbCrLf & _
'            "ON BM.MFHD_ID = MM.MFHD_ID " & vbCrLf & _
'        "INNER JOIN Location L " & vbCrLf & _
'            "ON MM.Location_ID = L.Location_ID " & vbCrLf & _
'        "WHERE L.Location_Code IN ('biill', 'lwill', 'mgill', 'yrill') " & vbCrLf & _
'        "--This EXISTS clause is much faster than using BH as base table instead of BT " & vbCrLf & _
'        "AND EXISTS " & vbCrLf & _
'        "(  SELECT * " & vbCrLf & _
'            "FROM Bib_History " & vbCrLf & _
'            "WHERE Bib_ID = BT.Bib_ID " & vbCrLf & _
'            "AND Action_Date < To_Date(Sysdate - 60) " & vbCrLf & _
'            "AND Action_Type_ID = 1 --create " & vbCrLf & _
'        ") " & vbCrLf & _
'        "ORDER BY BT.Bib_ID "
''*** BE SURE TO USE THE RIGHT SQL ***'
'
'    HolRS = GL.GetRS
'    ItemRS = GL.GetRS
'
'    BibDelCnt = 0
'    HolDelCnt = 0
'    ItemDelCnt = 0
'
'    SkeletonForm.lblStatus.Caption = "Executing SQL..."
'    DoEvents
'    'To be sure only ILL (location-based) records are deleted, iterate through MFHDs & delete (when possible)
'    'There shouldn't be non-ILL MFHDs on these records, but one never knows....
'    'Then, after deleting appropriate MFHDs, try to delete the bib.
'    With GL.Vger
'        .ExecuteSQL SQL, HolRS
'        Do While .GetNextRow(HolRS)
'            BibID = CLng(.CurrentRow(HolRS, 1))
'            HolID = CLng(.CurrentRow(HolRS, 2))
'            SkeletonForm.lblStatus.Caption = "Processing " & BibID
'            WriteLog GL.Logfile, "BibID: " & BibID
'            WriteLog GL.Logfile, vbTab & "HolID: " & HolID
'            .SearchItemNumbersForHold CStr(HolID), ItemRS
'            Do While .GetNextRow(ItemRS)
'                itemID = CLng(.CurrentRow(ItemRS, 1))
'                'WriteLog GL.Logfile, vbTab & vbTab & "ItemID: " & ItemID
'                ItemRC = GL.BatchCat.DeleteItemRecord(itemID)
'                If ItemRC = diSuccess Then
'                    WriteLog GL.Logfile, vbTab & vbTab & "Deleted ItemID " & itemID
'                    ItemDelCnt = ItemDelCnt + 1
'                Else
'                    WriteLog GL.Logfile, vbTab & vbTab & "Error deleting ItemID " & itemID & " : " & TranslateItemDeleteCode(ItemRC)
'                End If
'            Loop
'
'            HolRC = GL.BatchCat.DeleteHoldingRecord(HolID)
'            If HolRC = dhSuccess Then
'                WriteLog GL.Logfile, vbTab & "Deleted HolID " & HolID
'                HolDelCnt = HolDelCnt + 1
'            Else
'                WriteLog GL.Logfile, vbTab & "Error deleting HolID " & HolID & " : " & TranslateHoldingsDeleteCode(HolRC)
'            End If
'
'            BibRC = GL.BatchCat.DeleteBibRecord(BibID)
'            If BibRC = dbSuccess Then
'                WriteLog GL.Logfile, "Deleted BibID " & BibID
'                BibDelCnt = BibDelCnt + 1
'            Else
'                WriteLog GL.Logfile, "Error deleting BibID " & BibID & " : " & TranslateBibDeleteCode(BibRC)
'            End If
'
'            WriteLog GL.Logfile, ""
'            DoEvents
'            NiceSleep GL.Interval
'        Loop
'    End With
'    SkeletonForm.lblStatus.Caption = "Done!"
'
'    WriteLog GL.Logfile, "Deleted: " & BibDelCnt & " bibs, " & HolDelCnt & " hols, " & ItemDelCnt & " items"
'    WriteLog GL.Logfile, ""
'
'    GL.FreeRS ItemRS
'    GL.FreeRS HolRS
'
'End Sub
'
'
