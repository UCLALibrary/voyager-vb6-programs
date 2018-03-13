Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    
    Dim HolID As Long
    Dim itemID As Long
    
    Dim rs As Integer
    Dim ItemRS As Integer
    
    Dim HolRC As DeleteHoldingReturnCode
    Dim ItemRC As DeleteItemReturnCode
    
    SQL = GetTextFromFile(GL.InputFilename)
    
    rs = GL.GetRS
    ItemRS = GL.GetRS
    
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        Do While .GetNextRow(rs)
            HolID = CLng(.CurrentRow(rs, 1))
            itemID = CLng(.CurrentRow(rs, 2))
            SkeletonForm.lblStatus.Caption = "Processing item " & itemID
            
            WriteLog GL.Logfile, "HolID: " & HolID
            
            ItemRC = GL.BatchCat.DeleteItemRecord(itemID)
            If ItemRC = diSuccess Then
                WriteLog GL.Logfile, vbTab & "Deleted ItemID " & itemID
            Else
                .RetrieveItemRecord CStr(itemID)
                WriteLog GL.Logfile, vbTab & "Error deleting ItemID " & itemID & " (" & .ItemBarcode & ") : " & _
                    TranslateItemDeleteCode(ItemRC)
            End If
            
            'Are there still items on this holdings record?
            'If not, delete the holdings record - but not for Management.
            'Based on name of input file.
            If InStr(1, GL.InputFilename, "DeletePCCP_mgmt.sql", vbTextCompare) > 0 Then
                WriteLog GL.Logfile, "MGMT: HolID " & HolID & " not deleted"
            Else
                .SearchItemNumbersForHold CStr(HolID), ItemRS
                If Not .GetNextRow(ItemRS) Then
                    'Delete the holdings record
                    HolRC = GL.BatchCat.DeleteHoldingRecord(HolID)
                    If HolRC = dhSuccess Then
                        WriteLog GL.Logfile, "Deleted HolID " & HolID
                    Else
                        WriteLog GL.Logfile, "Error deleting HolID " & HolID & " : " & TranslateHoldingsDeleteCode(HolRC)
                        If HolRC = dhLineItemsCopyAttached Then
                            UpdatePCCPHoldings HolID
                        End If
                    End If
                End If
            End If
    
            'Get ready for the next one
            WriteLog GL.Logfile, ""
            DoEvents
            NiceSleep GL.Interval
        
        Loop 'next item record
    End With
    
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs
    GL.FreeRS ItemRS
End Sub

Public Sub UpdatePCCPHoldings(HolID As Long)
    'Holdings record can't be deleted, so update it as requested by Claudia Horning
    
    Dim HolRecord As Utf8MarcRecordClass
    Dim HolRC As UpdateHoldingReturnCode
    
    Set HolRecord = GetVgerHolRecord(CStr(HolID))
    With HolRecord
        '* Delete subfield and contents of 852 $h, $i, and $k
        '* Add new 852 $x as follows: Powell Weeding Project, items withdrawn, PO attached, holdings record suppressed [YYYYMMDD]
        .FldFindFirst "852"
        If .FldWasFound Then
            .SfdFindFirst "h"
            If .SfdWasFound Then
                .SfdDelete
            End If
            .SfdFindFirst "i"
            If .SfdWasFound Then
                .SfdDelete
            End If
            .SfdFindFirst "k"
            If .SfdWasFound Then
                .SfdDelete
            End If
            
            .SfdAdd "x", "Powell Weeding Project, items withdrawn, PO attached, holdings record suppressed " & Format(Now(), "yyyymmdd")
        End If
        
        '* If there is an 866-868 field, delete it.
        .FldFindFirst "866"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        .FldFindFirst "867"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        .FldFindFirst "868"
        Do While .FldWasFound
            .FldDelete
            .FldFindNext
        Loop
        
        '* Save changes and suppress the holdings record
        With GL.Vger
            HolRC = GL.BatchCat.UpdateHoldingRecord(HolID, HolRecord.MarcRecordOut, .HoldUpdateDateVB, GL.CatLocID, .HoldBibRecordNumber, .HoldLocationID, True)
            If HolRC = uhSuccess Then
                WriteLog GL.Logfile, "Updated and suppressed HolID: " & HolID
            Else
                WriteLog GL.Logfile, "ERROR updating HolID: " & HolID & " : " & HolRC
            End If
        End With
    End With
End Sub
