Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim BibRS As Integer
    
    Dim SQL As String
    Dim BibRecord As Utf8MarcRecordClass
    Dim BibID As Long
    Dim PrevBibID As Long
    
    Dim ItemBarcode As String
    Dim ItemEnum As String
    Dim ItemCaption As String
    Dim ItemLoc As String
    Dim ItemOwner As String
    Dim ItemType As String
    Dim ItemStatus As String
    Dim ItemStatuses() As String
    Dim Fld976 As String
    Dim OutFile As String
    Dim cnt As Integer
    
    SQL = GetTextFromFile(GL.InputFilename)
    OutFile = GL.BaseFilename & ".mrc"
    PrevBibID = 0
    BibRS = GL.GetRS
    
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, BibRS
        Do While .GetNextRow(BibRS)
            BibID = .CurrentRow(BibRS, 1)
            ItemBarcode = .CurrentRow(BibRS, 2)
            ItemEnum = .CurrentRow(BibRS, 3)
            ItemCaption = .CurrentRow(BibRS, 4)
            ItemLoc = .CurrentRow(BibRS, 5)
            ItemOwner = .CurrentRow(BibRS, 6)
            ItemType = .CurrentRow(BibRS, 7)
            'ItemStatus = .CurrentRow(BibRS, 8)
            ItemStatuses = Split(.CurrentRow(BibRS, 8), ", ")
            
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            DoEvents
            
            'Start of data for new record: write the previous record, and fetch the next bib from Voyager
            If BibID <> PrevBibID Then
                'Avoid fake starter bib
                If PrevBibID > 0 Then
                    WriteRawRecord OutFile, BibRecord.MarcRecordOut
WriteLog GL.Logfile, vbCrLf & BibRecord.TextFormatted
                End If
                Set BibRecord = GetVgerBibRecord(CStr(BibID))
                PrevBibID = BibID
            End If
            
            'Add item data to a new 976 field
            'First assemble the field by adding subfields (where data is present); barcode should always exist
            With BibRecord
                Fld976 = .SfdMake("a", ItemBarcode)
                If ItemEnum <> "" Then Fld976 = Fld976 & .SfdMake("b", ItemEnum)
                If ItemCaption <> "" Then Fld976 = Fld976 & .SfdMake("c", ItemCaption)
                If ItemLoc <> "" Then Fld976 = Fld976 & .SfdMake("d", ItemLoc)
                If ItemOwner <> "" Then Fld976 = Fld976 & .SfdMake("e", ItemOwner)
                If ItemType <> "" Then Fld976 = Fld976 & .SfdMake("f", ItemType)
                'TODO: Items can have multiple status; put each in its own subfield
                For cnt = 0 To UBound(ItemStatuses)
                    Fld976 = Fld976 & .SfdMake("g", ItemStatuses(cnt))
                Next
                'If ItemStatus <> "" Then Fld976 = Fld976 & .SfdMake("g", ItemStatus)
                .FldAddGeneric "976", "  ", Fld976, 3
            End With
            
            NiceSleep GL.Interval
        Loop 'BibRS
        'Finished final record, which still needs to be written
        WriteRawRecord OutFile, BibRecord.MarcRecordOut
WriteLog GL.Logfile, vbCrLf & BibRecord.TextFormatted

    End With
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS BibRS
    
End Sub

