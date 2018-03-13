Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    Dim SQL As String
    'Dim OldBibID As Long
    Dim CurrentBibID As Long
    Dim OldBibID As Long
    Dim rs As Integer
    Dim BibRecord As Utf8MarcRecordClass
    Dim Fld950 As String
    Dim Barcode As String
    Dim itemID As String
    Dim Location As String
    Dim CallNumber As String
    Dim ItemEnum As String
    Dim InternetArchiveID As String
    Dim ArkID As String
    Dim outfile As String
    
    SQL = GetTextFromFile(GL.BaseFilename & ".sql")
    outfile = GL.BaseFilename & ".mrc"
    rs = GL.GetRS
    SkeletonForm.lblStatus.Caption = "Executing SQL..."
    DoEvents
    With GL.Vger
        .ExecuteSQL SQL, rs
        OldBibID = 0
        Do While .GetNextRow(rs)
            CurrentBibID = .CurrentRow(rs, 1)
            Barcode = .CurrentRow(rs, 2)
            itemID = .CurrentRow(rs, 3)
            Location = .CurrentRow(rs, 4)
            CallNumber = .CurrentRow(rs, 5)
            ItemEnum = .CurrentRow(rs, 6)
            'Item sequence number is column 7 in query; not used in this program, only used in query for sorting
            
            'Somewhat fragile, probably should check names of column headers
            If .ColumnCount = 9 Then
                InternetArchiveID = .CurrentRow(rs, 8)  'Used only for Internet Archive extracts
                ArkID = .CurrentRow(rs, 9)              'Used only for Internet Archive extracts
            End If
            
            SkeletonForm.lblStatus.Caption = "Processing " & CurrentBibID & " (" & itemID & ")"
            DoEvents
            
            'Only retrieve bib from Voyager once
            If CurrentBibID <> OldBibID Then
                OldBibID = CurrentBibID

                'Get the new record
                Set BibRecord = GetVgerBibRecord(CStr(CurrentBibID))
            End If
                
            With BibRecord
                'Remove any pre-existing 950 fields from Voyager (or newly added one from previous iteration)
                .FldFindFirst "950"
                Do While .FldWasFound
                    .FldDelete
                    .FldFindNext
                Loop
            
                'Add holdings/item information for current item in new 950 field
                Fld950 = .SfdMake("b", Location)
                If ItemEnum <> "" Then
                    Fld950 = Fld950 & .SfdMake("c", ItemEnum)
                End If
                If CallNumber <> "" Then
                    Fld950 = Fld950 & .SfdMake("h", CallNumber)
                End If
                If Barcode <> "" Then
                    Fld950 = Fld950 & .SfdMake("i", Barcode)
                End If
                If itemID <> "" Then
                    Fld950 = Fld950 & .SfdMake("j", itemID)
                End If
                'Subfields used only for Internet Archive records - no data for Google records
                If InternetArchiveID <> "" Then
                    Fld950 = Fld950 & .SfdMake("p", InternetArchiveID)
                End If
                If ArkID <> "" Then
                    Fld950 = Fld950 & .SfdMake("q", ArkID)
                End If
                
                .FldAddGeneric "950", "  ", Fld950, 3
            End With 'Bibrecord
            
            'Write record
            WriteRawRecord outfile, BibRecord.MarcRecordOut
            
            NiceSleep GL.Interval
        Loop
    End With 'Vger
    SkeletonForm.lblStatus.Caption = "Done!"
    GL.FreeRS rs

End Sub

'20091214: temp copy of code used to embed all of a bib's items in the bib
'Apparently for IA, Paul Fogel wants each item embedded in its own copy of the bib, but saving this code for now just in case...
'    With GL.Vger
'        .ExecuteSQL SQL, rs
'        OldBibID = 0
'        Do While .GetNextRow(rs)
'            CurrentBibID = .CurrentRow(rs, 1)
'            Barcode = .CurrentRow(rs, 2)
'            Location = .CurrentRow(rs, 4)
'            CallNumber = .CurrentRow(rs, 5)
'            ItemEnum = .CurrentRow(rs, 6)
'
'            If CurrentBibID <> OldBibID Then
'                If OldBibID > 0 Then
'                    'Write old record to file
'                    WriteRawRecord outfile, BibRecord.MarcRecordOut
'                End If
'                OldBibID = CurrentBibID
'
'                'Get the new record
'                Set BibRecord = GetVgerBibRecord(CStr(CurrentBibID))
'WriteLog GL.Logfile, "Got new record " & CurrentBibID
'                SkeletonForm.lblStatus.Caption = "Processing " & CurrentBibID
'
'                'Remove any pre-existing 950 fields
'                With BibRecord
'                    .FldFindFirst "950"
'                    Do While .FldWasFound
'                        .FldDelete
'                    Loop
'                End With 'Bibrecord
'            End If
'
'            'Add holdings/item information in 950 field
'            With BibRecord
'                Fld950 = .SfdMake("b", Location)
'                If ItemEnum <> "" Then
'                    Fld950 = Fld950 & .SfdMake("c", ItemEnum)
'                End If
'                If CallNumber <> "" Then
'                    Fld950 = Fld950 & .SfdMake("h", CallNumber)
'                End If
'                If Barcode <> "" Then
'                    Fld950 = Fld950 & .SfdMake("i", Barcode)
'                End If
'                .FldAddGeneric "950", "  ", Fld950, 3
'            End With 'Bibrecord
'
'            NiceSleep GL.Interval
'        Loop
'        'Write final record
'        WriteRawRecord outfile, BibRecord.MarcRecordOut
'    End With 'Vger

