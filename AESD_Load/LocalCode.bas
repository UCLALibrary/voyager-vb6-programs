Attribute VB_Name = "LocalCode"
Option Explicit

Public Sub RunLocalCode()
    'This public procedure is called from SkeletonForm
    'It controls what happens for most projects
    'Global (GL) init/termination handled on SkeletonForm
    
    Dim BibID As String
    Dim SourceFile As Utf8MarcFileClass
    Dim DestFile As Integer 'File handle
    Dim BibRecord As Utf8MarcRecordClass
    Dim RawRecord As String
    Dim UCLA920 As Boolean

    DestFile = FreeFile
    Open GL.BaseFilename + ".fixed" For Binary As DestFile
    
    Set SourceFile = New Utf8MarcFileClass
    SourceFile.OpenFile GL.InputFilename

    Do While SourceFile.ReadNextRecord(RawRecord)
        DoEvents
        Set BibRecord = New Utf8MarcRecordClass
        With BibRecord
            'convert to Unicode for Voyager
            .CharacterSetIn = "U"
            .CharacterSetOut = "U"
            .IgnoreSfdOrder = True
            'THIS DOES NOTHING.....
            '.SetFldOrder "B", "000", "999", 9 'SORTINSTRUCTION_LeaveOrderAlone%
            .MarcRecordIn = RawRecord

            .FldFindFirst "001"
            If .FldWasFound Then
                BibID = Trim(.FldText)
            Else
                WriteLog GL.Logfile, "ERROR: no 001 field"
                WriteLog GL.Logfile, .TextRaw
                BibID = ""
            End If
            
            SkeletonForm.lblStatus.Caption = "Processing " & BibID
            
            'Use 920 to create 599 for SCP loading
            UCLA920 = False
            .FldFindFirst "920"
            Do While .FldWasFound
                .SfdFindFirst "a"
                If .SfdWasFound Then
                    If .SfdText = "UCLA" Then
                        UCLA920 = True
                    End If
                End If
                .FldFindNext
            Loop
            If UCLA920 = True Then
                .FldAddGeneric "599", "  ", .SfdMake("a", "UPD")
            Else
                .FldAddGeneric "599", "  ", .SfdMake("a", "NEW")
            End If
            
            'Strip out 793 fields with "Agricultural & environmental science database online journals"
            '2nd attempt: strip out *ALL* 793 fields
            .FldFindFirst "793"
            Do While .FldWasFound
'                .SfdFindFirst "a"
'                If .SfdWasFound Then
'                    If InStr(1, .SfdText, "Agricultural & environmental science database online journals", vbTextCompare) > 0 Then
'                        .FldDelete
'                    End If
'                End If
                .FldDelete
                .FldFindNext
            Loop
            
            '3) Modify AESD 856 fields:
            '   1. Change $3 to "Available issues" (serials only)
            '   2. Change $z to "Restricted to UCLA" (all)
            '   3. Change $x to "UCLA" (all)
            .FldFindFirst "856"
            Do While .FldWasFound
                .SfdFindFirst "z"
                If .SfdWasFound Then
                    If InStr(1, .SfdText, "Agricultural & environmental science database", vbTextCompare) > 0 Then
                        'Change this subfield, and others in the same field
                        .SfdText = "Restricted to UCLA"
                        'Change $3 (serials only)
                        If IsSerial(.GetLeaderValue(7, 1)) Then
                            .SfdFindFirst "3"
                            If .SfdWasFound Then
                                .SfdText = "Available issues"
                            End If
                        End If
                        'Change $x
                        .SfdFindFirst "x"
                        If .SfdWasFound Then
                            .SfdText = "UCLA"
                        Else
                            .SfdAdd "x", "UCLA"
                        End If
                    End If
                End If
                .FldFindNext
            Loop
            
            '4) Delete all 856 fields which now lack $x UCLA
            .FldFindFirst "856"
            Do While .FldWasFound
                .SfdFindFirst "x"
                If .SfdWasFound Then
                    If .SfdText <> "UCLA" Then
                        .FldDelete
                    End If
                Else
                    'no $x at all
                    .FldDelete
                End If
                .FldFindNext
            Loop
            
            'Write the updated record
            Put #DestFile, , .MarcRecordOut
            
        End With

        NiceSleep GL.Interval
    Loop

    SourceFile.CloseFile
    Close #DestFile
    SkeletonForm.lblStatus.Caption = "Done!"
End Sub
