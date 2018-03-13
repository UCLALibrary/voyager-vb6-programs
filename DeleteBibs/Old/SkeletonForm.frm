VERSION 5.00
Begin VB.Form SkeletonForm 
   Caption         =   "VGER Skeleton"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "SkeletonForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DBUG = True
Private Const HOME As Boolean = False
Private Const UBERLOGMODE As Boolean = True

Private Const ERROR_BAR As String = "*** ERROR ***"
Private Const MAX_RECORD_COUNT As Integer = 2000

'Form-level globals - let's keep this list short!
Private BaseFilename As String      'InputFilename without the file extension; we'll use this a lot
Private BatchCat As ClassBatchCat   'For updating the database
Private InputFilename As String     'Name of file containing records to process
Private Interval As Long            'Number of milliseconds to delay between records (via cmdline)
Private LogFile As Integer          'File handle for log file, available to all routines
Private ProductionMode As Boolean   '
Private Vger As VgerReaderClass6    'For searching the database
Private rs1 As Integer              'Vger Resultset #1
Private rs2 As Integer              'Vger Resultset #2

Private Sub Form_Load()
    'Main handles everything
    Main
    'Exit from VB if running as program
    End
End Sub

Private Sub Main()
    'This is the controlling procedure for this form
    Dim User As String
    Dim Tablespace As String
    Dim InputFileNum As Integer
    Dim Line As String
    Dim BibID As Long
    Dim RC As DeleteBibReturnCode
    
    'Parse the command-line to get arguments
    ParseArguments
        
    If ProductionMode Then
        Tablespace = "UCLADB"
    Else
        Tablespace = "UCLA_TESTDB"
    End If
    
    If InputFilename = "" Then InputFilename = "c:\projects\voyager\programs\deletebibs\bib5.lst"
    
    If InputFilename <> "" Then
        BaseFilename = GetBaseFilename(InputFilename)
        LogFile = OpenLog(BaseFilename)
        WriteLog LogFile, "Using " & Tablespace
        OpenVger Tablespace
        OpenBatchcat Tablespace
        InputFileNum = FreeFile
        Open InputFilename For Input As InputFileNum
        Do While Not EOF(InputFileNum)
            Sleep Interval
            Line Input #InputFileNum, Line
            BibID = CLng(Line)
            RC = DeleteBibRecord(BibID)
            If RC = dbSuccess Then
                WriteLog LogFile, "Bib #" & BibID & " deleted " & Now()
            Else
                WriteLog LogFile, "ERROR: Bib #" & BibID & " not deleted, error " & RC & " " & Now()
            End If
            DoEvents
        Loop
        Close InputFileNum
        CloseVger
    Else
        'error
    End If

CloseLog LogFile

End Sub

Private Sub ParseArguments()
    ' Handle command-line arguments, if any
    Dim cnt As Integer
    Dim Args() As String            'Dynamic array for unknown number of arguments
    Dim ArgsSize As Integer
    
    'Initialize globals in case not set by command-line
    Interval = 1000
    ProductionMode = False
    
    Args = Split(Command, " ")
    ArgsSize = UBound(Args)
    For cnt = 0 To ArgsSize
        If cnt <= ArgsSize Then
            Select Case Args(cnt)
                Case "-f"
                    InputFilename = Args(cnt + 1)
                    cnt = cnt + 1
                Case "-i"
                    Interval = CLng(Args(cnt + 1)) * 1000
                    cnt = cnt + 1
                Case "-p"
                    ProductionMode = True
            End Select
        End If
    Next
End Sub

Private Function DeleteBibRecord(BibID As Long) As DeleteBibReturnCode
    DeleteBibRecord = BatchCat.DeleteBibRecord(BibID)
End Function

Private Sub OpenBatchcat(Tablespace As String)
    'Sets form-global Batchcat
    Dim ReturnCode As Integer
    Dim IniPath As String
    Dim User As String
    Dim Password As String
    
    Const BASEPATH As String = "c:\voyager\extras\"
    
    If BatchCat Is Nothing Then
        Set BatchCat = New ClassBatchCat
    End If
    
    Select Case Tablespace
        Case "UCLA_TESTDB"
            User = "uclaloader"
            Password = "a9s8den1"
        Case "UCLADB"
            User = "uclaloader"
            Password = "a9s8den1"
    End Select
    
    If HOME Then
        IniPath = "d:\voyager\extras\" & Tablespace
    Else
        IniPath = BASEPATH & Tablespace
    End If
    
    ReturnCode = BatchCat.Connect(IniPath, User, Password)
    If ReturnCode <> 0 Then
        MsgBox "Batchcat return code: " & ReturnCode
    End If
End Sub

Private Sub OpenVger(Tablespace As String)
    'Sets form-global VgerReader
    Dim User As String
    Dim SQL As String
    
    If Vger Is Nothing Then
        Set Vger = New VgerReaderClass6
    Else
        Vger.Disconnect
    End If
    
    Select Case Tablespace
        Case "UCLA_TESTDB"
            User = "UCLA_TESTDBREAD"
        Case "UCLADB"
            User = "UCLA_PREADDB"
    End Select
    
    With Vger
        .DSN = "Voyager"
        .uID = User
        .PWD = User 'for read-only account user & pwd are the same
        .TableNamePrefix = Tablespace
    End With
    
    'Initialize form-global recordsets - ensure they exist, for later use
    SQL = "SELECT * FROM Dual"
    With Vger
        .ExecuteSQL SQL, 0
        rs1 = .ResultSet
        .ExecuteSQL SQL, 0
        rs2 = .ResultSet
    End With
End Sub

Private Sub CloseVger()
    'Closes form-global VgerReader
    Vger.Disconnect
End Sub

Private Function GetLocID(LocCode As String) As Long
    Dim LocID As String  'String for Vger function; need to convert to Long for Add/Update functions
    Dim LocName As String
    Dim LocDisplay As String
    Dim LocSpine As String
    Dim Suppressed As Boolean
        
    Vger.TranslateLocationCode LocCode, LocID, LocName, LocDisplay, LocSpine, Suppressed
    If IsNumeric(LocID) Then
        GetLocID = CLng(LocID)
    Else
        GetLocID = 0
    End If
End Function

Private Function GetItemTypeID(ItemType As String) As Long
    Dim ItemTypeID As Long
    Dim SQL As String
    
    With Vger
        SQL = "SELECT Item_Type_ID FROM " & .TableNamePrefix & "Item_Type" _
            & " WHERE Item_Type_Code = '" & ItemType & "'"
        .ExecuteSQL SQL, rs1
        .ResultSet = rs1
        If .GetNextRow Then
            ItemTypeID = .CurrentRow(1)
        End If
    End With
    
    GetItemTypeID = ItemTypeID
End Function

Private Function BarcodeExists(Barcode As String) As Boolean
    With Vger
        .ItemBarcodeIsNumeric = False
        .SearchItemBarcode Barcode, rs1
        .ResultSet = rs1
        If .GetNextRow Then
            BarcodeExists = True
        Else
            BarcodeExists = False
        End If
    End With
End Function

Private Function GetHolLocationCode(HolID As Long) As String
    Dim LocCode As String
    Dim SQL As String

    LocCode = ""
    With Vger
        SQL = "SELECT Location_Code FROM " & .TableNamePrefix & "Location" _
            & " WHERE Location_ID = (SELECT Location_ID FROM " & .TableNamePrefix & "MFHD_Master" _
                & " WHERE MFHD_ID = " & HolID & ")"
        .ExecuteSQL SQL, rs2
        .ResultSet = rs2
        If .GetNextRow Then
            LocCode = .CurrentRow(1)
        End If
    End With
    GetHolLocationCode = LocCode
End Function

Private Function GetHolLocationID(HolID As Long) As Long
    Dim LocID As Long
    Dim SQL As String

    LocID = 0
    With Vger
        SQL = "SELECT Location_ID FROM " & .TableNamePrefix & "Location" _
            & " WHERE Location_ID = (SELECT Location_ID FROM " & .TableNamePrefix & "MFHD_Master" _
                & " WHERE MFHD_ID = " & HolID & ")"
        .ExecuteSQL SQL, rs2
        .ResultSet = rs2
        If .GetNextRow Then
            LocID = CLng(.CurrentRow(1))
        End If
    End With
    GetHolLocationID = LocID
End Function

Private Function BarcodeMatchesHol(Barcode As String, HolID As Long) As Boolean
    Dim SQL As String
    With Vger
        SQL = "SELECT COUNT(*) FROM " & .TableNamePrefix & "item_barcode i" _
            & " WHERE item_barcode = '" & Barcode & "'" _
            & " AND EXISTS (SELECT * FROM " & .TableNamePrefix & "mfhd_item" _
            & " WHERE item_id = i.item_id AND mfhd_id = " & HolID & ")"
        .ExecuteSQL SQL, rs2
        .ResultSet = rs2
        .GetNextRow
        If .CurrentRow(1) = "0" Then
            BarcodeMatchesHol = False
        Else
            BarcodeMatchesHol = True
        End If
    End With
End Function
