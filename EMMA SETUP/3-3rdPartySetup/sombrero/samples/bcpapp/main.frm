VERSION 4.00
Begin VB.Form main 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BCP Windows"
   ClientHeight    =   3720
   ClientLeft      =   1035
   ClientTop       =   6045
   ClientWidth     =   8235
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   4410
   Icon            =   "MAIN.frx":0000
   Left            =   975
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8235
   Top             =   5415
   Width           =   8355
   Begin VB.CommandButton bulk_in 
      Caption         =   "Bulk In"
      Default         =   -1  'True
      Height          =   1215
      Left            =   6480
      TabIndex        =   15
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton bulk_out 
      Caption         =   "Bulk Out"
      Height          =   1215
      Left            =   4680
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bulk Processing Options"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   4335
      Begin VB.TextBox batch_size 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Text            =   "1000"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Number of rows between each Commit:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3315
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DOS File"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   3000
      TabIndex        =   16
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delimiters"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   4455
         Begin VB.CheckBox ck_rowCRLF 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Use CR/LF"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   11
            Top             =   480
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox del_row 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   10
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox del_col 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Row:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            TabIndex        =   9
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "C&olumn:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   690
         End
      End
      Begin VB.CommandButton cmd_select_file 
         BackColor       =   &H00000000&
         Caption         =   "Select &File"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lb_input_file 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         Caption         =   "(no file)"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SQL Server"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.CheckBox ck_sys_tables 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&System Tables"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox tb_list 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox db_list 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Table"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Database"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   825
      End
   End
   Begin SqlocxdbLibCtl.Sqlocxdb Sqlocxdb1 
      Left            =   5760
      Top             =   3600
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   6600
      Top             =   3600
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Menu mn_connection 
      Caption         =   "&Connection"
      Begin VB.Menu mn_login 
         Caption         =   "&Login"
      End
      Begin VB.Menu mn_logout 
         Caption         =   "L&ogout"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mn_help 
      Caption         =   "&Help"
      Begin VB.Menu help_usage 
         Caption         =   "&Usage"
      End
      Begin VB.Menu mn_about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z

Private Sub batch_size_LostFocus()
    If Val(batch_size.Text) = -1 Then
        MsgBox "Invalid Commit Batch size.  Please enter a number!", 48, ProgName$
        batch_size.SetFocus
    End If
    batch_size.Text = Str$(Val(batch_size.Text))
End Sub

Private Sub ck_rowCRLF_Click()
    del_row.Enabled = (ck_rowCRLF.Value = 0)
    If del_row.Enabled = False Then del_row.Text = ""
End Sub

Private Sub ck_sys_tables_Click()
    db_list_Click
End Sub

Private Sub cmd_select_file_Click()

'
'   Use the common dialog to get a file name from the user
'

    On Error Resume Next
    cmd1.DefaultExt = "DAT"
    cmd1.DialogTitle = "Select Bulk Input File"
    cmd1.Filter = "All Files (*.*)|*.*|Bulk Input Files (*.DAT)|*.DAT|"
    cmd1.FilterIndex = 2
    cmd1.CancelError = True
    cmd1.Flags = OFN_FILEMUSTEXIT Or OFN_HIDEREADONLY Or OFN_SHAREAWARE
    cmd1.Action = 1
    If Err <> 0 Then Exit Sub
    lb_input_file.Caption = cmd1.filename
End Sub

Private Sub bulk_in_Click()

'
'   User has clicked the Bulk In button
'   - ensure that the user has selected a column and row delimiter
'     or is using CR/LF for a row delimiter
'   - ensure that the user has selected a file name to copy from
'

    Dim Dummy$
    If Len(del_col.Text) <> 1 Then
        MsgBox "One character required for column delimiter!", 48, ProgName$
        del_col.SetFocus
        Exit Sub
    End If
    If ck_rowCRLF.Value <> 1 And Len(del_row.Text) <> 1 Then
        MsgBox "One character required for row delimiter or Check CR/LF to specify Carriage-Return / LineFeed combination!", 48, ProgName$
        del_row.SetFocus
    End If
    If Len(lb_input_file.Caption) = 0 Then
        MsgBox "Please select a DOS input file first!", 48, ProgName$
        Exit Sub
    End If
    
    ' check if CR/LF selected for ROW separator
    If ck_rowCRLF.Value = 1 Then
        Dummy$ = Chr$(13) + Chr$(10)
    Else
        Dummy$ = del_row.Text
    End If
    
 '
 '  Once all the checks have been made
 '  perform the Input Bulk Copy
 '
 
    PerformBCPIn RTrim$(tb_list.Text), (lb_input_file.Caption), Val(batch_size.Text), (del_col.Text), Dummy$
End Sub

Private Sub bulk_out_Click()
    
'
'   User has clicked the Bulk In button
'   - ensure that the user has selected a column and row delimiter
'     or is using CR/LF for a row delimiter
'   - ensure that the user has selected a file name to copy from
'
    Dim Dummy$
    If Len(del_col.Text) <> 1 Then
        MsgBox "One character required for column delimiter!", 48, ProgName$
        del_col.SetFocus
        Exit Sub
    End If
    If ck_rowCRLF.Value <> 1 And Len(del_row.Text) <> 1 Then
        MsgBox "One character required for row delimiter or Check CR/LF to specify Carriage-Return / LineFeed combination!", 48, ProgName$
        del_row.SetFocus
    End If
    If Len(lb_input_file.Caption) = 0 Then
        MsgBox "Please select a DOS input file first!", 48, ProgName$
        Exit Sub
    End If
    ' check if CR/LF selected for ROW separator
    If ck_rowCRLF.Value = 1 Then
        Dummy$ = Chr$(13) + Chr$(10)
    Else
        Dummy$ = del_row.Text
    End If
 '
 '  Once all the checks have been made
 '  perform the Output Bulk Copy
 '
    PerformBCPOut RTrim$(tb_list.Text), (lb_input_file.Caption), Val(batch_size.Text), (del_col.Text), Dummy$
End Sub

Private Sub db_list_Click()

'
'   If the list of databases is clicked then
'   change the database using the SqlUse variable
'   and the control connection
'

    Dim X%
    If db_list.Text = "" Then Exit Sub
    Screen.MousePointer = 11
    conobj.SqlUse = db_list.Text
    DatabaseName$ = conobj.SqlName()
    Caption = ProgName$ + " - " + ServerName$ + "/" + DatabaseName$
    X% = DoEvents()
    tb_list.Clear
    
'
'   Get a list of the tables available in the newly selected database
'

    If GetTables(tb_list) = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
End Sub


Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
'
'   Set an object with the SQL-Sombrero object dropped on the form
'   It is easier to use this object rather than having to use the
'   control name, since this object is defined as global and does
'   not require the form name if referenced from outside the main
'   form
'

    Set sesobj = Sqlocxdb1.object
#If Win32 Then
    sesobj.OCXInit (1)
#End If
'
'   Initialize the connection to SQL Server
'   and display the current version of DB-Library
'
    DBLIB_VERSION = sesobj.SqlInit(1)
        
    About.Show 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
'
'   Upon exiting from the application we must clean up our DB-Library
'   connections if any and then use the SqlWinExit and SqlExit functions
'   to release the DB-Library function from Windows
'

    If conobj Is Nothing Then
    Else
        sesobj.SqlFreeLogin (LoginRec%)
        conobj.SqlClose
        Set conobj = Nothing
        ServerName$ = ""
        DatabaseName$ = ""
    End If
    
'
'   Exit SQL, then exit the application.
'
   sesobj.SqlExit
   sesobj.SqlWinExit

End Sub



Private Sub help_usage_Click()
    HelpForm.Show 1
End Sub

Private Sub mn_about_Click()

'
'   Display the About Box
'

    About.Show 1
End Sub

Private Sub mn_connection_Click()
    mn_login.Enabled = conobj Is Nothing
    mn_logout.Enabled = Not (conobj Is Nothing)
End Sub

Private Sub mn_exit_Click()

'
'   Upon exiting from the application we must clean up our DB-Library
'   connections if any and then use the SqlWinExit and SqlExit functions
'   to release the DB-Library function from Windows
'

    If conobj Is Nothing Then
    Else
        conobj.SqlClose
        Set conobj = Nothing
        ServerName$ = ""
        DatabaseName$ = ""
    End If
    
'
'   Exit SQL, then exit the application.
'
   sesobj.SqlExit
   sesobj.SqlWinExit
   End
End Sub

Private Sub mn_login_Click()
  Login.Show 1
    
'
'   Check for a successful login by ensuring the the control connection
'   is valid (ie not nothing). If the login was successful then get
'   the list of databases attached to the server which was connected to
'

    If conobj Is Nothing Then
    Else
        Screen.MousePointer = 11
        If GetDatabases(db_list) = False Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        Screen.MousePointer = 0
    End If
End Sub

Private Sub mn_logout_Click()
    If Not conobj Is Nothing Then
        conobj.SqlClose
        Set conobj = Nothing
    End If
    mn_logout.Enabled = False
    db_list.Clear
    tb_list.Clear
End Sub

Private Sub SQLVBXDB1_ERROR(SqlConn As Integer, Severity As Integer, ErrorNum As Integer, OsError As Integer, errorstr As String, OsErrorStr As String, retcode As Integer)
    
'
'   Only display message if it's not a notification that there's a server error
'

    Dim X%
    If ErrorNum <> 10007 Then
        If ErrorNum = 10050 Or ErrorNum = 10051 Or ErrorNum = 10049 Then
        Else
            MsgBox "DBLibrary Error: " + Str$(ErrorNum) + " " + errorstr
        End If
    End If
End Sub

Private Sub SQLVBXDB1_MESSAGE(SqlConn As Integer, Message As Long, State As Integer, Severity As Integer, msgstr As String, ServerName As String, ProcName As String, LineNum As Integer)

'
'   Rem Only display the message if it's not a general msg or a change language message
'

    If Message <> 5701 And Message <> 5703 Then
        MsgBox "SQL Server Error: " + Str$(Message) + " " + msgstr
    End If

End Sub


Sub PerformBCPIn(db_tbl As String, DosFile As String, BatchSize As Long, ColSep As String, RowSep As String)
    Dim Result%, FileSize&, Dummy$, ColNumber As Integer, Flag%, RowsCopied&, a#, b#
    
'
'   Get Input File size, open the file then read the file looking for column separators
'   If a column seperator is found increment the count until the row separator is found
'   at which time we close the file keeping the number of columns in the file
'

    Open DosFile For Binary Access Read As #1
    FileSize& = LOF(1)
    ' Search number of column...
    Dummy$ = " "
    ColNumber = 0
    Flag% = False
    Do
        ' Check if byte read in the CR/LF test under...
        If Not Flag% Then
            Get #1, , Dummy$    ' Get one byte
        Else
            Flag% = False
        End If
        'Check if column separator
        If Dummy$ = ColSep Then
            ColNumber = ColNumber + 1
        Else
            ' Check if CR/LF as row separator
            If Len(RowSep) = 2 And Dummy$ = Left$(RowSep, 1) Then
                Flag% = True
                Get #1, , Dummy$
                If Dummy$ = Right$(RowSep, 1) Then
                    Exit Do
                End If
            End If
        End If
    Loop Until Dummy$ = RowSep
    ColNumber = ColNumber + 1
    Close #1
    
'
'   Create a BCP object for use in the Bulk Copy
'   The Servername, the database table, the Dos input file name, the file to receive all error messages
'   and the direction of the bulk copy are passed to the BCP object. If the object "bcpobj" is nothing
'   after the SqlBCPInit function then a problem was found during the attempt to make the connection and
'   the bulk copy operation will terminate
'

    Set bcpobj = sesobj.SqlBCPInit(ServerName$, "BCPWK", "BCPAPP", conobj.SqlUse, db_tbl, DosFile, "c:\bcp.err", DBIN%)
    If bcpobj Is Nothing Then
        MsgBox "Could not open BCP object for connection"
        Exit Sub
    End If
    
'
'   Set the batch size for the bulk copy. While this is an optional function this will ensure
'   that each time that each (BatchSize) number of records will result in a commit to the database
'   of the records being bulk copied
'

    Result% = bcpobj.SqlBCPControl(BCPBATCH%, (BatchSize))
    If Result% = FAIL Then
        MsgBox "SqlBCPControl Failed!", 48, ProgName$
        Exit Sub
    End If
    
'
'   Declare to the bulk copy the number of columns which will be found in the input file
'

    Result% = bcpobj.SqlBCPColumns(ColNumber)
    If Result% = FAIL Then
        MsgBox "SqlBCPColumns Failed!", 48, ProgName$
        Exit Sub
    End If
    
'
'   For each column tell the bulk copy the format of the input column. In this example the
'   format for each column is character
'

    For Flag% = 1 To ColNumber - 1
        Result% = bcpobj.SqlBCPColfmt(Flag%, SQLCHAR%, 0, -1, Chr$(Asc(ColSep)), 1, Flag%)
        If Result% = FAIL Then
            MsgBox "SqlBCPColfmt% Failed at column " + Str$(Flag%), 48, ProgName$
            Exit Sub
        End If
    Next
'
'   Set last column of the row with the row separator
'

    If Len(RowSep) = 2 Then
        Result% = bcpobj.SqlBCPColfmt%(Flag%, SQLCHAR%, 0, -1, Chr$(13) + Chr$(10), 2, Flag%)
    Else
        Result% = bcpobj.SqlBCPColfmt%(Flag%, SQLCHAR%, 0, -1, Chr$(Asc(RowSep)), 1, Flag%)
    End If
    If Result% = FAIL Then
        MsgBox "SqlBCPColfmt% Failed at column " + Str$(Flag%), 48, ProgName$
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    z% = DoEvents()
    a# = Now

'
'   Start the execution of the bulk copy. This operation cannot be interupted
'

    Result% = bcpobj.SqlBCPExec%(RowsCopied&)
    b# = Now
    Screen.MousePointer = 0
    z% = DoEvents()
    If Result% = FAIL Then
        Set bcpobj = Nothing
        MsgBox "Incomplete bulk-copy. Only " + Str$(RowsCopied&) + " rows copied.", 48, ProgName$
        Exit Sub
    End If
    MsgBox Str$(RowsCopied&) + " rows bulk copied in!", 48, ProgName$
    ShowTimeStat a#, b#, RowsCopied&

'
'   When the operation is complete close the bulk copy connection by setting the
'   BCP object to nothing. Note this also applies in the case of an error (see the
'   result = FAIL above
'

    Set bcpobj = Nothing
End Sub

Sub PerformBCPOut(db_tbl As String, DosFile As String, BatchSize As Long, ColSep As String, RowSep As String)
    Dim Result%, ColNumber As Integer, Flag%, RowsCopied&, a#, b#
    
'
'   Get number of columns for the table being bulk copied from. This is obtained
'   from the system table syscolumns.
'

    If ExecuteSQLCommand("Select count(*) from syscolumns where id=object_id('" + db_tbl + "')") = FAIL% Then
        MsgBox "Error Finding number of columns", 48, ProgName$
        Exit Sub
    Else
        If conobj.SqlResults() = FAIL% Then Exit Sub
        While conobj.SqlNextRow() <> NOMOREROWS%
            ColNumber = Val(conobj.SqlData(1))
        Wend
    End If
    If ColNumber = 0 Then
        MsgBox "Number of columns is 0!", 48, ProgName$
        Exit Sub
    End If
    
'
'   Create a BCP object for use in the Bulk Copy
'   The Servername, the database table, the Dos input file name, the file to receive all error messages
'   and the direction of the bulk copy are passed to the BCP object. If the object "bcpobj" is nothing
'   after the SqlBCPInit function then a problem was found during the attempt to make the connection and
'   the bulk copy operation will terminate
'
    
    Set bcpobj = sesobj.SqlBCPInit(ServerName$, "BCPWK", "BCPAPP", conobj.SqlUse, db_tbl, DosFile, "c:\bcp.err", DBOUT%)
    If bcpobj Is Nothing Then
        MsgBox "Could not open BCP object for connection"
        Exit Sub
    End If
    
'
'   Set the batch size for the bulk copy. While this is an optional function this will ensure
'   that each time that each (BatchSize) number of records will result in a commit to the file
'   of the records being bulk copied
'
    Result% = bcpobj.SqlBCPControl(BCPBATCH%, (BatchSize))

    If Result% = FAIL Then
        MsgBox "SqlBCPControl Failed!", 48, ProgName$
        Exit Sub
    End If

'
'   Declare to the bulk copy the number of columns which will be found in the database table
'

    Result% = bcpobj.SqlBCPColumns(ColNumber)
    If Result% = FAIL Then
        MsgBox "SqlBCPColumns Failed!", 48, ProgName$
        Exit Sub
    End If

'
'   For each column tell the bulk copy the format of the input column. In this example the
'   format for each column is character
'

    For Flag% = 1 To ColNumber - 1
        Result% = bcpobj.SqlBCPColfmt(Flag%, SQLCHAR%, 0, -1, Chr$(Asc(ColSep)), 1, Flag%)
        If Result% = FAIL Then
            MsgBox "SqlBCPColfmt Failed at column " + Str$(Flag%), 48, ProgName$
            Exit Sub
        End If
    Next

'
'   Set last column of the row with the row separator
'

    If Len(RowSep) = 2 Then
        Result% = bcpobj.SqlBCPColfmt(Flag%, SQLCHAR%, 0, -1, Chr$(13) + Chr$(10), 2, Flag%)
    Else
        Result% = bcpobj.SqlBCPColfmt(Flag%, SQLCHAR%, 0, -1, Chr$(Asc(RowSep)), 1, Flag%)
    End If
    If Result% = FAIL Then
        MsgBox "SqlBCPColfmt Failed at column " + Str$(Flag%), 48, ProgName$
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    a# = Now

'
'   Start the execution of the bulk copy. This operation cannot be interupted
'

    Result% = bcpobj.SqlBCPExec(RowsCopied&)

    b# = Now
    Screen.MousePointer = 0
    
    If Result% = FAIL Then
        Set bcpobj = Nothing
        MsgBox "Incomplete bulk-copy. Only " + Str$(RowsCopied&) + " rows copied.", 48, ProgName$
        Exit Sub
    End If
    MsgBox Str$(RowsCopied&) + " rows bulk copied out!", 48, ProgName$
    ShowTimeStat a#, b#, RowsCopied&
    
'
'   When the operation is complete close the bulk copy connection by setting the
'   BCP object to nothing. Note this also applies in the case of an error (see the
'   result = FAIL above
'
    
    Set bcpobj = Nothing
End Sub

Function GetTables(Database_Control As Control) As Integer

'
'   This routine gets the name of all the tables in the current database on SQL Server.
'   Fill each element in the combobox or list box which is passed into this procedure
'   execute the command.  Get each table name and fill the combobox.
'

    Dim Dummy$
    Dummy$ = "Select name from sysobjects where type='U' "
    If main.ck_sys_tables.Value = 1 Then
        Dummy$ = Dummy$ + "or type='S' "
    End If
    Dummy$ = Dummy$ + "order by name"
    If ExecuteSQLCommand(Dummy$) = FAIL% Then
        GetTables = FAIL
        Exit Function
    Else
        If conobj.SqlResults() = FAIL% Then Exit Function
        While conobj.SqlNextRow() <> NOMOREROWS%
            Database_Control.AddItem conobj.SqlData(1)
        Wend
    End If

    GetTables = SUCCEED

End Function

Function GetDatabases(Database_Control As Control) As Integer

'
'   Rem This routine gets the name of all the databases on the SQL Server.
'   Fill each element in the combobox or list box which is passed into this procedure
'   execute the command.  Get each database name and fill the combobox.
'

    If ExecuteSQLCommand("Select name from master..sysdatabases order by name") = FAIL% Then
        GetDatabases = FAIL
        Exit Function
    Else
        If conobj.SqlResults() = FAIL% Then Exit Function
        While conobj.SqlNextRow() <> NOMOREROWS%
            Database_Control.AddItem conobj.SqlData(1)
        Wend
    End If

    GetDatabases = SUCCEED
End Function

Sub ShowTimeStat(a#, b#, c&)
    If c& = 0 Then
        Exit Sub
    End If
    hd% = Hour(b#) - Hour(a#)
    md% = Minute(b#) - Minute(a#)
    sd% = Second(b#) - Second(a#) + 1
    If sd% = 60 Then
        md% = md% + 1
        sd% = 0
    End If
    If md% = 60 Then
        hd% = hd% + 1
        hd% = 0
    End If
    tmd# = (hd% * 60) + md%
    tsd# = (tmd# * 60) + sd%
    rps# = c& / tsd#
    Dummy$ = "Total seconds: " + Format$(tsd#, "###0") + "  Row/Sec: " + Format$(rps#, "###0") + "  Sec./Row: " + Format$(1 / rps#, "#0.###")
    MsgBox Dummy$, 48
End Sub

Function ExecuteSQLCommand(cmd As String) As Integer

'
'   This routine executes a command(s) and returns whether the
'   execute succeeded or failed.
'

    SQLStatus% = SUCCEED
    ExecuteSQLCommand = SUCCEED
    conobj.SqlCmd = cmd$
    If conobj.SqlExec() = FAIL% Then
        SQLStatus% = FAIL
        ExecuteSQLCommand = FAIL
    End If
    
End Function


Private Sub Sqlocxdb1_Error(ByVal SqlConn As Object, ByVal Severity As Integer, ByVal ErrorNum As Integer, ByVal OsError As Integer, ByVal errorstr As String, ByVal OsErrorStr As String, retcode As Integer)
    MsgBox errorstr
    retcode = INTCANCEL
End Sub



Private Sub Sqlocxdb1_Message(ByVal SqlConn As Object, ByVal Message As Long, ByVal State As Integer, ByVal Severity As Integer, ByVal msgstr As String, ByVal ServerName As String, ByVal ProcName As String, ByVal LineNum As Integer)
    If Message <> 5701 Then
        MsgBox msgstr
    End If
End Sub


