VERSION 4.00
Begin VB.Form mainform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Achive - Retrieve Documents"
   ClientHeight    =   4155
   ClientLeft      =   1830
   ClientTop       =   2415
   ClientWidth     =   8025
   ControlBox      =   0   'False
   Height          =   4560
   Left            =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   Top             =   2070
   Width           =   8145
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   720
      Picture         =   "MAINFORM.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archive / Retrieval Functions"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command6 
         Caption         =   "DELETE Server Item"
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Description_text 
         BackColor       =   &H00C0C0C0&
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   5775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ARCHIVE to Server"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "RETRIEVE from Server"
         Height          =   495
         Left            =   3960
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Filename 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Description :"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   0
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   735
         TabIndex        =   7
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Document Title :"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   0
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1560
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create Archive Table"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Logon to Server"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "MAINFORM.frx":018A
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      DefaultExt      =   "*"
      DialogTitle     =   "Archive Document"
      FileName        =   "*.*"
      Filter          =   "*.*"
   End
   Begin SqlocxdbLibCtl.Sqlocxdb Sqlocxdb1 
      Left            =   0
      Top             =   3480
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
End
Attribute VB_Name = "mainform"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Combo1_Click()

'
'   If a new entry is selected from the list of entries then retrieve all the information
'   about that entry and place it into the appropriate place on the form
'
    
    cobj.SqlCmd = "select description,filename from archive_table where title = '" & Combo1.Text & "'"
    ret = cobj.SqlExec()
    ret = cobj.SqlResults()
    ret = cobj.SqlNextRow()
    Description_text = cobj.SqlData(1)
    filename.Text = cobj.SqlData(2)
    ret = cobj.SqlNextRow()
End Sub

Private Sub Command1_Click()

'
'   show the logon form
'   If the logon succeeded then the cobj object will be initialized to a connection object
'   If the logon succeeded then get a list of all the archived entries in the database
'

    logon.Show 1
    If cobj Is Nothing Then
        Exit Sub
    End If
    Frame1.Visible = True
    Command2.Visible = True
    cobj.SqlCmd = "select title from archive_table"
    If cobj.retcode <> SUCCEED Then
        Unload Me
        Exit Sub
    End If
    ret = cobj.SqlExec()
    If cobj.retcode <> SUCCEED Then
        Unload Me
    End If
    ret = cobj.SqlResults()
    If cobj.retcode <> SUCCEED Then
        Unload Me
        Exit Sub
    End If
    While cobj.SqlNextRow() <> NOMOREROWS
        dt$ = cobj.SqlData(1)
        Combo1.AddItem dt$
    Wend
    Command3.Visible = True
    Command1.Visible = False
End Sub

Private Sub Command2_Click()

'
'   If the table "archive_table" does not exist then create it
'   An attempt to recreate it will result in an error message
'

    Screen.MousePointer = 11
    cobj.SqlCmd = "create table dbo.archive_table ("
    cobj.SqlCmd = "title varchar(30) not null, "
    cobj.SqlCmd = "filename varchar(255) not null, "
    cobj.SqlCmd = "description varchar (255) not null, "
    cobj.SqlCmd = "image_col image null)"
    ret = cobj.SqlExec()
    ret = cobj.SqlResults()
    Screen.MousePointer = 0
End Sub

Private Sub C4_Click()

'
'   The user wishes to create an new archive entry or to write over a previous
'   archive. This is not a real world application so there is no warning in the
'   case of overwriting a SQL Server entry.
'

    

' Open text/image file to load into SQL Server table
    
    Bitmap_file = filename.Text
    Open Bitmap_file For Binary As #1
    filelength& = LOF(1)
    
    Screen.MousePointer = 11
    
'
'   Delete any previous entry
'

    cobj.SqlCmd = "delete archive_table where title = '" + Combo1.Text + "'"
    ret = cobj.SqlExec()
    Do While cobj.SqlResults() <> NOMORERESULTS
        Do While cobj.SqlNextRow() <> NOMOREROWS
        Loop
    Loop
    
'
'   Insert the stub record into the database with a dummy value for the image column
'   The dummy value is required in order to get a valid timestamp in the next step
'

    cobj.SqlCmd = "insert into archive_table  values ('" & Combo1.Text & "', '" & filename.Text & "', '" & Description_text.Text & "',0x80)"
    ret = cobj.SqlExec()
    Do While cobj.SqlResults() <> NOMORERESULTS
        Do While cobj.SqlNextRow() <> NOMOREROWS
        Loop
    Loop
    
 '
 '  Select the image column from the newly created record
 '  in order to get the timestamp from that column
 '
 
    cobj.SqlCmd = "select image_col from archive_table where title = '" + Combo1.Text + "'"
    ret = cobj.SqlExec()
    Do While cobj.SqlResults() <> NOMORERESULTS
        Do While cobj.SqlNextRow() <> NOMOREROWS
            SqlPointer$ = cobj.SqlTxPtr(1)
            SqlTimestamp$ = cobj.SqlTxTimeStamp(1)
        Loop
    Loop
    
    
' Begin inserting text/image into image column in 4096 size chunks

    alen& = 0
    Table$ = "archive_table.image_col"
    
'
'   Execute the SqlWriteText with an empty string to allow data to be written using the
'   SqlMoreText function
'

    If cobj.SqlWriteText(Table$, SqlPointer$, SQLTXPLEN, SqlTimestamp$, 1, filelength&, "") <> FAIL Then
        If cobj.SqlOk() <> FAIL Then
            ret = cobj.SqlResults()
            Done% = False
            Do While filelength& > 0
                If filelength& < 4096 Then
                    bmpbuff$ = Space(filelength&)
                Else
                    bmpbuff$ = Space(4096)
                End If
                filelength& = filelength& - 4096
                Get #1, , bmpbuff$
                
'
'   While there is data to be written out use the SqlMoreText function
'
                ret = cobj.SqlMoreText(Len(bmpbuff$), bmpbuff$)
                alen& = alen& + Len(bmpbuff$)
            Loop
            Screen.MousePointer = 0

'
'   Once all data is written out use the SqlOk And SqlResults to commit the data to the database
'
            If cobj.SqlOk() <> FAIL Then
                If cobj.SqlResults() <> FAIL Then
                    MsgBox filename.Text & " Archived as " & Combo1.Text & "  " & Format(alen&) & " bytes"
                    Combo1.AddItem Combo1.Text
                End If
            End If
        End If
    End If

Close 1
    
End Sub


Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    If Combo1.Text = "" Then
        MsgBox "Please enter a Document Title"
        Exit Sub
    End If
    
    If Description_text.Text = "" Then
        MsgBox "Please enter a Document Description"
        Exit Sub
    End If
    CommonDialog1.Flags = FileMustExist + ReadOnly
    CommonDialog1.CancelError = True
    On Error GoTo cdialogcancel
    
    CommonDialog1.ShowOpen
    filename.Text = CommonDialog1.filename
    C4_Click
    Exit Sub
cdialogcancel:
    On Error GoTo 0
    Resume cidialogcancelexit
cidialogcancelexit:
End Sub

Private Sub Command5_Click()

    CommonDialog1.Flags = OverwritePrompt
    CommonDialog1.CancelError = True
    On Error GoTo cdialogcancel
    
    CommonDialog1.ShowSave
    filename.Text = CommonDialog1.filename
    On Error GoTo 0
    GoTo noerror
cdialogcancel:
    On Error GoTo 0
    Resume cidialogcancelexit
cidialogcancelexit:
    Exit Sub

noerror:
'
' This routine reads an image from the SQL Server
' Get length of bitmap image in image column
'
    Screen.MousePointer = 11

    cobj.SqlCmd = "select datalength(image_col) from archive_table where title = '" + Combo1.Text + "'"
    ret = cobj.SqlExec()
    Do While cobj.SqlResults() <> NOMORERESULTS
        Do While cobj.SqlNextRow() <> NOMOREROWS
            imagelen& = Val(cobj.SqlData(1))
        Loop
    Loop

'
'   Set the text size to the size required to read the text/image column
'

    cobj.SqlCmd = "set textsize " + Str$(imagelen&)
    ret = cobj.SqlExec()
    Do While cobj.SqlResults() <> NOMORERESULTS
        Do While cobj.SqlNextRow() <> NOMOREROWS
        Loop
    Loop

'   Retrieve image data in result rows and write to requested file. The example
'   uses the SqlReadText function which will execute until there is no more data
'   This function is used instead of SqlNextRow.  Using SqlNextRow in the example
'   will preclude SqlReadText from reading any data.
    
    cobj.SqlCmd = "select image_col from archive_table where title = '" + Combo1.Text + "'"
    If cobj.SqlExec() <> FAIL Then
        archive_out$ = filename.Text
        alen& = 0
        Open archive_out$ For Binary As #1
        filepos& = 1
        Do While cobj.SqlResults() = SUCCEED
                idone = False
                Do Until idone = True
                    niw$ = cobj.SqlReadText(4096)
                    If cobj.retcode = NOMOREROWS Then
                        idone = True
                    Else
                        alen& = alen& + Len(niw$)
                        Put 1, filepos&, niw$
                        filepos& = filepos& + Len(niw$)
                    End If
                Loop
        Loop
        Close 1
        Screen.MousePointer = 0
        MsgBox filename.Text & " Retrieved as " & Combo1.Text & "  " & Format(alen&) & " bytes"
    End If

End Sub

Private Sub Command6_Click()
    If Combo1.Text = "" Then
        MsgBox "You must select an archive entry to delete"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
'
'   Delete the required data row from the archive_table
'

    cobj.SqlCmd = "delete archive_table where title = '" + Combo1.Text + "'"
    ret = cobj.SqlExec()
    Do While cobj.SqlResults() <> NOMORERESULTS
        Do While cobj.SqlNextRow() <> NOMOREROWS
        Loop
    Loop
    
    MsgBox "Archive item " & Combo1.Text & " has been deleted"
    Combo1.RemoveItem Combo1.ListIndex
    Combo1.Text = ""
    Description_text.Text = ""
    filename.Text = ""
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    Set sobj = Sqlocxdb1.object
#If Win32 Then
    sobj.OCXInit (2)
#End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not cobj Is Nothing Then
        cobj.SqlClose
    End If
    sobj.SqlWinExit
    sobj.SqlExit
End Sub

Private Sub Image1_Click()
    CompressIt.Show 1
End Sub


Private Sub Picture1_Click()
    HelpForm.Show 1
End Sub

Private Sub Sqlocxdb1_Error(ByVal SqlConn As Object, ByVal Severity As Integer, ByVal ErrorNum As Integer, ByVal OsError As Integer, ByVal errorstr As String, ByVal OsErrorStr As String, retcode As Integer)
    MsgBox errorstr
    retcode = INTCANCEL
End Sub

Private Sub Sqlocxdb1_Message(ByVal SqlConn As Object, ByVal Message As Long, ByVal State As Integer, ByVal Severity As Integer, ByVal msgstr As String, ByVal ServerName As String, ByVal ProcName As String, ByVal LineNum As Integer)
    MsgBox msgstr
End Sub
