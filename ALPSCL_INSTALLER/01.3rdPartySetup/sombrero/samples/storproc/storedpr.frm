VERSION 4.00
Begin VB.Form storedproctest 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Stored Procedure Call"
   ClientHeight    =   3090
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   7770
   ControlBox      =   0   'False
   Height          =   3495
   Left            =   15
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7770
   Top             =   15
   Width           =   7890
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   720
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   6840
      Picture         =   "STOREDPR.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox resultstring 
      Height          =   285
      Index           =   3
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox resultstring 
      Height          =   285
      Index           =   2
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   12
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox resultstring 
      Height          =   285
      Index           =   1
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox resultstring 
      Height          =   285
      Index           =   0
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test Stored Procedure"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox iend 
      Height          =   285
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox istart 
      Height          =   285
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox istring 
      Height          =   285
      Left            =   2280
      MaxLength       =   60
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Login To Server"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Stored Procedure"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Return Status : "
      Height          =   195
      Left            =   1830
      TabIndex        =   13
      Top             =   2640
      Width           =   1110
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Results :"
      Height          =   195
      Left            =   5880
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Second Integer Parameter : "
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "First Integer Paramter : "
      Height          =   195
      Left            =   465
      TabIndex        =   3
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Input String :"
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   900
   End
   Begin SqlocxdbLibCtl.Sqlocxdb Sqlocxdb1 
      Left            =   120
      Top             =   120
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
End
Attribute VB_Name = "storedproctest"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'
'   Create the session and connection objects
'   which will be used in this sample program
'

Dim sesobj As Object
Dim conobj As Object

Private Sub Command1_Click()
'
'   This subroutine will create the sample stored procedure
'   used in this sample program
'   You must change the "pubs2" reference if you wish to test
'   using this sample in a database other than "pubs2"

    conobj.SqlUse = "pubs2"
    
'
'   This subroutine will always attempt to drop the existing
'   OCXTest stored procedure.  Do not worry if an error occurs
'   when creating the procedure for the first time
'

    conobj.SqlCmd = "drop procedure dbo.OCXTest"
    ret% = conobj.SqlExec()
    ret% = conobj.SqlResults()
    
'
'   Create the stored procedure for use in the sample program
'

    crlf = Chr(10) & Chr(13)
    
    conobj.SqlCmd = "create procedure dbo.OCXTest" & crlf
    conobj.SqlCmd = " @istring char(60)," & crlf
    conobj.SqlCmd = " @istart int ," & crlf
    conobj.SqlCmd = " @iend int " & crlf
    conobj.SqlCmd = " as" & crlf
    conobj.SqlCmd = " select @istring, convert(char(20),@istart), convert(char(20),@iend)" & crlf
    conobj.SqlCmd = " return 969"
    ret% = conobj.SqlExec()
    ret% = conobj.SqlResults()

'
'   Note that for the two above communications with the database we do
'   not attempt to return rows.  While this is legal only do so in the
'   case where no rows are certain to be returned.
'
End Sub

Private Sub Command2_Click()
 '
 '  This subroutine will initialize DB-Library and open
 '  a connection to the server.  You must change the userid
 '  password and server name to execute this stored procedure
 '
 
    sesobj.SqlInit (1)
    sesobj.userid = "sa"
    sesobj.password = ""
    Set conobj = sesobj.SqlOpenConnection("10nt", "TestStor", "TestStor")
    If conobj Is Nothing Then
        MsgBox "Connection could not be created"
        Exit Sub
    End If
    Command1.Visible = True
    Command2.Visible = False
    Command3.Visible = True
End Sub

Private Sub Command3_Click()
'
'   This is the example of how to call a stored procedure using the
'   DB-Library RPC commands.  You must first tell DB-Library which
'   stored procedure is being called using the SqlRpcInit function.
'   You then pass all necessary parameters using the SqlRpcParam function.
'   Note that all parameters are passed as strings.  SQL-Sombrero/OCX will
'   automatically convert those strings to the correct data type for the
'   parameter as defined in the SqlRpcParam call.
'   Finally you send the command to be processed using the SqlRpcSend function
'   then check for valid using the SqlOk and then you may process the returned
'   result sets as you would any other command
'
    conobj.SqlUse = "pubs2"
    ret% = conobj.SqlRpcInit("OCXTest", 0)
    istr$ = istring.Text
    ist$ = istart.Text
    ied$ = iend.Text
    ret% = conobj.SqlRpcParam("@istring", 0, SQLCHAR%, -1, Len(istr$), istr$)
    ret% = conobj.SqlRpcParam("@istart", 0, SQLINT1%, -1, -1, ist$)
    ret% = conobj.SqlRpcParam("@iend", 0, SQLINT1%, -1, -1, ied$)
    ret% = conobj.SqlRpcSend()
    ret% = conobj.SqlOk()
    If ret% = SUCCEED Then
        ret% = conobj.SqlResults()
        While conobj.SqlNextRow() <> -2
            resultstring(0).Text = conobj.SqlData(1)
            resultstring(1).Text = conobj.SqlData(2)
            resultstring(2).Text = conobj.SqlData(3)
        Wend
        If conobj.SqlHasRetStatus() Then
            resultstring(3).Text = conobj.SqlRetStatus()
        End If
    End If
End Sub


Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set sesobj = Sqlocxdb1.object
#If Win32 Then
    sesobj.OCXInit (1)
#End If
Set conobj = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not conobj Is Nothing Then
        conobj.SqlClose
        Set conobj = Nothing
    End If
    sesobj.SqlWinExit
    sesobj.SqlExit
End Sub



Private Sub Picture1_Click()
    HelpForm.Show 1
End Sub

Private Sub Sqlocxdb1_Error(ByVal SqlConn As Object, ByVal Severity As Integer, ByVal ErrorNum As Integer, ByVal OsError As Integer, ByVal errorstr As String, ByVal OsErrorStr As String, retcode As Integer)
    MsgBox errorstr
    retcode = INTCANCEL
End Sub

Private Sub Sqlocxdb1_Message(ByVal SqlConn As Object, ByVal Message As Long, ByVal State As Integer, ByVal Severity As Integer, ByVal msgstr As String, ByVal ServerName As String, ByVal ProcName As String, ByVal LineNum As Integer)
    If Message <> 5701 Then
        MsgBox msgstr
    End If
End Sub
