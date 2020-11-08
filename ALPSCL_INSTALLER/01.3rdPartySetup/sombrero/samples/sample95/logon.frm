VERSION 4.00
Begin VB.Form logon 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon For SQL Server System 10"
   ClientHeight    =   3435
   ClientLeft      =   1230
   ClientTop       =   2250
   ClientWidth     =   7560
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
   Height          =   3840
   Left            =   1170
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7560
   Top             =   1905
   Width           =   7680
   Begin VB.CommandButton cancelbut 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton logonbut 
      Caption         =   "Logon to Server"
      Default         =   -1  'True
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox dbase 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "pubs2"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox serverid 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "10NT"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox password 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox uid 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "sa"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "DataBase :"
      Height          =   195
      Left            =   200
      TabIndex        =   3
      Top             =   1200
      Width           =   1000
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Server Id :"
      Height          =   195
      Left            =   200
      TabIndex        =   2
      Top             =   840
      Width           =   1000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "PassWord :"
      Height          =   195
      Left            =   200
      TabIndex        =   1
      Top             =   480
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User Id :"
      Height          =   195
      Left            =   200
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
End
Attribute VB_Name = "logon"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cancelbut_Click()

'   User does not wish to connect

    Unload Me
End Sub

Private Sub logonbut_Click()
'   If the application has not already initialized the
'   DB-Library then use the SqlInit function to init
'   the DB-Library

    Set ses = Mainform!Sqlocxdb1.object
    If dblib = "" Then
        dblib = ses.SqlInit(1)
    End If
    MsgBox dblib
    
'   This form solicits from the user the following
'   information
'
'   Userid
'   Password
'   Database to be used -   this should be either pubs
'                           or pubs2 for this application
'                           to function property

    ses.SqlLoginTime = 5
    usid = uid.Text
    If usid = "" Then
        MsgBox "You must enter a userid"
        Exit Sub
    End If
    
    pword = password.Text
    sid = serverid.Text
'    If sid = "" Then
'        MsgBox "You must enter the server name"
'        Exit Sub
'    End If

    datbase = Dbase.Text
    If datbase = "" Then
        MsgBox "You must enter the database name - pubs for SqlServer 4.2 or pubs2 for System 10"
        Exit Sub
    End If

    
'   Once the information needed to open a connection is
'   obtained from the user open a connection with the server
'   using the SqlOpenConnection function. This function
'   returns a connection pointer if the connection is made.
'   if no connection is made then the connection pointer will
'   be zero (NULL)
    
    ses.userid = usid
    ses.password = pword
    Set cobj = ses.SqlOpenConnection(sid, "WorkStation", "SombreroApp1")
    
'   If the connection is made the next thing we need to do
'   is point to the correct database which has the data
'   required for the application. In this case the data is
'   in the authors table found in the pubs database for SQL
'   Server 4.2 or in the pubs2 database for SYBASE System 10

    If cobj Is Nothing Then
        Exit Sub
    End If
        
    cobj.SqlUse = datbase

'   If the connection has been made and the database changed
'   to pubs/pubs2 then hiding the connect button will allow
'   the application to run

    If cobj.retcode = 1 Then
        Mainform!multibut.Visible = False
        Unload Me
    End If
End Sub

