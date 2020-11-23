VERSION 4.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Server Login"
   ClientHeight    =   1980
   ClientLeft      =   555
   ClientTop       =   1575
   ClientWidth     =   4980
   ControlBox      =   0   'False
   Height          =   2415
   Left            =   495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4980
   Top             =   1200
   Width           =   5100
   Begin VB.CommandButton OK_BUTTON 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton CANCEL_BUTTON 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox PWord 
      Height          =   375
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Lid 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox SName 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Server:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Login ID:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1100
   End
End
Attribute VB_Name = "Login"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub CANCEL_BUTTON_Click()
 Unload Login
End Sub


Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub OK_BUTTON_Click()

' Get the server name, login Id, & password from the form
'
   ServerName$ = SName.Text
   LoginID$ = Lid.Text
   Password$ = PWord.Text
   
    If ServerName$ = "" Then
        MsgBox "Must supply a server name"
        Exit Sub
    End If
    If LoginID$ = "" Then
        MsgBox "Must supply a login id"
        Exit Sub
    End If

'
'   Connect to the server
'

Dim Result%, Status%
    LoginToServer = SUCCEED

'   Rem Check to see if the connection is live, if so, then close it
'   Set the max time to login to 10 seconds
'   Open the new connection
'   Change the caption of the application to reflect the server name and the database
'   Set the max time we will wait for a SQL Server response
'

    If conobj Is Nothing Then
    Else
        conobj.SqlClose
    End If
    
    sesobj.SqlLoginTime = 10
    sesobj.userid = LoginID$
    sesobj.Password = Password

'
'   Open the control connection. This connection will be used to retrieve information
'   about the table be BCP'd
'
    Screen.MousePointer = 11

    Set conobj = sesobj.SqlOpenConnection(ServerName$, ProgramName$, ProgramName$)
    If conobj Is Nothing Then
        DatabaseName$ = ""
        ServerName$ = ""
        LoginToServer = FAIL
    Else
        sesobj.SqlTime = QueryTimeout%
    End If
    main.Caption = ProgName$ + " - " + ServerName$ + "/" + DatabaseName$
    Screen.MousePointer = 0

    If LoginToServer = SUCCEED Then Unload Me
End Sub

