VERSION 4.00
Begin VB.Form HelpForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bulk Copy Help Information"
   ClientHeight    =   4815
   ClientLeft      =   870
   ClientTop       =   765
   ClientWidth     =   6720
   ControlBox      =   0   'False
   Height          =   5250
   Left            =   810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6720
   Top             =   390
   Width           =   6840
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"HelpForm.frx":0000
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub


