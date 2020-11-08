VERSION 4.00
Begin VB.Form HelpForm 
   Caption         =   "Sample Application Help Information"
   ClientHeight    =   6600
   ClientLeft      =   780
   ClientTop       =   675
   ClientWidth     =   6720
   Height          =   7005
   Left            =   720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   6720
   Top             =   330
   Width           =   6840
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"HelpForm.frx":0164
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3375
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


