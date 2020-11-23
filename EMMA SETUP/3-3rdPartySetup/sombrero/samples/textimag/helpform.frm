VERSION 4.00
Begin VB.Form HelpForm 
   Caption         =   "Archive Help"
   ClientHeight    =   4365
   ClientLeft      =   1125
   ClientTop       =   810
   ClientWidth     =   6720
   ControlBox      =   0   'False
   Height          =   4770
   Left            =   1065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6720
   Top             =   465
   Width           =   6840
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"HelpForm.frx":0000
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3615
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


