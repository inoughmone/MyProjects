VERSION 4.00
Begin VB.Form HelpForm 
   Caption         =   "Sample Application Help Information"
   ClientHeight    =   2955
   ClientLeft      =   855
   ClientTop       =   1455
   ClientWidth     =   6720
   Height          =   3360
   Left            =   795
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   6720
   Top             =   1110
   Width           =   6840
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Height          =   1695
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


