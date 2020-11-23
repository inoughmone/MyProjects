VERSION 4.00
Begin VB.Form HelpForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stored Procedure Help"
   ClientHeight    =   4290
   ClientLeft      =   315
   ClientTop       =   465
   ClientWidth     =   6720
   ControlBox      =   0   'False
   Height          =   4695
   Left            =   255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6720
   Top             =   120
   Width           =   6840
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label t1 
      Alignment       =   2  'Center
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub


