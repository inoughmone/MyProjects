VERSION 4.00
Begin VB.Form CompressIt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " CompressIT/OCX/VBX"
   ClientHeight    =   4590
   ClientLeft      =   225
   ClientTop       =   450
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Height          =   4995
   Left            =   165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6855
   Top             =   105
   Width           =   6975
   Begin VB.Image Image1 
      Height          =   720
      Left            =   3000
      Picture         =   "COMPRESS.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "819 778-5045 x152"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   6615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "For more information call SFI at "
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   6615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   $"COMPRESS.frx":030A
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sylvain Faust Inc. (SFI) offers the CompressIT Custom Control in both VBX and OCX/16/32 bit versions."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   $"COMPRESS.frx":03F0
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "CompressIt"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Image1_Click()
    Unload Me
End Sub

