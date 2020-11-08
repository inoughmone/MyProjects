VERSION 4.00
Begin VB.Form ShowRecord 
   Caption         =   "Update/Delete Record form"
   ClientHeight    =   8415
   ClientLeft      =   1680
   ClientTop       =   2400
   ClientWidth     =   6690
   Height          =   8820
   Left            =   1620
   LinkTopic       =   "Form1"
   ScaleHeight     =   35.062
   ScaleMode       =   4  'Character
   ScaleWidth      =   55.75
   Top             =   2055
   Width           =   6810
   Begin VB.CommandButton canbutton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton updbutton 
      Caption         =   "Update"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox iskey 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton delbutton 
      Caption         =   "Delete"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox columndata 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label columnname 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "ShowRecord"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub canbutton_Click()
    Unload Me
End Sub


Private Sub delbutton_Click()
    i = Me.Controls.Count
    cols = (i - 3) / 3
    keycount = 0
    For i = 0 To cols - 1
        keycount = keycount + iskey(i).Value
    Next
    If keycount = 0 Then
        MsgBox "You must check at least one field to be the key for deletion"
        Exit Sub
    End If
    sql$ = "where "
    For i = 0 To cols - 1
        If iskey(i).Value = 1 Then
            sql$ = sql$ + columnname(i) + " = '" + ColumnData(i) + "' and "
        End If
    Next
    sqllen = Len(sql$) - 5
    sql$ = Left$(sql$, sqllen)
End Sub


