VERSION 5.00
Begin VB.Form frmWinMenu 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   -9996
   ClientTop       =   1980
   ClientWidth     =   6684
   LinkTopic       =   "frmWinMenu"
   ScaleHeight     =   8280
   ScaleWidth      =   6684
   Begin VB.Menu mnuWindow 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "新建窗口(&N)"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "重叠(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "水平排列(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "垂直排列(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "排列图标(&A)"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmWinMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuWindowArrangeIcons_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
  MsgBox "New WIndow Code goes here!"
End Sub

Private Sub mnuWindowTileHorizontal_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  Me.Arrange vbTileVertical
End Sub
