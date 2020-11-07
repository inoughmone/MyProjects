VERSION 5.00
Begin VB.Form frmEditMenu 
   Caption         =   "Form1"
   ClientHeight    =   2868
   ClientLeft      =   -9996
   ClientTop       =   2880
   ClientWidth     =   4332
   LinkTopic       =   "Form1"
   ScaleHeight     =   2868
   ScaleWidth      =   4332
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "取消(&U)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "剪切(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "选择性粘贴(&S)..."
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "全部选定(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "反向选定(&I)"
      End
   End
End
Attribute VB_Name = "frmEditMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuEditCopy_Click()
  MsgBox "将复制的代码放在此处！"
End Sub

Private Sub mnuEditCut_Click()
  MsgBox "将剪切的代码放在此处！"
End Sub

Private Sub mnuEditDSelectAll_Click()
  MsgBox "将全部选择的代码放在此处！"
End Sub

Private Sub mnuEditInvertSelection_Click()
  MsgBox "将反向选择的代码放在此处！"
End Sub

Private Sub mnuEditPaste_Click()
  MsgBox "将粘贴的代码放在此处！"
End Sub

Private Sub mnuEditPasteSpecial_Click()
  MsgBox "将粘贴的特殊代码放在此处！"
End Sub

Private Sub mnuEditUndo_Click()
  MsgBox "将 Undo代码放在此处！"
End Sub
