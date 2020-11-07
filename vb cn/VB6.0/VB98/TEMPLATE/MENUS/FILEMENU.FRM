VERSION 5.00
Begin VB.Form frmFileMenu 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   -9996
   ClientTop       =   1980
   ClientWidth     =   6684
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   6684
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "关闭(&C)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "全部保存(&L)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "属性(&I)"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "打印设置(&U)..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "发送(&D)..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "frmFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFileClose_Click()
  MsgBox "Close Code goes here!"
End Sub

Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me
End Sub

Private Sub mnuFileNew_Click()
  MsgBox "New File Code goes here!"
End Sub

Private Sub mnuFileOpen_Click()
  MsgBox "Open Code goes here!"
End Sub

Private Sub mnuFilePrint_Click()
  MsgBox "Print Code goes here!"
End Sub

Private Sub mnuFilePrintPreview_Click()
  MsgBox "Print Preview Code goes here!"
End Sub

Private Sub mnuFilePrintSetup_Click()
  MsgBox "Print Setup Code goes here!"
End Sub

Private Sub mnuFileProperties_Click()
  MsgBox "Properties Code goes here!"
End Sub

Private Sub mnuFileSave_Click()
  MsgBox "Save File Code goes here!"
End Sub

Private Sub mnuFileSaveAll_Click()
  MsgBox "Save All Code goes here!"
End Sub

Private Sub mnuFileSaveAs_Click()
  MsgBox "Save As Code goes here!"
End Sub

Private Sub mnuFileSend_Click()
  MsgBox "Send Code goes here!"
End Sub
