VERSION 5.00
Begin VB.Form frmODBCLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC 登录"
   ClientHeight    =   3180
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Icon            =   "ODBC 登录.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   450
      Left            =   2520
      TabIndex        =   13
      Top             =   2655
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   450
      Left            =   915
      TabIndex        =   12
      Top             =   2655
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "连接值"
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4230
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1125
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         Left            =   1125
         TabIndex        =   5
         Top             =   930
         Width           =   3015
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   1125
         TabIndex        =   7
         Top             =   1260
         Width           =   3015
      End
      Begin VB.ComboBox cboDSNList 
         Height          =   315
         ItemData        =   "ODBC 登录.frx":000C
         Left            =   1125
         List            =   "ODBC 登录.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1125
         TabIndex        =   11
         Top             =   1935
         Width           =   3015
      End
      Begin VB.ComboBox cboDrivers 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1125
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1590
         Width           =   3015
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&DSN:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   285
         Width           =   390
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "&UID:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   630
         Width           =   330
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P):"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   4
         Top             =   975
         Width           =   735
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "数据库(&B):"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "驱动(&V):"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   1665
         Width           =   465
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S):"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   10
         Top             =   2010
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmODBCLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim sConnect    As String
    Dim sADOConnect As String
    Dim sDAOConnect As String
    Dim sDSN        As String
    
    If cboDSNList.ListIndex > 0 Then
        sDSN = "DSN=" & cboDSNList.Text & ";"
    Else
        sConnect = sConnect & "Driver=" & cboDrivers.Text & ";"
        sConnect = sConnect & "Server=" & txtServer.Text & ";"
    End If
    
    sConnect = sConnect & "UID=" & txtUID.Text & ";"
    sConnect = sConnect & "PWD=" & txtPWD.Text & ";"
    
    If Len(txtDatabase.Text) > 0 Then
        sConnect = sConnect & "Database=" & txtDatabase.Text & ";"
    End If
    
    sADOConnect = "PROVIDER=MSDASQL;" & sDSN & sConnect
    sDAOConnect = "ODBC;" & sDSN & sConnect
    
    MsgBox _
    "要打开一个ADO 连接，使用： " & vbCrLf & _
    "Set gConnection = New Connection" & vbCrLf & _
    "gConnection.Open """ & sADOConnect & """" & vbCrLf & vbCrLf & _
    "要打开一个DAO 数据库对象，使用： " & vbCrLf & _
    "设置 gDatabase = OpenDatabase(vbNullString, 0, 0, sDAOConnect)" & vbCrLf & vbCrLf & _
    "或打开一个 RDO 连接, 使用:" & vbCrLf & _
    "设置 gRDOConnection = rdoEnvironments(0).OpenConnection(sDSN, rdDriverNoPrompt, 0, sConnect)"
    
    '设置 gDatabase = OpenDatabase(vbNullString, 0, 0, sDAOConnect)
    'RDO:
    '设置 gRDOConnection = rdoEnvironments(0).OpenConnection(sDSN, rdDriverNoPrompt, 0, sConnect)
End Sub

Private Sub Form_Load()
    GetDSNsAndDrivers
End Sub

Private Sub cboDSNList_Click()
    On Error Resume Next
    If cboDSNList.Text = "(None)" Then
        txtServer.Enabled = True
        cboDrivers.Enabled = True
    Else
        txtServer.Enabled = False
        cboDrivers.Enabled = False
    End If
End Sub

Sub GetDSNsAndDrivers()
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         '环境句柄

    On Error Resume Next
    cboDSNList.AddItem "(None)"

    '获得 DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
                
            If sDSN <> Space(iDSNLen) Then
                cboDSNList.AddItem sDSN
                cboDrivers.AddItem sDRV
            End If
        Loop
    End If
    '删除重复项
    If cboDSNList.ListCount > 0 Then
        With cboDrivers
            If .ListCount > 1 Then
                i = 0
                While i < .ListCount
                    If .List(i) = .List(i + 1) Then
                        .RemoveItem (i)
                    Else
                        i = i + 1
                    End If
                Wend
            End If
        End With
    End If
    cboDSNList.ListIndex = 0
End Sub
