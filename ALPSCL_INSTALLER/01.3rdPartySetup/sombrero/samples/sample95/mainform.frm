VERSION 4.00
Begin VB.Form Mainform 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test SQL-Sombrero VBX - Author Application"
   ClientHeight    =   6585
   ClientLeft      =   1065
   ClientTop       =   1830
   ClientWidth     =   9255
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   0
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   7275
   Left            =   1005
   LinkTopic       =   "Form1"
   ScaleHeight     =   27.437
   ScaleMode       =   4  'Character
   ScaleWidth      =   77.125
   Top             =   1200
   Width           =   9375
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7800
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Execute"
      Height          =   615
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox sqltext 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.CommandButton multibut 
      Caption         =   "Logon"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton exitbut 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3960
      Picture         =   "mainform.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   5880
      Width           =   375
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   7920
      Top             =   5760
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      NumImages       =   1
      i1              =   "mainform.frx":018A
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   5760
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      NumImages       =   1
      i1              =   "mainform.frx":0681
   End
   Begin ComctlLib.ListView alist 
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   7223
      _StockProps     =   205
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Icons           =   "ImageList2"
      LabelEdit       =   1
      SmallIcons      =   "ImageList1"
      View            =   3
   End
   Begin SqlocxdbLibCtl.Sqlocxdb Sqlocxdb1 
      Left            =   480
      Top             =   5760
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
   Begin VB.Menu bview 
      Caption         =   "&View"
      Begin VB.Menu lview 
         Caption         =   "&Large Icons"
         Index           =   0
      End
      Begin VB.Menu lview 
         Caption         =   "&Small Icons"
         Index           =   1
      End
      Begin VB.Menu lview 
         Caption         =   "&List View"
         Index           =   2
      End
      Begin VB.Menu lview 
         Caption         =   "&Details"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public curItem As ListItem



Private Sub delbut_Click()
    updflag = 3
    updbut.Visible = True
    chgbut.Visible = False
    newrec.Visible = False
    cancelbut.Visible = True
    delbut.Visible = False
    author_list.Enabled = False
End Sub

Private Sub alist_ItemClick(ByVal Item As ListItem)
    Set curItem = Item
End Sub


Private Sub Command1_Click()
'   The first function is to place the SQL Statement required to get
'   this information from the server into the command buffer using
'   the SqlCmd function
    alist.ListItems.Clear
    alist.ColumnHeaders.Clear


    sql$ = SQLTEXT.Text
    cobj.SqlCmd = sql$
    If cobj.retcode = SUCCEED Then

'   The SqlExec function is then used to send the SQL Statement to the
'   server for execution. It is at this point that syntax checking is
'   performed

        cobj.SqlExec
        If cobj.retcode = SUCCEED Then
            
'   If the command syntax was correct the command will be executing at
'   the server now. The next function is the SqlResults function. This
'   function will return once the processing is complete on the server.
'
'   If the SqlCmd function had been passed more than one SQL command then
'   you must perform a SqlResults for each result set being sent back.
'   The end of result sets will be indicated by a NOMORERESUTLS(2) return
'   from SqlResults

            While cobj.SqlResults <> NOMORERESULTS
                If cobj.retcode = SUCCEED Then
                    Command2.Visible = True
                End If
                
'   Once the SqlResults returns with the indication that a result set
'   is available we can get the number of columns in the result set. In
'   this application this function is not required since we know how
'   many columns were requested. If the application allowed for AdHoc
'   SQL requests then this function is used to indicate the number of
'   columns of data available.

                cols% = cobj.SqlNumCols

'   To get the column headings we use the SqlColName function. This function
'   will return either the column name based on the internal column name or
'   if the syntax 'Column Name' = colname is used to override the internal
'   column name. In our example we have chosen to override the column name

                Dim ch As ColumnHeader
                For c% = 1 To cols%
                    colnam$ = cobj.SqlColName(c%)
                    Collen% = cobj.SqlColLen(c%)
                    If Len(colnam$) > Collen% Then
                        Collen% = Len(colnam$)
                    End If
                    Set ch = alist.ColumnHeaders.Add(, , colnam$, Collen%)
                Next

'   Once the result set is available for processing each row needs to be
'   retrieved. This is accomplished by calling SqlNextRow until the function
'   returns NOMOREROWS(-2)

                cobj.SqlNextRow
                brk = 0
                While cobj.retcode <> NOMOREROWS

'   For each column in the result set we call the function SqlData to
'   get the data returned for the column. The data returned is a string
'   representation of the data in the result column. The data is right
'   trimmed when returned. The data is loaded into the listview control
                    
                    Dim listx As ListItem
                    
                    If cobj.retcode <> REGROW Then
                        computeid = cobj.retcode
                        bylist$ = cobj.SqlByList(computeid)
                        numalts = cobj.SqlNumAlts(computeid)
                        For i% = 1 To numalts
                            Select Case cobj.SqlAltOp(computeid, i%)
                            Case SQLAOPSUM
                                t$ = "Sum"
                            Case SQLAOPMIN
                                t$ = "Min"
                            Case SQLAOPMAX
                                t$ = "Max"
                            Case SQLAOPCNT
                                t$ = "Cnt"
                            Case SQLAOPAVG
                                t$ = "Avg"
                            Case Else
                                t$ = "Other"
                            End Select
                            If Len(bylist$) = 0 Then
                                Set listx = alist.ListItems.Add(, , t$, 1)
                            Else
                                If (Asc(Mid(bylist$, Len(bylist$), 1)) = 1) Then
                                    Set listx = alist.ListItems.Add(, , t$, 1)
                                Else
                                    Set listx = alist.ListItems.Add(, , "", 1)
                                    listx.SubItems(Asc(Mid(bylist$, Len(bylist$), 1)) - 1) = t$
                                End If
                            End If
                            altcoldata$ = cobj.SqlAData(computeid, i%)
                            subitempos = cobj.SqlAltColId(computeid, i%) - 1
                            listx.SubItems(subitempos) = altcoldata$
                        Next
                    Else
                        For c% = 1 To cols%
                            t$ = cobj.SqlData(c%)
                            If c% = 1 Then
                                Set listx = alist.ListItems.Add(, , t$, 1)
                            Else
                                listx.SubItems(c% - 1) = t$
                            End If
                        Next
                    End If
                    brk = brk + 1
                    If brk = 10 Then
                        DoEvents
                        brk = 0
                    End If
                    cobj.SqlNextRow
                Wend
        Command2.Visible = False
    Wend
    End If
    End If
End Sub

Private Sub Command2_Click()
    cobj.SqlCancel
    Command2.Visible = False
End Sub


Private Sub exitbut_Click()
    If Not cobj Is Nothing Then
        cobj.SqlClose
    End If
    Sqlocxdb1.SqlWinExit
    Sqlocxdb1.SqlExit
    End
End Sub

Private Sub Form_Load()
    updflag = 0
    Sqlocxdb1.OCXInit (2)
End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
    i = 0
End Sub


Private Sub lview_Click(Index As Integer)
    alist.View = Index
End Sub

Private Sub multibut_Click()
    logon.Show 1
    If multibut.Visible = True Then
        Exit Sub
    End If
    SQLTEXT.Visible = True
    Command1.Visible = True
End Sub



Private Sub Picture1_Click()
    HelpForm.Show 1
End Sub

Private Sub Sqlocxdb1_Error(ByVal SqlConn As Object, ByVal Severity As Integer, ByVal ErrorNum As Integer, ByVal OsError As Integer, ByVal errorstr As String, ByVal OsErrorStr As String, retcode As Integer)
'   This is a SQL Server callback routine
'
'   When the server needs to inform the user of a error this callback is
'   called. For example if the user tries to logon to a server which does
'   not exist then the error message string (errorstr) will contain text
'   indicating that a connection with the requested server cannot be made.
'   It is also this routine which is called when syntax errors are discovered
'   in SQL commands submitted for execution.

    MsgBox Format(ErrorNum) + "  " + errorstr
    retcode = INTCANCEL%
End Sub

Private Sub Sqlocxdb1_Message(ByVal SqlConn As Object, ByVal Message As Long, ByVal State As Integer, ByVal Severity As Integer, ByVal msgstr As String, ByVal ServerName As String, ByVal ProcName As String, ByVal LineNum As Integer)
'   This is a SQL Server callback routine
'
'   When the server needs to inform the user of a status change this callback is
'   used. For example when the user changes databases using the SqlUse function
'   this event procedure will be called and the message string (msgstr) will
'   contain text indicating the new database name.
'   The variable severity can be used to filter the messages so that only messages
'   which need to be seen can be displayed to the user
    If Message <> 5701 Then
        MsgBox Format(Message) + "  " + msgstr
    End If
End Sub








