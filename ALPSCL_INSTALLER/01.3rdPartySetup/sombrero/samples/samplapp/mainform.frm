VERSION 4.00
Begin VB.Form Mainform 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test SQL-Sombrero VBX - Author Application"
   ClientHeight    =   4725
   ClientLeft      =   1320
   ClientTop       =   1590
   ClientWidth     =   7680
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
   Height          =   5130
   Left            =   1260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7680
   Top             =   1245
   Width           =   7800
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3000
      Picture         =   "mainform.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   3960
      Width           =   375
   End
   Begin VB.ComboBox author_list 
      BeginProperty Font 
         name            =   "Fixedsys"
         charset         =   0
         weight          =   700
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   120
      Width           =   7095
   End
   Begin VB.CommandButton delbut 
      Caption         =   "Delete Author"
      Height          =   615
      Left            =   6000
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cancelbut 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6240
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton newrec 
      Caption         =   "Add New"
      Height          =   615
      Left            =   6000
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton updbut 
      Caption         =   "Update Info"
      Height          =   615
      Left            =   6000
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton chgbut 
      Caption         =   "Change Info"
      Height          =   615
      Left            =   6000
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton exitbut 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton multibut 
      Caption         =   "Logon"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
   Begin SqlocxdbLibCtl.Sqlocxdb Sqlocxdb1 
      Left            =   240
      Top             =   2520
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   2295
      Left            =   5880
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1455
      Left            =   5880
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Postal Code :"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Country :"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "State :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "City :"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Phone # :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "First Name :"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Last Name :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Author Id :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim t(1 To 3) As String

Private Sub author_list_Click()
    show_author
    chgbut.Visible = True
    delbut.Visible = True
End Sub

Private Sub cancelbut_Click()
    author_list.Enabled = True
    If author_list.ListIndex <> -1 Then
        chgbut.Visible = True
        delbut.Visible = True
    End If
    updbut.Visible = False
    newrec.Visible = True
    cancelbut.Visible = False
    For i = 0 To 8
        text1(i).Enabled = False
    Next
    show_author
End Sub

Private Sub chgbut_Click()
    For i = 1 To 8
       text1(i).Enabled = True
    Next
    author_list.Enabled = False
    chgbut.Visible = False
    updbut.Visible = True
    newrec.Visible = False
    delbut.Visible = False
    cancelbut.Visible = True
    updflag = 1
    text1(1).SetFocus
End Sub

Private Sub delbut_Click()
    updflag = 3
    updbut.Visible = True
    chgbut.Visible = False
    newrec.Visible = False
    cancelbut.Visible = True
    delbut.Visible = False
    author_list.Enabled = False
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
#If Win32 Then
    Sqlocxdb1.OCXInit (2)
#End If
End Sub

Private Sub multibut_Click()
    logon.Show 1
    If multibut.Visible = True Then
        Exit Sub
    End If
    newrec.Visible = True
    
'   Once the user has logged on to the server the list box will be
'   populated with a list of the Author Id, Last Name, and First Name
'
'   The first function is to place the SQL Statement required to get
'   this information from the server into the command buffer using
'   the SqlCmd function

    cobj.SqlCmd = "select 'Author Id' = au_id , 'Last Name' = au_lname , 'First Name' = au_fname  from authors"
    If cobj.retcode = 1 Then

'   The SqlExec function is then used to send the SQL Statement to the
'   server for execution. It is at this point that syntax checking is
'   performed

        cobj.SqlExec
        If cobj.retcode = 1 Then
            
'   If the command syntax was correct the command will be executing at
'   the server now. The next function is the SqlResults function. This
'   function will return once the processing is complete on the server.
'
'   If the SqlCmd function had been passed more than one SQL command then
'   you must perform a SqlResults for each result set being sent back.
'   The end of result sets will be indicated by a NOMORERESUTLS(2) return
'   from SqlResults

            cobj.SqlResults
            If cobj.retcode = 1 Then

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

                For c% = 1 To cols%
                    colnam$ = cobj.SqlColName(c%)
                Next

'   Once the result set is available for processing each row needs to be
'   retrieved. This is accomplished by calling SqlNextRow until the function
'   returns NOMOREROWS(-2)

                cobj.SqlNextRow
                While cobj.retcode <> NOMOREROWS

'   For each column in the result set we call the function SqlData to
'   get the data returned for the column. The data returned is a string
'   representation of the data in the result column. The data is right
'   trimmed when returned. The data when return is being concatonated
'   with tabs (chr(9)) between each item to populate the drop down list
'   box with three columns
                    
                    For c% = 1 To cols%
                        t(c%) = cobj.SqlData(c%)
                    Next
                    aitem$ = t(1) & Space(15 - Len(t(1))) & Left(t(2) & Space(41 - Len(t(2))), 20) & t(3)
                    author_list.AddItem aitem$
                    cobj.SqlNextRow
                Wend
            End If
        End If
    End If
    If cobj.retcode <> NOMOREROWS Then
        Exit Sub
    End If
    For i = 0 To 8
        label1(i).Visible = True
        text1(i).Visible = True
    Next
    author_list.Visible = True
    newrec.Visible = True
    shape1.Visible = True
    shape2.Visible = True
End Sub

Private Sub newrec_Click()
    updflag = 2
    updbut.Visible = True
    chgbut.Visible = False
    newrec.Visible = False
    cancelbut.Visible = True
    delbut.Visible = False
    author_list.Enabled = False
    For i = 0 To 8
        text1(i).Enabled = True
        text1(i).Text = ""
    Next
    text1(0).SetFocus
End Sub

Private Function process_no_rows() As Integer

'   If the SQL Statement was successfully placed into the command buffer then
'   send the command to the server for execution.

    cobj.SqlExec
    If cobj.retcode = SUCCEED% Then
 
'   The SqlResults function will wait until the function is complete on
'   the server. A return of 1 indicates that the statement has been executed
'   on the server

        cobj.SqlResults
        If cobj.retcode = SUCCEED% Then

'   In the case of a delete there are no rows returned in the result set
'   the SqlRows command will confirm that no rows were returned.

            ret = cobj.SqlRows
            If ret = 1 Then
                MsgBox "This command should not return rows"
            Else
                If updflag = 1 Then
                    MsgBox "Information changed for au_id " & text1(0).Text
                Else
                    If updflag = 2 Then
                        MsgBox "Information added for au_id " & text1(0).Text
                    Else
                        MsgBox "Information deleted for au_id " & text1(0).Text
                    End If
                End If
            End If
        Else
            process_no_rows = cobj.retcode
            Exit Function
        End If
    Else
        process_no_rows = cobj.retcode
        Exit Function
    End If
    process_no_rows = cobj.retcode
End Function

Private Sub show_author()
    lc = author_list.ListIndex
    If lc = -1 Then
        Exit Sub
    End If
    d$ = RTrim(Left(author_list.List(lc), 15))
'
'   Once the author has been selected from the list of authors
'   process a SQL Statement which will select all the fields
'   from the authors table for the author selected from the list
'   by using the where clause on the primary key of the table
'
'   The checking the ret variable for 1 indicates a test for a
'   success. If the ret variable contains a 0 then the function
'   failed

    cobj.SqlCmd = "select * from authors where au_id = " & Chr(34) & d$ & Chr$(34)
    If cobj.retcode = 1 Then
        
'   If the command was accepted into the buffer in the statement above
'   then send it to the server for execution. If there are syntax
'   errors in the statement then they will be found in the SqlExec step

        cobj.SqlExec
        If cobj.retcode = 1 Then

'   If the command syntax was correct the command will be executing at
'   the server now. The next function is the SqlResults function. This
'   function will return once the processing is complete on the server.
'
'   If the SqlCmd function had been passed more than one SQL command then
'   you must perform a SqlResults for each result set being sent back.
'   The end of result sets will be indicated by a NOMORERESUTLS(2) return
'   from SqlResults

            cobj.SqlResults
            If cobj.retcode = 1 Then

'   Once the result set is available for processing each row needs to be
'   retrieved. This is accomplished by calling SqlNextRow until the function
'   returns NOMOREROWS(-2)

                cobj.SqlNextRow
                While cobj.retcode <> NOMOREROWS

'   For each column in the result set we call the function SqlData to
'   get the data returned for the column. The data returned is a string
'   representation of the data in the result column. The data is right
'   trimmed when returned.

                    For c% = 1 To 9
                        dat$ = cobj.SqlData(c%)
                        text1(c% - 1).Text = dat$
                    Next
                    cobj.SqlNextRow
                Wend
            End If
        End If
    End If
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

Private Sub Text1_KeyPress(index As Integer, keyascii As Integer)
    If updflag = 0 Then
        keyascii = 0
    End If
End Sub

Private Sub updbut_Click()
    Dim sql As String
    If updflag = 3 Then
        GoTo dontcheck
    End If
    If text1(0).Text = "" Then
        MsgBox "Author Id cannot be NULL"
        Exit Sub
    End If
    If text1(1).Text = "" Then
        MsgBox "Last name cannot be NULL"
        Exit Sub
    End If
    If text1(2).Text = "" Then
        MsgBox "First name cannot be NULL"
        Exit Sub
    End If
    If text1(3).Text = "" Then
        MsgBox "Phone # cannot be NULL"
        Exit Sub
    End If

dontcheck:

    If updflag = 1 Then

'   In this application if the updflag is set to 1 then we are modifing the
'   data for a particular row in the table authors. The following code will
'   create a SQL Statement to 'UPDATE' the authors table where the au_id field
'   is equal to the requested author id.

        sql = "update authors set "
        sql = sql & " au_lname = " & Chr(34) & text1(1).Text & Chr(34) & " ,"
        sql = sql & " au_fname = " & Chr(34) & text1(2).Text & Chr(34) & " ,"
        sql = sql & " phone = " & Chr(34) & text1(3).Text & Chr(34) & " ,"
        sql = sql & " address = " & Chr(34) & text1(4).Text & Chr(34) & " ,"
        sql = sql & " city = " & Chr(34) & text1(5).Text & Chr(34) & " ,"
        sql = sql & " state = " & Chr(34) & text1(6).Text & Chr(34) & " ,"
        sql = sql & " country = " & Chr(34) & text1(7).Text & Chr(34) & " ,"
        sql = sql & " postalcode = " & Chr(34) & text1(8).Text & Chr(34)
        sql = sql & " where au_id = " & Chr(34) & text1(0).Text & Chr(34)
    Else
        If updflag = 2 Then

'   In this application if the updflag is set to 2 then we are adding
'   data for a particular row in the table authors. The following code will
'   create a SQL Statement to 'INSERT' a row into the authors table where the
'   au_id field is equal to the requested author id.

        sql = "insert authors values ("
        For i = 0 To 7
            sql = sql & Chr(34) & text1(i).Text & Chr(34) & " , "
        Next
        sql = sql & "convert(bit," & Chr(34) & text1(8).Text & Chr(34) & "))"
          Else

'   In this application if the updflag is set to 3 then we are deleting
'   an author record from the table
'   The SQL statement will delete a row from the table based on the
'   primary key (au_id)

            sql = "delete from authors where au_id = " & Chr(34) & text1(0).Text & Chr(34)
    
        End If
    End If


    cobj.SqlCmd = sql
    If cobj.retcode = 1 Then

'   If the command was successfully placed in the command buffer then call
'   the subroutine to process a SQL Statement with no rows being returned

        ret = process_no_rows()
    End If

    If ret = 1 Then
        If updflag = 2 Then
            aitem$ = text1(0).Text & Space(15 - Len(text1(0).Text)) & Left(text1(1).Text & Space(41 - Len(text1(1).Text)), 20) & text1(2).Text
            author_list.AddItem aitem$
            author_list.ListIndex = author_list.ListCount - 1
        Else
            If updflag = 1 Then
                aitem$ = text1(0).Text & Space(15 - Len(text1(0).Text)) & Left(text1(1).Text & Space(41 - Len(text1(1).Text)), 20) & text1(2).Text
                lc = author_list.ListIndex
                author_list.List(lc) = aitem$
            Else
                lc = author_list.ListIndex
                author_list.RemoveItem lc
                author_list.ListIndex = 0
            End If
        End If
    Else
        Exit Sub
    End If

    author_list.Enabled = True
    chgbut.Visible = True
    updbut.Visible = False
    newrec.Visible = True
    cancelbut.Visible = False
    delbut.Visible = True
    For i = 0 To 8
        text1(i).Enabled = False
    Next
End Sub

