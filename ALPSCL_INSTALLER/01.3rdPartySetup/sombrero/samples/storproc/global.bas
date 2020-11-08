Attribute VB_Name = "GLOBAL"
' Global declarations for Visual Basic to DB-Library translation dll.
' Used by all apps linking with the SQL-Sombrero VBX (SQLVBXDB.VBX)

' Global return values for all RETCODE type functions
Global Const SUCCEED% = 1
Global Const FAIL% = 0

' Status code for dbresults(). Possible return values are
' SUCCEED, FAIL, and NO_MORE_RESULTS.

Global Const NOMORERESULTS = 2



' return values permitted in error handlers

Global Const INTEXIT% = 0
Global Const INTCONTINUE% = 1
Global Const INTCANCEL% = 2

Global Const MOREROWS = -1
Global Const NOMOREROWS = -2
Global Const REGROW = -1
Global Const BUFFULL = -3




' option values permitted in option setting/querying/clearing
' used by SqlSetOpt%(), SqlIsOpt%(), and SqlClrOpt%().

Global Const SQLBUFFER% = 0
Global Const SQLOFFSET% = 1
Global Const SQLROWCOUNT% = 2
Global Const SQLSTAT% = 3
Global Const SQLTEXTLIMIT% = 4
Global Const SQLTEXTSIZE% = 5
Global Const SQLARITHABORT% = 6
Global Const SQLARITHIGNORE% = 7
Global Const SQLNOAUTOFREE% = 8
Global Const SQLNOCOUNT% = 9
Global Const SQLNOEXEC% = 10
Global Const SQLPARSEONLY% = 11
Global Const SQLSHOWPLAN% = 12
Global Const SQLSTORPROCID% = 13
Global Const SQLANSItoOEM% = 14

' Data type token values.  Used for datatype determination for a column.

Global Const SQLTEXT% = &H23
Global Const SQLARRAY% = &H24
Global Const SQLVARBINARY% = &H25
Global Const SQLINTN% = &H26
Global Const SQLVARCHAR% = &H27
Global Const SQLBINARY% = &H2D
Global Const SQLIMAGE% = &H22
Global Const SQLCHAR% = &H2F
Global Const SQLINT1% = &H30
Global Const SQLBIT% = &H32
Global Const SQLINT2% = &H34
Global Const SQLINT4% = &H38
Global Const SQLMONEY% = &H3C
Global Const SQLDATETIME% = &H3D
Global Const SQLFLT8% = &H3E
Global Const SQLFLTN% = &H6D
Global Const SQLFLT4% = &H3B
Global Const SQLMONEYN% = &H6E
Global Const SQLDATETIMN% = &H6F
Global Const SQLAOPCNT% = &H4B
Global Const SQLAOPSUM% = &H4D
Global Const SQLAOPAVG% = &H4F
Global Const SQLAOPMIN% = &H51
Global Const SQLAOPMAX% = &H52
Global Const SQLAOPANY% = &H53
Global Const SQLAOPNOOP% = &H56
Global Const SQLMONEY4% = &H7A
Global Const SQLDATETIM4% = &H3A

'*** SYBASE System 10, SQL-Sombrero specific datatypes constants
Global Const SQLNUMERICN% = &H6C
Global Const SQLNUMERIC% = &H3F
Global Const SQLDECIMALN% = &H6A
Global Const SQLDECIMAL% = &H37


' error numbers SQL-Sombrero error codes that are passed to local error
' handler

Global Const SQLEMEM% = 10000
Global Const SQLENULL% = 10001
Global Const SQLENLOG% = 10002
Global Const SQLEPWD% = 10003
Global Const SQLECONN% = 10004
Global Const SQLEDDNE% = 10005
Global Const SQLNULLO% = 10006
Global Const SQLESMSG% = 10007
Global Const SQLEBTOK% = 10008
Global Const SQLENSPE% = 10009
Global Const SQLEREAD% = 10010
Global Const SQLECNOR% = 10011
Global Const SQLETSIT% = 10012
Global Const SQLEPARM% = 10013
Global Const SQLEAUTN% = 10014
Global Const SQLECOFL% = 10015
Global Const SQLERDCN% = 10016
Global Const SQLEICN% = 10017
Global Const SQLECLOS% = 10018
Global Const SQLENTXT% = 10019
Global Const SQLEDNTI% = 10020
Global Const SQLETMTD% = 10021
Global Const SQLEASEC% = 10022
Global Const SQLENTLL% = 10023
Global Const SQLETIME% = 10024
Global Const SQLEWRIT% = 10025
Global Const SQLEMODE% = 10026
Global Const SQLEOOB% = 10027
Global Const SQLEITIM% = 10028
Global Const SQLEDBPS% = 10029
Global Const SQLEIOPT% = 10030
Global Const SQLEASNL% = 10031
Global Const SQLEASUL% = 10032
Global Const SQLENPRM% = 10033
Global Const SQLEDBOP% = 10034
Global Const SQLENSIP% = 10035
Global Const SQLECNULL% = 10036
Global Const SQLESEOF% = 10037
Global Const SQLERPND% = 10038
Global Const SQLECSYN% = 10039
Global Const SQLENONET% = 10040
Global Const SQLEBTYP% = 10041
Global Const SQLEABNC% = 10042
Global Const SQLEABMT% = 10043
Global Const SQLEABNP% = 10044
Global Const SQLEBNCR% = 10045
Global Const SQLEAAMT% = 10046
Global Const SQLENXID% = 10047
Global Const SQLEIFNB% = 10048
Global Const SQLEKBCO% = 10049
Global Const SQLEBBCI% = 10050
Global Const SQLEKBCI% = 10051
Global Const SQLEBCWE% = 10052
Global Const SQLEBCNN% = 10053
Global Const SQLEBCOR% = 10054
Global Const SQLEBCPI% = 10055
Global Const SQLEBCPN% = 10056
Global Const SQLEBCPB% = 10057
Global Const SQLEVDPT% = 10058
Global Const SQLEBIVI% = 10059
Global Const SQLEBCBC% = 10060
Global Const SQLEBCFO% = 10061
Global Const SQLEBCVH% = 10062
Global Const SQLEBCUO% = 10063
Global Const SQLEBUOE% = 10064
Global Const SQLEBWEF% = 10065
Global Const SQLEBTMT% = 10066
Global Const SQLEBEOF% = 10067
Global Const SQLEBCSI% = 10068
Global Const SQLEPNUL% = 10069
Global Const SQLEBSKERR% = 10070
Global Const SQLEBDIO% = 10071
Global Const SQLEBCNT% = 10072
Global Const SQLEMDBP% = 10073
Global Const SQLEINIT% = 10074
Global Const SQLCRSINV% = 10075
Global Const SQLCRSCMD% = 10076
Global Const SQLCRSNOIND% = 10077
Global Const SQLCRSDIS% = 10078
Global Const SQLCRSAGR% = 10079
Global Const SQLCRSORD% = 10080
Global Const SQLCRSMEM% = 10081
Global Const SQLCRSBSKEY% = 10082
Global Const SQLCRSNORES% = 10083
Global Const SQLCRSVIEW% = 10084
Global Const SQLCRSBUFR% = 10085
Global Const SQLCRSFROWN% = 10086
Global Const SQLCRSBROL% = 10087
Global Const SQLCRSFRAND% = 10088
Global Const SQLCRSFLAST% = 10089
Global Const SQLCRSRO% = 10090
Global Const SQLCRSTAB% = 10091
Global Const SQLCRSUPDTAB% = 10092
Global Const SQLCRSUPDNB% = 10093
Global Const SQLCRSVIIND% = 10094
Global Const SQLCRSNOUPD% = 10095
Global Const SQLCRSOS2% = 10096
Global Const SQLEBCSA% = 10097
Global Const SQLEBCRO% = 10098
Global Const SQLEBCNE% = 10099
Global Const SQLEBCSK% = 10100

' The severity levels are defined here for error handlers

Global Const EXINFO% = 1
Global Const EXUSER% = 2
Global Const EXNONFATAL% = 3
Global Const EXCONVERSION% = 4
Global Const EXSERVER% = 5
Global Const EXTIME% = 6
Global Const EXPROGRAM% = 7
Global Const EXRESOURCE% = 8
Global Const EXCOMM% = 9
Global Const EXFATAL% = 10
Global Const EXCONSISTENCY% = 11

' Length of text timestamp and text pointer

Global Const SQLTXTSLEN% = 8          ' length of text timestamp
Global Const SQLTXPLEN% = 16          ' length of text pointer

Global Const OFF_SELECT% = &H16D
Global Const OFF_FROM% = &H14F
Global Const OFF_ORDER% = &H165
Global Const OFF_COMPUTE% = &H139
Global Const OFF_TABLE% = &H173
Global Const OFF_PROCEDURE% = &H16A
Global Const OFF_STATEMENT% = &H1CB
Global Const OFF_PARAM% = &H1C4
Global Const OFF_EXEC% = &H12C

Rem SQL Server data types print lengths.
Global Const PRINT4% = 11
Global Const PRINT2% = 6
Global Const PRINT1% = 3
Global Const PRFLT8% = 21
Global Const PRMONEY = 26
Global Const PRBIT% = 3
Global Const PRDATETIME% = 27
Global Const PRDATETIM4% = 20

' Bulk Copy Definitions (bcp)

Global Const DBIN% = 1              ' transfer from client to server
Global Const DBOUT% = 2             ' transfer from server to client

Global Const BCPMAXERRS% = 1        ' SqlBcpControl parameter
Global Const BCPFIRST% = 2          ' SqlBcpControl parameter
Global Const BCPLAST% = 3           ' SqlBcpControl parameter
Global Const BCPBATCH% = 4          ' SqlBcpControl parameter


' Remote Procedure Call function options

Global Const SQLRPCRECOMPILE% = 1 ' recompile the stored procedure
Global Const SQLRPCRETURN% = 1    ' return parameter

' The following values are passed to SqlServerEnum for searching criteria.

Global Const NETSEARCH% = 1
Global Const LOCSEARCH% = 2


' These constansts are the possible return values from SqlServerEnum.

Global Const ENUMSUCCESS% = 0
Global Const MOREDATA% = 1
Global Const NETNOTAVAIL% = 2
Global Const OUTOFMEMORY% = 4
Global Const NOTSUPPORTED% = 8


' User defined data type for SqlGetColumnInfo

Type ColumnData
   Coltype As Integer
   Collen As Long
   Colname As String * 30
   ColSqlType As String * 30
End Type

' User defined data type for SqlGetAltColInfo

Type AltColumnData
   ColID As Integer
   DataType As Integer
   MaxLen As Long
   AggType As Integer
   AggOpName As String * 30
End Type

' User defined data type for SqlBcpColumnFormat

Type BcpColData
    FType As Integer
    FPLen As Integer
    fColLen As Long
    FTerm As String * 30
    FTLen As Integer
    TCol As Integer
End Type

' User defined data type for SqlDateCrack

Type DateInfo
    Year As Integer
    Quarter As Integer
    Month As Integer
    DayOfYear As Integer
    Day As Integer
    Week As Integer
    WeekDay As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Millisecond As Integer
End Type
