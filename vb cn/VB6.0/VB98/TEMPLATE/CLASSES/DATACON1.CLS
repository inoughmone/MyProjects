VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComplexDataConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' 要做: 声明一个 recordset 变量来做为类的数据源。例如:

'Private WithEvents RS As ADODB.RecordSet

Public Property Get DataSource() As DataSource
    ' 要做: 返回类的数据源对象。例如:
   
    'Set DataSource = RS.DataSource
End Property
Public Property Set DataSource(ByVal objDataSource As DataSource)
    ' 要做: 设置类的数据源对象。例如:
    
    'Set RS.DataSource = objDataSource
    'PropertyChanged "DataSource"
End Property
Public Property Get DataMember() As DataMember
    ' 要做: 返回数据成员。例如:
    
    'Set DataMember = RS.DataMember
End Property
Public Property Let DataMember(ByVal DataMember As DataMember)
    ' 要做: 指向一个新的数据成员。例如:
    
    'Set RS.DataMember = DataMember
    'PropertyChanged "DataMember"
End Property

