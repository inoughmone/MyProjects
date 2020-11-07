Attribute VB_Name = "modDHTML"
'PutProperty: 通过调用此函数存储信息到 cookie 中。
'             请求输入的是您可能希望存储的
'             属性名称以及属性值。
'
'             选择输入的是:
'               过期 : 指定数据来定义属性的有效生存期
'                         一旦到达过期日期所指定的时间
'                         属性将不再被保存或者被释放。

Public Sub PutProperty(objDocument As HTMLDocument, strName As String, vntValue As Variant, Optional Expires As Date)

     objDocument.cookie = strName & "=" & CStr(vntValue) & _
        IIf(CLng(Expires) = 0, "", "; expires=" & Format(CStr(Expires), "ddd, dd-mmm-yy hh:mm:ss") & " GMT") ' & _

End Sub

'GetProperty: 通过调用此函数检索属性值。
'             请求输入的是属性名称，以及函数
'             返回值将作为当前属性的值。
'             如果属性没有找到或者已经过期,
'             那么返回值将为空字符串。
'
Public Function GetProperty(objDocument As HTMLDocument, strName As String) As Variant
    Dim aryCookies() As String
    Dim strCookie As Variant
    On Local Error GoTo NextCookie

    '将文档 cookie 对象分开并存入 cookie 数组。
    aryCookies = Split(objDocument.cookie, ";")
    For Each strCookie In aryCookies
        If Trim(VBA.Left(strCookie, InStr(strCookie, "=") - 1)) = Trim(strName) Then
            GetProperty = Trim(Mid(strCookie, InStr(strCookie, "=") + 1))
            Exit Function
        End If
NextCookie:
        Err = 0
    Next strCookie
End Function


