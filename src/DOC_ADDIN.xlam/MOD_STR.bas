Attribute VB_Name = "MOD_STR"
Option Explicit

Function P_SPLIT(ByVal txtRange, ByVal splitter As String, Optional ByVal get_index As Integer = 0)
'拆分字符串，
'get_index=0    则返回被拆分出来的第一个字符串
'get_index=1，2，3    则返回被拆分出来的第二、三、四个字符传
'get_index=-1，-2，-3    则返回被拆分出来的倒数第一、二、三个字符传
'越界返回 ""
    Dim a As Variant
    Dim b As Variant
    Dim c As Variant
    a = Split(txtRange, splitter)
    b = LBound(a)
    c = UBound(a)
    If get_index = -1 Then
    P_SPLIT = a(c)
    ElseIf get_index = 0 Then
    P_SPLIT = a(b)
    ElseIf get_index >= b And get_index <= c Then
    P_SPLIT = a(get_index)
    ElseIf get_index <= -1 And get_index >= -1 * c - 1 Then
    P_SPLIT = a(c + 1 + get_index)
    Else
    P_SPLIT = ""
    End If
End Function

