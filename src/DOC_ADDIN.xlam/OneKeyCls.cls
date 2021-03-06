VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneKeyCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'FileName : OneKeyCls.cls

Public key As String




'局部变量，保存集合
Private mCol As Collection

Public Function Add(key As String, Optional skey As String) As OneKeyCls
    '创建新对象
    Dim objNewMember As OneKeyCls
    Set objNewMember = New OneKeyCls
    
    
    '设置传入方法的属性
    objNewMember.key = key
    If Len(skey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, skey
    End If
    
    
    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
    
    
End Function

'下面 做了 错误处理！当 没有 关键字 的时候 就会返回Nothing了。
Public Property Get Item(vntIndexKey As Variant) As OneKeyCls
    Err.Clear
    On Error GoTo GetItemError
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
    Exit Property
    
GetItemError:
    Debug.Print Err.Source & " : Error Was found in Item(vntIndexKey as Variant) as oneKeyCls . " & vbCrLf & Err.Description & " : " & Err.Number
    Set Item = Nothing
End Property



Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

Public Function To_excel(fdn As String)
'本函数用于将数据到出到指定文件夹
Dim mfso As New CFSO
If mfso.FileExists(fdn) = False Then
mfso.CreateFolder fdn
End If
Set mfso = Nothing
Dim mokc0 As OneKeyCls
Dim wb As Workbook
Dim i As Integer
If Right(fdn, 1) <> "\" Then fdn = fdn & "\"
For Each mokc0 In mCol
Set wb = Workbooks.Add
wb.SaveAs fdn & Replace(mokc0.key, ":", "") & ".xlsx"

If mokc0.Count > 0 Then

For i = 1 To mokc0.Count
wb.Worksheets(1).Range("A" & i) = mokc0.Item(i).key
Next
End If

wb.Save
wb.Saved = True
wb.Close
Set wb = Nothing

Next mokc0
End Function
Public Function To_excel_ws(ws As Worksheet)
'不用迭代，写五层
Dim i As Long, j As Long
Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long, j1 As Long, j2 As Long, j3 As Long, j4 As Long
i = 1
j = 1
Dim mokc0 As OneKeyCls
Dim mokc1 As OneKeyCls
If mCol.Count > 0 Then
For Each mokc0 In mCol
If mokc0.Count > 0 Then
ws.Cells(i, j) = mokc0.key
i = i + 1
For i1 = 1 To mokc0.Count
If mokc0.Item(i1).Count > 0 Then
ws.Cells(i, 2) = mokc0.Item(i1).key
i = i + 1
For i2 = 1 To mokc0.Item(i1).Count
If mokc0.Item(i1).Item(i2).Count > 0 Then
ws.Cells(i, 3) = mokc0.Item(i1).Item(i2).key
'-------------
ws.Cells(i, 2) = mokc0.Item(i1).key
ws.Cells(i, 1) = mokc0.key
'-------------
i = i + 1
For i3 = 1 To mokc0.Item(i1).Item(i2).Count
ws.Cells(i, 4) = mokc0.Item(i1).Item(i2).Item(i3).key
i = i + 1
Next
Else
ws.Cells(i, 3) = mokc0.Item(i1).Item(i2).key
'-------------
ws.Cells(i, 2) = mokc0.Item(i1).key
ws.Cells(i, 1) = mokc0.key
'-------------

i = i + 1
End If
Next
Else
ws.Cells(i, 2) = mokc0.Item(i1).key
i = i + 1
End If
Next
Else
ws.Cells(i, j) = mokc0.key
i = i + 1
End If
Next mokc0
End If
End Function



