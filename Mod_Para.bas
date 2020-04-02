Attribute VB_Name = "Mod_Para"
Option Explicit
Function write_para(wb As Workbook, comment As String, Optional para_v As String = "") As String
    '在SETTING工作表A列中找备注为所给备注的单元格，要求输入
    Dim ws As Worksheet, str1 As String
    Set ws = get_ws(wb, "SETTING")
    Dim i As Integer, i_last As Integer
    i_last = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    For i = 1 To i_last
        If ws.Range("A" & i).comment Is Nothing Then
        Else
            If ws.Range("A" & i).comment.Text = comment Then
                ws.Range("A" & i) = para_v
                Exit Function
            End If
        End If
    Next
    add_comment comment, ws.Range("A" & i)
    ws.Range("A" & i) = para_v
End Function

Function get_para_rg(rg As Range, comment As String, Optional s_input As String = "Y") As String
    Dim rg1 As Range, rg2 As Range
    Dim b1 As Boolean, b2 As Boolean
    get_para_rg = ""
    b1 = False
    b2 = False
    For Each rg1 In rg
        If rg1.comment Is Nothing Then
            If b2 = False Then
                Set rg2 = rg1
                b2 = True
            End If
        Else
            If rg1.comment.Text = comment Then
                If s_input = "Y" Then
                    get_para_rg = InputBox(comment, "GET_PARA", rg1)
                    rg1 = get_para_rg
                    Exit Function
                Else
                    get_para_rg = rg1
                    Exit Function
                End If
            End If
        End If
    Next
    If s_input = "Y" Then
        If b2 Then
            add_comment comment, rg2
            get_para_rg = InputBox(comment, "GET_PARA", rg2)
            rg2 = get_para_rg
        Else
            MsgBox "NO Empty Cell for Para!"
            get_para_rg = ""
        End If
    End If
End Function

Function write_para_rg(rg As Range, comment As String, Optional para_v As String = "") As String
    Dim rg1 As Range, rg2 As Range
    Dim b1 As Boolean, b2 As Boolean
    b1 = False
    b2 = False
    For Each rg1 In rg
        If rg1.comment Is Nothing Then
            If b2 = False Then
                Set rg2 = rg1
                b2 = True
            End If
        Else
            If rg1.comment.Text = comment Then
                rg1 = para_v
                Exit Function
            End If
        End If
    Next
    If b2 Then
        add_comment comment, rg2
        rg2 = para_v
    Else
        MsgBox "NO Empty Cell for Para!"
        write_para_rg = ""
    End If
End Function

Private Function add_comment(ByVal comm_s As String, tar_rg As Range, Optional b_v As Boolean = False) As Boolean
    On Error GoTo errorhand
    If tar_rg.comment Is Nothing Then
        tar_rg.AddComment
    End If
    tar_rg.comment.Text Text:=comm_s
    tar_rg.comment.Visible = b_v
    Exit Function
errorhand:
    If Err.Number <> 0 Then MsgBox "get_ws function: " + Err.Description
    Err.Clear
End Function

