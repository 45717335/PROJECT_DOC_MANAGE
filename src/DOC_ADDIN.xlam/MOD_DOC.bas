Attribute VB_Name = "MOD_DOC"
Option Explicit

Private Const MODULE As String = "MOD_DOC"
Public s_def As String




Sub ReFresh()
    If s_def <> "Y" Then s_def = "N"
    Dim MyName As String: MyName = "ReFresh_"
    On Error GoTo Error
    'Dim INI As New clsIniFile
    Dim wb As Workbook
    Dim wb1 As Workbook
    Set wb1 = ActiveWorkbook
    Dim mokc_tkq As New OneKeyCls
    Dim j As Long

    Dim ws2 As Worksheet
    Dim rgselect As Range
    Set rgselect = Selection
    Set ws2 = rgselect.Worksheet
    Dim mfso As New CFSO
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String
    Dim para1 As String, para2 As String, para3 As String, para4 As String, para5 As String, para6 As String, para7 As String
    Dim rg As Range
    Dim i As Integer, i_last As Integer
    Dim ws3 As Worksheet
    If wb1.ActiveSheet.Name = "HELP" Then
        wb1.Activate
        wb1.Worksheets(1).Activate
    End If


    '所选表格和所选区域
    If get_xls_type(ws2, rgselect) = "DOC_TRACK_LIST" Then
        If mokc_tkq.Item("DOC_TRACK_LIST") Is Nothing Then
            para1 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST Root folder:(address in ws eg. A1.../ folder path)", s_def)
            If mfso.folderexists(para1) = True Then
                If Right(para1, 1) <> "\" Then para1 = para1 & "\"
            Else
                para1 = ws2.Range(para1)
                If mfso.folderexists(para1) = True Then
                    If Right(para1, 1) <> "\" Then para1 = para1 & "\"
                Else
                    para1 = ""
                End If
            End If
            para2 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST MID folder without ? or * :(ROW ?? / COL ?? )", s_def)
            If para2 Like "ROW*" Or para2 Like "COL*" Then
            Else
                MsgBox "must be ROW+NUMBER or COL+NUMBER"
                Exit Sub
            End If
            para3 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST END folder with ? or * :(ROW ?? / COL ?? )", s_def)
            If Left(para2, 3) = Left(para3, 3) Then
                MsgBox "HAVE two ROW or two COL MACRO END!"
                Exit Sub
            ElseIf para3 Like "ROW*" Or para3 Like "COL*" Then
            Else
                MsgBox "must be ROW+NUMBER or COL+NUMBER"
                Exit Sub
            End If
            para4 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST COLOUR THAT CAN CHANGE:(address in ws eg. A1.../ RGB(?,?,?))", s_def)
            If para4 Like "RGB(*,*,*)" Then
                para4 = Mid(para4, 5, Len(para4) - 5)
                para4 = CStr(rgb(P_SPLIT(para4, ",", -3), P_SPLIT(para4, ",", -2), P_SPLIT(para4, ",", -1)))


            ElseIf rg_exist(ws2, para4) Then
                para4 = CStr(ws2.Range(para4).Interior.Color)
            Else
                para4 = CStr(rgb(255, 255, 255))
                write_para_rg wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST COLOUR THAT CAN CHANGE:(address in ws eg. A1.../ RGB(?,?,?))", "RGB(255,255,255)"


            End If
            para5 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST TARGET COLOUR :(address in ws eg. A1.../ RGB(?,?,?))", s_def)
            If para5 Like "RGB(*,*,*)" Then
                para5 = Mid(para5, 5, Len(para5) - 5)
                para5 = CStr(rgb(P_SPLIT(para5, ",", -3), P_SPLIT(para5, ",", -2), P_SPLIT(para5, ",", -1)))
            ElseIf rg_exist(ws2, para5) Then
                para5 = CStr(ws2.Range(para5).Interior.Color)
            Else
                para5 = CStr(rgb(255, 255, 0))
                write_para_rg wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST TARGET COLOUR :(address in ws eg. A1.../ RGB(?,?,?))", "RGB(255,255,0)"

            End If
            mokc_tkq.Add "DOC_TRACK_LIST", "DOC_TRACK_LIST"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para1, "PARA1"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para2, "PARA2"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para3, "PARA3"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para4, "PARA4"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para5, "PARA5"
        End If



        ws2.Activate

        str1 = mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA1").key
        str2 = mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA2").key
        str3 = mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA3").key
        j = Selection.Count
        Dim my_selection
        On Error Resume Next
        Set my_selection = Application.InputBox("Select Range to update", "DOC_TOOL", Selection.Address, Type:=8)
        If my_selection Is Nothing Then Exit Sub
        On Error GoTo 0
        For Each rg In my_selection
            If Not (rg.EntireRow.Hidden Or rg.EntireColumn.Hidden) Then

                If Left(str2, 3) = "ROW" And Left(str3, 3) = "COL" Then
                    str4 = ws2.Cells(CInt(Right(str2, Len(str2) - 3)), rg.Column)
                    str5 = ws2.Cells(rg.Row, CInt(Right(str3, Len(str3) - 3)))
                    If Left(str4, 1) = "\" Then str4 = Right(str4, Len(str4) - 1)
                    If Right(str4, 1) <> "\" Then str4 = str4 & "\"


                ElseIf Left(str2, 3) = "COL" And Left(str3, 3) = "ROW" Then
                    str5 = ws2.Cells(rg.Row, CInt(Right(str2, Len(str2) - 3)))
                    str4 = ws2.Cells(CInt(Right(str3, Len(str3) - 3)), rg.Column)

                    If Left(str5, 1) = "\" Then str5 = Right(str5, Len(str5) - 1)
                    If Right(str5, 1) <> "\" Then str5 = str5 & "\"

                End If





                If have_file(str1, str5, str4) Then
                    If rg.Interior.Color = CLng(mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA4").key) Then
                        rg.Interior.Color = CLng(mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA5").key)
                    End If
                End If
            End If
            j = j - 1
            Application.StatusBar = j
            DoEvents
        Next
    End If
    '所选表格和所选区域
    Exit Sub
Error:
    RaiseError Err.Number, MyName, Err.Description

End Sub

Private Function get_xls_type(ws As Worksheet, rg As Range) As String
    On Error GoTo errhand
    get_xls_type = ""
    If InStr(ws.Name, "Docu tracking") > 0 Then
        get_xls_type = "DOC_TRACK_LIST"
    End If
    Exit Function
errhand:
    get_xls_type = ""
End Function

Private Function rg_exist(ws As Worksheet, s_address As String) As Boolean
    On Error GoTo errhand
    Dim str1 As String
    str1 = ws.Range(s_address)
    If Len(str1) > 0 Then
    End If
    rg_exist = True
    Exit Function
errhand:
    rg_exist = False
End Function
Private Function have_file(s_fd As String, s_mid As String, s_tpf As String) As Boolean

    Dim str1 As String, str2 As String
    Dim str3 As String

    Dim van_1 As Variant
    Dim mfso As New CFSO
    Dim i As Integer

    If Len(Trim(s_mid)) <= 1 Or Len(Trim(s_tpf)) = 0 Then
        have_file = False
        Exit Function
    End If
    If InStr(s_tpf, "?") = 0 And InStr(s_tpf, "*") = 0 Then
        have_file = False
        Exit Function
    End If




    str1 = P_SPLIT(s_tpf, "##", 0)
    If Len(s_tpf) > Len(str1) + 2 Then
        s_tpf = Right(s_tpf, Len(s_tpf) - Len(str1) - 2)
    Else
        s_tpf = ""
    End If
    If Left(str1, 1) = "\" Then str1 = Right(str1, Len(str1) - 1)
    str3 = P_SPLIT(P_SPLIT(str1, "?", 0), "*", 0)
    str3 = Left(str3, Len(str3) - Len(str3) + InStrRev(str3, "\"))
    If mfso.folderexists(s_fd & s_mid & str3) = False Then
        have_file = False
        Exit Function
    End If
    van_1 = mfso.GetFiles(s_fd & s_mid & str3, True, "f", True)
    For i = LBound(van_1) To UBound(van_1)
        str2 = van_1(i)
        If str2 Like s_fd & s_mid & str1 Then
            have_file = True
            Exit Function
        End If
    Next



    If Len(s_tpf) = 0 Then
        have_file = False
        Exit Function
    End If

    str1 = P_SPLIT(s_tpf, "##", 0)
    If Len(s_tpf) > Len(str1) + 2 Then
        s_tpf = Right(s_tpf, Len(s_tpf) - Len(str1) - 2)
    Else
        s_tpf = ""
    End If
    If Left(str1, 1) = "\" Then str1 = Right(str1, Len(str1) - 1)
    str3 = P_SPLIT(P_SPLIT(str1, "?", 0), "*", 0)
    str3 = Left(str3, Len(str3) - Len(str3) + InStrRev(str3, "\"))
    If mfso.folderexists(s_fd & s_mid & str3) = False Then
        have_file = False
        Exit Function
    End If
    van_1 = mfso.GetFiles(s_fd & s_mid & str3, True, "f", True)
    For i = LBound(van_1) To UBound(van_1)
        str2 = van_1(i)
        If str2 Like s_fd & s_mid & str1 Then
            have_file = True
            Exit Function
        End If
    Next





    have_file = False
End Function
Private Sub RaiseError(ByVal ErrNumber As Long, ByVal FunctionName As String, Optional ByVal msg As String)
    Dim foo As String
    foo = MODULE & "::" & FunctionName
    Dim ErrMsg As String
    ErrMsg = _
         "Error in: " & foo & vbCrLf & _
         "Errorcode: " & CStr(ErrNumber) & vbCrLf

    If Len(msg) Then
        ErrMsg = ErrMsg & vbCrLf & _
                "Additional error message: " & msg & vbCrLf
    End If
    ErrMsg = ErrMsg & vbCrLf & "Program is canceling now"
    MsgBox Prompt:=ErrMsg, _
           Buttons:=vbCritical + vbOKOnly
    Dim e As ErrObject
    e.Raise vbObjectError + 513, foo, ErrMsg
End Sub

Sub Doc_Heplp()
    Dim sPath As String
    sPath = "Z:\24_Temp\PA_Logs\TOOLS\ADD_IN_TOOL\HELP_DOC_ADDIN"
    Dim mfso As New CFSO
    mfso.CreateFolder sPath
    If mfso.folderexists(sPath) = False Then
        MsgBox "Can Not Open:" & sPath
    Else
        Shell "explorer.exe " & sPath, vbMaximizedFocus
    End If
End Sub

Sub open_doc_folder()
    '打开指定资料文件夹
    '1.判断当前ws是否为文档管理，不是则退出
    '2.选择单元格，打开文件夹
    On Error GoTo errorhand
    Dim para1 As String, para2 As String, para3 As String, para4 As String
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String
    Dim ws As Worksheet
    Dim rg_s
    
    On Error Resume Next
    Set rg_s = Application.InputBox("Select the cell to open document folder ", "DOC_TOOL", Selection.Address, Type:=8)
    If rg_s Is Nothing Then
        Exit Sub
    Else
        Set rg_s = rg_s.Resize(1, 1)
    End If
    On Error GoTo 0
    
    Set ws = ActiveWorkbook.ActiveSheet
    para1 = get_para_rg(ws.Range("A1:Z1"), "WORKSHEET_TYPE", "N")
    If para1 <> "DOC_TOOL" Then
        If MsgBox("SET :[" & ws.Name & "]" & " AS:" & "DOC_TOOL", vbYesNo) = vbYes Then
            write_para_rg ws.Range("A1:Z1"), "WORKSHEET_TYPE", "DOC_TOOL"
        Else
            Exit Sub
        End If
    End If
    para2 = get_para_rg(ws.Range("A1:Z1"), "DOC_TRACK_LIST Root folder:(address in ws eg. A1.../ folder path)", "N")
    para3 = get_para_rg(ws.Range("A1:Z1"), "DOC_TRACK_LIST MID folder without ? or * :(ROW ?? / COL ?? )", "N")
    para4 = get_para_rg(ws.Range("A1:Z1"), "DOC_TRACK_LIST END folder with ? or * :(ROW ?? / COL ?? )", "N")
    str2 = para3
    str3 = para4
    If Left(str2, 3) = "ROW" And Left(str3, 3) = "COL" Then
        str4 = ws.Cells(CInt(Right(str2, Len(str2) - 3)), rg_s.Column)
        str5 = ws.Cells(rg_s.Row, CInt(Right(str3, Len(str3) - 3)))
        If Left(str4, 1) = "\" Then str4 = Right(str4, Len(str4) - 1)
        If Right(str4, 1) <> "\" Then str4 = str4 & "\"
    ElseIf Left(str2, 3) = "COL" And Left(str3, 3) = "ROW" Then
        str5 = ws.Cells(rg_s.Row, CInt(Right(str2, Len(str2) - 3)))
        str4 = ws.Cells(CInt(Right(str3, Len(str3) - 3)), rg_s.Column)
        If Left(str5, 1) = "\" Then str5 = Right(str5, Len(str5) - 1)
        If Right(str5, 1) <> "\" Then str5 = str5 & "\"
        
        
    End If
    If InStr(str4, "*") > 0 And InStr(str4, "\") > 0 Then
        str4 = Left(str4, InStrRev(str4, "\"))
    End If
    If InStr(str4, "?") > 0 And InStr(str4, "\") > 0 Then
        str4 = Left(str4, InStrRev(str4, "\"))
    End If
    
    If Left(str4, 1) = "\" Then str4 = Right(str4, Len(str4) - 1)
    
    str6 = para2 & str5 & str4
    Dim mfso As New CFSO
    If mfso.folderexists(str6) Then
        Shell "explorer.exe " & str6, vbMaximizedFocus
    Else
        If MsgBox("Folder does not exist! Do you want create?(Y/N)" & Chr(10) & str6, vbYesNo) = vbYes Then
            mfso.CreateFolder str6
            Shell "explorer.exe " & str6, vbMaximizedFocus
        End If
    End If
    Exit Sub
errorhand:
End Sub

Sub Create_tree()
Dim mokc_tree As New OneKeyCls

   If s_def <> "Y" Then s_def = "N"
    Dim MyName As String: MyName = "ReFresh_"
    On Error GoTo Error
    'Dim INI As New clsIniFile
    Dim wb As Workbook
    Dim wb1 As Workbook
    Set wb1 = ActiveWorkbook
    Dim mokc_tkq As New OneKeyCls
    Dim j As Long

    Dim ws2 As Worksheet
    Dim rgselect As Range
    Set rgselect = Selection
    Set ws2 = rgselect.Worksheet
    Dim mfso As New CFSO
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String
    Dim para1 As String, para2 As String, para3 As String, para4 As String, para5 As String, para6 As String, para7 As String
    Dim rg As Range
    Dim i As Integer, i_last As Integer
    Dim ws3 As Worksheet
    If wb1.ActiveSheet.Name = "HELP" Then
        wb1.Activate
        wb1.Worksheets(1).Activate
    End If


    '所选表格和所选区域
    If get_xls_type(ws2, rgselect) = "DOC_TRACK_LIST" Then
        If mokc_tkq.Item("DOC_TRACK_LIST") Is Nothing Then
            para1 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST Root folder:(address in ws eg. A1.../ folder path)", s_def)
            If mfso.folderexists(para1) = True Then
                If Right(para1, 1) <> "\" Then para1 = para1 & "\"
            Else
                para1 = ws2.Range(para1)
                If mfso.folderexists(para1) = True Then
                    If Right(para1, 1) <> "\" Then para1 = para1 & "\"
                Else
                    para1 = ""
                End If
            End If
            para2 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST MID folder without ? or * :(ROW ?? / COL ?? )", s_def)
            If para2 Like "ROW*" Or para2 Like "COL*" Then
            Else
                MsgBox "must be ROW+NUMBER or COL+NUMBER"
                Exit Sub
            End If
            para3 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST END folder with ? or * :(ROW ?? / COL ?? )", s_def)
            If Left(para2, 3) = Left(para3, 3) Then
                MsgBox "HAVE two ROW or two COL MACRO END!"
                Exit Sub
            ElseIf para3 Like "ROW*" Or para3 Like "COL*" Then
            Else
                MsgBox "must be ROW+NUMBER or COL+NUMBER"
                Exit Sub
            End If
            para4 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST COLOUR THAT CAN CHANGE:(address in ws eg. A1.../ RGB(?,?,?))", s_def)
            If para4 Like "RGB(*,*,*)" Then
                para4 = Mid(para4, 5, Len(para4) - 5)
                para4 = CStr(rgb(P_SPLIT(para4, ",", -3), P_SPLIT(para4, ",", -2), P_SPLIT(para4, ",", -1)))


            ElseIf rg_exist(ws2, para4) Then
                para4 = CStr(ws2.Range(para4).Interior.Color)
            Else
                para4 = CStr(rgb(255, 255, 255))
                write_para_rg wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST COLOUR THAT CAN CHANGE:(address in ws eg. A1.../ RGB(?,?,?))", "RGB(255,255,255)"


            End If
            para5 = get_para_rg(wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST TARGET COLOUR :(address in ws eg. A1.../ RGB(?,?,?))", s_def)
            If para5 Like "RGB(*,*,*)" Then
                para5 = Mid(para5, 5, Len(para5) - 5)
                para5 = CStr(rgb(P_SPLIT(para5, ",", -3), P_SPLIT(para5, ",", -2), P_SPLIT(para5, ",", -1)))
            ElseIf rg_exist(ws2, para5) Then
                para5 = CStr(ws2.Range(para5).Interior.Color)
            Else
                para5 = CStr(rgb(255, 255, 0))
                write_para_rg wb1.ActiveSheet.Range("A1:Z1"), "DOC_TRACK_LIST TARGET COLOUR :(address in ws eg. A1.../ RGB(?,?,?))", "RGB(255,255,0)"

            End If
            mokc_tkq.Add "DOC_TRACK_LIST", "DOC_TRACK_LIST"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para1, "PARA1"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para2, "PARA2"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para3, "PARA3"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para4, "PARA4"
            mokc_tkq.Item("DOC_TRACK_LIST").Add para5, "PARA5"
        End If



        ws2.Activate

        str1 = mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA1").key
        str2 = mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA2").key
        str3 = mokc_tkq.Item("DOC_TRACK_LIST").Item("PARA3").key
        j = Selection.Count
        Dim my_selection
        On Error Resume Next
        Set my_selection = Application.InputBox("Select Range to update", "DOC_TOOL", Selection.Address, Type:=8)
        If my_selection Is Nothing Then Exit Sub
        On Error GoTo 0
        For Each rg In my_selection
            If Not (rg.EntireRow.Hidden Or rg.EntireColumn.Hidden) Then

                If Left(str2, 3) = "ROW" And Left(str3, 3) = "COL" Then
                    str4 = ws2.Cells(CInt(Right(str2, Len(str2) - 3)), rg.Column)
                    str5 = ws2.Cells(rg.Row, CInt(Right(str3, Len(str3) - 3)))
                    If Left(str4, 1) = "\" Then str4 = Right(str4, Len(str4) - 1)
                    If Right(str4, 1) <> "\" Then str4 = str4 & "\"


                ElseIf Left(str2, 3) = "COL" And Left(str3, 3) = "ROW" Then
                    str5 = ws2.Cells(rg.Row, CInt(Right(str2, Len(str2) - 3)))
                    str4 = ws2.Cells(CInt(Right(str3, Len(str3) - 3)), rg.Column)

                    If Left(str5, 1) = "\" Then str5 = Right(str5, Len(str5) - 1)
                    If Right(str5, 1) <> "\" Then str5 = str5 & "\"

                End If


                REC_FILE mokc_tree, str1, str5, str4
                
                
            End If
            j = j - 1
            Application.StatusBar = j
            DoEvents
        Next
    End If
    '所选表格和所选区域
  
    
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Add
    mokc_tree.To_excel_ws wb2.Worksheets(1)
    
    
    Exit Sub
Error:
    RaiseError Err.Number, MyName, Err.Description
    
End Sub
Function REC_FILE(mokc As OneKeyCls, str1 As String, str2 As String, str3 As String)

If Len(str2) = 0 Then Exit Function
If Len(str3) = 0 Then Exit Function

Dim v_1 As Variant
Dim mfso As New CFSO
Dim para1 As String, para2 As String, para3 As String, para4 As String

If InStr(str3, "*") > 0 And InStr(str3, "\") > 0 Then
para1 = Left(str3, InStrRev(str3, "\"))
Else
para1 = str3
End If

If InStr(para1, "?") > 0 And InStr(para1, "\") > 0 Then
para1 = Left(para1, InStrRev(para1, "\"))
End If

If mfso.folderexists(str1 & str2 & para1) = False Then Exit Function



If mokc.Item(str2) Is Nothing Then mokc.Add str2, str2


v_1 = mfso.GetFiles(str1 & str2 & para1, True, "f", True)


Dim i As Integer, j As Integer, i_last As Integer, j_last As Integer
For i = LBound(v_1) To UBound(v_1)
para2 = v_1(i)
para3 = P_SPLIT(para2, "\", -1)
para4 = mfso.get_flndatesize(para2)

If InStr(para4, "Thumbs.db") = 0 Then
If mokc.Item(str2).Item(para4) Is Nothing Then
 mokc.Item(str2).Add para4, para4
 
 para2 = Right(para2, Len(para2) - Len(str1))
 If mokc.Item(str2).Item(para4).Item(para2) Is Nothing Then
 mokc.Item(str2).Item(para4).Add para2, para2
 End If
 
 
End If
End If




Next
Set v_1 = Nothing



End Function
