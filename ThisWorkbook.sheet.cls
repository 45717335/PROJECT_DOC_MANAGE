Private Sub Workbook_Open()
    On Error GoTo errorhand
    Dim mfso As New CFSO
    Dim str1 As String, str2 As String
    str1 = ThisWorkbook.Fullname
    str2 = P_SPLIT(str1, "\", -1)
    If mfso.FileExists("Z:\24_Temp\PA_Logs\TOOLS\ADD_IN_TOOL\" & str2) = False Then
        Application.StatusBar = "LIB DOES NOT EXIST Z:\24_Temp\PA_Logs\TOOLS\ADD_IN_TOOL\" & str2
        Exit Sub
    End If
    If mfso.Datelastmodify("Z:\24_Temp\PA_Logs\TOOLS\ADD_IN_TOOL\" & str2) > mfso.Datelastmodify(str1) Then
        With ThisWorkbook
            .Saved = True
            .ChangeFileAccess xlReadOnly
            Kill .Fullname
            mfso.copy_file "Z:\24_Temp\PA_Logs\TOOLS\ADD_IN_TOOL\" & str2, str1
            ' .Close
        End With
    End If
    Application.StatusBar = "Update success!"
    Exit Sub
errorhand:
    Application.StatusBar = "Update fail!"
End Sub