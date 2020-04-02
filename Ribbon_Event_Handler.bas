Attribute VB_Name = "Ribbon_Event_Handler"
Dim xfg_ribbon As IRibbonUI
Dim bv_gp2 As Boolean, bv_gp3 As Boolean, bv_gp4 As Boolean, bv_gp5 As Boolean





Sub rf_rb()
On Error Resume Next

xfg_ribbon.Invalidate
End Sub







'Callback for customUI.onLoad
Sub BBAC_ADDIN_Loaded(ribbon As IRibbonUI)
Set xfg_ribbon = ribbon
End Sub

Sub DOC_ADDIN_Loaded(ribbon As IRibbonUI)
Set xfg_ribbon = ribbon
End Sub






'Callback for grp2 getVisible
Sub GetVisible_grp2(control As IRibbonControl, ByRef returnedVal)
returnedVal = bv_gp2
End Sub
'Callback for grp3 getVisible
Sub GetVisible_grp3(control As IRibbonControl, ByRef returnedVal)
returnedVal = bv_gp3
End Sub
'Callback for grp4 getVisible
Sub GetVisible_grp4(control As IRibbonControl, ByRef returnedVal)
returnedVal = bv_gp4
End Sub
'Callback for grp5 getVisible
Sub GetVisible_grp5(control As IRibbonControl, ByRef returnedVal)
returnedVal = bv_gp5
End Sub




'Callback for customUI.onLoad
Sub TBM_ADDIN_Loaded(ribbon As IRibbonUI)
Set xfg_ribbon = ribbon
End Sub

'Callback for btn1 onAction
Sub CB1_20200331(control As IRibbonControl)


bv_gp2 = True
bv_gp3 = True
bv_gp4 = True
If xfg_ribbon Is Nothing Then
Application.DisplayAlerts = False
Workbooks.Open "Z:\24_Temp\PA_Logs\TOOLS\ADD_IN_TOOL\DOC_ADDIN.xlam", False, True
Application.DisplayAlerts = True
End If
rf_rb

End Sub

'Callback for btn2 onAction
Sub CB2_20200331(control As IRibbonControl)
s_def = "Y"
ReFresh
s_def = "N"
End Sub


'Callback for btn3 onAction
Sub CB3_20200331(control As IRibbonControl)
s_def = "N"
ReFresh
s_def = "N"
End Sub


'Callback for btn4 onAction
Sub CB4_20200331(control As IRibbonControl)
Doc_Heplp
End Sub


'Callback for btn5 onAction
Sub CB5_20200115(control As IRibbonControl)
MsgBox "HELP"

End Sub

'Callback for btn5 onAction
Sub CB31_20200115(control As IRibbonControl)
 manual_match
End Sub
'Callback for btn5 onAction
Sub CB32_20200115(control As IRibbonControl)
force_match_spl
End Sub

'Callback for btn5 onAction
Sub CB33_20200115(control As IRibbonControl)
read_status_prlist
End Sub

'Callback for btn5 onAction
Sub CB34_20200115(control As IRibbonControl)
 Edit_status_prlist
End Sub

Sub CB35_20200331(control As IRibbonControl)
'打开指定工位的指定文件文件夹
open_doc_folder
End Sub

Sub CB36_20200331(control As IRibbonControl)
Create_tree
End Sub


