VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public myCon      As New ADODB.Connection
Public myRst      As New ADODB.Recordset

Function Connect_DB(FLFP_DB As String)

Dim re As String, ct As String

re = InStr(1, Application.OperatingSystem, "64-bit", vbTextCompare)
If re = 0 Then
    ct = "provider=Microsoft.jet.OLEDB.4.0;data source="
Else
    ct = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
End If

If re = 0 Then
'=======================================================================================================
myCon.Open ct & FLFP_DB & ";"

                                    
                       
                                    
'=======================================================================================================
Else
'=======================================================================================================
    myCon.Open ct & FLFP_DB

'=======================================================================================================
End If

End Function

Function Q(myTblName As String, sql As String)

Dim re As String, ct As String

re = InStr(1, Application.OperatingSystem, "64-bit", vbTextCompare)
If re = 0 Then
    ct = "provider=Microsoft.jet.OLEDB.4.0;data source="
Else
    ct = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
End If

If re = 0 Then
'=======================================================================================================

'myCon.Execute sql
 myRst.Open sql, myCon, adOpenKeyset, adLockOptimistic
 
 
'myRst2.Index = "PrimaryKey"
'myRst.Open Source:=myTblName, ActiveConnection:=myCon, _
'                                    CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
'                                    Options:=adCmdTableDirect
                                    
   
     
'=======================================================================================================
Else
'=======================================================================================================
    
    myRst.Open sql, myCon, adOpenKeyset, adLockOptimistic
'=======================================================================================================
End If

End Function


Function msg_file()
 myRst.MoveFirst
MsgBox myRst.Fields(2).Value
 

End Function

Function Table_exist(tbname As String) As Boolean
Table_exist = False
Set myRst = myCon.OpenSchema(adSchemaTables)
Do Until myRst.EOF
    If myRst!TABLE_TYPE = "TABLE" Then
       If tbname = myRst!TABLE_NAME Then
       Table_exist = True
       Exit Function
       End If
    End If
   myRst.MoveNext
Loop
End Function
