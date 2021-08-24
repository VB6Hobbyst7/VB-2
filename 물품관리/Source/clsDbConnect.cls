VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDbConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function cfOraConnect() As Boolean
Dim cCls As clsRegister

    On Error GoTo errOraConnect
    
    Set gOraCn = New ADODB.Connection
    If gOraCn.State = adStateOpen Then
        cfOraConnect = True
        Exit Function
    End If
    
    Set cCls = New clsRegister
    
    With gOraCn
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFile & ";Persist Security Info=False"
'        .ConnectionString = "Provider=MSDASQL;DRIVER={SQL Server};SERVER=" & cCls.sqlserver & ";DATABASE=" & cCls.sqldb & ";UID=" & cCls.sqluser & "; PWD=" & cCls.sqlpswd
        .ConnectionString = "Provider=MSDAORA.1;Persist Security Info=True;User ID=" & cCls.orauser & ";Password=" & cCls.orapswd & ";Data Source=" & cCls.oratns
        .ConnectionTimeout = 30
        .Open
    End With
    
    cfOraConnect = True
    Screen.MousePointer = vbDefault
    Exit Function

errOraConnect:
    MsgBox "검진시스템에 연결할 수 없습니다. (" & Err.Description & ")", vbCritical
    cfOraConnect = False

End Function

Public Function cfOraRecordSet(ByVal brSql As String, Optional ByVal brType As Boolean = False) As ADODB.Recordset

    On Error GoTo errOraRecordset
    Set cfOraRecordSet = New ADODB.Recordset
    cfOraRecordSet.CursorLocation = adUseClient
    If brType Then
        cfOraRecordSet.Open brSql, gOraCn, adOpenKeyset, adLockOptimistic
    Else
        cfOraRecordSet.Open brSql, gOraCn, adOpenStatic, adLockReadOnly
    End If
    Exit Function

errOraRecordset:
    Set cfOraRecordSet = New ADODB.Recordset
    MsgBox Err.Description, vbCritical

End Function

Public Function cfConnect() As Boolean
Dim cCls As clsRegister

    On Error GoTo errConnect
    
    Set cCls = New clsRegister
    Set gDbCn = New ADODB.Connection
    
    With gDbCn
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFile & ";Persist Security Info=False"
        .ConnectionString = "Provider=SQLOLEDB;DRIVER={SQL Server};SERVER=" & cCls.sqlserver & ",1433" & ";DATABASE=" & cCls.sqldb & ";UID=" & cCls.sqluser & "; PWD=" & cCls.sqlpswd
'        .ConnectionString = "Provider=SQLOLEDB;DRIVER={SQL Server};SERVER=" & cCls.sqlserver & ";DATABASE=" & cCls.sqldb & ";UID=" & cCls.sqluser & "; PWD=" & cCls.sqlpswd
'         .ConnectionString = "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=" & cCls.dbuser & ";Password=" & cCls.dbpswd & ";Data Source=" & cCls.dbtns
        .ConnectionTimeout = 10
        .Open
    End With
    
    cfConnect = True
    Exit Function

errConnect:
    MsgBox Err.Description, vbCritical
    cfConnect = False

End Function

Public Function cfRecordSet(ByVal brSql As String, Optional ByVal brType As Boolean = False) As ADODB.Recordset

    On Error GoTo errRecordset
    Set cfRecordSet = New ADODB.Recordset
    If brType Then
        cfRecordSet.Open brSql, gDbCn, adOpenKeyset, adLockOptimistic
    Else
        cfRecordSet.Open brSql, gDbCn, adOpenStatic, adLockReadOnly
    End If
    Exit Function

errRecordset:
    Set cfRecordSet = New ADODB.Recordset
    MsgBox Err.Description, vbCritical

End Function

Public Function cfExecute(ByVal brSql As String) As Boolean

    On Error GoTo errExecute
    gDbCn.Execute brSql
    cfExecute = True
    Exit Function

errExecute:
    cfExecute = False
    MsgBox Err.Description, vbCritical
   
End Function

Public Sub csBegin()

    gDbCn.BeginTrans

End Sub

Public Sub csCommit()

    gDbCn.CommitTrans
    
End Sub

Public Sub csRollback()

    gDbCn.RollbackTrans
    
End Sub

Private Sub Class_Initialize()

'    If gDbCn Is Nothing Then
'        If cfConnect = False Then
'            Set gDbCn = Nothing
'        End If
'    End If

End Sub





