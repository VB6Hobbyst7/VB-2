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

Public Function cfConnect() As Boolean
Dim mServerName As String, mDatabaseName As String, mUserName As String, mPassword As String
Dim mCentername As String

    On Error GoTo errConnect

    mServerName = GetString(HKEY_CURRENT_USER, REG_ORACLEDB_C, REG_SERVER)
    mUserName = GetString(HKEY_CURRENT_USER, REG_ORACLEDB_C, REG_USER_ID)
    mPassword = GetString(HKEY_CURRENT_USER, REG_ORACLEDB_C, REG_PASSWD)

    gAreaCd = GetString(HKEY_CURRENT_USER, REG_PIS, REG_AREACD)

'------------------
    If mServerName = "" Then
        mServerName = "orcl"
        Call SaveString(HKEY_CURRENT_USER, REG_ORACLEDB_C, REG_SERVER, mServerName)
    End If
    If mUserName = "" Then
        mUserName = "plis"
        Call SaveString(HKEY_CURRENT_USER, REG_ORACLEDB_C, REG_USER_ID, mUserName)
    End If
    If mPassword = "" Then
        mPassword = "plis"
        Call SaveString(HKEY_CURRENT_USER, REG_ORACLEDB_C, REG_PASSWD, mPassword)
    End If
'------------------
    Set gDbCn = New ADODB.Connection

    With gDbCn
       .ConnectionString = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;User ID=" & mUserName & ";Password=" & mPassword & ";Data Source=" & mServerName
       .CursorLocation = adUseClient
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




