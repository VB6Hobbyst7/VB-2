VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISDbCon"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISDbCon.cls
'   �ۼ���  : �̻��
'   ��  ��  : �����ͺ��̽��� �����ϴ� Ŭ����
'   �ۼ���  : 2003-12-02
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private WithEvents mDbSetup    As frmIISDbSetup
Attribute mDbSetup.VB_VarHelpID = -1

Public Property Get DbCon() As ADODB.Connection
    Set DbCon = mDbCon
End Property

Public Property Get CliCon() As ADODB.Connection
    Set CliCon = mCliCon
End Property

Public Property Get Error() As clsIISError
    Set Error = mError
End Property

Public Property Let Error(ByVal vData As clsIISError)
    Set mError = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : DB Connection
'   ��ȯ : Ture(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DbConnect() As Boolean
    Dim strCon      As String           'DB ���Ṯ�ڿ�
    Dim lngTryCnt   As Long             '��õ� ����ȸ��
    
    '## ������Ʈ���� ���������� ������ �װ� �ҷ�����
    '## ������ DB���������� �����ϴ� ���� ǥ���Ѵ�.
    If GetDbInfo = False Then Call DbConfig
    If mUserCancel Then Exit Function
    
On Error Resume Next
    Set mDbCon = New ADODB.Connection
    Set mError = New clsIISError
    
DbCon:
    '# DB���º� ���� ���ڿ� ����
    Select Case mDbType
        Case "0"    '## ORACLE
            '## MS Provider
            strCon = "Provider=MSDAORA.1;Persist Security Info=True;" & _
                     "Data Source=" & mSource & ";" & _
                     "User ID=" & mUid & ";" & _
                     "Password=" & mPwd
            
            '## Oracle Provider
'            strCon = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;" & _
'                     "Locale Identifier=1067;" & _
'                     "Data Source=" & mSource & ";" & _
'                     "User ID=" & mUid & ";" & _
'                     "Password=" & mPwd
                    
        Case "1"    '## SYBASE
            strCon = "Provider=Sybase.ASEOLEDBProvider.2;Persist Security Info=True;" & _
                     "Data Source=" & mSource & ";" & _
                     "Initial Catalog=" & mCatalog & ";" & _
                     "User ID=" & mUid & ";" & _
                     "Password=" & mPwd
        Case "2"    '## MSSQL
            strCon = "Provider=SQLOLEDB.1;Persist Security Info=True;" & _
                     "Data Source=" & mSource & ";" & _
                     "Initial Catalog=" & mCatalog & ";" & _
                     "User ID=" & mUid & ";" & _
                     "Password=" & mPwd
        Case "3"    '## ACCESS
            strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;" & _
                     "Data Source=" & mSource & ";" & _
                     "Jet OLEDB:Database Password=" & mPwd
    End Select

    '## DB Open
    mDbCon.CursorLocation = adUseClient
    mDbCon.Open strCon
    If mDbCon.Errors.Count > 0 Then
        If lngTryCnt = 3 Then GoTo Errors
        Call DbConfig
        If mUserCancel Then Exit Function
        lngTryCnt = lngTryCnt + 1
        GoTo DbCon
    End If
    
    Date = GetSysDate
    Time = GetSysDate
    
    '## ClientDb Connect
    If mDbType <> "3" Then
        If CliDbConnect = False Then Exit Function
    End If
    DbConnect = True
    Exit Function

Errors:
    '## �α����Ͽ� ���
    mError.SetLog App.EXEName, "clsIISDbCon", "DbConnect", Err.Description, Now
    DbConnect = False
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb Connection
'   ��ȯ : Ture(����), False(����)
'-----------------------------------------------------------------------------'
Public Function CliDbConnect() As Boolean
    Dim strCon As String            'ClientDb ���Ṯ�ڿ�
    
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;" & _
             "Data Source=" & mClientDbPath & ";" & _
             "Jet OLEDB:Database Password=plis"
    
    '## ClientDb Open
On Error Resume Next
    Set mCliCon = New ADODB.Connection
    
    mCliCon.Open strCon
    If mCliCon.Errors.Count > 0 Then GoTo Errors
    CliDbConnect = True
    Exit Function
    
Errors:
    mError.SetLog App.EXEName, "clsIISDbCon", "CliDbConnect", Err.Description, Now
    CliDbConnect = False
End Function

'-----------------------------------------------------------------------------'
'   ��� : DB Close
'   ��ȯ : Ture(����), False(����)
'-----------------------------------------------------------------------------'
Public Function DbClose() As Boolean
    If mDbCon.State = adStateOpen Then
        mDbCon.Close
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDB Close
'   ��ȯ : Ture(����), False(����)
'-----------------------------------------------------------------------------'
Public Function CliClose() As Boolean
    If mCliCon.State = adStateOpen Then
        mCliCon.Close
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : DB�� ���糯¥, �ð��� �ҷ���
'   ��ȯ : ���糯¥+�ð�
'-----------------------------------------------------------------------------'
Public Function GetSysDate() As Date
    Dim Rs  As ADODB.Recordset
    Dim SQL As String
    
    Select Case mDbType
        Case "0"    '# ORACLE
            SQL = "SELECT sysdate FROM dual"
        Case "1"    '# SYBASE
            SQL = "SELECT Convert(Char(10), GetDate(), 111) + ' '" & _
                  "+ Convert(Char(8), GetDate(), 108) sysdate"
        Case "2"    '# MSSQL
            SQL = "SELECT getdate() sysdate"
        Case "3"    '# ACCESS
            GetSysDate = Now
            Exit Function
    End Select
    
On Error GoTo Errors
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, mDbCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetSysDate = Rs.Fields("sysdate").Value
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    GetSysDate = Now
    Set Rs = Nothing
    mError.SetLog App.EXEName, "clsIISDbCon", "GetSysDate", Err.Description, Now
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������Ʈ���� DB���������� �ִ��� �Ǵ�, ������ ������� �ʱ�ȭ
'   ��ȯ : Ture(����), False(����)
'-----------------------------------------------------------------------------'
Private Function GetDbInfo() As Boolean
    '## ������Ʈ���� �ִ� ������ �ε�
On Error Resume Next
    mDbType = GetSetting(mAppName, cDBSERVER, cDBTYPE)
    mSource = GetSetting(mAppName, cDBSERVER, cSOURCE)
    mCatalog = GetSetting(mAppName, cDBSERVER, cCATALOG)
    mUid = GetSetting(mAppName, cDBSERVER, cUID)
    mPwd = GetSetting(mAppName, cDBSERVER, cPWD)
    
    '## ORACLE�� Catalog �׸�, ACCESS�� Catalog, Uid�׸��� �Է��� �ʿ����!!
    If mDbType = "" Then GoTo EndLine
    Select Case mDbType
        Case "0"        '# ORACLE
            If mUid = "" Then GoTo EndLine
        Case "1", "2"   '# SYBASE, MS-SQL
            If mCatalog = "" Or mUid = "" Then GoTo EndLine
        Case "3"        '# ACCESS
            If mSource = "" Or mPwd = "" Then GoTo EndLine
    End Select
    
    GetDbInfo = True
    Exit Function
    
EndLine:
    GetDbInfo = False
End Function

'-----------------------------------------------------------------------------'
'   ��� : DB���� �����츦 ǥ��
'   ��ȯ : Ture(����), False(����)
'-----------------------------------------------------------------------------'
Private Function DbConfig()
    Set mDbSetup = New frmIISDbSetup
    With mDbSetup
        .AppName = mAppName
        .DbType = mDbType
        .Source = mSource
        .Catalog = mCatalog
        .Uid = mUid
        .Pwd = mPwd
        .Show vbModal
    End With
    Set mDbSetup = Nothing
    If mUserCancel = False Then Call GetDbInfo
End Function

'-----------------------------------------------------------------------------'
'   ��� : ClientDb�� ����
'   ��ȯ : True(����) / False(����)
'-----------------------------------------------------------------------------'
Public Function ClientDbCompress() As Boolean
    Dim ClientDb    As JRO.JetEngine
    Dim strFileNm   As String   '�ӽ� MDB������ ���+���ϸ�
    Dim strCliCon   As String   'ClientDb�� ���Ṯ�ڿ�
    Dim strTmpCon   As String   '�ӽ� MDB�� ���Ṯ�ڿ�
    Dim lngTemp     As Long
    
    '## �ӽ� MDB������ ���+���ϸ��� ���� (ClientDb.mdb���ϰ� ����������)
    lngTemp = InStrRev(mClientDbPath, "\")
    strFileNm = Mid$(mClientDbPath, 1, lngTemp) & "Temp.mdb"
    
    '## ���Ṯ�ڿ�
    strCliCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mClientDbPath & ";" & _
                "Jet OLEDB:Database Password=plis"
    strTmpCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFileNm & ";" & _
                "Jet OLEDB:Database Password=plis"
    
    Call CliClose
                
On Error GoTo Errors
    Set ClientDb = New JRO.JetEngine
    
    '## MDB���� ����
    ClientDb.CompactDatabase strCliCon, strTmpCon
    Kill mClientDbPath
    Name strFileNm As mClientDbPath
    Set ClientDb = Nothing
    
    Call CliDbConnect
    ClientDbCompress = True
    Exit Function
    
Errors:
    Set ClientDb = Nothing
    Call CliDbConnect
    mError.SetLog App.EXEName, "clsIISDbCon", "ClientDbCompress", Err.Description, Now
End Function

Private Sub mDbSetup_UserExit()
    mUserCancel = True
End Sub

Public Sub UnloadObject()
    If Not (mCliCon Is Nothing) Then Call CliClose: Set mCliCon = Nothing
    If Not (mDbCon Is Nothing) Then Call DbClose: Set mDbCon = Nothing
    Set mDbCon = Nothing
    Set mError = Nothing
    Set mMainFrm = Nothing
End Sub
