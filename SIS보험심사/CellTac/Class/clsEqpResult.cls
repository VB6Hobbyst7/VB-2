VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEqpResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private AdoCn               As ADODB.Connection
Private AdoRs               As ADODB.Recordset
Private Call_Procedure      As String

Private mSPCID              As String
Private mEQPNUM             As String
Private mType               As String

Private Sub Class_Initialize()
    mSPCID = ""
    mEQPNUM = ""
    mType = ""
    Call DBErrClear
End Sub

Public Sub SetAdoCn(ByVal new_AdoCn As ADODB.Connection)
    Set AdoCn = new_AdoCn
End Sub

Public Function Set_EqpResultjet(ByVal Testcd As String, ByVal EqpRst As String) As Boolean
    Dim objPstmt    As New clsStrState
    Dim strSql      As String
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsEqpResult - Private Sub Set_EqpResoult(ByVal EqpRst As Collection)"
    
    With objPstmt
        .InitPreparedStmt "INSERT INTO INTERFACE003 (SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, EQUIPCD)" & _
                          " VALUES(?, ?, ?, ?, ?, ?, ?)"
        .SetString 1, SPCID
        .SetString 2, Testcd
        .SetString 3, EQPNUM
        .SetString 4, Format(Now, "YYYYMMDD")
        .SetString 5, Format(Now, "HHMM")
        .SetString 6, EqpRst
        .SetString 7, INS_CODE
    Call DBExec(AdoCn_Jet, .GetPreparedStmt, Call_Procedure)
    End With
    
    Set objPstmt = Nothing
    Set_EqpResultjet = True
Exit Function

ErrorTrap:
    Set objPstmt = Nothing
    Set_EqpResultjet = False
    Call ErrMsgProc(Call_Procedure)
End Function

Public Function Set_EqpResultsql(ByVal Testcd As String, ByVal ItemCd As String, ByVal EqpRst As String, ByVal EqpCD As String, ByVal ErrDes As String) As Boolean
    Dim objPstmt    As New clsStrState
    Dim strSql      As String
    Dim sqlRet      As Integer
    Dim AdoCn1                As ADODB.Connection
    
On Error GoTo ErrorTrap
    Call_Procedure = "clsEqpResult - Private Sub Set_EqpResoult(ByVal EqpRst As Collection)"
    
    AdoCn.Execute "Exec InterfaceResult_INSERT_sp '" & Testcd & "','" & ItemCd & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & EqpRst & "','" & EqpCD & "','" & ErrDes & "' ", sqlRet
    
    'strSql = "Exec InterfaceResult_INSERT_sp '" & Testcd & "','" & ItemCd & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & EqpRst & "','" & EqpCd & "','" & ErrDes & "' ",sqlret
    'AdoCn.Execute (strSql)
    
    If sqlRet = 0 Then
        Set_EqpResultsql = False
    Else
        Set_EqpResultsql = True
    End If

    Set AdoRs = Nothing
    
    
    Set objPstmt = Nothing
    Set_EqpResultsql = True
Exit Function

ErrorTrap:
    Set objPstmt = Nothing
    Set_EqpResultsql = False
    Call ErrMsgProc(Call_Procedure)

    
End Function


Public Function Spc_Exists(ByVal SpcNo As String) As Boolean
    Dim strSql      As String

On Error GoTo ErrorTrap
    Call_Procedure = "clsCommon - Public Function Spc_Exists(ByVal Spcno As String) As Boolean"
    strSql = "SELECT COUNT(SPCNO) AS CNT FROM LIMAS201 WHERE SPCNO = " & STS(SpcNo)

    Set AdoRs = New ADODB.Recordset
    If GetRecordset(AdoCn_SQL, strSql, AdoRs, Call_Procedure) Then
        If AdoRs("CNT") > 0 Then
            Spc_Exists = True
        Else
            Spc_Exists = False
        End If
    Else
        Spc_Exists = False
    End If
    
    Set AdoRs = Nothing
Exit Function

ErrorTrap:
    Set AdoRs = Nothing
    Spc_Exists = False
    Call ErrMsgProc(Call_Procedure)
End Function

Public Property Get EQPNUM() As String
    EQPNUM = Trim(mEQPNUM)
End Property

Public Property Let EQPNUM(ByVal vNewValue As String)
    mEQPNUM = vNewValue
End Property

Public Property Get SPCID() As String
    SPCID = Trim(mSPCID)
End Property

Public Property Let SPCID(ByVal vNewValue As String)
    mSPCID = vNewValue
End Property

Public Property Get SPCTYPE() As String
    SPCTYPE = Trim(mType)
End Property

Public Property Let SPCTYPE(ByVal vNewValue As String)
    mType = vNewValue
End Property


