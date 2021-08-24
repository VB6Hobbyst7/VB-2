VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSMLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Class명  : clsDSMLogOn
'|  2. 기  능   : 직원정보 및 로그인내역을 생성 또는 검색한다.
'|  3. 작성자   : 김미경
'|  4. 작성일   : 2000.10.23
'|
'|  Copyright(C) 2000 대련엠티에스
'+--------------------------------------------------------------------------------------+

Option Explicit

Private mvarProjectId As String
Private mvarDeptFg As String

'// COM010
Private mvarLogInID As String
Private mvarLogInPass As String
Private mvarGroupID As String

'//COM008
Private mvarGroupNm As String
Private mvarUserFg As String
Private mvarAPSFg As String
Private mvarBBSFg As String
Private mvarLISFg As String

Private mvarCanRead As Boolean
Private mvarCanWrite As Boolean
Private mvarCanPrint As Boolean

Private mvarIsManager As Boolean    'Manager
Private mvarIsDeveloper As Boolean  'Developer
Private mvarIsSupervisor As Boolean 'Supervisor
Private mvarIsEndUser As Boolean    'End User

'//COM006 : 직원마스터
Private mvarEmpID As String
Private mvarEmpLngNm As String
Private mvarDeptCd As String

Private mvarLoginExist As Boolean
Private mvarPermitted As Boolean

'// Form 사용권한 Collection
Private mcolAuthority As Collection
Private mvarCount As Integer

'공용변수 - 참조하는 프로젝트로부터 설정됨
'Public MyOraSE As OraSession      'OraSession
'Public MyOraDB As OraDatabase    'OraDatabase
Private MySe As Object
'Private MyDb As DrDatabase

Private mvarOldLoginId As String
Private mvarOldLogInPass As String
Private mvarOldGroupID As String

Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
    mvarDeptFg = Mid(vData, 1, 1)
End Property


Public Property Get DeptFg() As String
    DeptFg = mvarDeptFg
End Property

Public Property Get LoginId() As String
    LoginId = mvarLogInID
End Property

Public Property Get LogInPass() As String
    LogInPass = mvarLogInPass
End Property

Public Property Get GroupID() As String
    GroupID = mvarGroupID
End Property

Public Property Get OldLoginId() As String
    OldLoginId = mvarOldLoginId
End Property

Public Property Get OldLogInPass() As String
    OldLogInPass = mvarOldLogInPass
End Property

Public Property Get OldGroupID() As String
    OldGroupID = mvarOldGroupID
End Property

Public Property Get GroupNm() As String
    GroupNm = mvarGroupNm
End Property

Public Property Get UserFg() As String
    UserFg = mvarUserFg
End Property

Public Property Get APSFg() As String
    APSFg = mvarAPSFg
End Property

Public Property Get BBSFg() As String
    BBSFg = mvarBBSFg
End Property

Public Property Get LISFg() As String
    LISFg = mvarLISFg
End Property

Public Property Get EmpId() As String
    EmpId = mvarEmpID
End Property

Public Property Get EmpLngNm() As String
    EmpLngNm = mvarEmpLngNm
End Property

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

Public Property Get IsManager() As Boolean
    IsManager = mvarIsManager
End Property

Public Property Get IsDeveloper() As Boolean
    IsDeveloper = mvarIsDeveloper
End Property

Public Property Let IsDeveloper(ByVal vData As Boolean)
    mvarIsDeveloper = vData
End Property

Public Property Get IsSupervisr() As Boolean
    IsSupervisr = mvarIsSupervisor
End Property

Public Property Get IsEndUser() As Boolean
    IsEndUser = mvarIsEndUser
End Property

Public Property Get LoginExist() As Boolean
    LoginExist = mvarLoginExist
End Property

Public Property Get Permitted() As Boolean
    Permitted = mvarPermitted
End Property

Public Property Get Count() As Integer
    Count = mcolAuthority.Count
End Property

'Public Property Get Item(ByVal KeyIndex As Variant) As clsDSMAuthority
'    On Error Resume Next
'    Set Item = mcolAuthority(KeyIndex)
'End Property



Public Function LoginInfo(ByVal paraLogId As String) As Boolean
'[Method] - 로그인내역에서 직원정보 검색

    Dim strSQL As String
    Dim rs     As ADODB.Recordset
            
    strSQL = " Select a.loginid, a.loginpass, a.empid, a.logindesc, a.groupid, b.empnm, b.deptcd, " & _
             "       c.groupnm, c.groupdesc, c.userfg, c.apsfg, c.bbsfg, c.lisfg " & _
             " From  " & TB_COM010 & " a, " & TB_COM006 & " b, " & TB_COM008 & " c " & _
             " Where a.loginid " & DBS(paraLogId, 2) & _
             " and   b.empid = a.empid " & _
             " and   c.groupid = a.groupid "
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    With rs
        
        If .EOF Then
            LoginInfo = False: mvarLoginExist = False
            GoTo NoData
        End If
        
        mvarLoginExist = True
        
        mvarLogInID = "" & .Fields("LoginId").Value
        mvarEmpID = "" & .Fields("EmpId").Value
        mvarEmpLngNm = "" & .Fields("EmpNm").Value
        mvarDeptCd = "" & .Fields("DeptCd").Value
        mvarLogInPass = "" & .Fields("LoginPass").Value
        
        mvarGroupID = "" & .Fields("GroupId").Value
        mvarGroupNm = "" & .Fields("GroupNm").Value
        mvarUserFg = "" & .Fields("UserFg").Value
        
        mvarIsManager = IIf(mvarUserFg = "M", True, False)      'Manager
        mvarIsDeveloper = IIf(mvarUserFg = "D", True, False)    'Developer
        mvarIsSupervisor = IIf(mvarUserFg = "S", True, False)   'Supervisor
        mvarIsEndUser = IIf(mvarUserFg = "E", True, False)      'End User

        mvarAPSFg = "" & .Fields("ApsFg").Value
        mvarBBSFg = "" & .Fields("BbsFg").Value
        mvarLISFg = "" & .Fields("LisFg").Value
        
        Select Case mvarProjectId
            Case "APS": LoginInfo = mvarAPSFg   '해부병리
            Case "BBS": LoginInfo = mvarBBSFg   '혈액은행
            Case "LIS": LoginInfo = mvarLISFg   '임상병리
        End Select

        LoginInfo = True
    End With
    
NoData:
    rs.Close
    Set rs = Nothing
    
End Function

Public Sub GetAuthority()
'[Method] - Form 사용권한 읽어오기

    Dim strSQL  As String
    Dim rs      As ADODB.Recordset
    Dim clsData As clsDSMAuthority
   
    strSQL = " Select a.groupid, a.deptfg, a.formid, a.readfg, a.writefg, a.printfg " & _
             " From  " & TB_COM009 & " a " & _
             " Where  " & DBS(mvarGroupID, 2) & _
             " and    " & DBS(mvarDeptFg, 2)
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        mvarPermitted = True
'        GetAuthority = False
        GoTo NoData
    End If
    
    With rs
        
        Call clear
        While (Not .EOF)
            Set clsData = New clsDSMAuthority
            
            clsData.DeptFg = "" & .Fields("deptfg").Value
            clsData.FormId = "" & .Fields("formid").Value
            clsData.ReadFg = "" & .Fields("readfg").Value
            clsData.WriteFg = "" & .Fields("writefg").Value
            clsData.PrintFg = "" & .Fields("printfg").Value
            
            Call Add(clsData.FormId, clsData)
            
            .MoveNext
        Wend
    End With
    
    mvarPermitted = True
    Set clsData = Nothing
    
NoData:
    rs.Close
    Set rs = Nothing
    
End Sub

Public Function LogIn(ByVal paraPassWd As String) As Boolean
'[Method] - 로그인 한 ID 와 비밀번호 비교
    
    LogIn = (paraPassWd = mvarLogInPass)
    
End Function

Private Sub Class_Initialize()
      
    mvarCanRead = False
    mvarCanWrite = False
    mvarCanPrint = False
    
    mvarIsManager = False
    mvarIsSupervisor = False
    mvarIsDeveloper = False
    mvarIsEndUser = False
    
    mvarLoginExist = False
    
    Set mcolAuthority = New Collection
    mvarCount = 0
    
End Sub

Public Sub clear()
    Set mcolAuthority = Nothing
    Set mcolAuthority = New Collection
    mvarCount = 0
End Sub


Public Function Add(ByVal clsKey As Variant, ByVal clsData As clsDSMAuthority)
    Dim newItem As New clsDSMAuthority
    
    Set newItem = clsData
    
    mcolAuthority.Add newItem, clsKey
    Set Add = newItem
    mvarCount = mvarCount + 1
    
End Function


