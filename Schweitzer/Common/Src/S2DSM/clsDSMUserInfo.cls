VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSMUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#########################
'작성일자 : 2000.10      #
'작 성 자 : BY M.G.Choi  #
'#########################

'인적정보 변수선언=========
Private mvarLogInID As String
Private mvarLogInNm As String
Private mvarEmpID As String
Private mvarEmpNm As String
Private mvarDeptCd As String
Private mvarDeptNm As String
Private mvarLogInPass As String
Private mvarLoginDesc As String
Private mvarGroupID As String
Private mvarGroupNm As String
'==========================
Private mvarEdit As String

Public Property Let Edit(ByVal vData As String)
    mvarEdit = vData
End Property

Public Property Get Edit() As String
    Edit = mvarEdit
End Property

Public Property Let LoginId(ByVal vData As String)
    mvarLogInID = vData
End Property

Public Property Get LoginId() As String
    LoginId = mvarLogInID
End Property

Public Property Let LogInNm(ByVal vData As String)
    mvarLogInNm = vData
End Property

Public Property Get LogInNm() As String
    LogInNm = mvarLogInNm
End Property

Public Property Let EmpId(ByVal vData As String)
    mvarEmpID = vData
End Property

Public Property Get EmpId() As String
    EmpId = mvarEmpID
End Property

Public Property Let EmpNM(ByVal vData As String)
    mvarEmpNm = vData
End Property

Public Property Get EmpNM() As String
    EmpNM = mvarEmpNm
End Property

Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

Public Property Let DeptNm(ByVal vData As String)
    mvarDeptNm = vData
End Property

Public Property Get DeptNm() As String
    DeptNm = mvarDeptNm
End Property

Public Property Let LogInPass(ByVal vData As String)
    mvarLogInPass = vData
End Property

Public Property Get LogInPass() As String
    LogInPass = mvarLogInPass
End Property

Public Property Let LoginDesc(ByVal vData As String)
    mvarLoginDesc = vData
End Property

Public Property Get LoginDesc() As String
    LoginDesc = mvarLoginDesc
End Property

Public Property Let GroupID(ByVal vData As String)
    mvarGroupID = vData
End Property

Public Property Get GroupID() As String
    GroupID = mvarGroupID
End Property

Public Property Let GroupNm(ByVal vData As String)
    mvarGroupNm = vData
End Property

Public Property Get GroupNm() As String
    GroupNm = mvarGroupNm
End Property

'ComBo Box Control 에 List Up  (Control Name, 구분) 구분==> '0':ID(Code), '1':Name
Public Function ComBo_List(ByRef Obj As Object, ByVal i As Integer)
'    Dim strSQL As String
'    Dim Rs As New Recordset
'    Dim ObjCbo As ComboBox
'
'    Set ObjCbo = Obj
'
'    ObjCbo.clear
'    Select Case i
'    Case 0
'         strSQL = " SELECT groupid FROM " & T_COM008 & " "
'    Case 1
'         strSQL = " SELECT groupnm FROM " & T_COM008 & " "
'    End Select
'
'    Rs.Open strSQL, DBConn
'
'    With ObjCbo
'        While Rs.EOF = False
'           .AddItem Trim(Rs.Fields(0).Value & "")
'           '.ListIndex = -1
'
'           Rs.MoveNext
'        Wend
'    End With
'
'    Set Rs = Nothing
End Function

Public Function COM010_Insert(ByVal LoginId As String, ByVal LogInPass, _
                ByVal EmpId As String, ByVal LoginDesc As String, ByVal GroupID As String)
'    Dim strSQL As String
'
'    On Error GoTo ErrCOM010
'    DBConn.BeginTrans
'
'    strSQL = " INSERT INTO " & T_COM010 & "(loginid, loginpass, empid, logindesc, groupid) " _
'           & "             values(" & DBV("loginid", LoginId) & " , " & DBV("loginpass", LogInPass) & " , " _
'           & "                    " & DBV("empid", EmpId) & " , " & DBV("logindesc", LoginDesc) & " , " & _
'                                      DBV("groupid", GroupID) & " ) "
'
'    DBConn.Execute strSQL
'
'    DBConn.CommitTrans
'    MsgBox "저장되었습니다.", vbInformation, "저장"
'    Exit Function
'
'ErrCOM010:
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbExclamation

End Function

Public Function COM010_UpDate(ByVal LoginId As String, ByVal LogInPass, _
                ByVal EmpId As String, ByVal LoginDesc As String, ByVal GroupID As String)
'    Dim strSQL As String
'
'    On Error GoTo ErrCOM010
'    DBConn.BeginTrans
'
'    strSQL = " UPDATE " & T_COM010 & " SET " & DBW("loginpass=", LogInPass) & " , " & DBW("empid=", EmpId) & ", " _
'           & "                   " & DBW("logindesc=", LoginDesc) & " , " & DBW("groupid=", GroupID) & "  " _
'           & "  WHERE " & DBW("loginid=", LoginId)
'
'    DBConn.Execute strSQL
'
'    DBConn.Execute strSQL
'    DBConn.CommitTrans
'    MsgBox "수정되었습니다.", vbInformation, "수정"
'    Exit Function
'
'ErrCOM010:
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbExclamation
    
End Function

Public Function COM010_Delete(ByVal LoginId As String)
'    Dim strSQL As String
'
'    On Error GoTo ErrCOM010
'    DBConn.BeginTrans
'
'    strSQL = " DELETE " & T_COM010 & " where " & DBW("loginid=", LoginId)
'
'    DBConn.Execute strSQL
'
'    DBConn.Execute strSQL
'    DBConn.CommitTrans
'    MsgBox "삭제되었습니다.", vbCritical, "삭제"
'    Exit Function
'
'ErrCOM010:
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbExclamation
    
End Function

Public Function COM010_InitCheck(ByVal strLoginId) As Boolean
'    Dim strSQL As String
'    Dim Rs As New Recordset
'
'    strSQL = " select * from " & T_COM010 _
'           & "  where " & DBW("loginid=", Trim(strLoginId))
'
'    Rs.Open strSQL, DBConn
'
'    If Rs.EOF = False Then
'       COM010_InitCheck = True
'    Else
'       COM010_InitCheck = False
'    End If
'
'    Set Rs = Nothing
End Function

Public Function EmpInfo(ByVal strEmpID As String) As Boolean
'    Dim strSQL As String
'    Dim Rs As New Recordset
'
'    strSQL = " select empid, empnm, deptcd from " & T_COM006 & _
'             "  where  " & DBW("empid =", strEmpID)
'
'    Rs.Open strSQL, DBConn
'
'    With Rs
'    If .EOF = False Then
'       mvarEmpID = "" & .Fields("empid").Value
'       mvarEmpNm = IIf(IsNull("" & .Fields("empnm").Value) = True, "", "" & .Fields("empnm").Value)
'       mvarDeptCd = IIf(IsNull("" & .Fields("deptcd").Value) = True, "", "" & .Fields("deptcd").Value)
'    End If
'    End With
'
'    Set Rs = Nothing
End Function

Public Function UserInfo(ByVal strUserID As String) As Boolean
'    Dim strSQL As String
'    Dim Rs As New Recordset
'
'    strSQL = " select c.loginid, a.empid, a.empnm, a.deptcd, c.loginpass, " _
'           & "        b.groupid, b.groupnm, c.logindesc " _
'           & "   from " & T_COM006 & " a, " & T_COM010 & " c, " & T_COM008 & " b " _
'           & "  where a.empid = c.empid " _
'           & "    and  " & DBW("c.loginid =", strUserID) & _
'           " and c.groupid=b.groupid "
'
'    Rs.Open strSQL, DBConn
'    With Rs
'         If Rs.EOF = False Then
'            LoginId = "" & .Fields("loginid").Value
'            EmpId = "" & .Fields("empid").Value
'            EmpNM = IIf(IsNull(.Fields("empnm").Value) = True, "", "" & .Fields("empnm").Value)
'            DeptCd = IIf(IsNull(.Fields("deptcd").Value) = True, "", "" & .Fields("deptcd").Value)
'            LogInPass = IIf(IsNull(.Fields("loginpass").Value) = True, "", "" & .Fields("loginpass").Value)
'            GroupID = IIf(IsNull(.Fields("groupid").Value) = True, "", "" & .Fields("groupid").Value)
'            GroupNm = IIf(IsNull(.Fields("groupnm").Value) = True, "", "" & .Fields("groupnm").Value)
'            LoginDesc = IIf(IsNull(.Fields("logindesc").Value) = True, "", "" & .Fields("logindesc").Value)
'         Else
'            LoginId = ""
'            EmpId = ""
'            EmpNM = ""
'            DeptCd = ""
'            LogInPass = ""
'            GroupID = ""
'            GroupNm = ""
'            LoginDesc = ""
'         End If
'    End With
'
'    Set Rs = Nothing
End Function

Public Function MessageList(i As Integer) As Boolean
'    Dim Message As String
'
'    Select Case i
'        Case 0
'             Message = MsgBox("저장 하시겠습니까?", vbInformation + vbOKCancel)
'             If Message = vbCancel Then
'                MessageList = False
'                Exit Function
'             End If
'
'        Case 1
'             Message = MsgBox("수정 하시겠습니까?", vbExclamation + vbOKCancel)
'             If Message = vbCancel Then
'                MessageList = False
'                Exit Function
'             End If
'
'        Case 2
'             MsgBox "이미 사용되고 있습니다.", vbInformation
'
'        Case 3
'             MsgBox "사용할 수 있습니다.", vbInformation
'    End Select
'
'    MessageList = True
End Function

Public Function Set_ComboList(ByRef Obj As Object, strCbo As String)
'    Dim i As Integer
'    Dim Cnt As Integer
'    Dim ObjCbo As ComboBox
'
'    Set ObjCbo = Obj
'
'    With ObjCbo
'         Cnt = .ListCount - 1
'
'         For i = 0 To Cnt
'             If .List(i) = Trim(strCbo) Then
'                .ListIndex = i
'                Exit Function
'             End If
'         Next i
'         .ListIndex = -1
'    End With
'
'    Set ObjCbo = Nothing
End Function

Public Sub SelFocus(ByRef Obj As Object)
'    Obj.SelStart = 0:    Obj.SelLength = Len(Obj.Text)
End Sub

'ListView Control 적용 (listview control name, 쿼리문, item 갯수)
'<< 주의 : 쿼리문 작성시 사용할 Item 순서와 필드선택 맞추어야 함 >>
'if 필드 수=1 then item 갯수=0
Public Function Lvw_Set(ByRef Obj As Object, ByVal strSQL As String, ByVal ItemCnt As Integer)
'    Dim Rs As New Recordset
'    Dim ItemX As ListItem
'    Dim i As Integer
'
'    Obj.ListItems.clear
'
'    Rs.Open strSQL, DBConn
'
''    If Rs.DBerror = True Then
''       'Call DisplayErrors
''       Set Rs = Nothing
''       MsgBox Err.Description, vbCritical, "일반 오류"
''       Exit Function
''    End If
'
'    While Rs.EOF = False
'          Set ItemX = Obj.ListItems.Add()
'          With ItemX
'               .Text = "" & Rs.Fields(0).Value
'               For i = 1 To ItemCnt
'                   .SubItems(i) = IIf(IsNull(Rs.Fields(i).Value) = True, "", "" & Rs.Fields(i).Value)
'               Next i
'          End With
'
'          Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
End Function

'ListView 에서 선택한 Item 지우기
Public Function Lvw_Item_Remove(ByRef Obj As Object, strItem As String)
'    Dim i As Integer
'    Dim lvwObj As ListView
'
'    Set lvwObj = Obj
'
'    With lvwObj
'         For i = 1 To .ListItems.Count
'             If .ListItems(i).Text = strItem Then
'                .ListItems.Remove (i)
'                Exit Function
'             End If
'         Next i
'    End With
'
'    Set lvwObj = Nothing
End Function
