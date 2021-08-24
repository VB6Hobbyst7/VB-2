VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpecManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarLeg()
Private mvarRow()
Private mvarCol()
'보관장소에 관한 멤버변수
Private mvarLegCd As String
Private mvarRowCnt As Long
Private mvarColCnt As Long
Private mvarRmk As String


Public Property Get Legcd() As String
    Legcd = mvarLegCd
End Property

Public Property Let Legcd(ByVal vData As String)
    mvarLegCd = vData
End Property

Public Property Get RowCnt() As Long
    RowCnt = mvarRowCnt
End Property

Public Property Let RowCnt(ByVal vData As Long)
    mvarRowCnt = vData
End Property

Public Property Get ColCnt() As Long
    ColCnt = mvarColCnt
End Property

Public Property Let ColCnt(ByVal vData As Long)
    mvarColCnt = vData
End Property

Public Property Get rmk() As String
    rmk = mvarRmk
End Property

Public Property Let rmk(ByVal vData As String)
    mvarRmk = vData
End Property

Public Property Get leg(ByVal Index As Long) As String
    leg = mvarLeg(Index)
End Property
Public Property Get Row(ByVal Index As Long) As Long
    Row = mvarRow(Index)
End Property
Public Property Get Col(ByVal Index As Long) As Long
    Col = mvarCol(Index)
End Property

Public Function IsExistSpecmens(ByVal Centercd As String, ByVal leg As String, ByVal Row As Long, ByVal Col As Long) As Boolean
    Dim objSql As clsBBSMSTStatement
    
    Set objSql = New clsBBSMSTStatement
    IsExistSpecmens = objSql.IsExistSpecmens(Centercd, leg, Row, Col)
    Set objSql = Nothing
End Function

Public Function SavePointChk(ByVal Legcd As String, ByVal Rowno As String, ByVal Colno As String, _
                             Centercd As String) As Boolean
'사용확인
    '입력받은 leg/rowno/colno의 존재여부를 판단한다.
    
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = " SELECT * FROM " & T_BBS206 & _
           " WHERE " & _
                               DBW("centercd", Centercd, 2) & _
                     " AND " & DBW("legcd", Legcd, 2) & _
                     " AND " & DBW("rowno", Rowno, 2) & _
                     " AND " & DBW("colno", Colno, 2) & _
           " AND not stscd in(" & _
                      DBV("stscd", BBSSaveStatue.stsNotUsed, 1) & _
                      DBV("stscd", BBSSaveStatue.stsNotUsedPoss) & ")"
                   
    Set RS = New Recordset
'    If RS.DBerror Then
'        RS.RsClose:   Set RS = Nothing
'        'dbconn.DisplayErrors
'        SavePointChk = False
'        Exit Function
'    End If
    RS.Open SSQL, DBConn
    If RS.EOF = True Then
        MsgBox "해당 보관장소가 없거나 이미 사용중인 보관장소입니다.", vbCritical + vbOKOnly, "검체보관장소"
        SavePointChk = False
        Exit Function
    End If
    SavePointChk = True
    Set RS = Nothing
End Function
Public Function frm501_Save(cntercd As String, Legcd As String, Rowno As Long, Colno As Long, SPCYY As String, SPCNO As Long)
    frm501_Save = " update " & T_BBS206 & " set " & _
                                            DBW("spcyy", SPCYY, 3) & _
                                            DBW("spcno", SPCNO, 3) & _
                                            DBW("stscd", BBSSaveStatue.stsNotUsed, 2) & _
                  " WHERE " & _
                                      DBW("centercd", cntercd, 2) & _
                            " AND " & DBW("legcd", Legcd, 2) & _
                            " AND " & DBW("rowno", Rowno, 2) & _
                            " AND " & DBW("colno", Colno, 2)
End Function
Public Function Save_Spc_Search(icnt As Integer, Centercd As String, _
                                Optional Legcd As String = "") As Boolean
    'icnt=icnt의 갯수만큼 저장할 장소를 구한다.
    'legcd가 정해져 있지 않으면, 모든 leg에 대해서 저장할수 있는 leg를 찾아온다.
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim cnt     As Integer
    
    SSQL = "SELECT legcd ,rowno,colno FROM " & T_BBS206
    If Legcd = "" Then
        SSQL = SSQL & _
                      " WHERE " & DBW("centercd", Centercd, 2) & _
                      " AND not stscd in(" & _
                                  DBV("stscd", BBSSaveStatue.stsNotUsed, 1) & _
                                  DBV("stscd", BBSSaveStatue.stsNotUsedPoss) & ")" & _
                      " ORDER BY legcd,rowno,colno"
    Else
        
        SSQL = SSQL & " WHERE " & DBW("centercd", Centercd, 2) & _
                      " AND not stscd in(" & _
                                   DBV("stscd", BBSSaveStatue.stsNotUsed, 1) & _
                                   DBV("stscd", BBSSaveStatue.stsNotUsedPoss) & ")" & _
                      " AND " & DBW("legcd", Legcd, 2) & _
                      " ORDER BY legcd,rowno,colno"
    End If
    Set RS = New Recordset
'    If RS.DBerror = True Then
'        RS.RsClose: Set RS = Nothing
'        Save_Spc_Search = False
'        Exit Function
'    End If
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        If RS.RecordCount < icnt Then
            MsgBox "검체저장 공간이 부족합니다. 보관장소를 확인후 저장하세요.", vbCritical + vbOKOnly, "검체저장"
            Set RS = Nothing
            Save_Spc_Search = False
            Exit Function
        Else
            ReDim mvarLeg(icnt)
            ReDim mvarRow(icnt)
            ReDim mvarCol(icnt)
            
            cnt = 0
            While Not RS.EOF And cnt <= icnt - 1
                cnt = cnt + 1
                mvarLeg(cnt) = RS.Fields("legcd").Value & ""
                mvarRow(cnt) = RS.Fields("rowno").Value & ""
                mvarCol(cnt) = RS.Fields("colno").Value & ""
                
                RS.MoveNext
            Wend
        End If
    Else
        MsgBox "검체저장 공간이 부족합니다. 보관장소를 확인후 저장하세요.", vbCritical + vbOKOnly, "검체저장"
        Set RS = Nothing
        Save_Spc_Search = False
        Exit Function
    End If
    Save_Spc_Search = True
    Set RS = Nothing
End Function

Public Function GetLegList(ByVal Centercd As String, LegList() As String) As Long
    '----------------------------------------
    '현재 관리되는 Leg코드를 Array에 담는다.
    '코드 갯수 반환
    '----------------------------------------
    Dim i       As Long
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS003 & " " & _
           "WHERE " & DBW("centercd", Centercd, 2)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
'    If RS.DBerror = True Then
'        'dbconn.DisplayErrors
'        GetLegList = 0
'        Set RS = Nothing
'        Exit Function
'    End If
    
    If RS.RecordCount < 1 Then
'        RS.RsClose
        Set RS = Nothing
        GetLegList = 0
        Exit Function
    End If
    
    GetLegList = RS.RecordCount
    ReDim LegList(RS.RecordCount - 1)
    
    For i = 1 To GetLegList
        LegList(i - 1) = RS.Fields("legcd").Value & ""
        RS.MoveNext
    Next i
    
    Set RS = Nothing
End Function

Public Function ReadLegInfo(ByVal Centercd As String, ByVal Legcd As String) As Boolean
    '-----------------------------------------------
    '특정 Leg에 대한 정보를 읽어서 멤버변수에 담는다
    '-----------------------------------------------
    Dim SSQL As String
    Dim RS As Recordset
    
    SSQL = " SELECT * " & _
           " FROM " & T_BBS003 & " " & _
           " WHERE " & DBW("centercd", Centercd, 2) & " " & _
           " AND   " & DBW("legcd", Legcd, 2)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
'    If RS.DBerror = True Then
'        'dbconn.DisplayErrors
'        ReadLegInfo = False
'        Set RS = Nothing
'        Exit Function
'    End If
    
    mvarLegCd = Legcd
    mvarRowCnt = 0
    mvarColCnt = 0
    mvarRmk = ""
    
    If RS.RecordCount > 0 Then
        mvarLegCd = RS.Fields("legcd").Value & ""
        mvarRowCnt = RS.Fields("rowcnt").Value & ""
        mvarColCnt = RS.Fields("colcnt").Value & ""
        mvarRmk = RS.Fields("rmk").Value & ""
        
        ReadLegInfo = True
    Else
        ReadLegInfo = False
    End If
    Set RS = Nothing
    
End Function

Public Function GetSpcKeepSpace(ByVal Centercd As String, ByVal Legcd As String) As Recordset
    '--------------------------------------------------------
    '보관장소마스터에서 특정 leg에 대한 정보를 모두 읽어온다.
    '--------------------------------------------------------
    Dim SSQL As String
    
    SSQL = "SELECT legcd,rowno,colno,spcyy,spcno,stscd " & _
           "FROM " & T_BBS206 & " " & _
           "WHERE" & _
                   "    " & DBW("centercd", Centercd, 2) & " " & _
                   "AND " & DBW("legcd", Legcd, 2) & " " & _
                   "ORDER BY legcd,rowno,colno "
    Set GetSpcKeepSpace = New Recordset
    GetSpcKeepSpace.Open SSQL, DBConn
    
'    If GetSpcKeepSpace.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set GetSpcKeepSpace = Nothing
'    End If
    
End Function

Public Function OpenBBS201(ByVal SPCYY As String, ByVal SPCNO As String) As Recordset
    Dim SSQL As String
    
    SSQL = "SELECT * " & _
           "FROM " & T_BBS201 & " " & _
           "WHERE " & DBW("spcyy", SPCYY, 2) & " " & _
           "AND   " & DBW("spcno", SPCNO, 2)
    Set OpenBBS201 = New Recordset
    OpenBBS201.Open SSQL, DBConn
    
'    If OpenBBS201.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set OpenBBS201 = Nothing
'        Exit Function
'    End If
    
End Function

Public Function Expire(ByVal Centercd As String, ByVal SPCYY As String, ByVal SPCNO As String, Optional expfg As String = "1") As Boolean
    Dim Legcd   As String
    Dim Rowno   As Long
    Dim Colno   As Long
    Dim stscd   As String
    Dim SSQL    As String
    Dim asSql() As String
    Dim RS      As Recordset
    
    
    '검체번호로 저장장소를 알아낸다--------------------------------
    SSQL = " SELECT storeleg,storerno,storecno " & _
           " FROM " & T_BBS201 & " " & _
           " WHERE " & DBW("spcyy", SPCYY, 2) & " " & _
           " AND   " & DBW("spcno", SPCNO, 2) & " " & _
           " AND   " & DBW("buildcd", Centercd, 2)

    Set RS = New Recordset

'    If RS.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set RS = Nothing
'        Expire = False
'        Exit Function
'    End If
    RS.Open SSQL, DBConn
    If RS.RecordCount < 1 Then
        MsgBox "검체정보에 이상이 있읍니다.", vbCritical, "오류"
        Set RS = Nothing
        Expire = False
        Exit Function
    End If
    
    Legcd = RS.Fields("storeleg").Value & ""
    Rowno = RS.Fields("storerno").Value & ""
    Colno = RS.Fields("storecno").Value & ""
    
    Set RS = Nothing
    
    '폐기처리------------------------------------------------------
    '        채혈접수내역(BBS201)   expfg = '1'
    '        보관검체마스터(BBS206) stscd = if expfg=1 then '0'
    '                                       if expfg=2 then '2'
    '--------------------------------------------------------------
    If expfg = "1" Then
        stscd = "0"
    Else
        stscd = "2"
    End If
    
    ReDim asSql(1)
    asSql(0) = "update " & T_BBS201 & " " & _
               "set expfg" & DBV("expfg", "1", 2) & _
               "WHERE" & _
                       "    " & DBW("spcyy", SPCYY, 2) & " " & _
                       "AND " & DBW("spcno", SPCNO, 2) & " " & _
                       "AND " & DBW("buildcd", Centercd, 2)
                       
    asSql(1) = "update " & T_BBS206 & " " & _
               "set  " & DBW("stscd", stscd, 2) & " " & _
               "WHERE" & _
                       "    " & DBW("centercd", Centercd, 2) & " " & _
                       "AND " & DBW("legcd", Legcd, 2) & " " & _
                       "AND " & DBW("rowno", Rowno, 2) & " " & _
                       "AND " & DBW("colno", Colno, 2)
    
On Error GoTo Expire_error

    DBConn.BeginTrans
    DBConn.Execute asSql(0)
    DBConn.Execute asSql(1)
    DBConn.CommitTrans
    
    Expire = True
    Exit Function
    
Expire_error:

    DBConn.RollbackTrans
    
    Expire = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function ExpireAndKeep(ByVal Centercd As String, ByVal SPCYY As String, ByVal SPCNO As String) As Boolean
    ExpireAndKeep = Expire(Centercd, SPCYY, SPCNO, "2")
End Function

Public Function SpecMove(ByVal Centercd As String, ByVal SPCYY As String, ByVal SPCNO As String, ByVal Legcd As String, ByVal Rowno As String, ByVal Colno As String) As Boolean
    Dim orglegcd As String
    Dim orgrowno As String
    Dim orgcolno As String
    Dim orgstscd As String
    
    Dim SSQL As String
    Dim asSql() As String
    Dim RS As Recordset
    

    '이동해갈 장소에 다른 검체가 있는지 검사한다-------------------------------------
    SSQL = "SELECT spcyy,spcno,stscd " & _
           "FROM " & T_BBS206 & " " & _
           "WHERE " & DBW("centercd", Centercd, 2) & " " & _
           "AND " & DBW("legcd", Legcd, 2) & " " & _
           "AND " & DBW("rowno", Rowno, 2) & " " & _
           "AND " & DBW("colno", Colno, 2)
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
'    If RS.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set RS = Nothing
'        SpecMove = False
'        Exit Function
'    End If

    If RS.RecordCount < 1 Then
        MsgBox "존재하지 않는 보관장소로 이동을 시도하셨읍니다", vbCritical, "오류"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    If RS.Fields("stscd").Value & "" <> "0" Then
        MsgBox "이미 검체가 있는 장소입니다.", vbCritical, "오류"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    
    Set RS = Nothing
    

    '검체번호로 저장장소를 알아낸다-------------------------------------------------
    SSQL = "SELECT storeleg,storerno,storecno " & _
           "FROM " & T_BBS201 & " " & _
           "WHERE " & DBW("spcyy", SPCYY, 2) & " " & _
           "AND   " & DBW("spcno", SPCNO, 2) & " " & _
           "AND   " & DBW("buildcd", Centercd, 2)
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
'    If RS.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set RS = Nothing
'        SpecMove = False
'        Exit Function
'    End If
    
    If RS.RecordCount < 1 Then
        MsgBox "검체정보에 이상이 있읍니다.", vbCritical, "오류"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    orglegcd = RS.Fields("storeleg").Value & ""
    orgrowno = RS.Fields("storerno").Value & ""
    orgcolno = RS.Fields("storecno").Value & ""
    
    Set RS = Nothing
    
    
    '현 검체의 상태를 파악한다------------------------------------------------------
    SSQL = "SELECT stscd " & _
           "FROM " & T_BBS206 & " " & _
           "WHERE " & DBW("centercd", Centercd, 2) & " " & _
           "AND   " & DBW("legcd", orglegcd, 2) & " " & _
           "AND   " & DBW("rowno", orgrowno, 2) & " " & _
           "AND   " & DBW("colno", orgcolno, 2)

    Set RS = New Recordset
    RS.Open SSQL, DBConn
'    If RS.DBerror = True Then
'        'dbconn.DisplayErrors
'        Set RS = Nothing
'        SpecMove = False
'        Exit Function
'    End If
    
    If RS.RecordCount < 1 Then
        MsgBox "검체정보에 이상이 있읍니다.", vbCritical, "오류"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    orgstscd = RS.Fields("stscd").Value & ""
    
    Set RS = Nothing
    
    
    '검체이동-----------------------------------------------------------------------
    '        1. 채혈접수내역  (BBS201)의 보관장소변경
    '        2. 보관검체마스터(BBS206)의 원 보관장소 Clear
    '        3. 보관검체마스터(BBS206)의 새 보관장소에 Set
    '        4. 길병원의 경우 보관장소이동시 접수번호에 묶여있는 보관장소도 이동시킨다
    '-------------------------------------------------------------------------------
    ReDim asSql(2)
    asSql(0) = " update " & T_BBS201 & " " & _
               " set " & DBW("storeleg", Legcd, 3) & _
               "     " & DBW("storerno", Rowno, 3) & _
               "     " & DBW("storecno", Colno, 2) & _
               " WHERE " & DBW("spcyy", SPCYY, 2) & _
               " AND   " & DBW("spcno", SPCNO, 2)
               
    asSql(1) = " update " & T_BBS206 & " " & _
               " set " & DBW("spcyy", SPCYY, 3) & _
               "     " & DBW("spcno", SPCNO, 3) & _
               "     " & DBW("stscd", orgstscd, 2) & _
               " WHERE " & DBW("centercd", Centercd, 2) & _
               " AND   " & DBW("legcd", Legcd, 2) & _
               " AND   " & DBW("rowno", Rowno, 2) & _
               " AND   " & DBW("colno", Colno, 2)
               
    asSql(2) = " update " & T_BBS206 & " " & _
               " set " & DBW("spcyy", "", 3) & _
               "     " & DBW("spcno", "0", 3) & _
               "     " & DBW("stscd", "0", 2) & _
               " WHERE " & DBW("centercd", Centercd, 2) & _
               " AND   " & DBW("legcd", orglegcd, 2) & _
               " AND   " & DBW("rowno", orgrowno, 2) & " " & _
               " AND   " & DBW("colno", orgcolno, 2)
        
On Error GoTo SpecMove_error

    DBConn.BeginTrans
    DBConn.Execute asSql(0)
    DBConn.Execute asSql(1)
    DBConn.Execute asSql(2)

    DBConn.CommitTrans
    
    SpecMove = True
    Exit Function
    
SpecMove_error:

    DBConn.RollbackTrans
    
    SpecMove = False
    MsgBox Err.Description, vbExclamation
End Function

Public Function SetALL_Exp(ByVal Centercd As String, ByVal Legcd As String, ByVal Rowno As String, ByVal Colno As String) As String
    
    SetALL_Exp = " update " & T_BBS206 & " set " & DBW("stscd=", BBSSaveStatue.stsUsed, 1) & DBW("spcyy=", "", 1) & DBW("spcno=", "") & _
                 " WHERE " & _
                           DBW("centercd=", Centercd) & _
                 " AND " & DBW("legcd=", Legcd) & _
                 " AND " & DBW("rowno=", Rowno) & _
                 " AND " & DBW("colno=", Colno)
End Function
Public Function SetSpcExpString(ByVal SPCYY As String, ByVal SPCNO As String) As String
    SetSpcExpString = " update " & T_BBS201 & " set " & DBW("expfg=", "1") & _
                     " WHERE " & _
                            DBW("spcyy=", SPCYY) & " AND " & DBW("spcno=", SPCNO)
End Function
