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
'������ҿ� ���� �������
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
'���Ȯ��
    '�Է¹��� leg/rowno/colno�� ���翩�θ� �Ǵ��Ѵ�.
    
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
        MsgBox "�ش� ������Ұ� ���ų� �̹� ������� ��������Դϴ�.", vbCritical + vbOKOnly, "��ü�������"
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
    'icnt=icnt�� ������ŭ ������ ��Ҹ� ���Ѵ�.
    'legcd�� ������ ���� ������, ��� leg�� ���ؼ� �����Ҽ� �ִ� leg�� ã�ƿ´�.
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
            MsgBox "��ü���� ������ �����մϴ�. ������Ҹ� Ȯ���� �����ϼ���.", vbCritical + vbOKOnly, "��ü����"
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
        MsgBox "��ü���� ������ �����մϴ�. ������Ҹ� Ȯ���� �����ϼ���.", vbCritical + vbOKOnly, "��ü����"
        Set RS = Nothing
        Save_Spc_Search = False
        Exit Function
    End If
    Save_Spc_Search = True
    Set RS = Nothing
End Function

Public Function GetLegList(ByVal Centercd As String, LegList() As String) As Long
    '----------------------------------------
    '���� �����Ǵ� Leg�ڵ带 Array�� ��´�.
    '�ڵ� ���� ��ȯ
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
    'Ư�� Leg�� ���� ������ �о ��������� ��´�
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
    '������Ҹ����Ϳ��� Ư�� leg�� ���� ������ ��� �о�´�.
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
    
    
    '��ü��ȣ�� ������Ҹ� �˾Ƴ���--------------------------------
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
        MsgBox "��ü������ �̻��� �����ϴ�.", vbCritical, "����"
        Set RS = Nothing
        Expire = False
        Exit Function
    End If
    
    Legcd = RS.Fields("storeleg").Value & ""
    Rowno = RS.Fields("storerno").Value & ""
    Colno = RS.Fields("storecno").Value & ""
    
    Set RS = Nothing
    
    '���ó��------------------------------------------------------
    '        ä����������(BBS201)   expfg = '1'
    '        ������ü������(BBS206) stscd = if expfg=1 then '0'
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
    

    '�̵��ذ� ��ҿ� �ٸ� ��ü�� �ִ��� �˻��Ѵ�-------------------------------------
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
        MsgBox "�������� �ʴ� ������ҷ� �̵��� �õ��ϼ����ϴ�", vbCritical, "����"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    If RS.Fields("stscd").Value & "" <> "0" Then
        MsgBox "�̹� ��ü�� �ִ� ����Դϴ�.", vbCritical, "����"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    
    Set RS = Nothing
    

    '��ü��ȣ�� ������Ҹ� �˾Ƴ���-------------------------------------------------
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
        MsgBox "��ü������ �̻��� �����ϴ�.", vbCritical, "����"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    orglegcd = RS.Fields("storeleg").Value & ""
    orgrowno = RS.Fields("storerno").Value & ""
    orgcolno = RS.Fields("storecno").Value & ""
    
    Set RS = Nothing
    
    
    '�� ��ü�� ���¸� �ľ��Ѵ�------------------------------------------------------
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
        MsgBox "��ü������ �̻��� �����ϴ�.", vbCritical, "����"
        Set RS = Nothing
        SpecMove = False
        Exit Function
    End If
    
    orgstscd = RS.Fields("stscd").Value & ""
    
    Set RS = Nothing
    
    
    '��ü�̵�-----------------------------------------------------------------------
    '        1. ä����������  (BBS201)�� ������Һ���
    '        2. ������ü������(BBS206)�� �� ������� Clear
    '        3. ������ü������(BBS206)�� �� ������ҿ� Set
    '        4. �溴���� ��� ��������̵��� ������ȣ�� �����ִ� ������ҵ� �̵���Ų��
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