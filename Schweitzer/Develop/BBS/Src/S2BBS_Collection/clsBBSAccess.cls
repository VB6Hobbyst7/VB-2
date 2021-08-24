VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'������ ó����� ���� =======KJG(2001/12/12)
Private stsORDER   As String
Private stsCOLLECT As String
Private stsACCESS  As String

Private mvarLeg()
Private mvarRow()
Private mvarCol()

Public Property Get Leg(ByVal index As Long) As String
    Leg = mvarLeg(index)
End Property
Public Property Get Row(ByVal index As Long) As Long
    Row = mvarRow(index)
End Property
Public Property Get Col(ByVal index As Long) As Long
    Col = mvarCol(index)
End Property

Public Function Get_LegPos(ByVal Centercd As String) As String
    'Leg��
    Get_LegPos = "SELECT distinct legcd FROM " & T_BBS206 & " " & _
                 "WHERE " & DBW("centercd=", Centercd)
End Function

'==============================
'��ü�������� ����� DB SQL����
'==============================


Public Function Set_UpdateB201(ByVal spcyy As String, ByVal spcno As Long, _
                               ByVal rcvid As String, ByVal legcd As String, ByVal rowno As Long, _
                               ByVal colno As Long) As String
    
'------------------------------
'ä�������� ���������� �����Ѵ�
'------------------------------
    Dim strRcvdt As String
    Dim strrcvtm As String
    
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
    Set_UpdateB201 = " update " & T_BBS201 & " set " & _
                                               DBW("rcvdt", strRcvdt, 3) & _
                                               DBW("rcvtm", strrcvtm, 3) & _
                                               DBW("rcvid", rcvid, 3) & _
                                               DBW("storeleg", legcd, 3) & _
                                               DBW("storerno", rowno, 3) & _
                                               DBW("storecno", colno, 2) & _
                        " WHERE " & _
                                "     " & DBW("spcyy", spcyy, 2) & _
                                " AND " & DBW("spcno", spcno, 2)
                        
End Function

Public Function Set_UpdateB206(ByVal Centercd As String, ByVal legcd As String, _
                               ByVal rowno As Long, ByVal colno As Long, _
                               ByVal spcyy As String, ByVal spcno As Long) As String
'-------------------
'��ü������� update
'-------------------
    
    Set_UpdateB206 = " update " & T_BBS206 & " set " & _
                                               DBW("spcyy", spcyy, 3) & _
                                               DBW("spcno", spcno, 3) & _
                                               DBW("stscd", BBSSaveStatue.stsNotUsed, 2) & _
             " WHERE " & _
                                 DBW("centercd", Centercd, 2) & _
                       " AND " & DBW("legcd", legcd, 2) & _
                       " AND " & DBW("rowno", rowno, 2) & _
                       " AND " & DBW("colno", colno, 2)
             
End Function
Public Function Set_UpdateL101(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long) As String
'---------------
'ó����� Update
'---------------
'    MsgBox "ó����� ������Ʈ", vbExclamation
    
'    Set_UpdateL101 = " update " & T_LAB101 & " set " & _
'                                               DBW("donefg", stsACCESS, 2) & _
'                    " WHERE " & _
'                    "     " & DBW("ptid", ptid, 2) & _
'                    " AND " & DBW("orddt", orddt, 2) & _
'                    " AND " & DBW("ordno", OrdNo, 2)
    
    '���� ���������� ó�� ������ ���� ���
    Set_UpdateL101 = " update mdbldort set " & DBW("donefg", stsACCESS, 2) & _
                    " where " & DBW("patno=", ptid) & _
                    " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                    " and " & DBW("ordseqno=", OrdNo)
End Function
Public Function Set_UpdateL102(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long, _
                               ByVal ordseq As Long, ByVal accdt As String, ByVal accseq As Long) As String
'---------------
'ó��ٵ� update
'---------------
    Dim strRcvdt  As String
    Dim strrcvtm  As String
    
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
'    Set_UpdateL102 = " update " & T_LAB102 & " set " & _
'                                               DBW("rcvdt", strRcvdt, 3) & _
'                                               DBW("rcvtm", strrcvtm, 3) & _
'                                               DBW("donefg", stsACCESS, 3) & _
'                                               DBW("stscd", stsACCESS, 3) & _
'                                               DBW("workarea", C_WORKAREA, 3) & _
'                                               DBW("accdt", accdt, 3) & _
'                                               DBW("accseq", accseq, 2) & _
'                            " WHERE " & _
'                                    "     " & DBW("ptid", ptid, 2) & _
'                                    " AND " & DBW("orddt", orddt, 2) & _
'                                    " AND " & DBW("ordno", OrdNo, 2) & _
'                                    " AND " & DBW("ordseq", ordseq, 2)
             
    '���� ���������� ó�� ������ ���� ���
'�������� ������
'    Set_UpdateL102 = " UPDATE mdbldort SET " & _
'                                               " acptdate=to_date(" & strRcvdt & strrcvtm & ",'yyyymmdd hh24:mi:ss') ," & _
'                                               DBW("donefg", stsACCESS, 3) & _
'                                               DBW("stscd", stsACCESS, 3) & _
'                                               DBW("workarea", C_WORKAREA, 3) & _
'                                               DBW("accdt", accdt, 3) & _
'                                               DBW("accseq", accseq, 2) & _
'                            " WHERE " & DBW("patno=", ptid) & _
'                                    " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
'                                    " AND  " & DBW("ordseqno=", OrdNo)
'�������� ����
'OCS�� ���¸� C�� ������ �ӻ󺴸� Status E�� �ٸ�

    Set_UpdateL102 = " UPDATE mdbldort SET " & _
                                               " acptdate=to_date(" & strRcvdt & strrcvtm & ",'yyyymmdd hh24:mi:ss') ," & _
                                               DBW("donefg", stsACCESS, 3) & _
                                               DBW("stscd", stsACCESS, 3) & _
                                               DBW("workarea", C_WORKAREA, 3) & _
                                               DBW("accdt", accdt, 3) & _
                                               DBW("accseq", accseq, 3) & _
                                               DBW("procstat", "C", 2) & _
                            " WHERE " & DBW("patno=", ptid) & _
                                    " and orddate=to_date(" & orddt & " ,'yyyymmdd') " & _
                                    " AND  " & DBW("ordseqno=", OrdNo)
             
End Function
Public Function Set_UpdateB202(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long, ByVal ordseq As Long) As String
'---------------------------
'ó�� ���������� Update �Ѵ�
'---------------------------
    
    Set_UpdateB202 = " update " & T_BBS202 & " set " & DBW("stscd", stsACCESS, 3) & DBW("canceldt", "", 3) & DBW("canceltm", "", 3) & _
                                                       DBW("cancelid", "", 3) & DBW("cancelfg", "0", 3) & DBW("cancelrsn", "", 2) & _
                     " WHERE " & _
                                DBW("ptid", ptid, 2) & " AND " & DBW("orddt", orddt, 2) & " AND " & DBW("ordno", OrdNo, 2) & " AND " & DBW("ordseq", ordseq, 2)
                             
End Function
Public Function GetSpcSavePosition(ByVal spcyy As String, ByVal spcno As Long, ByRef strReg As String, ByRef strCol As Long, ByRef strRow As Long)
    Dim sSql As String
    Dim RS   As Recordset
    
    sSql = " SELECT storeleg,storecno,storerno FROM " & T_BBS201 & " WHERE " & DBW("spcyy=", spcyy) & " AND " & DBW("spcno=", spcno)
    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    If RS.EOF Then
        strReg = "": strCol = 0: strRow = 0
    Else
        strReg = RS.Fields("storeleg").Value & ""
        strCol = RS.Fields("storecno").Value & ""
        strRow = RS.Fields("storerno").Value & ""
    End If
    
    Set RS = Nothing
    
End Function


Public Function Set_InsertB202(ByVal accdt As String, ByVal accseq As Long, ByVal ptid As String, _
                               ByVal orddt As String, ByVal OrdNo As Long, _
                               ByVal ordseq As Long, ByVal rcvid As String) As String
'---------------------------
'ó�� ���������� Insert �Ѵ�
'---------------------------
    Dim strrcvtm As String
    Dim strRcvdt As String
    
    strrcvtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strRcvdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    Set_InsertB202 = " insert into " & T_BBS202 & "(" & _
                    " workarea,accdt,accseq,ptid,orddt,ordno,ordseq,rcvdt,rcvtm,rcvid,stscd," & _
                    " canceldt,canceltm,cancelid,cancelfg,cancelrsn)" & _
                    " values(" & _
                    DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & _
                    DBV("ptid", ptid, 1) & DBV("orddt", orddt, 1) & _
                    DBV("ordno", OrdNo, 1) & DBV("ordseq", ordseq, 1) & DBV("rcvdt", strRcvdt, 1) & _
                    DBV("rcvtm", strrcvtm, 1) & DBV("rcvid", rcvid, 1) & _
                    DBV("stscd", stsACCESS, 1) & DBV("canceldt", "", 1) & _
                    DBV("canceltm", "", 1) & DBV("cancelid", "", 1) & DBV("cancelfg", "0", 1) & _
                    DBV("cancelrsn", "") & ")"
                 


End Function
Public Function Set_AccessUpdate(accdt As String, accseq As Long) As String
'--------------------------
'��ȣ�ο����� update/insert
'--------------------------
    Dim RS As New Recordset
    Dim sSql As String
    
    sSql = " SELECT divcd2,seq " & _
           " FROM " & T_COM099 & _
           " WHERE  " & _
                                DBW("noindex", BN_ACC_NO, 2) & _
                    " AND   " & DBW("divcd1", C_WORKAREA, 2) & _
                    " AND   " & DBW("divcd2", accdt, 2) & _
                    " AND   " & DBW("divcd3", "0", 2)
           
    RS.Open sSql, DBConn
     
    If Not RS.EOF = True Then
        Set_AccessUpdate = " update " & T_COM099 & _
                           " set " & DBW("seq", accseq, 2) & _
                           " WHERE  " & _
                                       DBW("noindex", BN_ACC_NO, 2) & _
                           " AND   " & DBW("divcd1", C_WORKAREA, 2) & _
                           " AND   " & DBW("divcd2", accdt, 2) & _
                           " AND   " & DBW("divcd3", "0", 2)
    
    Else
        Set_AccessUpdate = "insert into " & T_COM099 & "( " & _
                            "noindex,divcd1,divcd2,divcd3,seq) " & _
                            "values(" & _
                            DBV("noindex", BN_ACC_NO, 1) & DBV("divcd1", C_WORKAREA, 1) & _
                            DBV("divcd2", accdt, 1) & DBV("divcd3", "0", 1) & _
                            DBV("seq", accseq) & ") "
    End If
    Set RS = Nothing
End Function


Public Function Get_AccessNo(ByVal accdt As String) As Long
'----------------------------
'��밡���� ������ȣ ��������
'----------------------------
    Dim sSql As String
    Dim RS As New Recordset
    
    sSql = "SELECT divcd2,seq " & _
           "FROM " & T_COM099 & " " & _
           "WHERE " & _
                               DBW("noindex", BN_ACC_NO, 2) & _
                    " AND  " & DBW("divcd1", C_WORKAREA, 2) & _
                    " AND  " & DBW("divcd2", accdt, 2) & _
                    " AND  " & DBW("divcd3", "0", 2)
    RS.Open sSql, DBConn
    
    If RS.EOF Then
        Get_AccessNo = 1
    Else
        Get_AccessNo = CLng(RS.Fields("seq").Value & "") + 1
    End If
    Set RS = Nothing

End Function
Public Function Get_Position(ByVal icnt As Integer, _
                             ByVal Centercd As String, _
                             Optional legcd As String = "") As Boolean
'----------------------------------------------
'��밡���� ������� ��������
'icnt�ǰ���ŭ ��밡���� ������Ҹ� ������ �´�
'----------------------------------------------
    Dim RS As New Recordset
    Dim strSql As String
    Dim cnt As Integer
    
    strSql = "SELECT legcd ,rowno,colno FROM " & T_BBS206
    If legcd = "" Then
        strSql = strSql & " WHERE " & DBW("centercd", Centercd, 2) & _
                          " AND not stscd in(" & _
                                        DBV("stscd", BBSSaveStatue.stsNotUsed, 1) & _
                                        DBV("stscd", BBSSaveStatue.stsNotUsedPoss) & ") " & _
                          " ORDER BY legcd,rowno,colno"
    Else
        strSql = strSql & " WHERE centercd" & DBV("centercd", Centercd, 2) & _
                          " AND not stscd in(" & _
                                        DBV("stscd", BBSSaveStatue.stsNotUsed, 1) & _
                                        DBV("stscd", BBSSaveStatue.stsNotUsedPoss) & ") " & _
                          " AND " & DBW("legcd", legcd, 2) & _
                          " ORDER BY legcd,rowno,colno"
    End If
    
    Set RS = New Recordset
    RS.Open strSql, DBConn
    
    If Not RS.EOF Then
        If RS.RecordCount < icnt Then
            MsgBox "��ü���� ������ �����մϴ�. ������Ҹ� Ȯ���� �����ϼ���.", vbCritical + vbOKOnly, "��ü����"
            Set RS = Nothing
            Get_Position = False
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
        Get_Position = False
        Exit Function
    End If
    Get_Position = True
    Set RS = Nothing
End Function
Public Function SavePointChk(ByVal legcd As String, _
                             ByVal rowno As Long, _
                             ByVal colno As Long, _
                             ByVal Centercd As String) As Boolean
'------------------------------
'��������� ��뿩�θ� �Ǵ��Ѵ�
'------------------------------
    
    Dim strSql As String
    Dim RS As New Recordset
    
    strSql = " SELECT * FROM " & T_BBS206 & _
           " WHERE " & _
                             DBW("centercd", Centercd, 2) & _
                   " AND " & DBW("legcd", legcd, 2) & _
                   " AND " & DBW("rowno", rowno, 2) & _
                   " AND " & DBW("colno", colno, 2) & _
           " AND not stscd in(" & _
                          DBV("stscd", BBSSaveStatue.stsNotUsed, 1) & _
                          DBV("stscd", BBSSaveStatue.stsNotUsedPoss) & ") "
    RS.Open strSql, DBConn
    If RS.EOF = True Then
        MsgBox "�ش� ������Ұ� ���ų� �̹� ������� ��������Դϴ�.", vbCritical + vbOKOnly, "��ü�������"
        Exit Function
    End If
    SavePointChk = True
    Set RS = Nothing
End Function
Public Function GetAccCancelSpcInfo(ByVal spcyy As String, ByVal spcno As String) As Recordset
    '-------------------------------------------------------------------
    '��ü��ȣ�� ������, ä���������� ȯ��Id,��������,�����ð��� �ҷ��´�.
    '-------------------------------------------------------------------
    
    Dim sSql As String
    Dim RS   As Recordset
    
    sSql = " SELECT ptid,rcvdt,rcvtm FROM " & T_BBS201 & " WHERE " & DBW("spcyy=", spcyy) & " AND " & DBW("spcno=", spcno)
    
    Set GetAccCancelSpcInfo = New Recordset
    GetAccCancelSpcInfo.Open sSql, DBConn

End Function
Public Function GetAccCancelOrdInfo(ByVal ptid As String, ByVal rcvdt As String, ByVal rcvtm As String) As String
'----------------------------------------------
'ó�������������� ������ҵ� ��ü�� �ִ��� ��ȸ
'----------------------------------------------

    Dim sSql As String
    
    sSql = " SELECT b.testnm,c.ordcd,c.unitqty,c.orddt,d.reqdt,c.ordno,c.ordseq,c.ptid,c.accdt,c.accseq,c.dcfg,d.orddiv FROM " & _
           T_LAB101 & " d," & T_LAB102 & " c," & T_BBS001 & " b," & T_BBS202 & " a " & _
           "WHERE " & _
                  DBW("a.ptid=", ptid) & " AND " & DBW("a.rcvdt=", rcvdt) & "  AND " & DBW("a.rcvtm=", rcvtm) & " AND  " & _
                  DBW("a.cancelfg=", BBSCancelStatus.stsCancel) & " AND " & DBW("a.stscd=", stsCOLLECT) & " AND " & _
                  "a.ptid=c.ptid AND a.orddt=c.orddt AND a.ordno=c.ordno AND a.ordseq=c.ordseq AND " & _
                  "c.ptid=d.ptid AND c.orddt=d.orddt AND c.ordno=d.ordno AND " & _
                  "c.ordcd=b.testcd AND " & _
                  "b.applydt=(SELECT max(e.applydt) FROM " & T_BBS001 & " e " & _
                             "WHERE  b.testcd = e.testcd AND " & _
                             DBW("e.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & " )"
    
     GetAccCancelOrdInfo = sSql

End Function


Public Function Spc_ExistChk(ByVal spcyy As String, ByVal spcno As Long) As Boolean
'--------------------------
'��ü�� �������θ� �Ǵ��Ѵ�
'--------------------------
    Dim sSql As String
    Dim RS As Recordset
    
    sSql = " SELECT * FROM " & T_BBS201 & _
           " WHERE " & _
                     "     " & DBW("spcyy", spcyy, 2) & _
                     " AND " & DBW("spcno", spcno, 2) & _
                     " AND (rcvdt is null or " & DBW("rcvdt", "", 2) & ")" & _
                     " AND " & DBW("expfg<>", BBSExpFg.stsExpfg)
            
    Set RS = New Recordset
    RS.Open sSql, DBConn
    
    If Not RS.EOF Then
        Spc_ExistChk = True
    End If
    Set RS = Nothing
End Function
Public Function Get_AccDtSeq(ByVal spcyy As String, ByVal spcno As Long) As String
'-----------------------------------
'��ü��ȣ�� ������ ������ȣ�� ���Ѵ�
'-----------------------------------
    Dim RS As New Recordset
    Dim sSql As String
    
    sSql = " SELECT accdt,accno FROM " & T_BBS207 & _
             " WHERE " & _
             "     " & DBW("spcyy", spcyy, 2) & _
             " AND " & DBW("spcno", spcno, 2)
    
    RS.Open sSql, DBConn
    If RS.EOF = False Then
        Get_AccDtSeq = Trim(RS.Fields("accdt").Value & "") & COL_DIV & Trim(RS.Fields("accno").Value & "")
    End If
    
    Set RS = Nothing
End Function


Public Function Get_SpcInFormation(ByVal spcyy As String, ByVal spcno As Long) As Recordset
'---------------------------------
'��ü��ȣ�� ��ü������ ������ �´�
'---------------------------------
    Dim sSql As String
    
           
    sSql = " SELECT a.ptid,a.coldt,a.coltm,a.addfg,a.buildcd," & _
           " b." & F_PTNM & " as ptnm," & F_SSN2("b") & " as ssn,c." & F_EMPNM & " as empnm" & _
           " FROM " & T_BBS201 & " a," & T_HIS001 & " b," & T_HIS005 & " c " & _
           " WHERE " & _
           "     " & DBW("a.spcyy", spcyy, 2) & _
           " AND " & DBW("a.spcno", spcno, 2) & _
           " AND a.ptid=b." & F_PTID & _
           " AND " & DBJ("a.colid*=c." & F_EMPID)
    Set Get_SpcInFormation = New Recordset
    Get_SpcInFormation.Open sSql, DBConn

End Function
Public Function Get_SpcOrderList(ByVal ptid As String, _
                                 Optional ByVal accdt As String = "", _
                                 Optional ByVal accseq As String = "") As Recordset
'-----------------------------------------------------------------------------------------------------
'������ȣ�� �ش��ϴ� ó�� ������ ������ �´�
'accdt="" : �Ϲ�ä���� ���Ѱ�ü�μ�, donefg�� '0' or '1'�̸�, reqdt�� ���ú��� ū ��¥�� ���� ��ȸ�Ѵ�
'accdt<>"": �߰�ä���� ���� ��ü�μ� �߰���ü�� �����Ͽ� �ش��ϴ� ������ ��ȸ�Ѵ�.
'-----------------------------------------------------------------------------------------------------
    Dim sSql As String
    
    sSql = " SELECT c.testnm,b.ordcd,b.unitqty,b.orddt,a.reqdt,b.ordno,b.ordseq," & _
           "        a.ptid,b.accdt,b.accseq,b.dcfg,a.orddiv" & _
           " FROM " & T_LAB101 & " a," & T_LAB102 & " b," & T_BBS001 & " c" & _
           " WHERE " & _
                     DBW("b.ptid", ptid, 2)
                     
    If accdt <> "" Then
        sSql = sSql & " AND " & DBW("b.workarea", C_WORKAREA, 2) & _
                      " AND " & DBW("b.accdt", accdt, 2) & _
                      " AND " & DBW("b.accseq", accseq, 2)
    Else
        sSql = sSql & " AND b.donefg in(" & _
                                    DBV("donefg", stsORDER, 1) & _
                                    DBV("donefg", stsCOLLECT) & ")"
    End If
    sSql = sSql & _
                " AND b.ptid=a.ptid" & _
                " AND b.orddt=a.orddt" & _
                " AND b.ordno=a.ordno" & _
                " AND c.applydt=(SELECT max(e.applydt)" & _
                "                FROM " & T_BBS001 & " e" & _
                "                WHERE c.testcd = e.testcd" & _
                "                AND   " & DBW("e.applydt<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & ")" & _
                " AND b.ordcd=c.testcd"
    
    Set Get_SpcOrderList = New Recordset
    Get_SpcOrderList.Open sSql, DBConn
    
End Function

Private Sub Class_Initialize()

    If TRANS_REQUIRE_USED = True Then
        stsORDER = BBSOrdStatus.stsORDER
        stsCOLLECT = BBSOrdStatus.stsCOLLECT
        stsACCESS = BBSOrdStatus.stsACCESS
    Else
        stsORDER = BBSOrderStatus.stsORDER
        stsCOLLECT = BBSOrderStatus.stsCOLLECT
        stsACCESS = BBSOrderStatus.stsACCESS
    End If

End Sub