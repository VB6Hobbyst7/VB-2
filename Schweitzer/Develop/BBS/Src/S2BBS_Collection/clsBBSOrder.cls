VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'병원별 처방상태 관리 =======KJG(2001/12/12)
Private stsORDER   As String
Private stsCOLLECT As String
Private stsACCESS  As String
Private mvarSpc    As String
Public Property Get SpcNum() As String
    SpcNum = mvarSpc
End Property


Public Function Get_OrdDiv(ByVal testcd As String) As String
'수혈처방 마스터에서 orddiv를 불러온다.
    Dim RS   As New Recordset
    Dim sSql As String
    
    sSql = " SELECT orddiv" & _
           " FROM " & T_BBS001 & _
           " WHERE" & _
                    "     " & DBW("testcd", testcd, 2) & _
                    " AND (expdt='' or expdt is null)" & _
                    " AND applydt=(SELECT max(applydt)" & _
                                 " FROM " & T_BBS001 & " z" & _
                                 " WHERE " & DBW("z.testcd", testcd, 2) & _
                                 " AND " & DBW("z.applydt<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & _
                                  ")"
    Set RS = New Recordset
    RS.Open sSql, DBConn
    If Not RS.EOF Then
        Get_OrdDiv = IIf(IsNull(RS.Fields("orddiv").Value & "") = True, "B", RS.Fields("orddiv").Value & "")
    Else
        Get_OrdDiv = "N"
    End If
    
    Set RS = Nothing
End Function

Public Function Get_TestDiv(ByVal testcd As String) As String
'수혈처방 마스터에서 newtestdiv를 불러온다.
    Dim RS   As New Recordset
    Dim sSql As String
    
    sSql = " SELECT newtestdiv" & _
           " FROM " & T_BBS001 & _
           " WHERE" & _
                    "     " & DBW("testcd", testcd, 2) & _
                    " AND (expdt='' or expdt is null)" & _
                    " AND applydt=(SELECT max(applydt)" & _
                                 " FROM " & T_BBS001 & " z" & _
                                 " WHERE " & DBW("z.testcd", testcd, 2) & _
                                 " AND " & DBW("z.applydt<", Format(GetSystemDate, PRESENTDATE_FORMAT), 2) & _
                                  ")"
    Set RS = New Recordset
    RS.Open sSql, DBConn
    If Not RS.EOF Then
        Get_TestDiv = IIf(IsNull(RS.Fields("newtestdiv").Value & "") = True, "N", RS.Fields("newtestdiv").Value & "")
    Else
        Get_TestDiv = "N"
    End If
    
    Set RS = Nothing
End Function
Private Function GetSpecSpace(ByVal build As String, ByVal savepos As String, _
                              Leg As String, Row As String, Col As String) As Boolean
'사용확인
    '--------------------------------------------------------------------------------------------
    '보관장소를 구하거나, 입력된 보관장소가 사용가능한지 검사한다.
    '--------------------------------------------------------------------------------------------
    'savepos = "1"  : 자동지정
    '        = "0"  : 수동지정
    '--------------------------------------------------------------------------------------------
    '반환값 = True  : 보관장소를 사용할 수 있다.
    '       = False : 보관장소를 사용할 수 없다.
    '--------------------------------------------------------------------------------------------
    Dim objSpec As clsSpecManagement
    
    
    '수동 지정일 경우 입력값 확인----------------------------------------------------------------
    If savepos = "0" Then
        If Leg = "" Or Row = "" Or Col = "" Then
            Call MsgBox("보관장소 누락", vbCritical + vbOKOnly, "보관장소 누락")
            GetSpecSpace = False
            Exit Function
        End If
    End If
    
    
    '메모리 할당---------------------------------------------------------------------------------
    Set objSpec = New clsSpecManagement
    
    
    '자동지정일 경우 보관장소를 구한다.----------------------------------------------------------
    If savepos = "1" Then
        objSpec.Save_Spc_Search 1, build, Leg
        Leg = objSpec.Leg(1)
        Row = objSpec.Row(1)
        Col = objSpec.Col(1)
    End If
        
    
    '선택된(구해진) 보관장소가 사용가능한지 검사-------------------------------------------------
    If objSpec.SavePointChk(Leg, Row, Col, build) = False Then
        MsgBox "이미 사용중인 보관장소이거나,사용할수 없는 보관장소입니다." & vbNewLine & _
               "확인후 진행하십시요.", vbCritical + vbOKOnly, "보관장소입력에러"
        GetSpecSpace = False
    Else
        GetSpecSpace = True
    End If


    '메모리 해제---------------------------------------------------------------------------------
    Set objSpec = Nothing
    
End Function

Private Function GetOrgDestBuilding(ByVal Bussdiv As String, _
                                    ByVal build As String, ByVal statfg As String, _
                                    ByVal Wardid As String, _
                                    orgbuildcd As String, TestBuild As String) As Boolean

    Dim ipos As Integer
    
    ipos = IIf(statfg = "1", 1, 2)
    '외래,병동에 따른 채혈 건물을 구한다.---------------------------------------------------------
    If Bussdiv = "1" Then
        orgbuildcd = build
    Else
        orgbuildcd = Get_Build(Wardid)
    End If
    
    '채혈건물에 따른 검사건물을 구한다.-----------------------------------------------------------
    TestBuild = medGetP(Get_TestBuild(orgbuildcd), ipos, COL_DIV)
    
    '검사건물 정보가 정확하게 있는지 검사한다.----------------------------------------------------
    If TestBuild = "" Then
        Call MsgBox("해당건물정보가 누락되었습니다.", vbCritical + vbOKOnly, "건물정보누락")
        GetOrgDestBuilding = False
    Else
        GetOrgDestBuilding = True
    End If
End Function

Public Function Set_Order(ByVal objHeader As clsDictionary, _
                          ByVal objBody As clsDictionary, _
                          Optional ByVal IsAccess As Boolean = False) As Boolean
    '-------------------------------------------------------------------------
    'IsAccess = False : 처방등록만 처리
    '         = True  : 처방,채혈,접수 처리
    '-------------------------------------------------------------------------
    'sql문장 저장용 변수
    Dim sSql        As String
    
    'DB변수
    Dim BedIndt     As String
    Dim Bussdiv     As String
    Dim reqdt       As String
    Dim reqtm       As String
    Dim Wardid      As String
    Dim BedID       As String
    Dim HosilID     As String
    Dim Orddoct     As String
    
    Dim entid       As String
    Dim DeptCd      As String
    Dim ocd         As String
    Dim reasoncd    As String
    Dim statfg      As String
    
    Dim receiptno   As Long
    
    Dim ptid        As String
    Dim OrdNo       As Long
    Dim ordseq      As Long
    Dim orddt       As String
    Dim ordtm       As String
    
    'Access가 True 일때 사용변수
    Dim savepos     As String
    Dim Leg         As String
    Dim Row         As String
    Dim Col         As String
    Dim colid       As Long
    Dim coldt       As String
    Dim coltm       As String
    Dim build       As String
    
    Dim testcd      As String
    Dim IrradFg     As String
    Dim FilterFg    As String
    Dim don_stsfg   As String
    Dim UnitQty     As Long
    Dim orgbuildcd  As String
    Dim TestBuild   As String
    Dim pheresis    As String
    Dim testdiv     As String
    Dim orddiv      As String
    
    Dim objNumbers  As clsBBSNumbers
    Dim accdt       As String
    Dim accseq      As Long
    Dim spcyy       As String
    Dim spcno       As Long
    
On Error GoTo Set_Order_error
    '해더값 추출----------------------------------------------------------------------------------
    With objHeader
        ptid = .Fields("ptid")
        Bussdiv = .Fields("bussdiv")
        BedIndt = .Fields("bedindt")
        reqdt = .Fields("reqdt")
        reqtm = .Fields("reqtm")
        DeptCd = .Fields("deptcd")
        Orddoct = .Fields("orddoct")
        entid = .Fields("entid")
        receiptno = Val(.Fields("receptno"))
        Wardid = .Fields("wardid")
        BedID = .Fields("bedid")
        HosilID = .Fields("hosilid")
        ocd = .Fields("ocd")
        reasoncd = .Fields("reasoncd")
        statfg = .Fields("statfg")
        orddt = Format(GetSystemDate, PRESENTDATE_FORMAT)
        ordtm = Format(GetSystemDate, PRESENTTIME_FORMAT)
        OrdNo = Get_Ordno(ptid, orddt)
        build = .Fields("buildcd")
        coldt = .Fields("coldt")
        coltm = .Fields("coltm")
        colid = Val(.Fields("colid"))
        savepos = .Fields("savepos")
        Leg = .Fields("leg")
        Row = Val(.Fields("row"))
        Col = Val(.Fields("col"))
        orddiv = .Fields("orddiv")
    End With
    don_stsfg = stsORDER
    
    
    
    '--------------------------------------------------------------------------------------------
    '처방 해더 생성
    '--------------------------------------------------------------------------------------------
    sSql = Set_OrderHeader(ptid, orddt, OrdNo, ordtm, Bussdiv, BedIndt, reqdt, reqtm, DeptCd, _
                           Orddoct, entid, don_stsfg, receiptno, Wardid, BedID, HosilID, orddiv)
    DBConn.Execute sSql
    
    
    '--------------------------------------------------------------------------------------------
    '바디 Dictionary를 읽으면서 처리
    '--------------------------------------------------------------------------------------------
    '처방 바디 생성
    '--------------------------------------------------------------------------------------------
    With objBody
        .MoveFirst
        Do Until .EOF
            '바디 변수 추출---------------------------------------------------------------------
            ordseq = Val(.Fields("seq"))
            testcd = .Fields("ordcd")
            UnitQty = Val(.Fields("unitqty"))
            IrradFg = .Fields("irradfg")
            FilterFg = .Fields("filterfg")
            pheresis = IIf(.Fields("phere") = "1", "1", "0")
            testdiv = .Fields("testdiv")
            
            sSql = Set_OrderBody(ptid, orddt, OrdNo, ordseq, testcd, don_stsfg, _
                                 statfg, don_stsfg, receiptno, UnitQty, IrradFg, FilterFg, testdiv)
            DBConn.Execute sSql
            '------------------------------------------------------------------------------------
            '수혈사유 저장(2001/06/01..수정)(성모자애용)
            '------------------------------------------------------------------------------------
            If AttributeBBS Then
                sSql = SetAttributeBBS(ptid, orddt, OrdNo, ordseq, reasoncd, ocd)
                DBConn.Execute sSql
            Else
            
            '----------------------------------------
            'lab103 테이블 저장시
            '----------------------------------------
                sSql = SetAttributeCom(ptid, orddt, OrdNo, BC2_REASONVAL, reasoncd)
                DBConn.Execute sSql
                sSql = SetAttributeCom(ptid, orddt, OrdNo, BC2_ORDVAL, ocd)
                DBConn.Execute sSql
            End If

            .MoveNext
        Loop
    End With
    
    '--------------------------------------------------------------------------------------------
    '처방을 읽으면서, 처방이나 채혈상태인 처방을 접수처리한다.
    '--------------------------------------------------------------------------------------------
    If IsAccess = True Then
        If CollectProcess(ptid, Bussdiv, _
                          coldt, coltm, _
                          orddt, ordtm, _
                          colid, entid, _
                          statfg, _
                          build, _
                          savepos, _
                          Wardid, _
                          Leg, Row, Col) = False Then
            Set_Order = False
            Exit Function
        End If
        If OrderAccessProcess(ptid, Bussdiv, orddt, ordtm, entid) = False Then
            Set_Order = False
            Exit Function
        End If
    End If
    
   ' DBConn.CommitTrans
    Set_Order = True
    Exit Function
    
Set_Order_error:
  '  DBConn.RollbackTrans
    Set_Order = False
End Function

Private Function CollectProcess(ByVal ptid As String, ByVal Bussdiv As String, _
                                ByVal coldt As String, ByVal coltm As String, _
                                ByVal orddt As String, ByVal ordtm As String, _
                                ByVal colid As String, ByVal entid As String, _
                                ByVal statfg As String, _
                                ByVal build As String, _
                                ByVal savepos As String, _
                                ByVal Wardid As String, _
                                ByVal Leg As String, ByVal Row As String, ByVal Col As String) As Boolean
    '----------------------------------------------------------------------------
    '처방을 읽으면서, 채혈처리를 한다.
    '----------------------------------------------------------------------------
    Dim RS          As Recordset
    Dim sSql        As String
    Dim i           As Long
    Dim orgbuildcd  As String
    Dim TestBuild   As String
    Dim spcyy       As String
    Dim spcno       As Long
    
    Dim objNumbers  As clsBBSNumbers
    Dim objsql      As clsBBSCollection
    
On Error GoTo CollectProcess_error

    '----------------------------------------------------------------------------
    '만일 이 환자에게 검체가 존재하면 채혈하지 않는다.
    '----------------------------------------------------------------------------
    Set objsql = New clsBBSCollection
    With objsql
        If .Blood_Existence(ptid, coldt, coltm) = False Then
            CollectProcess = True
            Set objsql = Nothing
            Exit Function
        End If
    End With
    Set objsql = Nothing
    

    
    '----------------------------------------------------------------------------
    '채혈되지 않은 처방이 있고, 그 처방의 검사방법이 None이 아니면 채혈을 한다.
    '----------------------------------------------------------------------------
    sSql = " SELECT b.* " & _
           " FROM " & _
                    T_LAB101 & " a," & T_LAB102 & " b," & T_BBS001 & " c " & _
           " WHERE" & _
                    " " & DBW("a.orddiv", C_WORKAREA, 2) & _
                    " AND " & DBW("a.ptid", ptid, 2) & _
                    " AND " & DBW("a.donefg", stsORDER, 2) & _
                    " AND " & DBW("a.bussdiv", Bussdiv, 2) & _
                    " AND a.ptid=b.ptid" & _
                    " AND a.orddt=b.orddt" & _
                    " AND a.ordno=b.ordno " & _
                    " AND " & DBW("b.stscd", stsORDER, 2) & _
                    " AND (b.dcfg is null or b.dcfg='' ) " & _
                    " AND b.ordcd=c.testcd " & _
                    " AND c.applydt=( SELECT max(d.applydt) " & _
                                    " FROM " & T_BBS001 & " d " & _
                                    " WHERE d.testcd = b.ordcd " & _
                                    " AND d.applydt<=b.orddt " & _
                                    ") " & _
                    " AND " & DBW("c.xmethod<>", XMethod.xmNONE)

    Set RS = New Recordset
    Call RS.Open(sSql, DBConn)
    With RS
        If .EOF Then GoTo CollectProcess_error
        If .RecordCount <= 0 Then
            '채혈대상이 없다.----------------------------------------------------
            CollectProcess = True
            Set RS = Nothing
            Exit Function
        End If
    
        '----------------------------------------------------------------------------
        '처방 헤더,바디 update
        '----------------------------------------------------------------------------
        For i = 1 To .RecordCount
            sSql = Set_AddCollectBody(ptid, .Fields("orddt").Value & "", .Fields("ordno").Value & "", .Fields("ordseq").Value & "")
            DBConn.Execute sSql
            
            sSql = Set_AddCollectOrder(ptid, .Fields("orddt").Value & "", .Fields("ordno").Value & "")
            DBConn.Execute sSql
            
            .MoveNext
        Next i
    End With
    Set RS = Nothing
    
    
    '--------------------------------------------------------------------------------------------
    '보관장소지정
    '검사건물 지정
    '--------------------------------------------------------------------------------------------
    If GetSpecSpace(build, savepos, Leg, Row, Col) = False Then Exit Function
    If GetOrgDestBuilding(Bussdiv, build, statfg, Wardid, orgbuildcd, TestBuild) = False Then Exit Function
    
    
    '----------------------------------------------------------------------------
    '검체번호를 구한다.
    '----------------------------------------------------------------------------
    Set objNumbers = New clsBBSNumbers
    With objNumbers
        spcyy = medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 1, COL_DIV)
        spcno = Val(medGetP(.Get_SpcYY_No(BC2_SPC_DEFAULT & Mid(Format(GetSystemDate, "yyyy"), 4)), 2, COL_DIV))
        '----------------------------------------------------------------------------
        '검체번호 update
        '----------------------------------------------------------------------------
        sSql = .Set_NumbersCom099(BN_SPC_NO, spcyy, spcno)
        DBConn.Execute sSql
        mvarSpc = spcyy & spcno
    End With
    
    
    '----------------------------------------------------------------------------
    '채혈내역 생성
    '----------------------------------------------------------------------------
    sSql = Set_Collect(spcyy, spcno, coldt, coltm, colid, entid, Leg, Row, Col, _
                       ptid, TestBuild, build, orddt, ordtm)
    DBConn.Execute sSql
    '----------------------------------------------------------------------------
    '처방 채혈 연결고리
    '----------------------------------------------------------------------------
    
    
    
    
    '----------------------------------------------------------------------------
    '보관장소 update
    '----------------------------------------------------------------------------
    sSql = Set_Position(build, Leg, Row, Col, spcyy, spcno)
    DBConn.Execute sSql
    
    CollectProcess = True
    Set objNumbers = Nothing
    Exit Function
    
CollectProcess_error:
    CollectProcess = False
    Set objNumbers = Nothing
    MsgBox Err.Description, vbExclamation
End Function

Private Function OrderAccessProcess(ByVal ptid As String, ByVal Bussdiv As String, _
                                    ByVal orddt As String, ByVal ordtm As String, _
                                    ByVal entid As String) As Boolean
    '----------------------------------------------------------------------------
    '처방을 읽으면서, 접수처리를 한다.
    '----------------------------------------------------------------------------
    Dim RsOrder     As Recordset
    Dim RS          As Recordset
    Dim sSql        As String
    Dim i           As Long
    
    Dim objNumbers  As clsBBSNumbers
    Dim accdt       As String
    Dim accseq      As Long
    
On Error GoTo OrderAccessProcess_error
    
    '처방을 읽는다.----------------------------------------------------------------------
    'Set RsOrder = New recordset
    Set objNumbers = New clsBBSNumbers
    With objNumbers
        accdt = .Get_AccdtFormat
        accseq = Val(.Get_AccDT_Seq(accdt))
    End With
    Set RsOrder = Get_Order_TarGet(ptid, Bussdiv)
    '접수번호를 구한다.-------------------------------------------------------------------
    With RsOrder
        If .RecordCount > 0 Then
        
            
            For i = 1 To .RecordCount
                '처방 바디 update
                sSql = Set_AddAcessBody(ptid, .Fields("orddt").Value & "", .Fields("ordno").Value & "", _
                                       .Fields("ordseq").Value & "", accdt, accseq, orddt, ordtm)
                DBConn.Execute sSql
                '처방 접수내역 생성
                sSql = Set_TransAcess(accdt, accseq, ptid, .Fields("orddt").Value & "", _
                                      CLng(.Fields("ordno").Value & ""), CLng(.Fields("ordseq").Value & ""), entid, _
                                      orddt, ordtm, .Fields("pherefg").Value & "")
                DBConn.Execute sSql
                
                sSql = Set_AddAcessOrder(ptid, .Fields("orddt").Value & "", CLng(.Fields("ordno").Value & ""))
                DBConn.Execute sSql
                

                Dim objCollect As New clsBBSCollection
                Dim SQLTmp As String
                SQLTmp = objCollect.Set_AccUnitSQL_203(ptid, accdt, CStr(accseq))
                sSql = medGetP(SQLTmp, 1, COL_DIV)
                DBConn.Execute sSql
                sSql = medGetP(SQLTmp, 2, COL_DIV)
                DBConn.Execute sSql
                Set objCollect = Nothing
                
                accseq = accseq + 1
                .MoveNext
            Next i
        End If
        
        Set RS = GetRTSingleOrder(ptid, Bussdiv)
        If Not RS.EOF Then
            For i = 1 To RS.RecordCount
                sSql = Set_AddAcessBody(ptid, RS.Fields("orddt").Value & "", RS.Fields("ordno").Value & "", RS.Fields("ordseq").Value & "", accdt, accseq, orddt, ordtm)
                DBConn.Execute sSql
                sSql = Set_AddAcessOrder(ptid, RS.Fields("orddt").Value & "", CLng(RS.Fields("ordno").Value & ""))
                DBConn.Execute sSql
                accseq = accseq + 1
                RS.MoveNext
            Next
        End If
        Set RS = Nothing
       
            '처방 헤더 update
'            .MoveFirst
'            sSql = Set_AddAcessOrder(ptid, .Fields("orddt").value & "", CLng(.Fields("ordno").value & ""))
'            DBConn.Execute sSql
            
            '접수번호 update-------------------------------------------------------------------
        sSql = objNumbers.Set_NumbersCom099(BN_ACC_NO, accdt, accseq - 1)
        DBConn.Execute sSql
            
    End With
    Set RsOrder = Nothing
    Set objNumbers = Nothing
    OrderAccessProcess = True
    Exit Function
    
OrderAccessProcess_error:
    Set RsOrder = Nothing
    Set objNumbers = Nothing
    OrderAccessProcess = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function Set_AddAcessBody(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As String, _
                                 ByVal ordseq As String, ByVal accdt As String, ByVal accseq As String, _
                                 ByVal rcvdt As String, ByVal rcvtm As String) As String
'    MsgBox "처방내역 업데이트", vbExclamation
    
'    Set_AddAcessBody = " update " & T_LAB102 & " set " & _
'                                                DBW("rcvdt", rcvdt, 3) & _
'                                                DBW("rcvtm", rcvtm, 3) & _
'                                                DBW("donefg", stsACCESS, 3) & _
'                                                DBW("stscd", stsACCESS, 3) & _
'                                                DBW("workarea", C_WORKAREA, 3) & _
'                                                DBW("accdt", accdt, 3) & _
'                                                DBW("accseq", accseq, 2) & _
'                       " WHERE " & _
'                                "     " & DBW("ptid", ptid, 2) & _
'                                " AND " & DBW("orddt", orddt, 2) & _
'                                " AND " & DBW("ordno", OrdNo, 2) & _
'                                " AND " & DBW("ordseq", ordseq, 2)
    '예수병원용
    Set_AddAcessBody = " update " & T_LAB102 & " set " & _
                                                " acptdate=to_date(" & rcvdt & rcvtm & ",'yyyymmdd hh24:mi:ss'), " & _
                                                DBW("donefg", stsACCESS, 3) & _
                                                DBW("stscd", stsACCESS, 3) & _
                                                DBW("workarea", C_WORKAREA, 3) & _
                                                DBW("accdt", accdt, 3) & _
                                                DBW("accseq", accseq, 2) & _
                       " WHERE " & _
                                "     " & DBW("patno", ptid, 2) & _
                                " AND orddate=to_date(" & orddt & ",'yyyymmdd')," & _
                                " AND " & DBW("ordseqno=", OrdNo)
                                
End Function
Private Function Set_AddAcessOrder(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long) As String
'    MsgBox "처방내역 업데이트", vbExclamation
    
'    Set_AddAcessOrder = " update " & T_LAB101 & " set " & _
'                                                DBW("donefg", stsACCESS, 2) & _
'                        " WHERE " & _
'                                      DBW("ptid", ptid, 2) & _
'                                " AND donefg in(" & _
'                                               DBV("donefg", stsORDER, 1) & _
'                                               DBV("donefg", stsCOLLECT) & ")" & _
'                                " AND " & DBW("orddt", orddt, 2) & _
'                                " AND " & DBW("ordno", OrdNo, 2)
    '예수병원용
    Set_AddAcessOrder = " update " & T_LAB101 & " set " & _
                                                DBW("donefg", stsACCESS, 2) & _
                        " WHERE " & _
                                      DBW("patno", ptid, 2) & _
                                " AND donefg in(" & _
                                               DBV("donefg", stsORDER, 1) & _
                                               DBV("donefg", stsCOLLECT) & ")" & _
                                " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                                " AND " & DBW("ordseqno", OrdNo, 2)

End Function

Private Function Set_AddCollectBody(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As String, ByVal ordseq As String) As String
'사용확인
'    MsgBox "처방내역 업데이트", vbExclamation
    
'    Set_AddCollectBody = " update " & T_LAB102 & " set " & _
'                                                   DBW("stscd", stsCOLLECT, 3) & _
'                                                   DBW("donefg", stsCOLLECT, 2) & _
'                         " WHERE" & _
'                                  "     " & DBW("ptid", ptid, 2) & _
'                                  " AND " & DBW("orddt", orddt, 2) & _
'                                  " AND " & DBW("ordno", OrdNo, 2) & _
'                                  " AND " & DBW("ordseq", ordseq, 2)
    '예수병원용
    Set_AddCollectBody = " update " & T_LAB102 & " set " & _
                                                   DBW("stscd", stsCOLLECT, 3) & _
                                                   DBW("donefg", stsCOLLECT, 2) & _
                         " WHERE" & _
                                  "     " & DBW("patno", ptid, 2) & _
                                  " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                                  " AND " & DBW("ordno=", OrdNo)
                                
End Function
Private Function Set_AddCollectOrder(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As String) As String
'사용확인
'    MsgBox "처방내역 업데이트", vbExclamation
    
'    Set_AddCollectOrder = " update " & T_LAB101 & " set " & _
'                                                    DBW("donefg", stsCOLLECT, 2) & _
'                          " WHERE" & _
'                                   "     " & DBW("ptid", ptid, 2) & _
'                                   " AND " & DBW("donefg", stsORDER, 2) & _
'                                   " AND " & DBW("orddt", orddt, 2) & _
'                                   " AND " & DBW("ordno", OrdNo, 2)
    '예수병원용
    Set_AddCollectOrder = " update " & T_LAB101 & " set " & _
                                                    DBW("donefg", stsCOLLECT, 2) & _
                          " WHERE" & _
                                   "     " & DBW("patno", ptid, 2) & _
                                   " AND " & DBW("donefg", stsORDER, 2) & _
                                   " AND orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                                   " AND " & DBW("ordseqno", OrdNo, 2)
End Function
Private Function GetRTSingleOrder(ByVal ptid As String, ByVal Bussdiv As String) As Recordset
    Dim sSql As String
    
    sSql = " SELECT b.ptid,b.orddt,b.ordno,b.ordseq  " & _
         " FROM " & T_BBS001 & " c," & T_LAB102 & " b," & T_LAB101 & " a" & _
         " WHERE " & _
                   DBW("a.ptid=", ptid) & _
         " AND " & DBW("a.bussdiv=", Bussdiv) & _
         " AND " & DBW("a.orddiv", C_WORKAREA, 2) & _
         " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno" & _
         " AND (b.dcfg='' or b.dcfg is null)" & _
         " AND b.donefg in (" & _
                          DBV("donefg", stsORDER, 1) & _
                          DBV("donefg", stsCOLLECT) & ") " & _
         " AND b.ordcd=c.testcd" & _
         " AND " & DBW("c.xmethod=", XMethod.xmNONE) & _
         " AND c.applydt=(SELECT max(z.applydt) FROM " & T_BBS001 & " z" & _
                         " WHERE " & _
                                      DBW("z.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
                                " AND z.testcd=b.ordcd" & _
                        ")"
    Set GetRTSingleOrder = New Recordset
    GetRTSingleOrder.Open sSql, DBConn
End Function
Private Function Get_Order_TarGet(ByVal ptid As String, ByVal Bussdiv As String) As Recordset
    Dim strSql As String
    
    
    
    
    strSql = " SELECT a.ptid,a.orddt,a.ordno,a.ordseq,d.pherefg " & _
             " FROM " & T_LAB102 & " a," & T_LAB101 & " b," & T_BBS001 & " c," & T_BBS006 & " d" & _
             " WHERE " & _
             "     " & DBW("a.ptid", ptid, 2) & _
             " AND " & DBW("b.bussdiv", Bussdiv, 2) & _
             " AND " & DBW("b.orddiv", C_WORKAREA, 2) & _
             " AND a.ptid=b.ptid" & _
             " AND a.orddt=b.orddt" & _
             " AND a.ordno=b.ordno" & _
             " AND (a.dcfg='' or a.dcfg is null)" & _
             " AND a.donefg in (" & _
                              DBV("donefg", stsORDER, 1) & _
                              DBV("donefg", stsCOLLECT) & ") " & _
             " AND a.ordcd=c.testcd" & _
             " AND c.compocd=d.compocd" & _
             " AND c.applydt=(SELECT max(z.applydt) FROM " & T_BBS001 & " z" & _
                             " WHERE " & _
                                           DBW("z.applydt<=", Format(GetSystemDate, PRESENTDATE_FORMAT)) & _
                                     " AND z.testcd=a.ordcd" & _
                             ")"
    
    
    Set Get_Order_TarGet = New Recordset
    Get_Order_TarGet.Open strSql, DBConn
    
End Function

Private Function Set_Position(ByVal cntercd As String, ByVal legcd As String, ByVal rowno As Long, _
                              ByVal colno As Long, ByVal spcyy As String, ByVal spcno As Long) As String
    
    Set_Position = " update " & T_BBS206 & " set " & _
                                             DBW("spcyy", spcyy, 3) & _
                                             DBW("spcno", spcno, 3) & _
                                             DBW("stscd", BBSSaveStatue.stsNotUsed, 2) & _
                   " WHERE" & _
                            "     " & DBW("centercd", cntercd, 2) & _
                            " AND " & DBW("legcd", legcd, 2) & _
                            " AND " & DBW("rowno", rowno, 2) & _
                            " AND " & DBW("colno", colno, 2) & ""
End Function

Private Function Set_Collect(ByVal spcyy As String, ByVal spcno As Long, ByVal coldt As String, _
                            ByVal coltm As String, ByVal colid As String, _
                            ByVal rcvid As String, ByVal storeleg As String, ByVal storerno As Long, _
                            ByVal storecno As Long, ByVal ptid As String, _
                            ByVal buildcd As String, ByVal orgbuildcd As String, _
                            ByVal rcvdt As String, ByVal rcvtm As String) As String
    Dim strAddFg As String

    strAddFg = BBSSpcAddFg.stsORDER
    
    Set_Collect = " insert into " & T_BBS201 & "(" & _
                  " spcyy,spcno,coldt,coltm,colid,rcvdt,rcvtm,rcvid,storeleg,storerno,storecno," & _
                  " ptid,buildcd,orgbuildcd,expfg,addfg)" & _
                  " values (" & _
                           DBV("spcyy", spcyy, 1) & DBV("spcno", spcno, 1) & DBV("coldt", coldt, 1) & _
                           DBV("coltm", coltm, 1) & _
                           DBV("colid", colid, 1) & DBV("rcvdt", rcvdt, 1) & DBV("rcvtm", rcvtm, 1) & _
                           DBV("rcvid", rcvid, 1) & _
                           DBV("storeleg", storeleg, 1) & DBV("storerno", storerno, 1) & _
                           DBV("storecno", storecno, 1) & DBV("ptid", ptid, 1) & _
                           DBV("buildcd", buildcd, 1) & DBV("orgbuildcd", orgbuildcd, 1) & DBV("expfg", "0", 1) & _
                           DBV("addfg", strAddFg) & ")"
            

End Function

Private Function Set_TransAcess(ByVal accdt As String, ByVal accseq As Long, ByVal ptid As String, _
                               ByVal orddt As String, ByVal OrdNo As Long, _
                               ByVal ordseq As Long, ByVal rcvid As String, ByVal rcvdt As String, _
                               ByVal rcvtm As String, ByVal pheresis As String) As String
'처방 접수내역에 Insert 한다.
    
    Set_TransAcess = " insert into " & T_BBS202 & "(" & _
                     " workarea,accdt,accseq,ptid,orddt,ordno,ordseq,rcvdt,rcvtm,rcvid,stscd," & _
                     " canceldt,canceltm,cancelid,cancelfg,cancelrsn,pheresisfg)" & _
                     " values(" & _
                             DBV("workarea", C_WORKAREA, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("ptid", ptid, 1) & _
                             DBV("orddt", orddt, 1) & DBV("ordno", OrdNo, 1) & DBV("ordseq", ordseq, 1) & DBV("rcvdt", rcvdt, 1) & _
                             DBV("rcvtm", rcvtm, 1) & DBV("rcvid", rcvid, 1) & DBV("stscd", stsACCESS, 1) & _
                             DBV("canceldt", "", 1) & DBV("canceltm", "", 1) & DBV("cancelid", "", 1) & DBV("cancelfg", "0", 1) & _
                             DBV("cancelrsn", "", 1) & DBV("pheresisfg", pheresis) & ")"
End Function

Private Function SetAttributeBBS(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As String, _
                                      ByVal ordseq As String, ByVal rsncd As String, ByVal opcd As String)
'사용확인
    '처방속성테이블대신 수혈 사유테이블에 저장함(BBS101)

    SetAttributeBBS = "insert into " & T_BBS101 & _
                         "(ptid,orddt,ordno,ordseq,bloodrsn,bloodop)" & _
                         " values(" & _
                                 DBV("ptid", ptid, 1) & DBV("orddt", orddt, 1) & DBV("ordno", OrdNo, 1) & DBV("ordseq", ordseq, 1) & _
                                 DBV("bloodrsn", rsncd, 1) & DBV("bloodop", opcd) & ")"
    
End Function



Private Function SetAttributeCom(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long, _
                                  ByVal Attrcd As String, ByVal AttrVal As String) As String
'사용확인
    '처방정보(속성) insert
    SetAttributeCom = " insert into " & T_LAB103 & _
                        " (ptid,orddt,ordno,attrcd,seq,attrval,attrtext,remark)" & _
                        " values (" & _
                                 DBV("ptid", ptid, 1) & DBV("orddt", orddt, 1) & DBV("ordno", OrdNo, 1) & _
                                 DBV("attrcd", Attrcd, 1) & DBV("seq", 1, 1) & DBV("attrval", AttrVal, 1) & DBV("attrtext", "", 1) & _
                                 DBV("remark", "") & ")"
               
End Function

Private Function Set_OrderHeader(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long, _
                                 ByVal ordtm As String, ByVal busidiv As String, ByVal BedIndt As String, _
                                 ByVal reqdt As String, ByVal reqtm As String, ByVal DeptCd As String, _
                                 ByVal Orddoct As String, ByVal entid As String, ByVal DoneFg As String, _
                                 ByVal Receptno As Long, ByVal Wardid As String, ByVal BedID As String, _
                                 ByVal HosilID As String, ByVal orddiv As String) As String
'사용확인
    '처방정보(헤더) insert
'    Set_OrderHeader = " insert into " & T_LAB101 & "(" & _
'                      " ptid,orddt,ordno,ordtm,bussdiv,bedindt,reqdt,reqtm,deptcd,orddoct," & _
'                      " majdoct,entid, entdt,enttm,orddiv,ordfg,orgaccno,sporddiv,donefg," & _
'                      " receptno,wardid,hosilid,bedid," & _
'                      " hoscd, repeatfg, roomid,trustdt) " & _
'                      " values( " & _
'                              DBV("ptid", ptid, 1) & DBV("orddt", orddt, 1) & DBV("ordno", OrdNo, 1) & DBV("ordtm", ordtm, 1) & _
'                              DBV("bussdiv", busidiv, 1) & DBV("bedindt", BedIndt, 1) & DBV("reqdt", reqdt, 1) & DBV("reqtm", reqtm, 1) & _
'                              DBV("deptcd", DeptCd, 1) & DBV("orddoct", Orddoct, 1) & DBV("majdoct", "", 1) & _
'                              DBV("entid", entid, 1) & DBV("entdt", orddt, 1) & DBV("enttm", ordtm, 1) & DBV("orddiv", orddiv, 1) & _
'                              DBV("ordfg", "1", 1) & DBV("orgaccno", "", 1) & DBV("sporddiv", "", 1) & DBV("donefg", DoneFg, 1) & DBV("receptno", Receptno, 1) & _
'                              DBV("wardid", Wardid, 1) & DBV("hosilid", BedID, 1) & DBV("bedid", HosilID, 1) & DBV("hoscd", "", 1) & _
'                              DBV("repeatfg", "", 1) & DBV("roomid", "", 1) & DBV("trustdt", "") & ") "
'예수병원 처방 테스트용

    Set_OrderHeader = " insert into oram1.mdbldort (patno, orddate, ordseqno, ordtime, patsect, meddate, " & _
                      " meddept, orddr, chadr, editid, editdate, wardno, roomno,patsite, ordgrp,slipcd,ordtype,ordkind,ordcd ) values (" & _
                      "'" & ptid & "'," & "to_date(" & orddt & ",'yyyy-mm-dd'), " & OrdNo & ",to_date(" & ordtm & ",'hh24:mi:ss'),'" & IIf(busidiv = "1", "O", "I") & "'," & IIf(BedIndt = "", "''", "to_date(" & BedIndt & ",'yyyy-mm-dd')") & _
                      ",'" & DeptCd & "','" & Orddoct & "','" & Orddoct & "','" & entid & "'," & "to_date(" & orddt & ordtm & ",'yyyy-mm-dd hh24:mi:ss') , " & _
                      "'" & Wardid & "','" & BedID & "','Y','1','1','1','1','1')"


End Function

Private Function Set_OrderBody(ByVal ptid As String, ByVal orddt As String, ByVal OrdNo As Long, _
                              ByVal ordseq As Long, ByVal OrdCd As String, ByVal stscd As String, _
                              ByVal statfg As String, ByVal DoneFg As String, ByVal Receptno As Long, _
                              ByVal UnitQty As Long, ByVal IrradFg As String, _
                              ByVal FilterFg As String, ByVal testdiv As String) As String
'사용확인
    '처방정보(내역) insert
'    Set_OrderBody = "insert into " & T_LAB102 & "(" & _
'                    " ptid,orddt,ordno,ordseq,ordcd,spccd,storecd,dcfg,dcdt,dcno," & _
'                    " attrcd,examdt,examtm,examdoct,rcvdt,rcvtm,stscd,statfg,donefg,insdiv," & _
'                    " workarea,accdt,accseq,receptno,ocsordno,ocsordseq,ocsordpumok,mesg,addfg,paydt," & _
'                    " unitqty,volume,irradfg,filterfg,newtestdiv )" & _
'                    " values (" & _
'                   DBV("ptid", ptid, 1) & DBV("orddt", orddt, 1) & DBV("ordno", OrdNo, 1) & DBV("ordseq", ordseq, 1) & DBV("ordcd", OrdCd, 1) & _
'                   DBV("spccd", "", 1) & DBV("storecd", "", 1) & DBV("dcfg", "", 1) & DBV("dcdt", "", 1) & DBV("dcno", "", 1) & _
'                   DBV("attrcd", "", 1) & DBV("examdt", "", 1) & DBV("examtm", "", 1) & DBV("examdoct", "", 1) & DBV("rcvdt", "", 1) & _
'                   DBV("rcvtm", "", 1) & DBV("stscd", stscd, 1) & DBV("statfg", statfg, 1) & DBV("donefg", DoneFg, 1) & DBV("insdiv", "", 1) & _
'                   DBV("workarea", "", 1) & DBV("accdt", "", 1) & DBV("accseq", "", 1) & DBV("receptno", Receptno, 1) & DBV("ocsordno", "", 1) & _
'                   DBV("ocsordseq", "", 1) & DBV("ocsordpumok", "", 1) & DBV("mesg", "", 1) & DBV("addfg", "", 1) & DBV("paydt", "0", 1) & _
'                   DBV("unitqty", UnitQty, 1) & DBV("volume", "", 1) & DBV("irradfg", IrradFg, 1) & DBV("filterfg", FilterFg, 1) & DBV("newtestdiv", testdiv) & ")"
    
    Set_OrderBody = " update oram1.mdbldort set ordcd='" & OrdCd & "',stscd='0',statfg='" & IIf(statfg = "1", "Y", "") & "',donefg='0',qty=" & UnitQty & ",irradyn='" & IIf(IrradFg = "1", "Y", "") & "', newtestdiv='" & testdiv & "'" & _
                    " where patno='" & ptid & "'" & _
                    " and orddate=to_date(" & orddt & ",'yyyymmdd')" & _
                    " and ordseqno=" & OrdNo
End Function


Private Function Get_Ordno(ByVal ptid As String, ByVal orddt As String) As Long
'사용확인
    '환자의 처방번호를 가지고 온다.
    Dim sSql As String
    Dim RS   As Recordset
    
    sSql = " SELECT max(ordno) as maxordno" & _
           " FROM " & T_LAB101 & _
           " WHERE" & _
                    "     " & DBW("ptid", ptid, 2) & _
                    " AND " & DBW("orddt", orddt, 2)
    
    Set RS = New Recordset
    RS.Open sSql, DBConn
    If RS.EOF Then
        Get_Ordno = 1
    Else
        Get_Ordno = Val("" & RS.Fields("maxordno").Value & "") + 1
    End If
    
    Set RS = Nothing
End Function
Public Function Get_Build(ByVal Wardid As String) As String
    Get_Build = ObjSysInfo.buildingcd
End Function

Public Function Get_TestBuild(ByVal blgcd As String) As String
    '건물코드를 가지고 검사가 수행될 건물을 찾는다.
    'True이면 응급, False이면 일반
    Dim sSql As String
    Dim RS As New Recordset
    
        sSql = " SELECT erbuilding,gbuilding" & _
               " FROM " & T_BBS004 & " " & _
               " WHERE" & _
                        " " & DBW("sbuilding", blgcd, 2)
    
    RS.Open sSql, DBConn
    
    If RS.EOF = False Then
        Get_TestBuild = RS.Fields("erbuilding").Value & "" & COL_DIV & RS.Fields("gbuilding").Value & ""
    Else
        Get_TestBuild = ObjSysInfo.buildingcd & COL_DIV & ObjSysInfo.buildingcd
    End If
    
    Set RS = Nothing
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
