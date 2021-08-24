VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISAccession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'+--------------------------------------------------------------------------------------+
'|  1. Class 명  : clsAccession
'|  2. 기  능    : 결과내역(LAB302)을 생성한다.
'|  3. 작성자    : 김미경
'|  4. 작성일    : 1999.06.18
'|  5. 수정일    : 2001.03.23
'|
'|  CopyRight(C) 1999 대련엠티에스
'+--------------------------------------------------------------------------------------+


'%  클래스 clsAccession의 Data Attributes

Private WorkArea As String  '/* Work Area(접수번호용) */
Private AccDt As String     '/* 접수일(접수번호용)    */
Private AccSeq As Long   '/* 접수순번(접수번호용)  */

Public TestCd As String     '/* 검사항목 코드   */
Public RstVal As String     '/* 결과(Numeric)   */
Public RstCd As String      '/* 결과코드(Alpha) */
Public RstUnit As String    '/* Unit            */
Public HLDiv As String      '/* High/Low(H:Hig,L:Low) */
Public DPDiv As String      '/* Delta/Panic(D:Delta,P:Panic) */
Public SpcCd As String      '/* 검체코드        */
Public StatFg As String     '/* 응급여부('0':부,'1':여) */
Public LastRst As String    '/* 최근결과        */
Public LastVfyDt As String  '/* 최근결과확인일  */
Public LastVfyTm As String  '/* 최근결과확인시간 */
Public LastVfyId As Integer '/* 최근결과확인자  */
Public LastVfyId1 As String '/* 최근결과확인자  */
Public VfyDt As String      '/* 결과확인일자    */
Public VfyTm As String      '/* 결과확인시간    */
Public VfyId As Integer     '/* 결과확인자      */
Public VfyId1 As String     '/* 결과확인자      */
Public AttrCd As String     '/* 속성코드   */
Public MfyFg As String      '/* 수정여부('0':부,'1':여) */
Public GrpFg As String      '/* 그래프 결과여부('0':부,'1':여) */
Public TxtFg As String      '/* TEXT 결과여부('0':부,'1':여) */
Public AutoFg As String
Public RstType As String    '/* 결과유형(N,Alpha,비율,Free) */
Public DetailFg As String   '/* 상세항목Group Seq */
Public RstDiv As String     '/Alternative,Required 여부 */
Public EqpCd As String
Public Ptid As String       '/* 환자ID  */
Public OrdDt As String      '/* 처방일 */
Public OrdNo As Integer     '/* 처방번호 */
Public OrdSeq As Integer    '/* 처방Seq */
Public SenFg As String      '/* 감수성결과 여부 */
Public OutLabCd As String

Private TestDiv As String

Private L201_TestDiv As String
Private L201_MultiFg As String
Private L201_SpcCd As String
Private L203_SpcSeq As String

Private ItemIndex As Integer
'Private Type ResultTable
'    TestCd As String     '/* 검사항목 코드   */
'    RstVal As String     '/* 결과(Numeric)   */
'    RstCd As String      '/* 결과코드(Alpha) */
'    RstUnit As String    '/* Unit            */
'    HLDiv As String      '/* High/Low(H:Hig,L:Low) */
'    DPDiv As String      '/* Delta/Panic(D:Delta,P:Panic) */
'    SpcCd As String      '/* 검체코드        */
'    StatFg As String     '/* 응급여부('0':부,'1':여) */
'    LastRst As String    '/* 최근결과        */
'    LastVfyDt As String  '/* 최근결과확인일  */
'    LastVfyTm As String  '/* 최근결과확인시간 */
'    LastVfyId As Integer '/* 최근결과확인자  */
'    VfyDt As String      '/* 결과확인일자    */
'    VfyTm As String      '/* 결과확인시간    */
'    VfyId As Integer     '/* 결과확인자      */
'    AttrCd As String     '/* 속성코드   */
'    MfyFg As String      '/* 수정여부('0':부,'1':여) */
'    GrpFg As String      '/* 그래프 결과여부('0':부,'1':여) */
'    ValFg As String      '/* 수치 결과여부('0':부,'1':여)  - 기타검사에서만 사용*/
'    TxtFg As String      '/* TEXT 결과여부('0':부,'1':여) */
'    RstType As String    '/* 결과유형(N,Alpha,비율,Free) */
'    RstDiv As String     '/Alternative,Required 여부 */
'    DetailFg As String   '/* 상세항목Group Seq */
'    PtId As String       '/* 환자 ID */
'    OrdDt As String      '/* 처방일 */
'    OrdNo As Integer     '/* 처방번호 */
'    OrdSeq As Integer    '/* 처방Seq */
'    SenFg As String      '/* 감수성결과 여부 */
'    AutoFg As String
'    EqpCd As String
'
'    TestDiv As String
'End Type

Private OutLabIndex As Integer
'Private Type OutLabTable
'    WorkArea As String  '/* Work Area(접수번호용) */
'    AccDt As String     '/* 접수일(접수번호용)    */
'    AccSeq As Integer   '/* 접수순번(접수번호용)  */
'    TestCd As String    '/* 검사항목        */
'    PtId As String      '/* 환자ID  */
'    SpcCd As String     '/* 검체코드  */
'    OutLabCd As String  '/* 기관코드        */
'    StsCd As String     '/* STATUS          */
'    RcvDt As String     '/* 접수일자        */
'    SendDt As String    '/* SEND Date       */
'    ChargeDt As String  '/* 청구일자        */
'    SendId As String    '/* 전송자 ID   */
'End Type

Private MyResult As New Collection    ' ResultTable
Private MyOutLab As New Collection    ' OutLabTable
Private MyOrders() As String
Private DetailSeq As Integer
Private AltTotCnt As Integer
Private ReqTotCnt As Integer
Private StsCd As String
Private RcvDt As String
Private RcvTm As String
Private RcvId As String
Private PanelFg As String
Private ColDone As Boolean

Private ABOResult_Created As Boolean

Private objAccSql As New clsLISSqlAccession
'Private MyCollect As New clsCollectioin

'% 접수Procedure를 수행하여 결과내역을 생성한다.
Public Function DoAccession(ByVal pWorkArea As String, ByVal pAccDt As String, _
                            ByVal pAccSeq As Integer, ByVal pRcvId As String) As Boolean

    Dim i As Integer, j As Integer
    Dim ItemCnt As Integer
    Dim SqlStmt As String
    Dim DetailItem As Recordset
    Dim BodyRs As Recordset
    Dim tmpMultiRs As Recordset
    Dim tmpSpcRs As Recordset
    Dim tmpSpcCd As String
    Dim Success As Boolean
    
    '/* 전주예수병원 추가변수 */ ====================================
    Dim clsTransfer As New clsLISTransfer
    '================================================================
    
    WorkArea = pWorkArea
    AccDt = pAccDt
    AccSeq = pAccSeq
    RcvId = pRcvId
    
    ItemIndex = 0           '결과내역 갯수
    OutLabIndex = 0         '외부검사내역 갯수
    AltTotCnt = 0
    ReqTotCnt = 0
    DetailSeq = 0
    ABOResult_Created = False
    
    Set MyResult = Nothing
    Set MyOutLab = Nothing
    Set MyResult = New Collection
    Set MyOutLab = New Collection
   
    SqlStmt = objAccSql.SqlOrderItems(WorkArea, AccDt, AccSeq)       '처방Body검색
    Set BodyRs = New Recordset
    BodyRs.Open SqlStmt, DBConn
   
    If BodyRs.EOF Then      '이런 경우는 있을수 없지만.. 처방Body내역이 없는경우..
        DoAccession = False
        Exit Function
    End If
   
    i = 0      '----> 처방Body 갯수
    While (Not BodyRs.EOF)
      
        i = i + 1
        
        Call KeepData(BodyRs)
      
        Select Case TestDiv
        
            Case enTestDiv.TST_RouTest, enTestDiv.TST_MicTest, enTestDiv.TST_AboTest:       '일반검사/미생물검사/혈액형검사
      
                '// 그룹처방인 경우 //
            If PanelFg = PN_Group Then    'PanelFg = 'G'
            
                SpcCd = L201_SpcCd                      '채혈내역의 검체코드
                tmpSpcCd = BodyRs.Fields("SpcCd").Value '처방Body의 검체코드
            
            
                ColDone = True  '전체 채혈 완료
                If Trim(L201_MultiFg) <> "" Then    ' MyCollect.MultiFg <> "" Then
                    
                    '>> 연속검사인 경우 채혈접수내역의 검체Seq를 Key로 해당항목 검색
                    SqlStmt = objAccSql.SqlPanelTest(Ptid, LC2_Panel, TestCd, SpcCd, L203_SpcSeq)
               
                Else
            
                    '>> 일반 그룹처방인 경우는 모든 상세항목 검색
                    SqlStmt = objAccSql.SqlPanelTest(Ptid, LC2_Panel, TestCd, SpcCd)
            
                End If
            
                '상세항목 검색
                Set DetailItem = Nothing
                Set DetailItem = New Recordset
                DetailItem.Open SqlStmt, DBConn
                
                With DetailItem
                    While (Not .EOF)
                        TestCd = "" & .Fields("TestCd").Value               '그룹처방의 상세항목
                        RstDiv = "" & .Fields("RstDiv").Value               'Alternative/Required
                        RstType = "" & .Fields("RstType").Value             '결과유형
                        RstUnit = "" & .Fields("RstUnit").Value             '결과단위
                        PanelFg = "" & .Fields("PanelFg").Value             '그룹처방여부
                        OutLabCd = "" & .Fields("OutLabCd").Value           '외부의뢰기관
                  
                        LastRst = Trim("" & .Fields("LastRst").Value)
                        LastVfyDt = Trim("" & .Fields("LastVfyDt").Value)
                        LastVfyTm = Trim("" & .Fields("LastVfyTm").Value)
                        'LastVfyId = Val("" & .Fields("LastVfyId").Value)
                        LastVfyId1 = Trim("" & .Fields("LastVfyId").Value)
      
                        Call ProcessOneItem     '결과내역 Row 생성
                        DetailItem.MoveNext
                    Wend
                End With
            
            '// 그룹처방이 아닌 경우 //
            Else
                ColDone = True
                Call ProcessOneItem             '결과내역 Row 생성
            End If
              
            Case enTestDiv.TST_SpeTest:     '기타검사
      
                PanelFg = ""
                ColDone = True
                Call ProcessOneItem          '결과내역 Row 생성
      
        End Select
      
        '// 처방Body Status Update //
        ReDim Preserve MyOrders(4, i)
        If ColDone Then
           MyOrders(1, i) = objAccSql.SqlStatusUpdate(Ptid, OrdDt, OrdNo, OrdSeq, StsCd_LIS_Accession, _
                            Format(GetSystemDate, CS_DateDbFormat), Format(GetSystemDate, CS_TimeDbFormat), _
                            StsCd_LIS_Accession)
        Else
           MyOrders(1, i) = objAccSql.SqlStatusUpdate(Ptid, OrdDt, OrdNo, OrdSeq, StsCd_LIS_Accession, _
                            Format(GetSystemDate, CS_DateDbFormat), Format(GetSystemDate, CS_TimeDbFormat))
        End If
        MyOrders(2, i) = objAccSql.SqlHeaderUpdate(Ptid, OrdDt, OrdNo, StsCd_LIS_Accession, 1)
        MyOrders(3, i) = objAccSql.SqlHeaderUpdate(Ptid, OrdDt, OrdNo, StsCd_LIS_Accession, 2)
        
        '-- 예수병원 Status Follow Up ======================================================
        ' - Eorror 처리 누락
        Call clsTransfer.OCS_Status_Update(Ptid, OrdDt, OrdNo, _
                                         Format(GetSystemDate, CS_DateDbFormat), Format(GetSystemDate, CS_TimeDbFormat), _
                                         StsCd_LIS_Accession)
        
        '===================================================================================
        
        BodyRs.MoveNext
   
    Wend
   
    Set BodyRs = Nothing
    Set DetailItem = Nothing
    Set clsTransfer = Nothing
    
    DoAccession = SaveData
End Function

Public Function DoAccession_New(ByVal pWorkArea As String, ByVal pAccDt As String, _
                            ByVal pAccSeq As Long, ByVal pRcvId As String) As Boolean

    Dim i As Integer, j As Integer
    Dim ItemCnt As Integer
    Dim SqlStmt As String
    Dim DetailItem As Recordset
    Dim BodyRs As Recordset
    Dim tmpMultiRs As Recordset
    Dim tmpSpcRs As Recordset
    Dim tmpSpcCd As String
    Dim Success As Boolean
    
    '/* 전주예수병원 추가변수 */ ====================================
    Dim clsTransfer As New clsLISTransfer
    '================================================================
    
    WorkArea = pWorkArea
    AccDt = pAccDt
    AccSeq = pAccSeq
    RcvId = pRcvId
    
    ItemIndex = 0           '결과내역 갯수
    OutLabIndex = 0         '외부검사내역 갯수
    AltTotCnt = 0
    ReqTotCnt = 0
    DetailSeq = 0
    ABOResult_Created = False
    
    Set MyResult = Nothing
    Set MyOutLab = Nothing
    Set MyResult = New Collection
    Set MyOutLab = New Collection
   
    SqlStmt = objAccSql.SqlOrderItems(WorkArea, AccDt, AccSeq)       '처방Body검색
    Set BodyRs = New Recordset
    BodyRs.Open SqlStmt, DBConn
   
    If BodyRs.EOF Then      '이런 경우는 있을수 없지만.. 처방Body내역이 없는경우..
        DoAccession_New = False
        Exit Function
    End If
   
    i = 0      '----> 처방Body 갯수
    While (Not BodyRs.EOF)
      
        i = i + 1
        
        Call KeepData(BodyRs)
      
        Select Case TestDiv
        
            Case enTestDiv.TST_RouTest, enTestDiv.TST_MicTest, enTestDiv.TST_AboTest:       '일반검사/미생물검사/혈액형검사
      
                '// 그룹처방인 경우 //
            If PanelFg = PN_Group Then    'PanelFg = 'G'
            
                SpcCd = L201_SpcCd                      '채혈내역의 검체코드
                tmpSpcCd = BodyRs.Fields("SpcCd").Value '처방Body의 검체코드
            
            
                ColDone = True  '전체 채혈 완료
                If Trim(L201_MultiFg) <> "" Then    ' MyCollect.MultiFg <> "" Then
                    
                    '>> 연속검사인 경우 채혈접수내역의 검체Seq를 Key로 해당항목 검색
                    SqlStmt = objAccSql.SqlPanelTest(Ptid, LC2_Panel, TestCd, SpcCd, L203_SpcSeq)
               
                Else
            
                    '>> 일반 그룹처방인 경우는 모든 상세항목 검색
                    SqlStmt = objAccSql.SqlPanelTest(Ptid, LC2_Panel, TestCd, SpcCd)
            
                End If
            
                '상세항목 검색
                Set DetailItem = Nothing
                Set DetailItem = New Recordset
                DetailItem.Open SqlStmt, DBConn
                
                With DetailItem
                    While (Not .EOF)
                        TestCd = "" & .Fields("TestCd").Value               '그룹처방의 상세항목
                        RstDiv = "" & .Fields("RstDiv").Value               'Alternative/Required
                        RstType = "" & .Fields("RstType").Value             '결과유형
                        RstUnit = "" & .Fields("RstUnit").Value             '결과단위
                        PanelFg = "" & .Fields("PanelFg").Value             '그룹처방여부
                        OutLabCd = "" & .Fields("OutLabCd").Value           '외부의뢰기관
                  
                        LastRst = Trim("" & .Fields("LastRst").Value)
                        LastVfyDt = Trim("" & .Fields("LastVfyDt").Value)
                        LastVfyTm = Trim("" & .Fields("LastVfyTm").Value)
                        'LastVfyId = Val("" & .Fields("LastVfyId").Value)
                        LastVfyId1 = Trim("" & .Fields("LastVfyId").Value)
      
                        Call ProcessOneItem     '결과내역 Row 생성
                        DetailItem.MoveNext
                    Wend
                End With
            
            '// 그룹처방이 아닌 경우 //
            Else
                ColDone = True
                Call ProcessOneItem             '결과내역 Row 생성
            End If
              
            Case enTestDiv.TST_SpeTest:     '기타검사
      
                PanelFg = ""
                ColDone = True
                Call ProcessOneItem          '결과내역 Row 생성
      
        End Select
      
        '// 처방Body Status Update //
        ReDim Preserve MyOrders(4, i)
        If ColDone Then
           MyOrders(1, i) = objAccSql.SqlStatusUpdate(Ptid, OrdDt, OrdNo, OrdSeq, StsCd_LIS_Accession, _
                            Format(GetSystemDate, CS_DateDbFormat), Format(GetSystemDate, CS_TimeDbFormat), _
                            StsCd_LIS_Accession)
        Else
           MyOrders(1, i) = objAccSql.SqlStatusUpdate(Ptid, OrdDt, OrdNo, OrdSeq, StsCd_LIS_Accession, _
                            Format(GetSystemDate, CS_DateDbFormat), Format(GetSystemDate, CS_TimeDbFormat))
        End If
        MyOrders(2, i) = objAccSql.SqlHeaderUpdate(Ptid, OrdDt, OrdNo, StsCd_LIS_Accession, 1)
        MyOrders(3, i) = objAccSql.SqlHeaderUpdate(Ptid, OrdDt, OrdNo, StsCd_LIS_Accession, 2)
        
        '-- 예수병원 Status Follow Up ======================================================
        ' - Eorror 처리 누락
        Call clsTransfer.OCS_Status_Update(Ptid, OrdDt, OrdNo, _
                                         Format(GetSystemDate, CS_DateDbFormat), Format(GetSystemDate, CS_TimeDbFormat), _
                                         StsCd_LIS_Accession)
        
        '===================================================================================
        
        BodyRs.MoveNext
   
    Wend
   
    Set BodyRs = Nothing
    Set DetailItem = Nothing
    Set clsTransfer = Nothing
    
    DoAccession_New = SaveData
End Function

Private Function SaveData() As Boolean
    Dim tmpRs               As Recordset
    Dim Rs                  As Recordset
    Dim SqlStmt             As String
    Dim strSEX              As String
    Dim strAgeDay           As String
    Dim strBirth            As String
    Dim strOcsOrdNo         As String
    Dim strBussdiv          As String
    Dim ABOResult_Created   As Boolean
    Dim i                   As Integer
    
    DBConn.BeginTrans           '*** Transaction 시작 ***'
   
    On Error GoTo Err_Dup302    '--> 처방코드 Duplicate 발생시 Ignore 처리
      
        ABOResult_Created = False
        
        '결과내역(LAB302) 생성
        For i = 1 To ItemIndex
            SqlStmt = objAccSql.CreateSql_Result(WorkArea, AccDt, AccSeq, MyResult(i))
            
            Call DBConn.Execute(SqlStmt)
            
            
            '** 2001.2.14 김미경 추가 : ABO검사일 경우 h7bbs303 생성
    On Error GoTo Err_Dup303   ' : ABO검사 Duplicate 발생시 Ignore 처리
            If MyResult(i).TestDiv = TST_AboTest And (Not ABOResult_Created) Then  '3
                SqlStmt = objAccSql.CreateSql_ABOResult(WorkArea, AccDt, AccSeq, MyResult(i))
                Call DBConn.Execute(SqlStmt)
                ABOResult_Created = True
            End If
           
        Next
      
    On Error GoTo Err_Dup205    '--> 처방코드 Duplicate 발생시 Ignore 처리
      
        '외부의뢰내역(LAB205) 생성
        For i = 1 To OutLabIndex
            SqlStmt = objAccSql.CreateSql_OutLab(WorkArea, AccDt, AccSeq, MyOutLab(i))
            Call DBConn.Execute(SqlStmt)    'Sql 실행
        Next
      
    On Error GoTo Err_Trap      '--> Update 에러 발생시 Rollback Trans
      
        '채혈접수내역(LAB201) Update
        ReqTotCnt = GetTotCnt(WorkArea, AccDt, AccSeq)
        SqlStmt = objAccSql.CreateSql_ColUpdate(WorkArea, AccDt, AccSeq, StsCd_LIS_Accession, ReqTotCnt, _
                                                Format(GetSystemDate, CS_DateDbFormat), _
                                                Format(GetSystemDate, CS_TimeDbFormat), RcvId)
        Call DBConn.Execute(SqlStmt)
      
        '처방 Status
        For i = 1 To UBound(MyOrders, 2)
            '처방Body(LAB102) Status Update
            Call DBConn.Execute(MyOrders(1, i))
            '처방Header(LAB101) Status Update
            Set tmpRs = New Recordset
            tmpRs.Open MyOrders(2, i), DBConn
            
            If tmpRs.EOF Then
                Set tmpRs = Nothing
                Call DBConn.Execute(MyOrders(3, i))
            Else
                Set tmpRs = Nothing
            End If
        Next
        
        ' 접수시 환자마스터의 성별과 접수내역의 성별을 동일하게 업데이터 한다.
        ' 이유: 채혈이후에 성별의 변화가 있을수 있으므로....
        
        SqlStmt = " SELECT " & F_SEX & " AS sex, " & F_DOB & " AS dob " & _
                  "FROM " & T_HIS001 & _
                  " WHERE " & DBW(F_PTID, Ptid, 2)
        
        Set Rs = New Recordset
        Rs.Open SqlStmt, DBConn
        
        If Not Rs.EOF Then
            strSEX = Rs.Fields("sex").Value & ""
            If IsNumeric(strSEX) Then
                strSEX = Choose((Val(strSEX) Mod 2) + 1, "F", "M")
            End If
            
            '** 원본 ------------------------------------------------------------------------------------------------
            'SqlStmt = " UPDATE " & T_LAB201 & " SET " & DBW("sex", strSEX, 2) & "  WHERE " & DBW("ptid=", Ptid)
            'DBConn.Execute SqlStmt
            '--------------------------------------------------------------------------------------------------------
            
            '** 전산실 요청으로 수정 함 By M.G.Choi 2007.02.05
            If strSEX <> "" Then
                SqlStmt = " UPDATE " & T_LAB201 & " SET " & DBW("sex", strSEX, 2) & _
                          "  where workarea = " & DBS(WorkArea) & _
                          "    and accdt = " & DBS(AccDt) & _
                          "    and accseq = " & DBN(AccSeq)
                
                DBConn.Execute SqlStmt
            End If
            '-------------------------------------------------------------------------------------------------
                      
            '** 예수병원 추가루틴 By M.G.Choi ==================================
            ' * 외래/병동 무관하게 처리한다.
            strBirth = Rs.Fields("dob").Value & ""
            
            If IsDate(strBirth) = True Then
                strAgeDay = DateDiff("y", strBirth, GetSystemDate)
            Else
                strAgeDay = "0"
            End If
            
            SqlStmt = " UPDATE " & T_LAB201 & _
                      "    SET ageday = " & DBN(strAgeDay) & _
                      "  WHERE " & DBW("workarea=", WorkArea) & _
                      "    AND " & DBW("accdt=", AccDt) & _
                      "    AND " & DBW("accseq=", AccSeq)
                      
            DBConn.Execute SqlStmt
            '===================================================================
        End If
        
        Set Rs = Nothing
        
        '접수시 OCS 관련 Table 에 상태Flag를 업뎃해준다.
        
        SqlStmt = ""
        
'        SqlStmt = " SELECT a.ocsordno,b.bussdiv " & _
'                  " FROM " & T_LAB101 & " b," & T_LAB102 & " a" & _
'                  " WHERE " & dbw("a.ptid =", Ptid) & _
'                  " AND " & dbw("a.orddt=", OrdDt) & _
'                  " AND " & dbw("a.ordno=", OrdNo) & _
'                  " AND " & dbw("a.ordseq=", OrdSeq) & _
'                  " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno"
'
'        Set Rs = Nothing
'        Set Rs = New Recordset
'        Rs.Open SqlStmt, dbconn
'
'        If Not Rs.EOF Then
'            strOcsOrdNo = Val(Trim(Rs.Fields("ocsordno").Value & ""))
'            strBussdiv = Trim(Rs.Fields("bussdiv").Value & "")
'            If strOcsOrdNo <> "" Then
'                '병동은 ipd_order_dmc,ipd_order_update_dmc 업데이트
'                '외래는 opd_order_dmc 업데이트
'                If strBussdiv = enBussDiv.BussDiv_InPatient Then
'                    SqlStmt = " UPDATE med_ocs.ipd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'                    dbconn.Execute SqlStmt
'                    SqlStmt = " UPDATE med_ocs.ipd_order_update_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'                Else
'                    SqlStmt = " UPDATE med_ocs.opd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'                    dbconn.Execute SqlStmt
'                End If
'            End If
'        End If
'        Set Rs = Nothing
        
    DBConn.CommitTrans
    SaveData = True
    Set tmpRs = Nothing
    Exit Function
   
Err_Dup302:
    If Trim(MyResult(i).DetailFg) = "" Or MyResult(i).DetailFg = "*" Then ReqTotCnt = ReqTotCnt - 1
    On Error GoTo Err_Dup302    '--> 처방코드 Duplicate 발생시 Ignore 처리
    Resume Next
    
Err_Dup205:
    On Error GoTo Err_Dup205    '--> 처방코드 Duplicate 발생시 Ignore 처리
    Resume Next

Err_Dup303:
    On Error GoTo Err_Dup303   ' : ABO검사 Duplicate 발생시 Ignore 처리
    Resume Next
    
Err_Trap:
    DBConn.RollbackTrans
    
    MsgBox Err.Description, vbExclamation
    SaveData = False
    ''Resume Next
    Set tmpRs = Nothing

End Function


Private Sub KeepData(ByVal Rs As Object)
   
    Ptid = Trim("" & Rs.Fields("PtId").Value)
    OrdDt = Trim("" & Rs.Fields("OrdDt").Value)
    OrdNo = Val("" & Rs.Fields("OrdNo").Value)
    OrdSeq = Val("" & Rs.Fields("OrdSeq").Value)
    TestCd = Trim("" & Rs.Fields("OrdCd").Value)
    SpcCd = Trim("" & Rs.Fields("SpcCd").Value)
    AttrCd = Trim("" & Rs.Fields("AttrCd").Value)
    StatFg = Trim("" & Rs.Fields("StatFg").Value)
    
    TestDiv = Trim("" & Rs.Fields("TestDiv").Value)
    RstDiv = Trim("" & Rs.Fields("RstDiv").Value)
    RstType = Trim("" & Rs.Fields("RstType").Value)
    RstUnit = Trim("" & Rs.Fields("RstUnit").Value)
    PanelFg = Trim("" & Rs.Fields("PanelFg").Value)
    OutLabCd = Trim("" & Rs.Fields("OutLabCd").Value)
    
    LastRst = Trim("" & Rs.Fields("LastRst").Value)
    LastVfyDt = Trim("" & Rs.Fields("LastVfyDt").Value)
    LastVfyTm = Trim("" & Rs.Fields("LastVfyTm").Value)
    LastVfyId = 0   'Val("" & rs.Fields("LastVfyId").Value)
    LastVfyId1 = Trim("" & Rs.Fields("LastVfyId").Value)
    
    RstVal = "null"
    RstCd = ""
    HLDiv = ""
    DPDiv = ""
    VfyDt = ""
    VfyTm = ""
    VfyId = 0
    VfyId1 = ""
    MfyFg = "0"
    GrpFg = "0"
    TxtFg = "0"
'    ValFg = "0"
    SenFg = ""
    AutoFg = ""
    EqpCd = ""

    L201_TestDiv = Trim("" & Rs.Fields("L201TestDiv").Value)
    L201_MultiFg = Trim("" & Rs.Fields("MultiFg").Value)
    L201_SpcCd = Trim("" & Rs.Fields("ColSpcCd").Value)
    L203_SpcSeq = Trim("" & Rs.Fields("SpcSeq").Value)

End Sub
               
Private Sub ProcessOneItem()

    Dim tmpRs As New Recordset
    Dim SqlStmt As String
    Dim DetailCd As String
    Dim i As Integer
   
    DetailFg = ""
   
    '// 상세항목인 경우... //
    If PanelFg = PN_Detail Then   'PanelFg = 'D'
        RstDiv = "*"
        DetailFg = CStr(DetailSeq)
      
        Call SaveOneItem(TestCd)
        ReqTotCnt = ReqTotCnt + 1
        '상세항목검색
        SqlStmt = objAccSql.SqlPanelTest(Ptid, LC2_Detail, TestCd, SpcCd)
        Set tmpRs = New Recordset
        tmpRs.Open SqlStmt, DBConn
        
        While (Not tmpRs.EOF)
            With tmpRs
                DetailCd = "" & .Fields("TestCd").Value         '상세항목코드
            
                RstDiv = "" & .Fields("RstDiv").Value           'Alternative/Required
                RstType = "" & .Fields("RstType").Value         '결과유형
                RstUnit = "" & .Fields("RstUnit").Value         '결과단위
                PanelFg = "" & .Fields("PanelFg").Value         '그룹처방여부
                OutLabCd = "" & .Fields("OutLabCd").Value       '외부의뢰기관
                DetailFg = CStr(DetailSeq)                      '세부항목Seq
            
                LastRst = Trim("" & .Fields("LastRst").Value)
                LastVfyDt = Trim("" & .Fields("LastVfyDt").Value)
                LastVfyTm = Trim("" & .Fields("LastVfyTm").Value)
                'LastVfyId = Val("" & .Fields("LastVfyId").Value)
                LastVfyId1 = Trim("" & .Fields("LastVfyId").Value)
      
                If TestDiv = enTestDiv.TST_MicTest Then Call GetMicLastResult(Ptid, DetailCd, SpcCd)           '가장 최근에 Verify된 결과,결과확인일,결과확인자
                Call SaveOneItem(DetailCd)
            
                .MoveNext
            End With
        Wend
        Set tmpRs = Nothing
        DetailSeq = DetailSeq + 1
   
    '// 일반항목인 경우... //
    Else
        If TestDiv = enTestDiv.TST_MicTest Then Call GetMicLastResult(Ptid, TestCd, SpcCd)          '가장 최근에 Verify된 결과,결과확인일,결과확인자
        Call SaveOneItem(TestCd)
        ReqTotCnt = ReqTotCnt + 1
      
    End If

    '// 외부검사인 경우 //
    If Trim(OutLabCd) <> "" Then
        
        Dim objOutLab As New clsLISOutLab
        
        OutLabIndex = OutLabIndex + 1
        'ReDim Preserve MyOutLab(OutLabIndex)
        With objOutLab
            .WorkArea = WorkArea
            .AccDt = AccDt
            .AccSeq = AccSeq
            .TestCd = TestCd
            .Ptid = Ptid
            .SpcCd = SpcCd
            .OutLabCd = OutLabCd
            .StsCd = STS_OUTLAB_ACCESSION
            .RcvDt = Format(Now, CS_DateDbFormat)
            .SendDt = ""
            .ChargeDt = ""
            .SendId = ""
        End With
        
        MyOutLab.Add objOutLab  ', OutLabIndex
        
        Set objOutLab = Nothing
        
    End If
      
End Sub


'% 환자별 검사항목별로 가장 최근에 Verify된 결과를 검색한다.
Private Sub GetLastResult(ByVal PID As String, ByVal TCD As String)

    Dim tmpRs As Recordset
    Dim SqlStmt As String
   
    LastRst = ""
    LastVfyDt = ""
    LastVfyTm = ""
    'LastVfyId = 0
    LastVfyId1 = ""
   
    SqlStmt = objAccSql.SqlLastResult(PID, TCD)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
   
    With tmpRs
       If Not .EOF Then
          LastRst = Trim("" & .Fields("RstCd").Value)
          LastVfyDt = Trim("" & .Fields("VfyDt").Value)
          LastVfyTm = Trim("" & .Fields("VfyTm").Value)
          'LastVfyId = Val("" & .Fields("VfyId").Value)
          LastVfyId1 = Trim("" & .Fields("VfyId").Value)
       End If
    End With
    
    Set tmpRs = Nothing
End Sub

Private Sub GetMicLastResult(ByVal PID As String, ByVal TCD As String, ByVal SCD As String)

    Dim tmpRs As Recordset
    Dim SqlStmt As String
   
    LastRst = ""
    LastVfyDt = ""
    LastVfyTm = ""
    'LastVfyId = 0
    LastVfyId1 = ""
   
    SqlStmt = objAccSql.SqlMicLastResult(PID, TCD, SCD)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn

    If Not tmpRs.EOF Then
        LastRst = Trim("" & tmpRs.Fields("RstCd").Value)
        LastVfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
        LastVfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
        'LastVfyId = Val("" & tmpRs.Fields("VfyId").Value)
        LastVfyId1 = Trim("" & tmpRs.Fields("VfyId").Value)
    Else
        SqlStmt = objAccSql.SqlMicLastResult(PID, TCD, "")
        Set tmpRs = Nothing
        Set tmpRs = New Recordset
        tmpRs.Open SqlStmt, DBConn
        
        If Not tmpRs.EOF Then
            LastRst = Trim("" & tmpRs.Fields("RstCd").Value)
            LastVfyDt = Trim("" & tmpRs.Fields("VfyDt").Value)
            LastVfyTm = Trim("" & tmpRs.Fields("VfyTm").Value)
            'LastVfyId = Val("" & tmpRs.Fields("VfyId").Value)
            LastVfyId1 = Trim("" & tmpRs.Fields("VfyId").Value)
        End If
    End If
    
    Set tmpRs = Nothing
End Sub


Private Function GetTotCnt(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As Long

    Dim strSQL As String
    Dim tmpRs As New Recordset
    
    GetTotCnt = 0
        
    tmpRs.Open objAccSql.SqlGetTotCnt(pWorkArea, pAccDt, pAccSeq), DBConn
    
    While (Not tmpRs.EOF)
        GetTotCnt = GetTotCnt + Val(tmpRs.Fields("TotCnt").Value)
        tmpRs.MoveNext
    Wend
    Set tmpRs = Nothing
End Function

Private Sub SaveOneItem(ByVal pTestCd As String)

    Dim objAccItem As New clsLISAccItem
    
    ItemIndex = ItemIndex + 1     '결과내역Row 증가
    
    'ReDim Preserve MyResult(ItemIndex)
    With objAccItem
        .TestCd = pTestCd
        .Ptid = Ptid
        .OrdDt = OrdDt
        .OrdNo = OrdNo
        .OrdSeq = OrdSeq
        .SpcCd = SpcCd
        .AttrCd = AttrCd
        .StatFg = StatFg
      
        .RstVal = RstVal
        .RstCd = RstCd
        .HLDiv = HLDiv
        .DPDiv = DPDiv
        .VfyDt = VfyDt
        .VfyTm = VfyTm
        '.VfyId = VfyId
        .VfyId1 = VfyId1
        .MfyFg = MfyFg
        .GrpFg = GrpFg
        .TxtFg = TxtFg
'        .ValFg = ValFg
        .SenFg = SenFg
      
        .RstUnit = RstUnit
        .LastRst = LastRst
        .LastVfyDt = LastVfyDt
        .LastVfyTm = LastVfyTm
        '.LastVfyId = LastVfyId
        .LastVfyId1 = LastVfyId1
        .RstType = RstType
        .RstDiv = RstDiv
        .DetailFg = DetailFg
        .AutoFg = AutoFg
        .EqpCd = EqpCd
        
        .TestDiv = TestDiv
      
    End With
    
    MyResult.Add objAccItem ',  ItemIndex
    
    Set objAccItem = Nothing

End Sub

Private Sub Class_Terminate()
    Set objAccSql = Nothing
    Set MyResult = Nothing
    Set MyOutLab = Nothing
End Sub
