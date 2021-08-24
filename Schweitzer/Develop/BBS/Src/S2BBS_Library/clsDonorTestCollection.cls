VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDonorTestCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'+--------------------------------------------------------------------------------------+
'|  clsDonorBusiCollection의 대용입니다.
'|  S2BBS_Library 의 호환성을 유지하기 위해 하나 더 만들었읍니다.
'|
'|  wooil
'+--------------------------------------------------------------------------------------+
'|  1. Class 명 : clsCollection(DLL용)
'|  2. 기    능 : 채혈/접수내역(LAB201)을 생성한다.
'|  3. 작 성 자 : 김미경
'|
'|  CopyRight(C) 1999 대련엠티에스
'|
'+--------------------------------------------------------------------------------------+


'%  클래스 clsCollection의 Data Attributes
Public spcyy        As String       '/* 검체번호-년도   */
Public spcno        As String       '/* 검체번호        */
Public PtId         As String       '/* 환자 ID         */
Public PtNm         As String       '/* 환자명         */
Public AgeDay       As Long         '/* 환자 일령       */
Public sex          As String       '/* 환자 성별   */
Public BedIndt      As String       '/* 입원일          */
Public DeptCd       As String       '/* 진료과          */
Public Orddoct      As String       '/* 처방의          */
Public MajDoct      As String       '/* 주치의          */
Public WorkArea     As String       '/* Work Area(접수번호용) */
Public accdt        As String       '/* 접수일(접수번호용)    */
Public accseq       As String       '/* 접수순번(접수번호용)  */
Public stscd        As String       '/* STATUS - '0 ':처방,1:채혈,2:접수,3:In-Process,4:결과,5:확인,6:수정 */
Public ReqTotCnt    As Long         '/* Required Item Total Count */
Public ReqInputCnt  As Long         '/* Required Item Input Count */
Public VfyDt        As String       '/* 결과확인일자    */
Public VfyTm        As String       '/* 결과확인시간    */
Public VfyId        As String       '/* 결과확인자      */
Public coldt        As String       '/* 채혈일자        */
Public coltm        As String       '/* 채혈시간        */
Public colid        As String       '/* 채혈자          */
Public rcvdt        As String       '/* 접수일자        */
Public rcvtm        As String       '/* 접수시간        */
Public rcvid        As String       '/* 접수자          */
Public EntDt        As String       '/* 입력일자        */
Public EntTm        As String       '/* 입력시간        */
Public entid        As String       '/* 입력자          */
Public SpcCd        As String       '/* 검체코드        */
Public SpcNm        As String       '/* 검체명        */
Public MultiFg      As String       '/* 복수검체여부  */
Public OrgAccNo     As String       '/* 원 접수번호     */
Public wardid       As String       '/* 병동코드        */
Public RoomId       As String       '/* 병실코드     */
Public BedID        As String       '/* 침상코드    */
Public HosilID      As String       '/* 호실코드    */
Public FootNoteFg   As String       '/* FOOTNOTE여부('0':부,'1':여) */
Public StoreCd      As String       '/* 검체보관 코드   */
Public RptFg        As String       '/* REPORT출력여부('0':부,'1':여) */
Public testdiv      As String       '/* 검사구분 - 1:기타검사,2:미생물검사 */
Public QcFg         As String       '/* QC여부('0':부,'1':여) */
Public RmkCd        As String       '/* 검체 Remark 코드 */
Public statfg       As String       '/* 응급여부    */
Public BuildCd      As String       '/* 건물코드    */
Public BuildNm      As String       '/* 건물코드명    */
Public orgbuildcd   As String       '/* 채혈된 건물코드    */
Public TestNames    As String       '/* Barcode에 출력될 검사명들 */

Public orddt        As String
Public OrdNo        As String
Public OrdSeq       As String
Public OrdCd        As String

Public SpcGrp       As String     '/* 검체군  */
Public spcseq       As String     '/* 복수검체의 Seq */

'= Lab 채혈접수 관련 Data Type ='
Private Type TpLabGroup
   BuildCd  As String
   testdiv  As String
   WorkArea As String
   SpcCd    As String
   StoreCd  As String
   statfg   As String
   ColDate  As String
   MultiFg  As String
   SpcGrp   As String
   orddt    As String
   OrdNo    As String
   OrdSeq   As String
   OrdCd    As String
   DeptCd   As String    '/* 진료과          */
   Orddoct  As String    '/* 처방의          */
   MajDoct  As String    '/* 주치의          */
   LabelCnt As Long
   TestNm   As String    '/* 검사 약어명   */
End Type

Private Type LabNumbers
   WorkArea As String
   accdt    As String
   accseq   As String
   BuildCd  As String
End Type
   
Private Type BarcodeLabel
   WorkArea     As String
   accdt        As String
   accseq       As String
   coldt        As String
   coltm        As String
   orddt        As String
   statfg       As String
   spcno        As String
   Location     As String
   PtNm         As String
   PtId         As String
   SpcNm        As String
   StoreCd      As String
   wardid       As String
   TestNames    As String
   LabelCnt     As String
End Type
   
   
Public ErrMsg As String
Public ColCount As Integer

'공용변수 - 참조하는 프로젝트로부터 설정됨

Public CollectDone As Boolean

Private blnBatchCol As Boolean
Private ColDone As Boolean  '복수검체 채혈완료 여부
Private ReadyToCollect As Boolean
Private KeyCount As Integer
Private MultiSpcFlag As String
Private BarCount As Integer
Private mvarSortList As Object
Private LabGroup() As TpLabGroup
Private CreateLabNo() As LabNumbers
Private BarcodeBuffer() As BarcodeLabel
Private SqlBuffer() As String
Private objMySQL As New clsBBSSQLStatement

Private sWorkDt As String
Private sWorkTm As String
Private iWorkSeq As Integer

Private ColSqlStmt() As String

'Private DbConn As DrDatabase

Private mvarDonorID As String
Private mvarDonorAccdt As String
Private mvarDateTime As Date    '오늘 날짜를 넣는다.

Private objBarCode As clsBarcode

Public Property Let DateTime(ByVal vData As Date)
    mvarDateTime = vData
End Property

Public Property Get DateTime() As Date
    DateTime = mvarDateTime
End Property

Public Property Let donoraccdt(ByVal vData As String)
    mvarDonorAccdt = vData
End Property

Public Property Get donoraccdt() As String
    donoraccdt = mvarDonorAccdt
End Property

Public Property Let Donorid(ByVal vData As String)
    mvarDonorID = vData
End Property

Public Property Get Donorid() As String
    Donorid = mvarDonorID
End Property

'Public Sub setDbConn(ByRef pDbConn As DrDatabase)
''디비 커넥션
'
'    Set DbConn = pDbConn
'End Sub

Public Property Get SortList() As Object
    Set SortList = mvarSortList
End Property

Public Property Set SortList(ByVal vData As Object)
    Set mvarSortList = vData
End Property

'% 채혈내역을 생성하기 위한 기초데이타를 저장한다.
Public Sub AddLabCollect(ParamArray tmpData() As Variant)
   
   Dim tmpKeyString As String
   
   KeyCount = KeyCount + 1
   ReDim Preserve LabGroup(KeyCount)
   With LabGroup(KeyCount)
      .BuildCd = "" & tmpData(0)
      .testdiv = "" & tmpData(1)
      .WorkArea = "" & tmpData(2)
      .SpcCd = "" & tmpData(3)
      .StoreCd = "" & tmpData(4)
      .statfg = "" & tmpData(5)
      .ColDate = "" & tmpData(6)
      .MultiFg = "" & tmpData(7)
      .SpcGrp = "" & tmpData(8)
      .orddt = "" & tmpData(9)
      .OrdNo = Val("" & tmpData(10))
      .OrdSeq = Val("" & tmpData(11))
      .OrdCd = "" & tmpData(12)
      .DeptCd = "" & tmpData(13)
      .Orddoct = Val("" & tmpData(14))
      .MajDoct = Val("" & tmpData(15))
      .TestNm = "" & tmpData(16)
      .LabelCnt = Val("" & tmpData(17))
      tmpKeyString = .BuildCd & .WorkArea & .SpcCd & .StoreCd & .statfg & .ColDate
      mvarSortList.AddItem tmpKeyString & Chr(9) & KeyCount
   End With
   ReadyToCollect = True
   CollectDone = False
End Sub

'Public Sub SetWardCol(ByRef pWorkDt As String, ByRef pWorkTm As String)
'
'    Dim tmpRs As RECORDSET
'
'    blnBatchCol = True    '병동별 일괄채혈 구분
'
'    sWorkDt = Format(Now, CS_DateDbFormat)
'    sWorkTm = Format(Now, CS_TimeDbFormat)
'
'    pWorkDt = sWorkDt
'    pWorkTm = sWorkTm
'
'    Set tmpRs = objMySQL.OpenRecordSet("SELECT max(seq) as Seq FROM " & T_LAB204 & " WHERE workdt = '" & sWorkDt & _
'                                            "' AND wardid = '" & wardid & "' AND worktm = '" & sWorkTm & "' ")
'    If tmpRs.EOF Then
'        iWorkSeq = 0
'    Else
'        iWorkSeq = Val(tmpRs.Fields("Seq").Value)
'    End If
'    tmpRs.RsClose
'    Set tmpRs = Nothing
'End Sub

'% 채혈Procedure를 수행하여 채혈접수내역을 생성한다.
'% ( 복수검체의 경우 한꺼번에 모두 채혈처리 한다 --> 아래 DoCollection_Backup은 한번에 하나씩.. )
Public Function DoCollection(Optional ByRef ProgressBar As Variant) As Boolean
   
    Dim i As Integer, j As Integer
    Dim tmpStr As String
    Dim SubKey As String, SubIndex As Integer
    Dim NextKey As String
    Dim tmpSql As String
    Dim tmpSpcRs As Recordset
    Dim tmpMultiRs As Recordset
   
    If Not ReadyToCollect Then
       DoCollection = False
       Exit Function
    End If
       
    ReDim ColSqlStmt(0)

'    SaveKey = ""
    ColCount = 0
   
    stscd = "1" 'Status
    ReqTotCnt = 0       'Required Item Total Count
    ReqInputCnt = 0     'Required Item Input Count
    FootNoteFg = "0"    'FootNote구분
    RptFg = "0"         '레포트 출력 여부
    QcFg = "0"          'QC여부
    VfyDt = ""          '결과확인일
    VfyTm = ""          '결과확인시간
    VfyId = ""           '결과확인자
    rcvdt = ""          '접수일
    rcvtm = ""          '접수시간
    rcvid = ""           '접수자
    RmkCd = ""          '검체리마크
    TestNames = ""      '검사약어명String
       
    For i = 0 To mvarSortList.ListCount - 1
      
        tmpStr = mvarSortList.List(i)
        SubKey = medShift(tmpStr, vbTab)
        SubIndex = Val(tmpStr)
      
        If i < mvarSortList.ListCount - 1 Then
            tmpStr = mvarSortList.List(i + 1)
            NextKey = medShift(tmpStr, vbTab)
        Else
            NextKey = ""
        End If
        
        TestNames = TestNames & LabGroup(SubIndex).TestNm & ","      '** Barcode 라벨에 출력될 검사명들...
        '>> 기준이 달라질때마다 채혈내역 한건 생성...
        If NextKey <> SubKey Then
           
           BarCount = 0  '바코드출력장수 초기화
           
           With LabGroup(SubIndex)
            
                MultiSpcFlag = .MultiFg
                
                WorkArea = .WorkArea
                coldt = Format(mvarDateTime, CS_DateDbFormat)     '채혈일
                coltm = Format(mvarDateTime, CS_TimeDbFormat)    '채혈시간
                SpcCd = .SpcCd                          '검체
                MultiFg = .MultiFg                      '복수검체여부
                StoreCd = .StoreCd                      '검체보관구분
                testdiv = .testdiv                      '검사구분
                statfg = .statfg                        '응급여부
                SpcGrp = .SpcGrp                        '검체군
                BuildCd = .BuildCd                      '** 건물코드
                
                orddt = .orddt
                OrdNo = .OrdNo
                OrdSeq = .OrdSeq
                
                Orddoct = .Orddoct
                MajDoct = .MajDoct
                DeptCd = .DeptCd
                
                BuildNm = ObjSysInfo.Buildingnm
                
                
                
                spcno = GetSpcNo(spcyy)                  '검체번호
                accseq = GetLabNo(WorkArea, accdt)       'AccSeq
                If accseq = 0 Then GoTo Err_Trap      'AccSeq를 제대로 못가져온 경우...
                MultiFg = ""
                ColDone = True
                If BarCount < LabGroup(SubIndex).LabelCnt Then BarCount = LabGroup(SubIndex).LabelCnt   '출력장수
                Call SaveData(i, SubKey)  '저장
                '** Barcode 출력내용을 Buffering 한다.
                Call PrintBarcode(ColCount, .ColDate)
                TestNames = ""
                If Not CollectDone Then GoTo Err_Trap
          End With
      End If
      'Progress Bar 증가...
      If Not IsMissing(ProgressBar) Then
            If ProgressBar.Value = ProgressBar.Max Then ProgressBar.Max = ProgressBar.Max + 1
            ProgressBar.Value = ProgressBar.Value + 1
      End If
      DoEvents
   Next
   
   DoCollection = True
   Exit Function
   
Err_Trap:
   DoCollection = False
   Set tmpSpcRs = Nothing
   Set tmpMultiRs = Nothing
End Function

'% 클래스의 채혈 내역을 DB로 저장한다.
Public Sub SaveData(ByVal ii As Integer, ByVal KeyString As String)
   
   Dim SqlStmt As String
   Dim tmpStr As String
   Dim tmpRs As Recordset
   Dim i As Integer
   Dim lngRequestSeq As Long
   
   'ColCount 증가
   ColCount = ColCount + 1
   ReDim Preserve CreateLabNo(ColCount)

   '--채혈접수 동시 수행을 위해 생성된 LabNo를 Keep한다.
   CreateLabNo(ColCount).WorkArea = WorkArea
   CreateLabNo(ColCount).accdt = accdt
   CreateLabNo(ColCount).accseq = accseq
   CreateLabNo(ColCount).BuildCd = BuildCd
      
   '검사의뢰 순번의 최고값을 불러온다.
    
   lngRequestSeq = GetMaxRequestSeq(mvarDonorID, mvarDonorAccdt, Format(mvarDateTime, PRESENTDATE_FORMAT)) + ColCount
    
   ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
   
   On Error GoTo Err_Trap

'--채혈접수내역 생성
   ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_Collection   '// Sql문장 생성
'   Debug.Print "Colsqlstmt(" & UBound(ColSqlStmt) & ") : " & ColSqlStmt(UBound(ColSqlStmt))
   
'--검사의뢰내역 SQL
    
    ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
    
    ColSqlStmt(UBound(ColSqlStmt)) = objMySQL.SetTestRequest(mvarDonorID, mvarDonorAccdt, _
                            Format(mvarDateTime, PRESENTDATE_FORMAT), lngRequestSeq, CreateLabNo(ColCount).WorkArea, CreateLabNo(ColCount).accdt, CreateLabNo(ColCount).accseq)

'--헌혈자 접수 내역 업데이트
    
    ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
    ColSqlStmt(UBound(ColSqlStmt)) = objMySQL.SetDonorAccHistoryUpdateByTmpID(mvarDonorID, mvarDonorAccdt, PtId)

'--처방 Status Follow Up
   
   For i = 0 To mvarSortList.ListCount - 1
      tmpStr = mvarSortList.List(i)
      If KeyString <> medShift(tmpStr, Chr(9)) Then GoTo Skip
      With LabGroup(Val(tmpStr))
         If ColDone Then            '--> 연속검사일 경우 모두 채혈수행이 되야 처방Status Follow Up.
            'LAB102의 DoneFg = '1' : 채혈완료
            ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
            
            ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_UpdateLabNo(PtId, .orddt, .OrdNo, .OrdSeq, "1", "1") '// Sql문장 생성
         Else
            'LAB102의 DoneFg = '0' : 채혈미완료
            
            ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
            
            ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_UpdateLabNo(PtId, .orddt, .OrdNo, .OrdSeq, "1")  '// Sql문장 생성
         End If
            
         SqlStmt = CreateSql_HeaderUpdate(PtId, .orddt, .OrdNo, "1", 1)
        
         Set tmpRs = New Recordset
         tmpRs.Open SqlStmt, DBConn
         If tmpRs.EOF Then
            Set tmpRs = Nothing
            
            ReDim Preserve ColSqlStmt(UBound(ColSqlStmt) + 1)
            
            ColSqlStmt(UBound(ColSqlStmt)) = CreateSql_HeaderUpdate(PtId, .orddt, .OrdNo, "1", 2)
         Else
            Set tmpRs = Nothing
         End If
      End With
Skip:
   Next
     
   CollectDone = True
   Exit Sub

Err_Trap:
   CollectDone = False
   Set tmpRs = Nothing
End Sub

Private Function GetMaxRequestSeq(ByVal Donorid As String, ByVal donoraccdt As String, _
                                  ByVal orddt As String) As Long
'DonorID, DonorAccDt, OrdDt 가 같은 조건에서
'검사의뢰번호의 최대값을 얻어온다.
                                  
    Dim RsMaxRSeq As Recordset
    Dim objMaxRSeq As clsBBSSQLStatement
    
    Set objMaxRSeq = New clsBBSSQLStatement
    With objMaxRSeq
'        .setDbConn DbConn
        Set RsMaxRSeq = New Recordset
        RsMaxRSeq.Open .GetTestRequestMaxSeq(Donorid, donoraccdt, orddt), DBConn
    End With
    
    If RsMaxRSeq.EOF Then
        GetMaxRequestSeq = 0
    Else
        If IsNull(RsMaxRSeq.Fields("maxseq").Value) Then
            GetMaxRequestSeq = 0
        Else
            GetMaxRequestSeq = RsMaxRSeq.Fields("maxseq").Value & ""
        End If
    End If
    
    Set RsMaxRSeq = Nothing
    Set objMaxRSeq = Nothing
End Function

Public Function ExecuteSqlStmt() As Boolean
    
    Dim i As Long

On Error GoTo ErrExecute

    For i = 1 To UBound(ColSqlStmt)
'        Debug.Print " i  = " & i & " : " & ColSqlStmt(i)
        DBConn.Execute ColSqlStmt(i)
    Next
        
    ExecuteSqlStmt = True
        
    Exit Function
ErrExecute:
    ExecuteSqlStmt = False
End Function

'% 검체번호를 부여한다.
Private Function GetSpcNo(ByVal spcyy As String) As Long
   
   Dim SpcNoRs As Recordset
   
   Dim tmpSql As String
   
   GetSpcNo = 0
   
   '먼저 Lock을 걸기위해 의미없는 update문 실행.
   'select는 Lock이 안걸리기 때문에 극히 짧은 순간에도
   '검체번호가 중복될수 있음.. 을 방지.    2000.3.10  by 김미경
   tmpSql = CreateSql_SpcNo(spcyy, 4)
   
   DBConn.BeginTrans
   
   On Error GoTo Err_Trap
   
   DBConn.Execute tmpSql   'Lock 걸림
   
   tmpSql = CreateSql_SpcNo(spcyy, 1)
   Set SpcNoRs = New Recordset
   SpcNoRs.Open tmpSql, DBConn
   If SpcNoRs.EOF Then
      GetSpcNo = 1
      tmpSql = CreateSql_SpcNo(spcyy, 2, GetSpcNo)
   Else
      GetSpcNo = Val("" & SpcNoRs.Fields("Seq").Value) + 1
      tmpSql = CreateSql_SpcNo(spcyy, 3, GetSpcNo)
   End If
   Set SpcNoRs = Nothing
   
   DBConn.Execute tmpSql
   DBConn.CommitTrans
   
   Exit Function
   
Err_Trap:
    DBConn.RollbackTrans
'    DBConn.DisplayErrors
   Set SpcNoRs = Nothing
   
End Function

'% 접수번호를 부여한다.
Private Function GetLabNo(ByVal pWorkArea As String, ByRef pAccDt As String) As Integer
   
   Dim tmpRs As Recordset
   Dim tmpSql As String
   Dim LabDiv As String
   Dim tmpStr As String
   Dim tmpRng1 As Integer, tmpRng2 As Integer
   Dim tmpSpcGrp As String
   
   GetLabNo = 0
   '// Sql문'접수번호 부여..
   tmpSql = objMySQL.SqlCommonCode(T_LAB032, CD2_WORKAREA, pWorkArea)
   '//
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   LabDiv = "" & tmpRs.Fields("Field2").Value
   
   tmpRng1 = 1
   tmpRng2 = 9999
   tmpSpcGrp = "0"
   Select Case LabDiv
   Case "1":       '일단위
      pAccDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
   Case "2":       '월단위
      pAccDt = Format(GetSystemDate, "yyyyMM")
   Case "3":       '년단위
      pAccDt = Format(GetSystemDate, "yyyy")
   Case "4":       '검체단위
      pAccDt = Format(GetSystemDate, "yyyyMM")
      tmpSpcGrp = SpcGrp
      '// Sql문
      tmpSql = objMySQL.SqlCommonCode(T_LAB032, CD2_SGROUP, tmpSpcGrp)
      '//
      Set tmpRs = New Recordset
      tmpRs.Open tmpSql, DBConn
      '검체단위인데 검체군 마스터가 없는 경우... (이런일은 발생하면 안되지...)
      If Not tmpRs.EOF Then
         tmpStr = "" & tmpRs.Fields("Field2").Value '검체군별 Lab No 범위
         tmpRng1 = Val(medGetP(tmpStr, 2, ";"))     '검체군별 Lab No 시작번호
         tmpRng2 = Val(medGetP(tmpStr, 3, ";"))     '검체군별 Lab No 끝번호
      End If
   Case Else:
      pAccDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
   End Select
   
   '먼저 Lock을 걸기위해 의미없는 update문 실행.
   'select는 Lock이 안걸리기 때문에 극히 짧은 순간에도
   '접수번호가 중복될수 있음.. 을 방지.    2000.3.10  by 김미경
   tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 4)
   
   On Error GoTo Err_Trap
   
   DBConn.BeginTrans
   DBConn.Execute tmpSql   'Lock 걸림
   
   '// Sql문장 생성
   tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 1)
   '//
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   If tmpRs.EOF Then
      GetLabNo = tmpRng1
      tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 2, GetLabNo)
   Else
      GetLabNo = Val(tmpRs.Fields("Seq").Value & "") + 1
      If GetLabNo > tmpRng2 Then
         MsgBox "접수번호가 Range(" & tmpRng1 & "-" & tmpRng2 & ")를 벗어났읍니다. : " & GetLabNo
         GoTo Err_Trap
      End If
      tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 3, GetLabNo)
   End If
   Set tmpRs = Nothing
      
   DBConn.Execute tmpSql
   DBConn.CommitTrans
   
   Exit Function
   
Err_Trap:
   DBConn.RollbackTrans
'   DBConn.DisplayErrors
   Set tmpRs = Nothing
   Resume Next
   GetLabNo = 0
   Exit Function

End Function

Public Sub GetLabNumbers(ByVal AryIndex As Integer, ByRef pWorkArea As String, _
                                        ByRef pAccDt As String, ByRef pAccSeq As Integer, Optional ByRef pBuildCd As Variant)
   
   If CollectDone Then
      With CreateLabNo(AryIndex)
         pWorkArea = .WorkArea
         pAccDt = .accdt
         pAccSeq = .accseq
         pBuildCd = .BuildCd
      End With
   End If
End Sub


'% Method 1 : CollectQuery
'%                 Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                 클래스 clsCollection의 Data Attribute에 저장한다.

Public Function CollectQuery(ByVal pWorkArea As String, _
                                         ByVal pAccDt As String, ByVal pAccSeq As Integer) As Boolean

   Dim SqlStmt As String
   Dim MyOraDS As Recordset
   
   SqlStmt = " SELECT a.* " & _
                 " FROM " & T_LAB201 & " a " & _
                 " WHERE " & DBW("a.workarea", pWorkArea, 2) & _
                 " AND   " & DBW("a.accdt", pAccDt, 2) & _
                 " AND   " & DBW("a.accseq", pAccSeq, 2)
   
   Set MyOraDS = New Recordset
   MyOraDS.Open SqlStmt, DBConn
   If MyOraDS.EOF Then
      CollectQuery = False
      Exit Function
   End If
   
   With MyOraDS
      
      .MoveFirst
         
      spcyy = Trim("" & .Fields("SpcYy").Value)
      spcno = Val("" & .Fields("SpcNo").Value)
      PtId = Trim("" & .Fields("PtId").Value)
      sex = Trim("" & .Fields("Sex").Value)
      AgeDay = Val("" & .Fields("AgeDay").Value)
      BedIndt = Trim("" & .Fields("BedInDt").Value)
      DeptCd = Trim("" & .Fields("DeptCd").Value)
      Orddoct = Val("" & .Fields("OrdDoct").Value)
      MajDoct = Val("" & .Fields("MajDoct").Value)
      WorkArea = Trim("" & .Fields("WorkArea").Value)
      accdt = Trim("" & .Fields("AccDt").Value)
      accseq = Val("" & .Fields("AccSeq").Value)
      stscd = Trim("" & .Fields("StsCd").Value)
      ReqTotCnt = Val("" & .Fields("ReqTotCnt").Value)
      ReqInputCnt = Val("" & .Fields("ReqInputCnt").Value)
      VfyDt = Trim("" & .Fields("VfyDt").Value)
      VfyTm = Trim("" & .Fields("VfyTm").Value)
      VfyId = Val("" & .Fields("VfyId").Value)
      coldt = Trim("" & .Fields("ColDt").Value)
      coltm = Trim("" & .Fields("ColTm").Value)
      colid = Val("" & .Fields("ColId").Value)
      rcvdt = Trim("" & .Fields("RcvDt").Value)
      rcvtm = Trim("" & .Fields("RcvTm").Value)
      rcvid = Val("" & .Fields("RcvId").Value)
      EntDt = Trim("" & .Fields("EntDt").Value)
      EntTm = Trim("" & .Fields("EntTm").Value)
      entid = Val("" & .Fields("EntId").Value)
      SpcCd = Trim("" & .Fields("SpcCd").Value)
      MultiFg = Trim("" & .Fields("MultiFg").Value)
      OrgAccNo = Trim("" & .Fields("OrgAccNo").Value)
      wardid = Trim("" & .Fields("WardId").Value)
      RoomId = Trim("" & .Fields("RoomId").Value)
      BedID = Trim("" & .Fields("BedId").Value)
      HosilID = Trim("" & .Fields("HosilId").Value)
      FootNoteFg = Trim("" & .Fields("FootNoteFg").Value)
      StoreCd = Trim("" & .Fields("StoreCd").Value)
      RptFg = Trim("" & .Fields("RptFg").Value)
      testdiv = Trim("" & .Fields("TestDiv").Value)
      QcFg = Trim("" & .Fields("QcFg").Value)
      RmkCd = Trim("" & .Fields("RmkCd").Value)
      statfg = Trim("" & .Fields("StatFg").Value)
      BuildCd = Trim("" & .Fields("BuildCd").Value)
      
   End With
   
End Function

Private Function CreateSql_Collection() As String

   CreateSql_Collection = "Insert into " & _
                        "" & T_LAB201 & " (spcyy, spcno, ptid, sex, ageday, bedindt, deptcd, orddoct, majdoct, " & _
                        "            workarea, accdt, accseq, stscd, reqtotcnt, reqinputcnt, " & _
                        "            vfydt, vfytm, vfyid, coldt, coltm, colid, rcvdt, rcvtm, " & _
                        "            rcvid, entdt, enttm, entid, spccd, multifg, orgaccno, wardid, roomid, " & _
                        "            bedid, footnotefg, storecd, rptfg, testdiv, qcfg, rmkcd, statfg, buildcd, orgbuildcd, hosilid) " & _
                        "Values  (" & DBV("spcyy", spcyy, 1) & DBV("spcno", spcno, 1) & DBV("ptid", PtId, 1) & DBV("sex", sex, 1) & DBV("ageday", AgeDay, 1) _
                                    & DBV("bedindt", BedIndt, 1) & DBV("deptcd", DeptCd, 1) & DBV("orddoct", Orddoct, 1) & DBV("majdoct", MajDoct, 1) _
                                    & DBV("workarea", WorkArea, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("stscd", stscd, 1) _
                                    & DBV("reqtotcnt", ReqTotCnt, 1) & DBV("reqinputcnt", ReqInputCnt, 1) & DBV("vfydt", VfyDt, 1) & DBV("vfytm", VfyTm, 1) & DBV("vfyid", VfyId, 1) _
                                    & DBV("coldt", coldt, 1) & DBV("coltm", coltm, 1) & DBV("colid", colid, 1) & DBV("rcvdt", rcvdt, 1) & DBV("rcvtm", rcvtm, 1) _
                                    & DBV("rcvid", rcvid, 1) & DBV("entdt", EntDt, 1) & DBV("enttm", EntTm, 1) & DBV("entid", entid, 1) & DBV("spccd", SpcCd, 1) _
                                    & DBV("multifg", MultiFg, 1) & DBV("orgaccno", OrgAccNo, 1) & DBV("wardid", wardid, 1) & DBV("roomid", RoomId, 1) & DBV("bedid", BedID, 1) _
                                    & DBV("footnotefg", FootNoteFg, 1) & DBV("storecd", StoreCd, 1) & DBV("rptfg", RptFg, 1) & DBV("testdiv", testdiv, 1) & DBV("qcfg", QcFg, 1) _
                                    & DBV("rmkcd", RmkCd, 1) & DBV("statfg", statfg, 1) & DBV("buildcd", BuildCd, 1) & DBV("orgbuildcd", orgbuildcd, 1) & DBV("hosilid", HosilID) & ") "

End Function


Private Function CreateSql_MultiSpc() As String

   CreateSql_MultiSpc = "Insert into " & T_LAB203 & " ( ptid, orddt, ordno, ordseq, spcseq, spccd, workarea, accdt, accseq ) " & _
                        "Values ( " & DBV("ptid", PtId, 1) & DBV("orddt", orddt, 1) & DBV("ordno", OrdNo, 1) & DBV("ordseq", OrdSeq, 1) & DBV("spcseq", spcseq, 1) _
                                    & DBV("spccd", SpcCd, 1) & DBV("workarea", WorkArea, 1) & DBV("accdt", accdt, 1) & DBV("accseq", accseq) & " )"

End Function

Private Function CreateSql_WardCol() As String

   CreateSql_WardCol = "Insert into " & T_LAB204 & " ( workdt, wardid, worktm, seq, workarea, accdt, accseq, colid, buildcd, orgbuildcd ) " & _
                       "Values ( " & DBV("workdt", sWorkDt, 1) & DBV("wardid", wardid, 1) & DBV("worktm", sWorkTm, 1) & DBV("seq", iWorkSeq, 1) & DBV("workarea", WorkArea, 1) & _
                                     DBV("accdt", accdt, 1) & DBV("accseq", accseq, 1) & DBV("colid", colid, 1) & DBV("buildcd", BuildCd, 1) & DBV("orgbuildcd", orgbuildcd) & " )"

End Function

Private Function CreateSql_UpdateLabNo(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As Integer, _
                                                           ByVal OrdSeq As String, ByVal stscd As String, Optional ByVal DoneFg As Variant) As String
   Dim tmpStr As String
   
   If IsMissing(DoneFg) Then
      tmpStr = ""
   Else
      tmpStr = ", donefg " & DBV("donefg", DoneFg, 2) & " "
   End If
   CreateSql_UpdateLabNo = "Update S2ORD102 " & _
                            "Set    " & DBW("workarea=", WorkArea) & ", " & _
                            "       " & DBW("accdt   =", accdt) & ", " & _
                            "       " & DBW("accseq  =", accseq) & ", " & _
                            "       " & DBW("stscd   =", stscd) & " " & tmpStr & _
                            "WHERE  " & DBW("ptid    =", PtId) & " " & _
                            "AND    " & DBW("orddt   =", orddt) & " " & _
                            "AND    " & DBW("ordno   =", OrdNo) & " " & _
                            "AND    " & DBW("ordseq  =", OrdSeq)
'   '예수병원 요
'   CreateSql_UpdateLabNo = "Update mdbldort " & _
'                            "Set    " & DBW("workarea=", WorkArea) & ", " & _
'                            "       " & DBW("accdt   =", accdt) & ", " & _
'                            "       " & DBW("accseq  =", accseq) & ", " & _
'                            "       " & DBW("stscd   =", stscd) & " " & tmpStr & _
'                            "WHERE  " & DBW("patno=", PtId) & " " & _
'                            "AND     orddate=to_date(" & orddt & ",'yyyymmdd')" & _
'                            "AND    " & DBW("ordseqno   =", OrdNo)
End Function

Private Function CreateSql_HeaderUpdate(ByVal PtId As String, ByVal orddt As String, ByVal OrdNo As Integer, _
                                                            ByVal DoneFg As String, ByVal flag As Integer)
   
    If flag = 1 Then
        CreateSql_HeaderUpdate = "SELECT * FROM S2ORD102 " & _
                                 "WHERE  " & DBW("ptid=", PtId) & " " & _
                                 "AND    " & DBW("orddt=", orddt) & " " & _
                                 "AND    " & DBW("ordno=", OrdNo) & " " & _
                                 "AND    " & DBW("dcfg=", "") & " " & _
                                 "AND    " & DBW("donefg<", DoneFg)
    Else
        CreateSql_HeaderUpdate = "Update S2ORD101 " & _
                                 "Set    " & DBW("donefg=", DoneFg) & " " & _
                                 "WHERE  " & DBW("ptid=", PtId) & " " & _
                                 "AND    " & DBW("orddt=", orddt) & " " & _
                                 "AND    " & DBW("ordno=", OrdNo)
'    '예수병원 용
'   CreateSql_HeaderUpdate = "Update mdbldort " & _
'                            "Set    " & DBW("donefg=", DoneFg) & " " & _
'                            "WHERE  " & DBW("patno=", PtId) & " " & _
'                            "AND    orddate=to_date(" & orddt & ",'yyyymmdd')" & _
'                            "AND    " & DBW("ordseqno=", OrdNo)
    End If
End Function

Private Function CreateSql_SpcNo(ByVal spcyy As String, ByVal flag As Integer, Optional ByVal seq As Variant) As String
   
    
    '나머지는 새로운 기준의 번호부여 SQL문 사용
        Select Case flag
            Case 1:
                CreateSql_SpcNo = " SELECT seq " & _
                                  " FROM   " & T_COM099 & _
                                  " WHERE  " & DBW("noindex =", COM99_LIS_Specimen) & _
                                  " AND    " & DBW("divcd1  =", LIS_ORDDIV) & _
                                  " AND    " & DBW("divcd2  =", spcyy) & _
                                  " AND    " & DBW("divcd3  =", "0")
            Case 2:
                CreateSql_SpcNo = " Insert Into " & T_COM099 & " (noindex, divcd1, divcd2, divcd3, seq) " & _
                                  " Values (" & DBV("noindex", COM99_LIS_Specimen, 1) & DBV("divcd1", LIS_ORDDIV, 1) _
                                              & DBV("divcd2", spcyy, 1) & DBV("divcd3", "0", 1) & DBV("seq", seq) & ")"
            Case 3:
                CreateSql_SpcNo = " Update " & T_COM099 & _
                                  " Set    " & DBW("seq =", seq) & _
                                  " WHERE  " & DBW("noindex =", COM99_LIS_Specimen) & _
                                  " AND    " & DBW("divcd1  =", LIS_ORDDIV) & _
                                  " AND    " & DBW("divcd2  =", spcyy) & _
                                  " AND    " & DBW("divcd3  =", "0")
            Case 4:
                CreateSql_SpcNo = " Update " & T_COM099 & _
                                  " Set    seq = seq " & _
                                  " WHERE  " & DBW("noindex =", COM99_LIS_Specimen) & _
                                  " AND    " & DBW("divcd1  =", LIS_ORDDIV) & _
                                  " AND    " & DBW("divcd2  =", spcyy) & _
                                  " AND    " & DBW("divcd3  =", "0")
        End Select


End Function

Private Function CreateSql_LabNo(ByVal WorkArea As String, ByVal accdt As String, ByVal SpcGrp As String, _
                                                 ByVal flag As Integer, Optional ByVal seq As Variant) As String
    
'나머지는 새로운 기준의 번호부여 SQL문 사용
    Select Case flag
        Case 1:
            CreateSql_LabNo = " SELECT seq " & _
                              " FROM   " & T_COM099 & " " & _
                              " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                              " AND    " & DBW("divcd1 =", WorkArea) & _
                              " AND    " & DBW("divcd2 =", accdt) & _
                              " AND    " & DBW("divcd3 =", SpcGrp)
        Case 2:
            CreateSql_LabNo = "Insert Into " & T_COM099 & " (noindex, divcd1, divcd2, divcd3, seq) " & _
                              "Values (" & DBV("noindex", COM99_LIS_LabNo, 1) & DBV("divcd1", WorkArea, 1) _
                                         & DBV("divcd2", accdt, 1) & DBV("divcd3", SpcGrp, 1) & DBV("seq", seq) & ")"
        Case 3:
            CreateSql_LabNo = " Update " & T_COM099 & _
                              " Set    " & DBW("seq =", seq) & _
                              " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                              " AND    " & DBW("divcd1 =", WorkArea) & _
                              " AND    " & DBW("divcd2 =", accdt) & _
                              " AND    " & DBW("divcd3 =", SpcGrp)
        Case 4:
            CreateSql_LabNo = " Update " & T_COM099 & " Set seq = seq " & _
                              " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                              " AND    " & DBW("divcd1 =", WorkArea) & _
                              " AND    " & DBW("divcd2 =", accdt) & _
                              " AND    " & DBW("divcd3 =", SpcGrp)
    End Select
End Function

Private Sub Class_Initialize()
   blnBatchCol = False
   Set objMySQL = New clsBBSSQLStatement
   
    Set objBarCode = New clsBarcode
    Set objBarCode.TableInfo = New clsTables
    Set objBarCode.FieldInfo = New clsFields
   
'   Call InitRtn
End Sub

Private Sub Class_Terminate()
   Set objMySQL = Nothing
   Set objBarCode = Nothing
End Sub

Public Sub InitRtn()
   KeyCount = 0
   ColCount = 0
   mvarSortList.Clear
   ReadyToCollect = False
   CollectDone = False
   ColDone = False
   Erase SqlBuffer
   ErrMsg = ""
End Sub

'---------------------------
'바코드 출력 항목을 담아논다
'---------------------------
Public Function PrintBarcode(ByVal ColCnt As Integer, ByVal pColTm As String) As String

   Dim tmpSql As String
   Dim tmpRs As Recordset

   ReDim Preserve BarcodeBuffer(ColCnt)

   tmpSql = objMySQL.SqlCommonCode(T_LAB032, CD2_SPECIMEN, SpcCd)
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   SpcNm = "" & tmpRs.Fields("Field5").Value
   Set tmpRs = Nothing

   With BarcodeBuffer(ColCnt)
      .Location = BuildNm
      .WorkArea = WorkArea
      .orddt = Mid(orddt, 5, 2) & "/" & Mid(orddt, 7, 2)
      .coldt = Mid(accdt, 3)
      .coltm = Mid(pColTm, 12, 5)
      .accseq = Format(accseq, "@@@@@")
      .statfg = statfg
      
      .spcno = spcyy & Format(spcno, LIS_BarFormat)
      
      .PtNm = Mid(PtNm, 1, 5)
      .PtId = PtId
      .SpcNm = SpcNm
      .StoreCd = StoreCd
      If Trim(wardid) <> "" Then
         .wardid = HosilID  'WardId
      Else
         .wardid = DeptCd
      End If
      .TestNames = TestNames
      .LabelCnt = BarCount
   End With

End Function

'---------------------------
'실제 바코드가 출력되는 함수
'---------------------------
Public Function GetBarcodeLabel(ByVal AryIndex As Integer, Optional ByVal AccFg As Boolean = False) As String
'    MsgBox "바코드 출력하는 곳이네..... 이거 쓰는데 없으면 지워.... ", vbCritical
    
    If CollectDone Then
        With BarcodeBuffer(AryIndex)
'            Call ObjBBSComCode.BarInfo.Label_PrintOut("임상", .WorkArea, .coldt, .accseq, .spcno, .PtId, .PtNm, .SpcNm, .StoreCd, .statfg, .wardid, .orddt, .coltm, .TestNames, .LabelCnt)
            Call objBarCode.Label_PrintOut("임상", .WorkArea, .coldt, .accseq, .spcno, .PtId, .PtNm, .SpcNm, .StoreCd, .statfg, .wardid, .orddt, .coltm, .TestNames, .LabelCnt)
        End With
    End If
End Function
