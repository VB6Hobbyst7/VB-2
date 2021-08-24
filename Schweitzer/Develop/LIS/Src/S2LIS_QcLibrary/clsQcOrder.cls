VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQcOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SpcYY As String      '/* 검체번호-년도   */
Public SpcNo As Long        '/* 검체번호        */
Public PtId As String       '/* 환자 ID         */
Public PtNm As String       '/* 환자명          */
Public AgeDay As Long       '/* 환자 일령       */
Public Sex As String        '/* 환자 성별       */
Public BedInDt As String    '/* 입원일          */
Public DeptCd As String     '/* 진료과          */
Public OrdDoct As String   '/* 처방의          */

Public MajDoct As String   '/* 주치의          */
Public WorkArea As String   '/* Work Area(접수번호용) */
Public AccDt As String      '/* 접수일(접수번호용)    */
Public AccSeq As Integer    '/* 접수순번(접수번호용)  */
Public StsCd As String      '/* STATUS - '0 ':처방,1:채혈,2:접수,3:In-Process,4:결과,5:확인,6:수정 */
Public ReqTotCnt As Integer '/* Required Item Total Count */
Public ReqInputCnt As Integer  '/* Required Item Input Count */
Public VfyDt As String      '/* 결과확인일자    */
Public VfyTm As String      '/* 결과확인시간    */
Public VfyId As String     '/* 결과확인자      */
Public ColDt As String      '/* 채혈일자        */
Public ColTm As String      '/* 채혈시간        */
Public ColId As String     '/* 채혈자          */
Public RcvDt As String      '/* 접수일자        */
Public RcvTm As String      '/* 접수시간        */
Public RcvId As String     '/* 접수자          */
Public EntDt As String      '/* 입력일자        */
Public EntTm As String      '/* 입력시간        */
Public EntId As String     '/* 입력자          */
Public SpcCd As String      '/* 검체코드        */
Public SpcNm As String      '/* 검체명          */
Public MultiFg As String    '/* 복수검체여부    */
Public OrgAccNo As String   '/* 원 접수번호     */
Public WardId As String     '/* 병동코드        */
Public RoomId As String     '/* 병실코드        */
Public BedId As String      '/* 침상코드        */
Public HosilId As String    '/* 침상코드        */
Public FootNoteFg As String '/* FOOTNOTE여부('0':부,'1':여) */
Public StoreCd As String    '/* 검체보관 코드   */
Public RptFg As String      '/* REPORT출력여부('0':부,'1':여) */
Public TestDiv As String    '/* 검사구분 - 1:기타검사,2:미생물검사 */
Public QcFg As String       '/* QC여부('0':부,'1':여) */
Public RmkCd As String      '/* 검체 Remark 코드*/
Public StatFg As String     '/* 응급여부        */
Public BuildCd As String    '/* 건물코드        */
Public BuildNm As String    '/* 건물명        */
Public OrgBuildCd As String '/* 채혈된 건물코드 */
Public TestNames As String  '/* Barcode에 출력될 검사명들 */

Public Controlcd As String  'Control코드
Public ControlNm As String  'COntrol명
Public EqpNm As String      '장비명
Public LevelCd As String    'Level코드
Public Lotno As String      'Lot Number

Public TestCd As String
Public RstVal As String
Public RstCd As String
Public RstDiv As String
Public RaDiv As String
Public RstUnit As String
Public RstType As String
Public DetailFg As String
Public MfyFg As String
Public TxtFg As String
Public AutoFg As String
Public EqpCd As String
Public Method As String

Public OrdDt As String
Public OrdNo As Integer
Public OrdSeq As Integer
Public OrdCd As String

Public SpcGrp As String     '/* 검체군          */
Public SpcSeq As String     '/* 복수검체의 Seq  */
Public ColCount As Integer
Public BarCount As Integer

Private ColDate As String
Private SqlCreate201 As String      '접수내역
Private SqlCreate026() As String    '결과내역
Private mItemCol() As String
Private mvarItemCount As Integer '--Item 카운트

Private DetailSeq As Integer

Private Type BarcodeLabel
    WorkArea As String
    AccDt As String
    AccSeq As Integer
    ColDt As String
    ColTm As String
    OrdDt As String
    StatFg As String
    SpcNo As String
    Location As String
    PtNm As String
    PtId As String
    SpcNm As String
    StoreCd As String
    WardId As String
    TestNames As String
    LabelCnt As String
End Type
   
Private mvarMyQc As clsQcMst
Private BarcodeBuffer() As BarcodeLabel
Private CollectDone As Boolean

Private objBarCode As clsBarcode

Private mvarTrans As Boolean    '외부에서 트랙잭션을 받을지 말지를 판단.(QC자동처방에서는 트랜잭션을 외부에서 통합관리한다.)

Public Property Set MyQc(ByVal vNewValue As clsQcMst)
    Set mvarMyQc = vNewValue
End Property

Public Property Let Trans(ByVal vData As Boolean)
    mvarTrans = vData
End Property

Public Property Get Trans() As Boolean
     Trans = mvarTrans
End Property

Public Function GetTestNames(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByRef pCount As Integer) As String

    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlTestNames(pWorkArea, pAccDt, pAccSeq)
    Set objSqlQc = Nothing
    
    GetTestNames = ""
    pCount = 0
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    While (Not Rs.EOF)
        pCount = pCount + 1
        GetTestNames = GetTestNames & "" & Rs.Fields("AbbrNm5").Value & vbTab
        Rs.MoveNext
    Wend
    
Nodata:
    Set Rs = Nothing
    
End Function

Public Function GetLastLotNo(ByVal pCtrlCd As String, ByVal pLevelCd As String) As String

    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    Set objSqlQc = New clsLISSqlQc
    
    SqlStmt = objSqlQc.SqlLastLotNo(pCtrlCd, pLevelCd, _
                                    Format(GetSystemDate, "yyyyMMdd"), _
                                    Format(GetSystemDate, "yyyyMMdd"))
                                    
    Set Rs = New Recordset
    Rs.Open SqlStmt
    
    If Rs.EOF Then GoTo Nodata
    
    GetLastLotNo = "" & Rs.Fields("LotNo").Value
    
Nodata:
    Set Rs = Nothing
    Set objSqlQc = Nothing
End Function

'Public Function SqlLastLotNo(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
'                             ByVal pFrDt As String, ByVal pToDt As String) As String
'
'    SqlLastLotNo = "select  a.lotno " & _
'                   "from  " & T_LAB023 & " a " & _
'                   "where " & DBW("a.ctrlcd  = ", pCtrlCd) & " " & _
'                   "and   " & DBW("a.levelcd = ", pLevelCd) & " " & _
'                   "and   " & DBW("a.opendt <= ", pFrDt) & " " & _
'                   "and   " & DBW("a.expdt  >= ", pToDt) & " " & _
'                   "order by a.opendt desc "
'End Function


Public Function GetLabNumbers(ByVal pCtrlCd As String, ByVal pLevelCd As String, _
                              ByVal pLotNo As String, ByVal pWorkArea As String, _
                              ByVal pFrDt As String, ByVal pToDt As String) As Recordset
                              
    Dim objSqlQc As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim LabDiv As String
    Dim lngLength As Integer
    
    
    SqlStmt = SqlCommonCode(T_LAB032, LC3_WorkArea, pWorkArea)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Not Rs.EOF Then
        LabDiv = "" & Rs.Fields("Field2").Value
    End If
    
    Set Rs = Nothing
    
    Select Case LabDiv
        Case "1": lngLength = 8   '일단위
        Case "2": lngLength = 6   '월단위
        Case "3": lngLength = 4   '년단위
        Case "4": lngLength = 6   '검체단위
    End Select
    
    Set objSqlQc = New clsLISSqlQc
    SqlStmt = objSqlQc.SqlLabNumbers(pCtrlCd, pLevelCd, pLotNo, pWorkArea, pFrDt, pToDt, lngLength)
    Set objSqlQc = Nothing
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    Set GetLabNumbers = Rs
    
End Function

'검사항목 컬렉션
Public Property Get ItemCount() As Integer
    ItemCount = mvarItemCount
End Property

Public Property Get Item(ByVal KeyIndex As Variant) As String
    Item = mItemCol(KeyIndex)
End Property

Private Sub Class_Initialize()
    ReDim mItemCol(0)
    mvarItemCount = 0
    Call InitRtn
    
    mvarTrans = False
    
    Set objBarCode = New clsBarcode
    
    Set objBarCode.TableInfo = New clsTables
    Set objBarCode.FieldInfo = New clsFields
End Sub

Public Sub ItemClear()
    Erase mItemCol
    ReDim mItemCol(0)
    mvarItemCount = 0
End Sub

Public Function addItem(ByVal clsData As String)
   
    mvarItemCount = mvarItemCount + 1
    ReDim Preserve mItemCol(mvarItemCount)
    mItemCol(mvarItemCount) = clsData
   
End Function

'% 채혈Procedure를 수행하여 채혈접수내역을 생성한다.
'% ( 복수검체의 경우 한꺼번에 모두 채혈처리 한다 --> 아래 DoCollection_Backup은 한번에 하나씩.. )
Public Function DoCollection(Optional ByRef ProgressBar As Variant) As Boolean
   
    Dim i As Long, j As Long
    Dim tmpStr As String
    Dim SqlIndex As Integer
    Dim NextKey As String
    Dim tmpSql As String
    Dim Success As Boolean
          
    ReDim SqlCreate026(0)

'    SaveKey = ""
    ColCount = 0
    DetailSeq = 0
   
    StsCd = enStsCd.StsCd_LIS_Accession   'Status
    ReqTotCnt = 0       'Required Item Total Count
    ReqInputCnt = 0     'Required Item Input Count
    FootNoteFg = "0"    'FootNote구분
    RptFg = "0"         '레포트 출력 여부
    VfyDt = ""          '결과확인일
    VfyTm = ""          '결과확인시간
    VfyId = 0           '결과확인자
    ColDt = Format(GetSystemDate, CS_DateDbFormat)    '채혈일
    ColTm = Format(GetSystemDate, CS_TimeDbFormat)    '채혈시간
    RcvDt = ColDt       '접수일
    RcvTm = ColTm       '접수시간
    RcvId = ColId       '접수자
    RmkCd = ""          '검체리마크
    TestNames = ""      '검사약어명String
    
    ColDate = Format(GetSystemDate, "YYYY-MM-DD HH:MM")
    StoreCd = ""                            '검체보관구분
    StatFg = ""                             '응급여부
    SpcGrp = ""                             '검체군
    
    
    If BuildCd = EmergencyLab Then
        BuildNm = "응급"
    ElseIf BuildCd = CentralLab Then
        BuildNm = "본원"
    ElseIf BuildCd = AneLab Then
        BuildNm = "안이"
    Else
        BuildNm = "" 'BuildingNm
    End If
    
    SpcNo = GetSpcNo(SpcYY)                  '검체번호
    AccSeq = GetLabNo(WorkArea, AccDt)       'AccSeq
    
    BarCount = 1  '바코드출력장수 초기화
    ColCount = 1
        
    For i = 1 To mvarMyQc.ItemCount
              
        If Not mvarMyQc.Item(i).Selected Then GoTo Skip
        
        TestCd = mvarMyQc.Item(i).TestCd
        RstVal = ""
        RstCd = ""
        RstDiv = mvarMyQc.Item(i).RstDiv
        RaDiv = ""
        RstUnit = mvarMyQc.Item(i).RstUnit
        RstType = mvarMyQc.Item(i).RstType
        MfyFg = ""
        TxtFg = ""
        AutoFg = ""
        
        TestNames = TestNames & mvarMyQc.Item(i).AbbrNm & ","        '** Barcode 라벨에 출력될 검사명들...
        
        Call ProcessOneItem(i)
           
Skip:
   Next
   
   SqlCreate201 = CreateSql_Collection  '접수내역
   
   Success = SaveData    '저장
   
   If Success Then
        Call PrintBarcode(ColCount, ColDate)    '바코드에 출력할 정보저장
        
        DoCollection = True
   Else
        DoCollection = False
   End If
   
   Exit Function
   
Err_Trap:
   DoCollection = False

End Function

Private Sub ProcessOneItem(ByVal Index As Integer)

   Dim tmpRs As Recordset
   Dim SqlStmt As String
   Dim SqlIndex As Integer
   Dim i As Integer
   Dim DetailCd As String
   
   
   DetailFg = ""
   
   With mvarMyQc.Item(Index)
        
        '// 상세항목인 경우... //
        
        If .PanelFg = PN_Detail Then    'PanelFg = 'D'
           
           RstDiv = "*"
           DetailFg = CStr(DetailSeq)
           
           SqlIndex = UBound(SqlCreate026) + 1
           ReDim Preserve SqlCreate026(SqlIndex)
           SqlCreate026(SqlIndex) = CreateSql_Result   '결과내역
           
           ReqTotCnt = ReqTotCnt + 1
           
           '상세항목검색
           SqlStmt = SqlPanelTest(Controlcd, LevelCd, Lotno, TestCd)
           Set tmpRs = New Recordset
           tmpRs.Open SqlStmt, DBConn
           
           While (Not tmpRs.EOF)
              
              With tmpRs
                 
                 TestCd = "" & .Fields("TestCd").Value              '상세항목코드
                 
                 RstDiv = "" & .Fields("RstDiv").Value              'Alternative/Required
                 RstType = "" & .Fields("RstType").Value            '결과유형
                 RstUnit = "" & .Fields("RstUnit").Value            '결과단위
'                 PanelFg = "" & .Fields("PanelFg").Value            '그룹처방여부
                 DetailFg = CStr(DetailSeq)                         '세부항목Seq
                 
                 SqlIndex = UBound(SqlCreate026) + 1
                 ReDim Preserve SqlCreate026(SqlIndex)
                 SqlCreate026(SqlIndex) = CreateSql_Result          '결과내역
                 
                 .MoveNext
              
              End With
           Wend
           Set tmpRs = Nothing
           DetailSeq = DetailSeq + 1
        
        '// 일반항목인 경우... //
        Else
           
           SqlIndex = UBound(SqlCreate026) + 1
           ReDim Preserve SqlCreate026(SqlIndex)
           SqlCreate026(SqlIndex) = CreateSql_Result   '결과내역
              
           ReqTotCnt = ReqTotCnt + 1
           
        End If
    
    End With
      
End Sub


'% 클래스의 채혈 내역을 DB로 저장한다.
Public Function SaveData() As Boolean
   
   Dim i As Integer
   
   On Error GoTo Err_Trap
   
   SaveData = True
   
   If mvarTrans = False Then DBConn.BeginTrans    '-------------------------<< Transaction Start >>

   '--채혈접수내역 생성
   Call DBConn.Execute(SqlCreate201)
   '--결과내역 생성
   For i = 1 To UBound(SqlCreate026)
        Call DBConn.Execute(SqlCreate026(i))
   Next
   
   If mvarTrans = False Then DBConn.CommitTrans   '-------------------------<< Transaction End >>
   
   CollectDone = True
   Exit Function

Err_Trap:
   SaveData = False
   If mvarTrans = False Then MsgBox Err.Description, vbExclamation
   CollectDone = False

End Function

'외부에서 트랙잭션을 컨트롤할 수도 있으므로 새로 생성했음
'Append by Legends 2003/10/09
Public Function SaveCollection(Optional ByVal pTrans As Boolean = True) As Boolean
   Dim i As Long
   
   On Error GoTo Err_Trap
   
   SaveCollection = True
   
   If pTrans Then DBConn.BeginTrans      '-------------------------<< Transaction Start >>

   '--채혈접수내역 생성
   Call DBConn.Execute(SqlCreate201)
   '--결과내역 생성
   For i = 1 To UBound(SqlCreate026)
        Call DBConn.Execute(SqlCreate026(i))
   Next
   
   If pTrans Then DBConn.CommitTrans     '-------------------------<< Transaction End >>
   
   CollectDone = True
   Exit Function

Err_Trap:
   SaveCollection = False
   If pTrans Then DBConn.RollbackTrans
   
   CollectDone = False
End Function

'% 검체번호를 부여한다.
Private Function GetSpcNo(ByVal SpcYY As String) As Long
   
   Dim SpcNoRs As Recordset
   
   Dim tmpSql As String
   
   GetSpcNo = 0
   
   '먼저 Lock을 걸기위해 의미없는 update문 실행.
   'select는 Lock이 안걸리기 때문에 극히 짧은 순간에도
   '검체번호가 중복될수 있음.. 을 방지.    2000.3.10  by 김미경
   tmpSql = CreateSql_SpcNo(SpcYY, 4)
   
   If mvarTrans = False Then DBConn.BeginTrans
   
   On Error GoTo Err_Trap
   
   DBConn.Execute tmpSql, , adCmdText   'Lock 걸림
   
   tmpSql = CreateSql_SpcNo(SpcYY, 1)
   Set SpcNoRs = New Recordset
   SpcNoRs.Open tmpSql, DBConn
   
   If SpcNoRs.EOF Then
      GetSpcNo = 1
      tmpSql = CreateSql_SpcNo(SpcYY, 2, GetSpcNo)
   Else
      GetSpcNo = Val("" & SpcNoRs.Fields("Seq").Value) + 1
      tmpSql = CreateSql_SpcNo(SpcYY, 3, GetSpcNo)
   End If
   Set SpcNoRs = Nothing
   
   DBConn.Execute tmpSql, , adCmdText
   If mvarTrans = False Then DBConn.CommitTrans
   
   Exit Function
   
Err_Trap:
   If mvarTrans = False Then MsgBox Err.Description, vbExclamation
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
   '// Sql문
   tmpSql = SqlCommonCode(T_LAB032, LC3_WorkArea, pWorkArea)
   '//
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   LabDiv = "" & tmpRs.Fields("Field2").Value
   Set tmpRs = Nothing
   
   tmpRng1 = 1
   tmpRng2 = 9999
   tmpSpcGrp = "0"
   Select Case LabDiv
   Case enLabDiv.LabDiv_ByDay:       '일단위
      pAccDt = Format(GetSystemDate, LIS_LabDayFormat)
   Case enLabDiv.LabDiv_ByMonth:       '월단위
      pAccDt = Format(GetSystemDate, LIS_LabMonthFormat)
   Case enLabDiv.LabDiv_ByYear:       '년단위
      pAccDt = Format(GetSystemDate, LIS_LabYearFormat)
   Case enLabDiv.LabDiv_BySpc:       '검체단위
      pAccDt = Format(GetSystemDate, LIS_LabMonthFormat)
      tmpSpcGrp = SpcGrp
      '// Sql문
      tmpSql = SqlCommonCode(T_LAB032, LC3_SGroup, tmpSpcGrp)
      '//
      Set tmpRs = Nothing
      Set tmpRs = New Recordset
      tmpRs.Open tmpSql, DBConn
      '검체단위인데 검체군 마스터가 없는 경우... (이런일은 발생하면 안되지...)
      If Not tmpRs.EOF Then
         tmpStr = "" & tmpRs.Fields("Field2").Value '검체군별 Lab No 범위
         tmpRng1 = Val(medGetP(tmpStr, 2, ";"))      '검체군별 Lab No 시작번호
         tmpRng2 = Val(medGetP(tmpStr, 3, ";"))     '검체군별 Lab No 끝번호
      End If
      Set tmpRs = Nothing
   Case Else:
      pAccDt = Format(GetSystemDate, LIS_LabDayFormat)
   End Select
   
   '먼저 Lock을 걸기위해 의미없는 update문 실행.
   'select는 Lock이 안걸리기 때문에 극히 짧은 순간에도
   '접수번호가 중복될수 있음.. 을 방지.    2000.3.10  by 김미경
   tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 4)
   
   On Error GoTo Err_Trap
   
   If mvarTrans = False Then DBConn.BeginTrans
   DBConn.Execute tmpSql, , adCmdText   'Lock 걸림
   
   '// Sql문장 생성
   tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 1)
   '//
   Set tmpRs = Nothing
   Set tmpRs = New Recordset
   tmpRs.Open tmpSql, DBConn
   If tmpRs.EOF Then
      GetLabNo = tmpRng1
      tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 2, GetLabNo)
   Else
      GetLabNo = Val("" & tmpRs.Fields("Seq").Value) + 1
      If GetLabNo > tmpRng2 Then
         MsgBox "접수번호가 Range(" & tmpRng1 & "-" & tmpRng2 & ")를 벗어났읍니다. : " & GetLabNo
         GetLabNo = 0
         Set tmpRs = Nothing
         Exit Function
      End If
      tmpSql = CreateSql_LabNo(pWorkArea, pAccDt, tmpSpcGrp, 3, GetLabNo)
   End If
   Set tmpRs = Nothing
      
   DBConn.Execute tmpSql, , adCmdText
   If mvarTrans = False Then DBConn.CommitTrans
   
   Exit Function
   
Err_Trap:
   If mvarTrans = False Then MsgBox Err.Description, vbExclamation
   Set tmpRs = Nothing
   Exit Function

End Function

Public Function PrintBarcode(ByVal ColCnt As Integer, ByVal pColTm As String) As String

    Dim tmpSql As String
   
    ReDim Preserve BarcodeBuffer(ColCnt)
     
    SpcNm = ControlNm
   
    With BarcodeBuffer(ColCnt)
        .Location = BuildNm
        .WorkArea = WorkArea
        .OrdDt = Mid(OrdDt, 5, 2) & "/" & Mid(OrdDt, 7, 2)
        .AccDt = Mid(AccDt, 3)
        .ColDt = Mid(AccDt, 3)
        .ColTm = Mid(pColTm, 12, 5)
        .AccSeq = Format(AccSeq, "@@@@@")
        .StatFg = ""
        .SpcNo = SpcYY & Format(SpcNo, CS_BarFormat)
        .PtNm = Controlcd   '환자명 대신 컨트롤명
        .PtId = ""
        .SpcNm = LevelCd    '검체명 대신 컨트롤레벨
        .StoreCd = ""
        .WardId = ""
        .TestNames = TestNames
        .LabelCnt = BarCount
    End With

End Function

Public Function PrintBarcodeLabel(ByVal pRowIndex As Long) As Boolean
'QC 자동처방 화면에서 사용되는 메소드 위의 함수와 기능은 같으나 의미없는 인수를 제거하였음.
'이진 호환성이 깨질경우 위의 함수는 제거해도 됨.
'Append By legends 2003/10/14

    PrintBarcodeLabel = True
    
    On Error GoTo ErrTrap   '바코드 출력메소드가 리턴값이 없어서 에러를 판단하기 힘들지만.. 혹시 몰라서..
    
    With BarcodeBuffer(pRowIndex)
'        Call ObjLISComCode.BarInfo.Label_PrintOut(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .PtId, .PtNm, .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt)
        Call objBarCode.Label_PrintOut(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .PtId, .PtNm, .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt)
    End With
    
    Exit Function
ErrTrap:
    PrintBarcodeLabel = False
End Function

Public Function GetBarcodeString(ByVal AryIndex As Integer, Optional ByVal AccFg As Boolean = False) As String
   
   If CollectDone Then
      With BarcodeBuffer(AryIndex)
         'Call MyBar.LabelPrint(.WorkArea, .ColDt, .AccSeq, .StatFg, .SpcNo, .PtId, .PtNm, .SpcNm, .StoreCd, .WardId, .Location, .TestNames, .LabelCnt)
'         GetBarcodeString = ObjLISComCode.BarInfo.Label_String(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .PtId, .PtNm, .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt)
         GetBarcodeString = objBarCode.Label_String(.Location, .WorkArea, .AccDt, .AccSeq, .SpcNo, .PtId, .PtNm, .SpcNm, .StoreCd, .StatFg, .WardId, .OrdDt, .ColTm, .TestNames, .LabelCnt)
      End With
   End If
   
End Function

Public Function BarFormFeed()
    Call objBarCode.Label_FormFeed
'    Call ObjLISComCode.BarInfo.Label_FormFeed
End Function

'% Method 1 : CollectQuery
'%                 Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                 클래스 clsCollection의 Data Attribute에 저장한다.

Public Function CollectQuery(ByVal pWorkArea As String, _
                                         ByVal pAccDt As String, ByVal pAccSeq As Integer) As Boolean

   Dim objSqlQc As clsLISSqlQc
   Dim SqlStmt As String
   Dim MyOraDS As Recordset
   
   Set objSqlQc = New clsLISSqlQc
   SqlStmt = objSqlQc.SqlCollectQuery(pWorkArea, pAccDt, pAccSeq)
   Set objSqlQc = Nothing
   
   Set MyOraDS = New Recordset 'Sql 실행
   MyOraDS.Open SqlStmt, DBConn
   
   If MyOraDS.EOF Then
      CollectQuery = False
      Exit Function
   End If
   
   With MyOraDS
      
      .MoveFirst
         
      SpcYY = Trim("" & .Fields("SpcYy").Value)
      SpcNo = Val("" & .Fields("SpcNo").Value)
      PtId = Trim("" & .Fields("PtId").Value)
      Sex = Trim("" & .Fields("Sex").Value)
      AgeDay = Val("" & .Fields("AgeDay").Value)
      BedInDt = Trim("" & .Fields("BedInDt").Value)
      DeptCd = Trim("" & .Fields("DeptCd").Value)
      OrdDoct = Val("" & .Fields("OrdDoct").Value)
      MajDoct = Val("" & .Fields("MajDoct").Value)
      WorkArea = Trim("" & .Fields("WorkArea").Value)
      AccDt = Trim("" & .Fields("AccDt").Value)
      AccSeq = Val("" & .Fields("AccSeq").Value)
      StsCd = Trim("" & .Fields("StsCd").Value)
      ReqTotCnt = Val("" & .Fields("ReqTotCnt").Value)
      ReqInputCnt = Val("" & .Fields("ReqInputCnt").Value)
      VfyDt = Trim("" & .Fields("VfyDt").Value)
      VfyTm = Trim("" & .Fields("VfyTm").Value)
      VfyId = Val("" & .Fields("VfyId").Value)
      ColDt = Trim("" & .Fields("ColDt").Value)
      ColTm = Trim("" & .Fields("ColTm").Value)
      ColId = Val("" & .Fields("ColId").Value)
      RcvDt = Trim("" & .Fields("RcvDt").Value)
      RcvTm = Trim("" & .Fields("RcvTm").Value)
      RcvId = Val("" & .Fields("RcvId").Value)
      EntDt = Trim("" & .Fields("EntDt").Value)
      EntTm = Trim("" & .Fields("EntTm").Value)
      EntId = Val("" & .Fields("EntId").Value)
      SpcCd = Trim("" & .Fields("SpcCd").Value)
      MultiFg = Trim("" & .Fields("MultiFg").Value)
      OrgAccNo = Trim("" & .Fields("OrgAccNo").Value)
      WardId = Trim("" & .Fields("WardId").Value)
      RoomId = Trim("" & .Fields("RoomId").Value)
      BedId = Trim("" & .Fields("BedId").Value)
      HosilId = Trim("" & .Fields("HosilId").Value)
      FootNoteFg = Trim("" & .Fields("FootNoteFg").Value)
      StoreCd = Trim("" & .Fields("StoreCd").Value)
      RptFg = Trim("" & .Fields("RptFg").Value)
      TestDiv = Trim("" & .Fields("TestDiv").Value)
      QcFg = Trim("" & .Fields("QcFg").Value)
      RmkCd = Trim("" & .Fields("RmkCd").Value)
      StatFg = Trim("" & .Fields("StatFg").Value)
      BuildCd = Trim("" & .Fields("BuildCd").Value)
      
      OrdDt = Trim("" & .Fields("EntDt").Value)

      
   End With
   CollectDone = True
   Set MyOraDS = Nothing
End Function

Private Function CreateSql_Collection() As String

   CreateSql_Collection = "Insert into " & T_LAB201 & _
                               "(spcyy, spcno, ptid, sex, ageday, bedindt, deptcd, orddoct, majdoct, " & _
                               "workarea, accdt, accseq, stscd, reqtotcnt, reqinputcnt, vfydt, vfytm, vfyid, " & _
                               "coldt, coltm, colid, rcvdt, rcvtm, rcvid, entdt, enttm, entid, spccd, multifg, " & _
                               "orgaccno, wardid, roomid, bedid, footnotefg, storecd, rptfg, testdiv, qcfg, rmkcd, " & _
                               "statfg, buildcd, orgbuildcd, hosilid) "
   CreateSql_Collection = CreateSql_Collection & _
                        "Values  (" & DBV("spcyy", SpcYY, 1) & DBV("spcno", SpcNo, 1) & DBV("ptid", PtId, 1) _
                                    & DBV("sex", Sex, 1) & DBV("ageday", AgeDay, 1) _
                                    & DBV("bedindt", BedInDt, 1) & DBV("deptcd", DeptCd, 1) _
                                    & DBV("orddoct", OrdDoct, 1) & DBV("majdoct", MajDoct, 1) _
                                    & DBV("workarea", WorkArea, 1) & DBV("accdt", AccDt, 1) _
                                    & DBV("accseq", AccSeq, 1) & DBV("stscd", StsCd, 1) _
                                    & DBV("reqtotcnt", ReqTotCnt, 1) & DBV("reqinputcnt", ReqInputCnt, 1) _
                                    & DBV("vfydt", VfyDt, 1) & DBV("vfytm", VfyTm, 1) _
                                    & DBV("vfyid", VfyId, 1) & DBV("coldt", ColDt, 1) _
                                    & DBV("coltm", ColTm, 1) & DBV("colid", ColId, 1) _
                                    & DBV("rcvdt", RcvDt, 1) _
                                    & DBV("rcvtm", RcvTm, 1) & DBV("rcvid", RcvId, 1) _
                                    & DBV("entdt", Format(GetSystemDate, "yyyyMMdd"), 1) & DBV("enttm", Format(GetSystemDate, "hhMMss"), 1) _
                                    & DBV("entid", EntId, 1) & DBV("spccd", SpcCd, 1) _
                                    & DBV("multifg", MultiFg, 1) & DBV("orgaccno", OrgAccNo, 1) _
                                    & DBV("wardid", WardId, 1) _
                                    & DBV("roomid", RoomId, 1) & DBV("bedid", BedId, 1) _
                                    & DBV("footnotefg", FootNoteFg, 1) & DBV("storecd", StoreCd, 1) _
                                    & DBV("rptfg", RptFg, 1) & DBV("testdiv", TestDiv, 1) _
                                    & DBV("qcfg", QcFg, 1) & DBV("rmkcd", RmkCd, 1) & DBV("statfg", StatFg, 1) _
                                    & DBV("buildcd", BuildCd, 1) & DBV("orgbuildcd", OrgBuildCd, 1) _
                                    & DBV("hosilid", HosilId) & ") "

End Function


Private Function CreateSql_Result() As String

    CreateSql_Result = "insert into " & T_LAB026 & _
                                    "(workarea,accdt,accseq,testcd,rstval," & _
                                     "rstcd,rstunit,rsttype,rstdiv,radiv,detailfg,vfydt,vfytm,vfyid,mfyfg," & _
                                     "txtfg,autofg,eqpcd,method,ctrlcd,levelcd,lotno) " & _
                        "values (" & _
                                  DBV("workarea", WorkArea, 1) & DBV("accdt", AccDt, 1) & _
                                  DBV("accseq", AccSeq, 1) & DBV("testcd", TestCd, 1) & _
                                  DBV("rstval", RstVal, 1) & DBV("rstcd", RstCd, 1) & _
                                  DBV("rstunit", RstUnit, 1) & DBV("rsttype", RstType, 1) & _
                                  DBV("rstdiv", RstDiv, 1) & DBV("radiv", RaDiv, 1) & _
                                  DBV("detailfg", DetailFg, 1) & DBV("vfydt", VfyDt, 1) & _
                                  DBV("vfytm", VfyTm, 1) & DBV("vfyid", VfyId, 1) & _
                                  DBV("mfyfg", MfyFg, 1) & DBV("txtfg", TxtFg, 1) & _
                                  DBV("autofg", AutoFg, 1) & DBV("eqpcd", EqpCd, 1) & _
                                  DBV("method", Method, 1) & DBV("ctrlcd", Controlcd, 1) & _
                                  DBV("levelcd", LevelCd, 1) & DBV("lotno", Lotno) & ")"
    
End Function

Private Function CreateSql_UpdateLabNo(ByVal PtId As String, ByVal OrdDt As String, ByVal OrdNo As Integer, _
                                                           ByVal OrdSeq As String, ByVal StsCd As String, Optional ByVal DoneFg As Variant) As String
   Dim tmpStr As String
   
   If IsMissing(DoneFg) Then
      tmpStr = ""
   Else
      tmpStr = ", " & DBW("donefg ", DoneFg, 2) & " "
   End If
   CreateSql_UpdateLabNo = "Update " & T_LAB102 & " " & _
                           "Set " & DBW("workarea = ", WorkArea) & ", " & _
                                    DBW("accdt    = ", AccDt) & ", " & _
                                    DBW("accseq   = ", AccSeq) & ", " & _
                                    DBW("stscd    = ", StsCd) & " " & tmpStr & _
                           "Where " & DBW("ptid   = ", PtId) & " " & _
                           "and   " & DBW("orddt  = ", OrdDt) & " " & _
                           "and   " & DBW("ordno  = ", OrdNo) & " " & _
                           "and   " & DBW("ordseq = ", OrdSeq)

End Function


Private Function CreateSql_SpcNo(ByVal SpcYY As String, ByVal flag As Integer, Optional ByVal Seq As Variant) As String

    Select Case flag
        Case 1:
           CreateSql_SpcNo = "Select seq " & _
                             "From  " & T_COM099 & " " & _
                             "Where " & DBW("noindex = ", COM99_LIS_Specimen) & " " & _
                             "and   " & DBW("divcd1  = ", LIS_ORDDIV) & " " & _
                             "and   " & DBW("divcd2  = ", SpcYY) & " " & _
                             "and   " & DBW("divcd3  = ", "0")
        Case 2:
           CreateSql_SpcNo = "Insert Into " & T_COM099 & " (noindex, divcd1, divcd2, divcd3, seq) " & _
                             "Values (" & _
                                        DBV("noindex", COM99_LIS_Specimen, 1) & _
                                        DBV("divcd1", LIS_ORDDIV, 1) & _
                                        DBV("divcd2", SpcYY, 1) & _
                                        DBV("divcd3", "0", 1) & _
                                        DBV("seq", Seq) & ")"
        Case 3:
           CreateSql_SpcNo = "Update " & T_COM099 & " " & _
                             "Set   " & DBW("seq     = ", Seq) & " " & _
                             "Where " & DBW("noindex = ", COM99_LIS_Specimen) & " " & _
                             "and   " & DBW("divcd1  = ", LIS_ORDDIV) & " " & _
                             "and   " & DBW("divcd2  = ", SpcYY) & " " & _
                             "and   " & DBW("divcd3  = ", "0")
        Case 4:
           CreateSql_SpcNo = "Update " & T_COM099 & " " & _
                             "Set seq = seq " & _
                             "Where " & DBW("noindex = ", COM99_LIS_Specimen) & " " & _
                             "and   " & DBW("divcd1  = ", LIS_ORDDIV) & " " & _
                             "and   " & DBW("divcd2  = ", SpcYY) & " " & _
                             "and   " & DBW("divcd3  = ", "0")
    End Select
   
'   "SPCNO" divcd1
End Function

Private Function CreateSql_LabNo(ByVal WorkArea As String, ByVal AccDt As String, ByVal SpcGrp As String, _
                                                 ByVal flag As Integer, Optional ByVal Seq As Variant) As String
    Select Case flag
        Case 1:
            CreateSql_LabNo = " SELECT seq " & _
                              " FROM   " & T_COM099 & " " & _
                              " WHERE " & DBW("noindex=", COM99_LIS_LabNo) & _
                              " AND " & DBW("divcd1=", WorkArea) & _
                              " AND " & DBW("divcd2=", AccDt) & _
                              " AND " & DBW("divcd3=", SpcGrp)
        Case 2:
            CreateSql_LabNo = " INSERT INTO " & T_COM099 & " (noindex, divcd1, divcd2, divcd3, seq)" & _
                              " VALUES (" & DBV("noindex", COM99_LIS_LabNo, 1) & DBV("divcd1", WorkArea, 1) _
                                    & DBV("divcd2", AccDt, 1) & DBV("divcd3", SpcGrp, 1) & DBV("seq", Seq) & ")"
        Case 3:
            CreateSql_LabNo = " UPDATE " & T_COM099 & _
                              " SET    " & DBW("seq =", Seq) & _
                              " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                              " AND    " & DBW("divcd1 =", WorkArea) & _
                              " AND    " & DBW("divcd2 =", AccDt) & _
                              " AND    " & DBW("divcd3 =", SpcGrp)
        Case 4:
            CreateSql_LabNo = " UPDATE " & T_COM099 & " Set seq = seq " & _
                              " WHERE  " & DBW("noindex =", COM99_LIS_LabNo) & _
                              " AND    " & DBW("divcd1 =", WorkArea) & _
                              " AND    " & DBW("divcd2 =", AccDt) & _
                              " AND    " & DBW("divcd3 =", SpcGrp)
    End Select
End Function

Public Function SqlPanelTest(ByVal Controlcd As String, ByVal LevelCd As String, ByVal Lotno As String, _
                             ByVal TestCd As String) As String
   Dim tmpStr As String
   
   SqlPanelTest = "Select a.cdval2 as testcd, b.rstdiv, b.rsttype, b.panelfg, c.rstunit " & _
                "From  " & T_LAB031 & " a, " & T_LAB001 & " b, " & T_LAB024 & " c " & _
                "where " & DBW("a.cdindex = ", LC2_Detail) & " " & _
                "and   " & DBW("a.cdval1 = ", TestCd) & " " & _
                "and     b.testcd = a.cdval2 " & _
                "and     b.applydt = (select max(applydt) from " & T_LAB001 & " where testcd = b.testcd) " & _
                "and   " & DBW("c.ctrlcd = ", Controlcd) & " " & _
                "and   " & DBW("c.levelcd = ", LevelCd) & " " & _
                "and   " & DBW("c.lotno = ", Lotno) & " " & _
                "and     c.testcd = a.cdval2 "
                         
End Function

Public Function SqlCommonCode(ByVal TblName As String, ByVal CdIndex As String, _
                                                Optional ByVal CdVal1 As Variant, Optional ByVal CdVal2 As Variant) As String
   Dim tmpStr As String
   
   tmpStr = ""
   If Not IsMissing(CdVal1) Then tmpStr = tmpStr & "and " & DBW("cdval1 = ", CdVal1) & " "
   If Not IsMissing(CdVal2) Then tmpStr = tmpStr & "and " & DBW("cdval2 = ", CdVal2) & " "
   
   SqlCommonCode = "Select * " & _
                   "From  " & TblName & " " & _
                   "Where " & DBW("cdindex = ", CdIndex) & " " & tmpStr
End Function

Private Sub Class_Terminate()
'   Set MySql = Nothing
'   Set SortList = Nothing
    Set objBarCode = Nothing
End Sub

Public Sub InitRtn()
   'SortList.Clear
'   KeyCount = 0
   ColCount = 0
'   ReadyToCollect = False
   CollectDone = False
   'Erase SqlBuffer
End Sub

Public Sub LoadCtrlForOrder(ByRef lstList As Object, ByVal pBuildCd As String)

    Dim SqlStmt As String
    Dim Rs As Recordset
    
    SqlStmt = "select distinct a.ctrlcd, a.levelcd, a.lotno, b.ctrlnm " & _
              "from " & T_LAB023 & " a, " & T_LAB021 & " b " & _
              "where " & DBW("a.opendt <= ", Format(GetSystemDate, CS_DateDbFormat)) & _
              "and   " & DBW("a.expdt >= ", Format(GetSystemDate, CS_DateDbFormat)) & _
              "and   b.ctrlcd = a.ctrlcd " & _
              "and   b.levelcd = a.levelcd " & _
              "and   " & DBW("b.buildcd = ", pBuildCd)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    While (Not Rs.EOF)
        lstList.addItem "" & Rs.Fields("CtrlCd").Value & vbTab & _
                        "" & Rs.Fields("CtrlNm").Value & vbTab & _
                        "" & Rs.Fields("LevelCd").Value & vbTab & _
                        "" & Rs.Fields("LotNo").Value
        Rs.MoveNext
    Wend
    
Nodata:
    Set Rs = Nothing
    
End Sub

Public Sub LoadTestItems(ByRef tblItemList As Object)

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim i As Integer


'쿼리 수정
'2002/09/26 Modify By Legends
'Detailfg ="*"이고 Panelfg <>"G"인데이터를 불러와야 됨
'단일항목이거나 상세 새끼코드만 로드

    SqlStmt = " select a.testcd, a.testnm, b.avalval, b.rstunit " & _
              " from " & _
                       T_LAB001 & " a, " & T_LAB004 & " b " & _
              " where exists (select b.ctrlcd" & _
                            " from " & _
                                     T_LAB024 & " b, " & T_LAB023 & " c " & _
                            " where " & _
                                    "     b.testcd = a.testcd" & _
                                    " and c.ctrlcd = b.ctrlcd" & _
                                    " and c.levelcd = b.levelcd" & _
                                    " and c.lotno = b.lotno" & _
                                    " and (" & _
                                             DBW("c.expdt =", "") & " or " & _
                                             DBW("c.expdt >", Format(GetSystemDate, CS_DateDbFormat)) & _
                                         ")" & _
                                    " and " & DBW("c.opendt <=", Format(GetSystemDate, CS_DateDbFormat)) & _
                            ")" & _
                     " and (a.expdt ='' or a.expdt is null)" & _
                     " and (a.detailfg ='*' or (a.detailfg is null or a.detailfg = '' ) )" & _
                     " and ( " & DBW("a.panelfg <>", PN_Group) & " or a.panelfg is null ) " & _
                     " and b.testcd = a.testcd " & _
                     " and b.seq = (select max(seq) from " & T_LAB004 & _
                                  " where testcd = b.testcd" & _
                                     " and (expdt ='' or expdt is null) " & ")" & _
              " order by testnm"
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    With tblItemList
        .MaxRows = Rs.RecordCount
        For i = 1 To Rs.RecordCount
            .Row = i
            .Col = 2: .Value = "" & Rs.Fields("TestNm").Value
            .Col = 5: .Value = "" & Rs.Fields("TestCd").Value
            .Col = 10: .Value = "" & Rs.Fields("AvalVal").Value
            .Col = 11: .Value = "" & Rs.Fields("RstUnit").Value
            Rs.MoveNext
        Next
    End With
    Set Rs = Nothing

End Sub


Public Sub LoadItems(ByRef lstList As Object, ByRef lstList1 As Object, _
                     Optional ByVal pWorkArea As String = "", Optional ByVal pSpcFg As Boolean = False)

    Dim Rs As Recordset
    Dim objItem As clsLISSqlQc
    Dim tmpStr As String
    Dim strSQL As String
    Dim i As Long
    
    Set objItem = New clsLISSqlQc
    
    With objItem
        Set Rs = New Recordset
        Rs.Open .GetTestItem(pWorkArea, pSpcFg), DBConn
    End With
    
    lstList.Clear
    lstList1.Clear
    If Rs.EOF Then GoTo Nodata
    
    For i = 1 To Rs.RecordCount
        tmpStr = "" & Rs.Fields("testcd").Value & Space(9)
        tmpStr = Mid(tmpStr, 1, 10) & "" & Rs.Fields("testnm").Value & Space(40)
        If Not pSpcFg Then
            lstList.addItem tmpStr  ', 1, 50) & _
                            rs.Fields("TestNm").Value
        Else
            lstList.addItem Mid(tmpStr, 1, 50) & vbTab & _
                            "" & Rs.Fields("rsttype").Value & vbTab & _
                            "" & Rs.Fields("rstdiv").Value & vbTab & _
                            "" & Rs.Fields("rstunit").Value & vbTab & _
                            "" & Rs.Fields("avalval").Value
        End If
        lstList1.addItem "" & Rs.Fields("TestNm").Value & vbTab & _
                         "" & Rs.Fields("TestCd").Value  'Index List
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    Set objItem = Nothing
End Sub

Public Sub LoadPartCd(ByRef lstList As Object)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim i%

    SqlStmt = " select cdval1, field1" & _
              " from " & T_LAB032 & _
              " where " & DBW("cdindex =", LC3_PartCd)
              
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn

    If Rs.EOF Then GoTo Nodata

    lstList.Clear
    For i = 1 To Rs.RecordCount
        lstList.addItem "" & Rs.Fields("cdval1").Value & vbTab & _
                        "" & Rs.Fields("field1").Value
        Rs.MoveNext
    Next i

Nodata:
    Set Rs = Nothing
End Sub

Public Sub LoadHospCd(ByRef lstList As Object)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim i%

    SqlStmt = " select cdval1, field1, field2, field3" & _
              " from " & T_LAB032 & _
              " where " & DBW("cdindex =", LC3_HospCd)
              
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn

    If Rs.EOF Then GoTo Nodata

    lstList.Clear
    For i = 1 To Rs.RecordCount
        lstList.addItem "" & Rs.Fields("cdval1").Value & vbTab & _
                        "" & Rs.Fields("field1").Value & vbTab & _
                        "" & Rs.Fields("field2").Value & vbTab & _
                        "" & Rs.Fields("field3").Value
        Rs.MoveNext
    Next i

Nodata:
    Set Rs = Nothing
End Sub

Public Function GetControlComboString(ByVal pTestCd As String) As String

    Dim objQcSql As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    Set objQcSql = New clsLISSqlQc
    
    SqlStmt = objQcSql.SqlGetControlList(pTestCd)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    GetControlComboString = ""
    While (Not Rs.EOF)
        GetControlComboString = GetControlComboString & _
                                "" & Rs.Fields("CtrlCd").Value & "-" & _
                                "" & Rs.Fields("CtrlNm").Value & "-" & _
                                "" & Rs.Fields("WorkArea").Value & vbTab
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    Set objQcSql = Nothing
End Function


Public Function GetLotnoComboString(ByVal pCtrlCd As String) As String

    Dim objQcSql As clsLISSqlQc
    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim strCboString As String
               
    Set objQcSql = New clsLISSqlQc
    
    SqlStmt = objQcSql.SqlGetLotNoList(pCtrlCd)
                      
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    strCboString = ""
    While (Not Rs.EOF)
        If "" & Rs.Fields("LevelCd").Value = QC_LEVEL_HIGH Then
            strCboString = strCboString & "" & Rs.Fields("LevelCd").Value & "-High-"
        
        ElseIf "" & Rs.Fields("LevelCd").Value = QC_LEVEL_NORMAL Then
            strCboString = strCboString & "" & Rs.Fields("LevelCd").Value & "-Normal-"
        
        Else
            strCboString = strCboString & "" & Rs.Fields("LevelCd").Value & "-Low-"
        End If
        
        strCboString = strCboString & "" & Rs.Fields("LotNo").Value & "-" & _
                        "" & Rs.Fields("OpenDt").Value & "-" & _
                        "" & Rs.Fields("ExpDt").Value & "-" & _
                        "" & Rs.Fields("EqpCd").Value & "-" & _
                        "" & Rs.Fields("EqpNm").Value & vbTab
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    Set objQcSql = Nothing
    
    GetLotnoComboString = strCboString
    
End Function
            
 
Public Function GetWorkAreaForOutQc() As String

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim objQcSql As clsLISSqlQc
    
    Set objQcSql = New clsLISSqlQc
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlGetOutQcWorkArea, DBConn

    GetWorkAreaForOutQc = ""

    If Rs.EOF Then
        MsgBox "외부정도관리 WorkArea가 정의되지 않았습니다.", vbExclamation, "메세지"
    Else
        GetWorkAreaForOutQc = "" & Rs.Fields("CdVal1").Value & "-" & "" & Rs.Fields("Field1").Value
    End If

    Set Rs = Nothing
    Set objQcSql = Nothing
    
End Function


Public Function GetLastSeq(ByVal pWorkArea As Integer, ByVal pOutDiv As Integer) As Integer

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim objQcSql As clsLISSqlQc
    
    Set objQcSql = New clsLISSqlQc
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlGetOutQcLastSeq(pWorkArea, pOutDiv), DBConn
    
    If Rs.EOF Then
        GetLastSeq = 0
    Else
        GetLastSeq = Val("" & Rs.Fields("Seq").Value)
    End If
    Set Rs = Nothing
    Set objQcSql = Nothing

End Function

'호환안깨드리려구 만들었음.
Public Function GetLastSeq_New(ByVal pWorkArea As String, ByVal pOutDiv As Integer) As Integer

    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim objQcSql As clsLISSqlQc
    
    Set objQcSql = New clsLISSqlQc
    Set Rs = New Recordset
    Rs.Open objQcSql.SqlGetOutQcLastSeq(pWorkArea, pOutDiv), DBConn
    If Rs.EOF Then
        GetLastSeq_New = 0
    Else
        GetLastSeq_New = Val("" & Rs.Fields("Seq").Value)
    End If
    Set Rs = Nothing
    Set objQcSql = Nothing

End Function

