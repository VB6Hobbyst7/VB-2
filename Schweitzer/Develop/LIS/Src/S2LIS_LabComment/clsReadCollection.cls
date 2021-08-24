VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollectioin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Class 명 : clsCollection(DLL용)
'|  2. 기    능 : 채혈/접수내역(LAB201)을 생성한다.
'|  3. 작 성 자 : 김미경
'|
'|  CopyRight(C) 1999 대련엠티에스
'|
'+--------------------------------------------------------------------------------------+


'%  클래스 clsCollection의 Data Attributes
Public SpcYY As String      '/* 검체번호-년도   */
Public SpcNo As Long        '/* 검체번호        */
Public PtId As String       '/* 환자 ID         */
Public PtNm As String       '/* 환자명         */
Public AgeDay As Long       '/* 환자 일령       */
Public Sex As String        '/* 환자 성별   */
Public BedinDt As String    '/* 입원일          */
Public DeptCd As String     '/* 진료과          */
Public OrdDoct As Integer   '/* 처방의          */
Public MajDoct As Integer   '/* 주치의          */
Public WorkArea As String   '/* Work Area(접수번호용) */
Public AccDt As String      '/* 접수일(접수번호용)    */
Public AccSeq As Integer    '/* 접수순번(접수번호용)  */
Public StsCd As String      '/* STATUS - '0 ':처방,1:채혈,2:접수,3:In-Process,4:결과,5:확인,6:수정 */
Public ReqTotCnt As Integer '/* Required Item Total Count */
Public ReqInputCnt As Integer  '/* Required Item Input Count */
Public VfyDt As String      '/* 결과확인일자    */
Public VfyTm As String      '/* 결과확인시간    */
Public VfyId As Integer     '/* 결과확인자      */
Public ColDt As String      '/* 채혈일자        */
Public ColTm As String      '/* 채혈시간        */
Public ColId As Integer     '/* 채혈자          */
Public RcvDt As String      '/* 접수일자        */
Public RcvTm As String      '/* 접수시간        */
Public RcvId As Integer     '/* 접수자          */
Public EntDt As String      '/* 입력일자        */
Public EntTm As String      '/* 입력시간        */
Public EntId As Integer     '/* 입력자          */
Public SpcCd As String      '/* 검체코드        */
Public SpcNm As String      '/* 검체명        */
Public MultiFg As String    '/* 복수검체여부  */
Public OrgAccNo As String   '/* 원 접수번호     */
Public WardId As String     '/* 병동코드        */
Public RoomId As String     '/* 병실코드     */
Public BedId As String      '/* 침상코드    */
Public HosilId As String    '/* 호실코드    */
Public FootNoteFg As String '/* FOOTNOTE여부('0':부,'1':여) */
Public StoreCd As String    '/* 검체보관 코드   */
Public RptFg As String      '/* REPORT출력여부('0':부,'1':여) */
Public TestDiv As String    '/* 검사구분 - 1:기타검사,2:미생물검사 */
Public QcFg As String       '/* QC여부('0':부,'1':여) */
Public RmkCd As String      '/* 검체 Remark 코드 */
Public StatFg As String     '/* 응급여부    */
Public BuildCd As String    '/* 건물코드    */
Public BuildNm As String    '/* 건물코드명    */
Public OrgBuildCd As String '/* 채혈된 건물코드    */
Public TestNames As String  '/* Barcode에 출력될 검사명들 */

Public OrdDt As String
Public OrdNo As Integer
Public OrdSeq As Integer
Public OrdCd As String

Public SpcGrp As String     '/* 검체군  */
Public SpcSeq As String     '/* 복수검체의 Seq */

'= Lab 채혈접수 관련 Data Type ='
Private Type TpLabGroup
   BuildCd As String
   TestDiv As String
   WorkArea As String
   SpcCd As String
   StoreCd As String
   StatFg As String
   ColDate As String
   MultiFg As String
   SpcGrp As String
   OrdDt As String
   OrdNo As Integer
   OrdSeq As Integer
   OrdCd As String
   DeptCd As String     '/* 진료과          */
   OrdDoct As Integer   '/* 처방의          */
   MajDoct As Integer   '/* 주치의          */
   LabelCnt As Integer
   TestNm As String     '/* 검사 약어명   */
End Type

Private Type LabNumbers
   WorkArea As String
   AccDt As String
   AccSeq As Integer
   BuildCd As String
End Type
   
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
   
   
Public ErrMsg As String
Public ColCount As Integer

'공용변수 - 참조하는 프로젝트로부터 설정됨
'Public MyOraSE As Object     'OraSession
Public MyDb As Object     'OraDatabase

Public CollectDone As Boolean

Private blnBatchCol As Boolean
Private ColDone As Boolean  '복수검체 채혈완료 여부
Private ReadyToCollect As Boolean
Private KeyCount As Integer
Private MultiSpcFlag As String
Private BarCount As Integer
Private SortList As ListBox
Private LabGroup() As TpLabGroup
Private CreateLabNo() As LabNumbers
Private BarcodeBuffer() As BarcodeLabel
Private SqlBuffer() As String
Private MySql As New clsLISSqlStatement
'Private MyBar As New clsBarcode

Private sWorkDt As String
Private sWorkTm As String
Private iWorkSeq As Integer



'% Method 1 : CollectQuery
'%                 Parameter로 받은 Sql을 실행하고, 각 필드의 값을
'%                 클래스 clsCollection의 Data Attribute에 저장한다.

Public Function CollectQuery(ByVal pWorkArea As String, _
                                         ByVal pAccDt As String, ByVal pAccSeq As Integer) As Boolean

   Dim SqlStmt As String
   Dim MyOraDS As Object
   
   SqlStmt = "Select a.* " & _
                 "From " & T_LAB201 & " a " & _
                 "Where a.workarea = '" & pWorkArea & "' " & _
                 "and     a.accdt = '" & pAccDt & "' " & _
                 "and     a.accseq = " & pAccSeq
   
   Set MyOraDS = OpenRecordSet(SqlStmt)  'Sql 실행
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
      BedinDt = Trim("" & .Fields("BedInDt").Value)
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
      
   End With
   
End Function


Private Sub Class_Initialize()
   blnBatchCol = False
   Set MySql = New clsLISSqlStatement
   Set SortList = frmControls.lstList
   Call InitRtn
End Sub

Private Sub Class_Terminate()
   Set MySql = Nothing
   Set SortList = Nothing
End Sub

Public Sub InitRtn()
   SortList.Clear
   KeyCount = 0
   ColCount = 0
   ReadyToCollect = False
   CollectDone = False
   ColDone = False
   Erase SqlBuffer
   ErrMsg = ""
End Sub










