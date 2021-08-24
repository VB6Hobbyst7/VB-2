VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISAccItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd As String        '/* 검사항목 코드   */
Public RstVal As String        '/* 결과(Numeric)   */
Public RstCd As String         '/* 결과코드(Alpha) */
Public RstUnit As String       '/* Unit            */
Public HLDiv As String         '/* High/Low(H:Hig,L:Low) */
Public DPDiv As String         '/* Delta/Panic(D:Delta,P:Panic) */
Public SpcCd As String         '/* 검체코드        */
Public StatFg As String        '/* 응급여부('0':부,'1':여) */
Public LastRst As String       '/* 최근결과        */
Public LastVfyDt As String     '/* 최근결과확인일  */
Public LastVfyTm As String     '/* 최근결과확인시간 */
Public LastVfyId As Integer    '/* 최근결과확인자  */
Public LastVfyId1 As String    '/* 최근결과확인자  */
Public VfyDt As String         '/* 결과확인일자    */
Public VfyTm As String         '/* 결과확인시간    */
Public VfyId As Integer        '/* 결과확인자      */
Public VfyId1 As String        '/* 결과확인자      */
Public AttrCd As String        '/* 속성코드   */
Public MfyFg As String         '/* 수정여부('0':부,'1':여) */
Public GrpFg As String         '/* 그래프 결과여부('0':부,'1':여) */
Public ValFg As String         '/* 수치 결과여부('0':부,'1':여)  - 기타검사에서만 사용*/
Public TxtFg As String         '/* TEXT 결과여부('0':부,'1':여) */
Public RstType As String       '/* 결과유형(N,Alpha,비율,Free) */
Public RstDiv As String        '/Alternative,Required 여부 */
Public DetailFg As String      '/* 상세항목Group Seq */
Public Ptid As String          '/* 환자 ID */
Public OrdDt As String         '/* 처방일 */
Public OrdNo As Integer        '/* 처방번호 */
Public OrdSeq As Integer       '/* 처방Seq */
Public SenFg As String         '/* 감수성결과 여부 */
Public AutoFg As String
Public EqpCd As String
    
Public TestDiv As String
