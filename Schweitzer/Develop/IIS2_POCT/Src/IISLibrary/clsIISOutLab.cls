VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISOutLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WorkArea As String     '/* Work Area(접수번호용) */
Public AccDt As String        '/* 접수일(접수번호용)    */
Public AccSeq As Integer      '/* 접수순번(접수번호용)  */
Public TestCd As String       '/* 검사항목        */
Public PtId As String         '/* 환자ID  */
Public SpcCd As String        '/* 검체코드  */
Public OutLabCd As String     '/* 기관코드        */
Public StsCd As String        '/* STATUS          */
Public RcvDt As String        '/* 접수일자        */
Public SendDt As String       '/* SEND Date       */
Public ChargeDt As String     '/* 청구일자        */
Public SendId As String       '/* 전송자 ID   */


