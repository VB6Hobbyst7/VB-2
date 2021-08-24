VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLab001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd As String         ' 검사항목 코드
Public ApplyDt As String        ' 적용일
Public TestNm As String         ' 검사명
Public AbbrNm5 As String        ' 약어명(5자리)
Public AbbrNm10 As String       ' 약어명(10자리)
Public WorkArea As String       ' Work Area
Public RstType As String        ' 결과유형
Public TestDiv As String        ' 기타검사,미생물검사 구분
Public RptSeq As Integer        ' Report Seq
Public PanelFg As String        ' Pannel(Group,Detail) 처방여부('0':부,'1':여)
Public DetailFg As String       ' 상세항목여부('*')
Public RstDiv As String         ' Alternative /Required 여부
Public OutLabCd As String       ' 외부검사 기관
Public GrpFg As String          ' Graph 결과여부('0':부,'1':여)
Public TxtType As String        ' 텍스트결과유형('0':부,'1':여)
Public WorkUnit As Integer      ' Work Load Unit
Public AttrCd As String         ' 속성 패턴 코드
Public MatchCd As String
Public ExpDt As String          ' 폐기일
Public ItemSeq As Integer       ' 통계출력순서
Public GroupCd As String        ' 통계Group

