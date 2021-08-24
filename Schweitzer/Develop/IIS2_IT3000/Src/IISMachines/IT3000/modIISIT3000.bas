Attribute VB_Name = "modIISIT3000"
'-----------------------------------------------------------------------------'
'   파일명 : modIISIT3000.bas
'   작성자 : 오세원
'   내  용 : IT3000 장비의 옵션저장 모듈
'   작성일 : 2014-07-28
'   버  전 :
'-----------------------------------------------------------------------------'

Option Explicit

Public mOrderPath     As String   '오더파일 생성경로
Public mResultPath    As String   '결과파일 생성경로
Public mOrderFileNm   As String   '오더파일명
Public mResultFileNm  As String   '결과파일명 확장자
Public mOrderRefresh  As String   '오더파일 Refresh time(sec)
Public mResultRefresh As String   '결과파일 Refresh time(sec)
Public mDB            As String
Public mUID           As String
Public mPW            As String

