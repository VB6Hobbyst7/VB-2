Attribute VB_Name = "modIISLASCHST"
'-----------------------------------------------------------------------------'
'   파일명 : modIISLASCHST.bas
'   작성자 : 이상대
'   내  용 : LASC-HST 장비의 옵션저장 모듈
'   작성일 : 2005-09-15
'   버  전 :
'-----------------------------------------------------------------------------'

Option Explicit

Public mPort        As Integer  'LASC-HST 사용포트
Public mBaudRate    As String   'LASC-HST Baud Rate
Public mDataBit     As String   'LASC-HST Data Bit
Public mStopBit     As String   'LASC-HST Stop Bit
Public mParityBit   As String   'LASC-HST Parity Bit
Public mInterval    As Long     '오더전송 시간간격

