VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLab004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd       As String               ' 검사항목 코드
Public SpcCd        As String               ' 검체 코드
Public Seq          As Integer              ' Seq
Public ApplyDt      As String               ' 적용일
Public SpcGrpCd     As String               ' 검체군 코드
Public LabelCnt     As Integer              ' 바코드라벨 출력장수
Public RstUnit      As String               ' 결과 단위
Public RndFg        As String               ' Round 채혈여부('0':부,'1':여)
Public StatFlags    As String               ' 각 건물별 Stat 여부
Public StatFg       As String               ' Stat 여부('0':부,'1':여)
Public AvalVal      As Integer              ' 유효 숫자
Public PanicFg      As String               ' Panic Check여부('0':부,'1':여)
Public PanicFrVal   As Double               ' Panic Range(FROM)
Public PanicToVal   As Double               ' Panic Range(To)
Public DeltaFg      As String               ' Delta Check여부('0':부,'1':여)
Public DeltaVal1    As Integer              ' Delta Value
Public DeltaVal2    As Integer              ' Delta Value
Public TestCost     As String               ' 수가코드
Public StoreCd      As String               ' 보관구분
Public TatAvg       As Integer              ' TAT소요시간
Public SpcQty       As Double               ' 검체 체취량(TLA)
Public SpcUnit      As String               ' 검체 채취 단위
Public ExpDt        As String               ' 폐기일
Public TATS         As String               ' 검사소요시간
Public ArletFg      As String               ' Arlet Check여부('0':부,'1':여)
Public ArletFrVal   As Double               ' Arlet Range(FROM)
Public ArletToVal   As Double               ' Arlet Range(To)

