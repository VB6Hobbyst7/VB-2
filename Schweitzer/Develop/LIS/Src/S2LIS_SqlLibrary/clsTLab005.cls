VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLab005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd       As String           ' 검사항목 코드
Public SpcCd        As String           ' 검체 코드
Public ApplyDt      As String           ' 적용일
Public ApplySex     As String           ' 적용성별
Public AgeFrom      As Long             ' 적용나이(FROM)
Public AgeTo        As Long             ' 적용나이(To)
Public RefValFrom   As Double           ' 참고치(FROM)
Public RefValTo     As Double           ' 참고치(To)
Public RefCd        As String           ' 참고치(Alpha)
Public RefText      As String           ' 다중참고치(Alpha)
Public ARefValFrom  As Double           ' AutoVerify 참고치(FROM)
Public ARefValTo    As Double           ' AutoVerify 참고치(To)
Public ExpDt        As String           ' 폐기일
Public PanicFrVal   As Double           ' Panic Range(FROM)
Public PanicToVal   As Double           ' Panic Range(To)
Public ArletFrVal   As Double
Public ArletToVal   As Double
Public RefDiv       As String
Public AMRFrVal     As Double           ' AMR Range(FROM)
Public AMRToVal     As Double           ' AMR Range(To)
