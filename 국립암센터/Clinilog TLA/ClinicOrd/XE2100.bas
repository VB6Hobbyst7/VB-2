Attribute VB_Name = "XE2100"
Option Explicit

'통신설정
Type COMConfig
    ComPort       As String
    Speed      As String
    Parity     As String
    DataBit    As String
    StopBit    As String
    StartBit   As String
    RTSEnable  As String
    DTREnable  As String
    ExamUID    As String
    Gubun      As String
    ConnectFlag As Boolean
    UseEquip   As String
    Protocol    As String
End Type
Public IPU1 As COMConfig
Public IPU2 As COMConfig

Public gArrExam()
