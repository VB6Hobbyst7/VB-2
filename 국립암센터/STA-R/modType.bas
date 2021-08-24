Attribute VB_Name = "modType"
Option Explicit

Type gTb006
    EqpCd As String
    EqpNm As String
    Model As String
    PurChdt As String
    BaudRate As String
    ParityBit As String
    StopBit As String
    DataBit As String
End Type

Type gTb001
    TestCd As String
    TestNm As String
    ApplyDt As String
    WorkArea As String
    PanelFg As String
    WorkAreaNm As String
End Type

Type gTb015
    EmpId As String
    EmpNm As String
    LoginId As String
    PassWord As String
    EmpLngNm As String
End Type

Type gTb032
    CdIndex As String
    CdVal As String
    Field1 As String
    Field2 As String
    Field3 As String
    Field4 As String
End Type

Type gTbWorkList
    WorkArea As String
    AccDt As String
    AccSeq As Integer
    RcvTm As String
    PtId As String
    PtNm As String
    AGE As Integer
    SEX As String
    WorkAreaNm As String
    ColDt As String
    ColTm As String
    SpcYy As String
    SpcCd As Long
End Type

Type gTb703
    EqpCd As String
    Eqpseq As Integer
    TestCd As String
    SpcCd As String
    Intbase As String
    Intnm As String
    Prtord As Integer
End Type
