Attribute VB_Name = "Comm"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'통신설정
Type config
    gPort       As String
    gSpeed      As String
    gParity     As String
    gDataBit    As String
    gStopBit    As String
    gStartBit   As String
    gRTSEnable  As String
    gDTREnable  As String
End Type
Public gSetup As config

Public gGubun As Integer
Public gEquip As String
Public gEquipSlip As String
Public gExamUID As String
Public gOrderSelect As Integer
Public gVersion As String

'Public gSugaCode As String
Public gExamCode As String
Public gResult As String

Type DB_Parm
    Driver  As String
    User    As String
    Passwd  As String
    Server  As String
    DB      As String
    HostName    As String
End Type
Public gDB_Parm As DB_Parm
Public gLocalDBPath As String

Public raw_data As String

Public gCurDate As String
Public gCurMsgCnt As String
Public gOrdCnt As String

Public gOrderMessage As String

Public comState As String
Public comsignal As String
Public comSend As String

Public gOrderCnt As Integer
Public gNACKCnt As Integer
Public gENQCnt As Integer

Public gPreSpecID As String
Public gPreRow As Long
Public glRow As Long
Public gOrdRow As Long
Public gSndState As String
Public gSndMessage As String
Public gOrdNum As Integer


Public gHeader As String    'UltraM에서 사용
Public gPatient As String

Public gPatCnt As String    'Axsym
Public gOrder As String

'Public gOrder As String
Public gResCnt As String    'Axsym
Public gMsgEnd As String

Public gMsgFlag As String
Public gPreMsg As String

Public gEquipCode As String
Public gHeadRecode As String    'UltraM
Public gRecodeType As String    'Axsym, ABL50

Public gState As String
'장비코드 당 검사코드가 하나라면 무조건 배열에 가지고 있는 부분
'CD3000에서 사용
Public gArr_Exam() As String

'CD3000외에서 사용
Public gArr_ExamCode() As String

Public gAllExam As String
Public gSelExam As String   'WorkList에서 사용

Public gGetSetup As Integer
Public gSelFlag As Integer

Type sOrder
    OrderText  As String
    OrderCnt    As Integer
    ExamCode As String
    SampleType1 As String
    SampleType2 As String
End Type
Public gOrd As sOrder

Public Function chrSTX() As String
    chrSTX = Chr(2)
End Function

Public Function chrETX() As String
    chrETX = Chr(3)
End Function

Public Function chrEOT() As String
    chrEOT = Chr(4)
End Function

Public Function chrENQ() As String
    chrENQ = Chr(5)
End Function

Public Function chrACK() As String
    chrACK = Chr(6)
End Function

Public Function chrTAB() As String
    chrTAB = Chr(9)
End Function

Public Function chrLF() As String
    chrLF = Chr(10)
End Function

Public Function chrCR() As String
    chrCR = Chr(13)
End Function

Public Function chrNACK() As String
    chrNACK = Chr(15)
End Function

Public Function chrSPC() As String
    chrSPC = Chr(20)
End Function

Public Function chrETB() As String
    chrETB = Chr(23)
End Function

Function GetEquipInfo() As Boolean
'ini아닌 서버에서 불러오기
    
'    SQL = " Select Port, Baud, PARITY, DATABIT, STOPBIT, STARTBIT " & _
'          " From Equip " & _
'          " Where HID = '111' " & _
'          " And EquipCode = '" & Trim(gEquip) & "' "
'
'    res = db_select_Col(gServer, SQL)
'
'    If res = 1 And gEquip <> "" Then
'        gSetup.gPort = Trim(gReadBuf(0))
'        gSetup.gSpeed = Trim(gReadBuf(1))
'        gSetup.gParity = Trim(gReadBuf(2))
'        gSetup.gDataBit = Trim(gReadBuf(3))
'        gSetup.gStopBit = Trim(gReadBuf(4))
'        gSetup.gStartBit = Trim(gReadBuf(5))
'
'        Select Case Trim(gReadBuf(6))
'        Case "T"
'            gSetup.gRTSEnable = True
'        Case "F"
'            gSetup.gRTSEnable = False
'        End Select
'
'        Select Case Trim(gReadBuf(7))
'        Case "T"
'            gSetup.gDTREnable = True
'        Case "F"
'            gSetup.gDTREnable = False
'        End Select
'    End If
    If Not IsNumeric(gSetup.gPort) Then
        gSetup.gPort = 1
        gSetup.gSpeed = 9600
        gSetup.gParity = "N"
        gSetup.gDataBit = 8
        gSetup.gStopBit = 1
        gSetup.gStartBit = 1
        gSetup.gRTSEnable = True
        gSetup.gDTREnable = True
    End If

End Function

Public Function CheckSum(ByVal CheSum As String) As String
    Dim Tot  As Integer
    Dim sStr As String
    Dim i As Integer
    
    For i = 1 To Len(CheSum)
        Tot = Tot + Asc(Mid(CheSum, i, 1))
    Next i
    
    sStr = Hex(Tot)
    If Len(sStr) = 1 Then
        sStr = "0" & sStr
    End If
    
    CheckSum = Right(sStr, 2)
End Function

