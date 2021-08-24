Attribute VB_Name = "Comm"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
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
    gRTSEnable  As Boolean
    gDTREnable  As Boolean
End Type
Public gSetup As config

Public gGubun As Integer
Public gEquip As String
Public gExamUID As String

'Public gSugaCode As String
Public gExamCode As String
Public gResult As String

Type DB_Parm
    DSN As String
    UID As String
    PWD As String
End Type
Public gDB_Ser As DB_Parm
Public gDB_OCS As DB_Parm

Public raw_data As String

Public gCurDate As String
Public gCurMsgCnt As String
Public gOrdCnt As String


Public gHeader As String    'UltraM에서 사용
Public gPatient As String
Public gPatCnt As String    'Axsym
Public gOrder As String
Public gResCnt As String    'Axsym
Public gMsgEnd As String

Public gMsgFlag As String
Public gPreMsg As String

Public gEQUIPCODE As String
Public gHeadRecode As String    'UltraM
Public gRecodeType As String    'Axsym, ABL50

Public gState As String

Public gPreData As String

Public glRow As Long

'Hitachi7170*****************
Public comState As String
Public comsignal As String
Public comSend As String

Public gOrderMessage As String
Public gOrderCnt As Integer
Public gNACKCnt As Integer

Type sOrder
    OrderText  As String
    OrderCnt    As Integer
    ExamCode As String
End Type
Public gOrd As sOrder
'****************************

'장비코드 당 검사코드가 하나라면 무조건 배열에 가지고 있는 부분
'CD3000에서 사용
Public gArr_Exam(1 To 22, 1 To 3) As String

'CD3000외에서 사용
Public gArr_ExamCode() As String

Public gAllExam As String
Public gResultFlag As String

Public Function chrSOH() As String
    chrSOH = Chr(1)
End Function
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
    
    SQL = " SELECT PORT, BAUD, PARITY, DATABIT, STOPBIT, STARTBIT, RTS, DTR " & CR & _
          " From Equip " & CR & _
          " WHERE HID = '117' " & CR & _
          " And EQUIPCODE = '" & Trim(gEquip) & "' "
    res = db_SELECT_Col(gServer, SQL)

    If res = 0 Then
        frmConfig.Show
    ElseIf res = 1 And gEquip <> "" Then
        gSetup.gPort = Trim(gReadBuf(0))
        gSetup.gSpeed = Trim(gReadBuf(1))
        gSetup.gParity = Trim(gReadBuf(2))
        gSetup.gDataBit = Trim(gReadBuf(3))
        gSetup.gStopBit = Trim(gReadBuf(4))
        gSetup.gStartBit = Trim(gReadBuf(5))

        Select Case Trim(gReadBuf(6))
        Case "T"
            gSetup.gRTSEnable = True
        Case "F"
            gSetup.gRTSEnable = False
        End Select

        Select Case Trim(gReadBuf(7))
        Case "T"
            gSetup.gDTREnable = True
        Case "F"
            gSetup.gDTREnable = False
        End Select

    End If

End Function

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False
    
    Call GetPrivateProfileString("CONFIG", "Equip", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "ComPort", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "BaudRate", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gSpeed = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gParity = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDataBit = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gStopBit = Trim(frmInterface.txtTemp)

'    db_tmp = ""
'    Call GetPrivateProfileString("config", "StartBit", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmInterface.txtTemp = Trim(db_tmp)
'    gSetup.gStartBit = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gRTSEnable = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDTREnable = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_DSN", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Ser.DSN = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_UID", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Ser.UID = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_PWD", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Ser.PWD = Trim(frmInterface.txtTemp)
    


    GetSetup = True

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

Public Function ASTM_CSum(ByVal CheSum As String) As String
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
    
    ASTM_CSum = Right(sStr, 2)
End Function

Public Function MOR() As String
    MOR = Chr(2) & ">" & Chr(3) & "3E" & Chr(13)
End Function

Public Function REP() As String
    REP = Chr(2) & "?" & Chr(3) & "3F" & Chr(13)
End Function

