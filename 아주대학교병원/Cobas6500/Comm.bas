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
    gRTSEnable  As Boolean
    gDTREnable  As Boolean
End Type
Public gSetup As config

Public gGubun As Integer
Public gEquip As String
Public gEquipName As String
Public gExamUID As String
Public gArchitect As String
Public gCobas As String
Public gImagePath As String
Public gEquipIP As String


'Public gSugaCode As String
Public gExamCode As String
Public gResult As String

Type DB_Parm
    User    As String
    EnvPath As String
    Server As String
    Passwd As String
    DB  As String
    
End Type

Type Slip
    Slip    As String
    Slip1   As String
    Slip2   As String
    Slip3   As String
End Type

Public gSlip As Slip


Public gUserID As String

Public gDB_Parm As DB_Parm
Public gDB_LParm As DB_Parm
Public gLocalExpDate As Long


Public raw_data As String

Public gCurDate As String
Public gCurMsgCnt As String
Public gOrdCnt As String
Public SinCnt As Integer
Public comSend As String

Public gHeader As String    'UltraM에서 사용
Public gPatient As String
Public gPatCnt As String    'Axsym
Public gOrder As String
Public gResCnt As String    'Axsym
Public gMsgEnd As String

Public gSMP As String
Public gSAC As String
Public gTQ1 As String

Public gMsgFlag As String
Public gPreMsg As String

Public gEquipCode As String
Public gHeadRecode As String    'UltraM
Public gRecodeType As String    'Axsym, ABL50

Public gState As String
'장비코드 당 검사코드가 하나라면 무조건 배열에 가지고 있는 부분
'CD3000에서 사용
Public gArr_Exam(1 To 200, 1 To 4) As String

'CD3000외에서 사용
Public gArr_ExamCode() As String

Public gAllExam As String
Public gAllExam_NAF As String
Public gAllExam_Micro As String
Public gOrderExam As String
Public gReceExam As String
Public gAllExam1 As String
Public comsignal As String
Public gNACKCnt As Integer
Public gOrderCnt As Integer
Public gOrderMessage As String
Public gSndState As String
Public gENQFlag As String
Public gPreSpecID As String
Public gPreRow As String
Public gOrdRow As String
Public gPatFlag As Integer
Public gSpecID As String
Public gSampleInfo As String

Public gtestid As String
Public gCmtFlag As String
Public gResultRes As String

Public gBarCode As String
Public gWorkList As String
Type sOrder
    OrderText  As String
    OrderCnt    As Integer
    ExamCode As String
    SampleType As String
End Type
Public gOrd As sOrder



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

Public Function chrVT() As String
    chrVT = Chr(11)
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

Public Function chrFS() As String
    chrFS = Chr(28)
End Function

'
'Function GetEquipInfo() As Boolean
''ini아닌 서버에서 불러오기
'
'    SQL = " Select Port, baud, PARITY, DATABIT, STOPBIT, STARTBIT, RTS, DTR " & _
'          " From Equip " & _
'          " Where EquipCode = '" & Trim(gEquip) & "' "
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
'
'    End If
'
'End Function

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    

    GetSetup = False
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "Equip", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "EquipName", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquipName = Trim(frmInterface.txtTemp)
    
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
    Call GetPrivateProfileString("config", "UID", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gExamUID = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "Architect", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gArchitect = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "Cobas", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gCobas = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "DeleteDate", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gLocalExpDate = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "WorkList", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gWorkList = Trim(frmInterface.txtTemp)
    
'''    Call GetPrivateProfileString("code", "EquipCode", "", db_tmp, 20, App.Path & "\Interface.ini")
'''    gEquip = Mid(db_tmp, 1, 5)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "IMAGEPATH", "", db_tmp, 50, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gImagePath = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "EQUIPIP", "", db_tmp, 50, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquipIP = Trim(frmInterface.txtTemp)
    
    'Server DB Connect info
    
    
    db_tmp = ""
    Call GetPrivateProfileString("TMAX", "EnvPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.EnvPath = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("TMAX", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("SLIP", "SLIP", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSlip.Slip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("SLIP", "SLIP1", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSlip.Slip1 = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("SLIP", "SLIP2", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSlip.Slip2 = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("SLIP", "SLIP3", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSlip.Slip3 = Trim(frmInterface.txtTemp)
    
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_DSN", "", db_tmp, 30, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_UID", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(frmInterface.txtTemp)
    
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_PWD", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_DB", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(frmInterface.txtTemp)
    
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

