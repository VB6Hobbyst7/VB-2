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
Public gExamUID As String

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
    LocalDB As String
End Type
Public gDB_Parm As DB_Parm

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

Public gPreMsg As String

Public gEquipCode As String
Public gHeadRecode As String    'UltraM
Public gRecodeType As String    'Axsym, ABL50

Public gTxMsgFlag As String
Public gCurTxCnt As Integer         '0~7이고, 8이면 다시 0부터 시작함
Public gOrderMessage As String
Public gPreData As String
Public gNACKCnt As Integer

Public gState As String
'장비코드 당 검사코드가 하나라면 무조건 배열에 가지고 있는 부분
'CD3000에서 사용
Public gArr_Exam(1 To 50, 1 To 4) As String

'CD3000외에서 사용
Public gArr_ExamCode() As String

Public gAllExam As String
Public gAllOcsExam As String

Public gSleepSec As String

Public gOrderSelect As String

Public gOrderRow As Long    'WorkList - Order 전송시 사용
Public gWorkFlag As Long

Public comState As String
Public comsignal As String
Public comSend As String

'보건소 적용
Public gUID As String
Public gHPID As String
Public gHPEquip As String
Public gAddr As String

Public Function Centaur_Str() As String
    Centaur_Str = Chr(240)
End Function

Public Function Centaur_End() As String
    Centaur_End = Chr(248)
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
    
    SQL = " Select Port, baud, PARITY, DATABIT, STOPBIT, STARTBIT, RTS, DTR " & _
          " From Equip " & _
          " Where EquipCode = '" & Trim(gEquip) & "' "
    SaveQuery SQL
    res = db_select_Col(gServer, SQL)

    If res = 1 And gEquip <> "" Then
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
    
    Call GetPrivateProfileString("code", "EquipCode", "", db_tmp, 20, App.Path & "\Interface.ini")
    gEquip = Mid(db_tmp, 1, 5)
    
    'Comport
    db_tmp = ""
    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gSpeed = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gParity = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDataBit = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gStopBit = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gStartBit = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gRTSEnable = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gDTREnable = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "WaitTime", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSleepSec = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "OrderSelect", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOrderSelect = Trim(frmInterface.txtTemp)
    
    'DATABASE
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 20, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "hostname", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.HostName = Trim(frmInterface.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "SiteID", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gHPID = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "EquipCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gHPEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "UserID", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    frmInterface.txtUID = Trim(frmInterface.txtTemp)
    gUID = Trim(frmInterface.txtUID)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Address", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gAddr = Trim(frmInterface.txtTemp)
    
    '로컬DB 설정
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "LocalDB", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.LocalDB = Trim(frmInterface.txtTemp)
    
    GetSetup = True

End Function


'Function GetSetup() As Boolean
''---------------------------------------------------------------------------------------------------------------------
''                       Setup  File을 읽어온다.
''---------------------------------------------------------------------------------------------------------------------
'    Dim db_tmp As String * 20
'
'    db_tmp = ""
'
'    GetSetup = False
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "Port", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gPort = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gSpeed = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gParity = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gDataBit = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gStopBit = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "StartBit", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gStartBit = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gRTSEnable = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gSetup.gDTREnable = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString(gEquip, "ExamUID", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gExamUID = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gDB_Parm.Driver = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gDB_Parm.User = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gDB_Parm.Passwd = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gDB_Parm.Server = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gDB_Parm.DB = Trim(frmCD3000.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("DATABASE", "hostname", "", db_tmp, 20, App.Path & "\Interface.ini")
'    frmCD3000.txtTemp = Trim(db_tmp)
'    gDB_Parm.HostName = Trim(frmCD3000.txtTemp)
'
'    GetSetup = True
'
'End Function

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

