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
    ACKUse      As String

End Type
Public gSetup As config
Public gPart As String
Public gGubun As Integer
Public gEquip As String
Public gEquipCode As String
Public gIP As String
Public gOrderExam As String

Public Const dpGumjin1 = 0
Public Const dpOCS = 1
Public Const dpGumjin2 = 2

Public gHeadRecode As String    'UltraM
Public gRecodeType As String    'Axsym, ABL50

Public gTxMsgFlag As String
Public gCurTxCnt As Integer         '0~7이고, 8이면 다시 0부터 시작함
Public gOrderMessage As String
Public gPreData As String
Public gNACKCnt As Integer
Public gOrderRow As Long


Public gQCEquip As String

Public gEquipID As String

Type DB_Parm
    Driver  As String
    User    As String
    Passwd  As String
    Server  As String
    db      As String
    hostname    As String
    Localdb As String
End Type

Public gDB_Parm As DB_Parm
Public gOCSDB_Parm As DB_Parm
Public gLocalDB_Parm As DB_Parm


Public comState As String
Public comsignal As String
Public comSend As String

Public gOrderCnt As Integer
Public gPreMsg As String
Public gACKSig As Integer
Public gCurDate As String
Public gCurMsgCnt As String
Public gENQCnt As Integer
Public sOCnt As Integer
Public gHeader As String    'UltraM에서 사용
Public gPatient As String
Public gPatCnt As String    'Axsym
Public gOrder As String
Public gResCnt As String    'Axsym
Public gMsgEnd As String


Public gAllExam As String
Public gAllExam_Ocs As String

Public gAllExam1 As String
Public gAllExam_Ocs1 As String

Public gArrEquip() As String
Public gArrEquip_Ocs() As String

Public gMachPath  As String

Public Function chrSTX() As String
    chrSTX = Chr(2)
End Function

Public Function chrETX() As String
    chrETX = Chr(3)
End Function

Public Function chrSOH() As String
    chrSOH = Chr(1)
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

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False
    

    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "EquipCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquipCode = Trim(frmInterface.txtTemp)
    
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
    
    
    '검진 접속
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.db = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    '진료접속
    db_tmp = ""
    Call GetPrivateProfileString("DATABASEOCS", "driver", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    
    Call GetPrivateProfileString("DATABASEOCS", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOCSDB_Parm.Server = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    
    Call GetPrivateProfileString("DATABASEOCS", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOCSDB_Parm.db = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASEOCS", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOCSDB_Parm.User = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASEOCS", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gOCSDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    '로컬접속
    db_tmp = ""
    Call GetPrivateProfileString("DATABASELOCAL", "driver", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gLocalDB_Parm.Driver = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    
    Call GetPrivateProfileString("DATABASELOCAL", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gLocalDB_Parm.Server = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    
    Call GetPrivateProfileString("DATABASELOCAL", "database", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gLocalDB_Parm.db = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASELOCAL", "uid", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gLocalDB_Parm.User = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASELOCAL", "pwd", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gLocalDB_Parm.Passwd = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gMachPath = Trim(frmInterface.txtTemp)
    
    
    
    GetSetup = True

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

Public Function CS(ByVal CheSum As String) As String
    Dim Tot  As Integer
    Dim sStr As String
    Dim i As Integer
    
    For i = 1 To Len(CheSum)
        Tot = Tot + Asc(Mid(CheSum, i, 1))
    Next i
    
    Tot = 256 - (Tot Mod 256)
    sStr = Hex(Tot)
    If Len(sStr) = 1 Then
        sStr = "0" & sStr
    End If
    
    CS = Right(sStr, 2)
End Function

Public Function MOR() As String
    MOR = Chr(2) & ">" & Chr(3) & "3E" & Chr(13)
End Function

Public Function REP() As String
    REP = Chr(2) & "?" & Chr(3) & "3F" & Chr(13)
End Function
Public Sub Save_Raw_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(CDate(frmInterface.dtpToday), "yyyy-mm-dd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, Format(Time, "hh:nn:ss") & " " & argSQL
    Close FilNum
End Sub

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

