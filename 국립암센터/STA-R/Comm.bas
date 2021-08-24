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
Public gEquipID As String

Type DB_Parm
    Driver  As String
    User    As String
    Passwd  As String
    Server  As String
    db      As String
    HostName    As String
    LocalDB As String
End Type
Public gDB_Parm As DB_Parm

Public ComState As String
Public comsignal As String
Public comSend As String

Public gOrderCnt As Integer
Public gNACKCnt As Integer
Public gPreMsg As String

Public gTxMsgFlag As String
Public gCurTxCnt As Integer         '0~7이고, 8이면 다시 0부터 시작함
Public gOrderMessage As String
Public gPreData As String
Public gHeader As String
Public gPatient As String
Public gOrder As String
Public gMsgEnd As String
Public gOrderCode As String


Public gDate As String              'Header에서 받은 날짜시간

Public gArrEquip() As String

Public gAllExam As String

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

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False
    

    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gSpeed = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gParity = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gDataBit = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gStopBit = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gStartBit = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gRTSEnable = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gSetup.gDTREnable = Trim(frmLogin.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmLogin.txtTemp = Trim(db_tmp)
    gServerPath = Trim(frmLogin.txtTemp)
    
'    db_tmp = ""
'    Call GetPrivateProfileString("Server", "IFUser", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gIFUser = Trim(frmLogin.txtTemp)


    GetSetup = True

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

Public Function MOR() As String
    MOR = Chr(2) & ">" & Chr(3) & "3E" & Chr(13)
End Function

Public Function REP() As String
    REP = Chr(2) & "?" & Chr(3) & "3F" & Chr(13)
End Function


