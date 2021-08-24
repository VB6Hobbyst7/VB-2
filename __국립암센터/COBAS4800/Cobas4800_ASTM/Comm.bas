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
    gPort2      As String
    gTestWay    As String
    gTestIdName As String
    gTestIdList As String
End Type

Public gSetup As config
Public gPart As String
Public gGubun As Integer
Public gEquip As String
Public gEquipID As String

Public gEGFRMDCMNT As String
Public gEGFRNDCMNT As String

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

Public comState As String
Public comsignal As String
Public comSend As String

Public gOrderMessage As String
Public gOrderCnt    As Integer
Public gNACKCnt     As Integer
Public gPreMsg      As String
Public gACKSig      As Integer

Public gSndState    As String

Public gMsgFlag     As String

Public gHeadRecode  As String
Public gRecodeType  As String

Public gHeader      As String
Public gPatient     As String

Public gPatCnt      As String
Public gOrder       As String
Public gResCnt      As String
Public gMsgEnd      As String

Public gPreSpecID   As String
Public gPreRow      As Long
Public gOrdRow      As Long
Public gPreData     As String

Public gVersion     As String       '장비버전

Public gCurMsgCnt   As String

Public gArrEquip()  As String

Public gTimer       As String

Public gENQCnt As Integer

Public gXMLResultFileName As String

Type sOrder
    OrderText  As String
    OrderCnt    As Integer
    ExamCode As String
    SampleType1 As String
    SampleType2 As String
End Type
Public gOrd As sOrder

Public gAllExam As String
Public gArr_Exam(1 To 22, 1 To 3) As String

'CD3000외에서 사용
Public gArr_ExamCode() As String




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


Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gEquip = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gPort = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gSpeed = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gParity = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gDataBit = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gStopBit = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gStartBit = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gRTSEnable = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gSetup.gDTREnable = Trim(frmLogin.txtTemp)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "Timer", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gTimer = Trim(frmLogin.txtTemp)
'    If gTimer = "" Then
'        gTimer = 30000
'    End If
'
'    db_tmp = ""
'    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
'    frmLogin.txtTemp = Trim(db_tmp)
'    gServerPath = Trim(frmLogin.txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gEquip = Trim(frmInterface.txtTemp)

    db_tmp = ""
    
    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gPort = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    
    Call GetPrivateProfileString("config", "gTestWay", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gTestWay = Trim(frmInterface.txtTemp)

    
    db_tmp = ""
    
    Call GetPrivateProfileString("config", "gTestIdName", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gTestIdName = Trim(frmInterface.txtTemp)

    db_tmp = ""
    
    Call GetPrivateProfileString("config", "gTestIdList", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gTestIdList = Trim(frmInterface.txtTemp)

   
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "gPort2", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gSetup.gPort2 = Trim(frmInterface.txtTemp)

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
    Call GetPrivateProfileString("config", "Timer", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gTimer = Trim(frmInterface.txtTemp)
    If gTimer = "" Then
        gTimer = 30000
    End If

    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gServerPath = Trim(frmInterface.txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "IFUser", "", db_tmp, 100, App.Path & "\Interface.ini")
    frmInterface.txtTemp = Trim(db_tmp)
    gIFUser = Trim(frmInterface.txtTemp)
    
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

Public Sub Save_XML_File(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    argSQL = Replace(argSQL, "癤?", "<")
    FilNum = FreeFile
    
    If Dir(App.Path & "\ResultXML", vbDirectory) <> "ResultXML" Then
        MkDir (App.Path & "\ResultXML")
    End If
    
    sFileName = gXMLResultFileName
    
    Open App.Path & "\ResultXML\" & sFileName & ".xml" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Public Sub Save_Raw_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(Date, "YYYYMMDD")
    
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

