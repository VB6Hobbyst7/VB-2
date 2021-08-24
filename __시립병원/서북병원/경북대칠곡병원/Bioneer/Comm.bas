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

Type DB_Parm
    DSN As String
    UID As String
    PWD As String
End Type
Public gDB_Ser As DB_Parm


Type ExamRes
    res As String
    RefLow As String
    RefHigh As String
    RefFlag As String
    EquipCode  As String
    ExamCode  As String
    ExamName As String
    SeqNo As String
    EquipGubun As String
End Type

Public gArrExamRes() As ExamRes

Type Remote
    RemoteHost As String
    RemotePort As String
End Type
Public gRemote As Remote

Public ComState As Boolean

Public gTxMsgFlag As String
Public gCurTxCnt As Integer
Public gOrderMessage As String
Public gPreData As String
Public gHeader As String
Public gPatient As String
Public gOrder As String
Public gMsgEnd As String

Public gArrEquip() As String

Public gAllExam As String

Public gUID As String
Public gHPEQUIP As String

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

Public Function chrFS() As String
    chrFS = Chr(28)
End Function

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 20

    db_tmp = ""
    
    GetSetup = False
    

'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gPort", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gPort = Trim(Form_Main.Text_ini)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gSpeed", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gSpeed = Trim(Form_Main.Text_ini)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gParity", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gParity = Trim(Form_Main.Text_ini)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gDataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gDataBit = Trim(Form_Main.Text_ini)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gStopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gStopBit = Trim(Form_Main.Text_ini)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gStartBit", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gStartBit = Trim(Form_Main.Text_ini)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gRTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gRTSEnable = Trim(Form_Main.Text_ini)
'
'    db_tmp = ""
'    Call GetPrivateProfileString("config", "gDTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
'    Form_Main.Text_ini = Trim(db_tmp)
'    gSetup.gDTREnable = Trim(Form_Main.Text_ini)

    db_tmp = ""
    Call GetPrivateProfileString("config", "RemoteHost", "", db_tmp, 100, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gRemote.RemoteHost = Trim(Form_Main.Text_ini)
    
    db_tmp = ""
    Call GetPrivateProfileString("config", "RemotePort", "", db_tmp, 100, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gRemote.RemotePort = Trim(Form_Main.Text_ini)

    db_tmp = ""
    Call GetPrivateProfileString("config", "gubn", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    If IsNumeric(Trim(Form_Main.Text_ini)) = False Then
        gGubun = 0
    Else
        gGubun = CInt(Trim(Form_Main.Text_ini))
    End If

    db_tmp = ""
    Call GetPrivateProfileString("config", "Equip", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gEquip = Trim(Form_Main.Text_ini)

    db_tmp = ""
    Call GetPrivateProfileString("config", "EQUIPCD", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gHPEQUIP = Trim(Form_Main.Text_ini)
    


    'DATABASE
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_DSN", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gDB_Ser.DSN = Trim(Form_Main.Text_ini)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_UID", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gDB_Ser.UID = Trim(Form_Main.Text_ini)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "Server_PWD", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gDB_Ser.PWD = Trim(Form_Main.Text_ini)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "UID", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    gUID = Trim(Form_Main.Text_ini)
    
    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "MODE", "", db_tmp, 20, App.Path & "\Interface.ini")
    Form_Main.Text_ini = Trim(db_tmp)
    If Trim(Form_Main.Text_ini) = "1" Then
        Form_Main.optOption(0).Value = True
    Else
        Form_Main.optOption(1).Value = True
    End If
    
    GetSetup = True

End Function

Public Function CheckSum(ByVal CheSum As String) As String
    Dim Tot  As Double
    Dim sStr As String
    Dim i As Integer
    
    For i = 1 To Len(CheSum)
        Tot = Tot + Asc(Mid(CheSum, i, 1))
        
        Debug.Print Mid(CheSum, i, 1) & vbTab & Asc(Mid(CheSum, i, 1)) & vbTab & Tot
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

Public Sub Save_Raw_Data(asData As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
        MkDir (App.Path & "\Result")
    End If
    
    sFileName = Format(Now, "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open App.Path & "\Result\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, asData
    Close FilNum
End Sub

