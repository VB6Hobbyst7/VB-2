Attribute VB_Name = "modMMIF"
Option Explicit

'/레지스트 정보 변수-----------------------------------------------------------------------/
Public Const REG_MAKER              As String = "MEDIMATE"          '/제작사
Public Const REG_PRODUCT            As String = "SM"                '/제품명
Public Const REG_EQUIP              As String = "INTEGRA400PLUS"    '/장비코드

Public Const REG_DB_CONSTR_LOCAL    As String = "DB_CONSTR_LOCAL"   '/Local DB 연결문자 정보
Public Const REG_DB_CONSTR_HIS      As String = "DB_CONSTR_HIS"     '/HIS DB 연결문자 정보
Public Const REG_DB_CONSTR_QC       As String = "DB_CONSTR_QC"      '/QC DB 연결문자 정보
Public Const REG_SERIALPORT         As String = "SERIALPORT"
Public Const REG_SERIALBAUD         As String = "SERIALBAUD"
Public Const REG_SERIALDATABIT      As String = "SERIALDATABIT"
Public Const REG_SERIALSTARTBIT     As String = "SERIALSTARTBIT"
Public Const REG_SERIALSTOPBIT      As String = "SERIALSTOPBIT"
Public Const REG_SERIALPARITY       As String = "SERIALPARITY"
Public Const REG_SERIALRTS          As String = "SERIALRTS"
Public Const REG_SERIALDTR          As String = "SERIALDTR"
Public Const REG_EQ_NAME            As String = "EQ_NAME"           '/장비명(화면상의 Title로 활용)
Public Const REG_PG_WORKLIST        As String = "PG_WORKLIST"       '/Work List 사용여부(1.사용, 2.미사용)
Public Const REG_PG_QC              As String = "PG_QC"             '/QC 업무사용여부(1.사용, 2.미사용)
Public Const REG_PG_WAITTIME        As String = "PG_WAITTIME"       '/Integra400 장비의 Request Delay Time

'/프로그램 환경 변수
Type REG_INFO
    EQUIPCD         As String   '/장비코드
    EQUIPSEQ        As Long     '/장비일련번호(동일장비가 여러대일 경우 사용할 목적임/기본값 1) 레지스터 구성상수는 나중에 정의할 예정임(오덕진)
    DB_CONSTR_LOCAL As String   '/Local DB 연결문자 정보
    DB_CONSTR_HIS   As String   '/HIS DB 연결문자 정보
    DB_CONSTR_QC    As String   '/QC DB 연결문자 정보
    SERIALPORT      As String
    SERIALBAUD      As String
    SERIALDATABIT   As String
    SERIALSTARTBIT  As String
    SERIALSTOPBIT   As String
    SERIALPARITY    As String
    SERIALRTS       As String
    SERIALDTR       As String
    PG_EQ_NAME      As String   '/장비명(화면상의 Title로 활용)
    PG_WORKLIST     As String   '/WorkList 사용여부(1.사용, 2.미사용)
    PG_QC           As String   '/QC 업무사용여부(1.사용, 2.미사용)
    PG_WAITTIME     As String   '/Integra400 장비의 Request Delay Time
End Type

Public gtypREG_INFO  As REG_INFO

'/임시 변수
Public intX                     As Integer
Public intY                     As Integer
Public intZ                     As Integer
Public strTemp                  As String

'''Public Function Centaur_Str() As String
'''    Centaur_Str = Chr(240)
'''End Function
'''
'''Public Function Centaur_End() As String
'''    Centaur_End = Chr(248)
'''End Function
'''
'''Public Function chrSOH() As String
'''    chrSOH = Chr(1)
'''End Function
'''
'''Public Function chrSTX() As String
'''    chrSTX = Chr(2)
'''End Function
'''
'''Public Function chrETX() As String
'''    chrETX = Chr(3)
'''End Function
'''
'''Public Function chrEOT() As String
'''    chrEOT = Chr(4)
'''End Function
'''
'''Public Function chrENQ() As String
'''    chrENQ = Chr(5)
'''End Function
'''
'''Public Function chrACK() As String
'''    chrACK = Chr(6)
'''End Function
'''
'''Public Function chrTAB() As String
'''    chrTAB = Chr(9)
'''End Function
'''
'''Public Function chrLF() As String
'''    chrLF = Chr(10)
'''End Function
'''
'''Public Function chrCR() As String
'''    chrCR = Chr(13)
'''End Function
'''
'''Public Function chrNACK() As String
'''    chrNACK = Chr(15)
'''End Function
'''
'''Public Function chrSPC() As String
'''    chrSPC = Chr(20)
'''End Function
'''
'''Public Function chrETB() As String
'''    chrETB = Chr(23)
'''End Function

Public Sub GET_REGIST()
    With gtypREG_INFO
        .EQUIPCD = REG_EQUIP
        .EQUIPSEQ = 1
        .DB_CONSTR_LOCAL = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_DB_CONSTR_LOCAL)
        .DB_CONSTR_HIS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_DB_CONSTR_HIS)
        .DB_CONSTR_QC = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_DB_CONSTR_QC)
        .SERIALPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALPORT)
        .SERIALBAUD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALBAUD)
        .SERIALDATABIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALDATABIT)
        .SERIALSTARTBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALSTARTBIT)
        .SERIALSTOPBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALSTOPBIT)
        .SERIALPARITY = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALPARITY)
        .SERIALRTS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALRTS)
        .SERIALDTR = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_SERIALDTR)
        .PG_EQ_NAME = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_EQ_NAME)
        .PG_WORKLIST = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_PG_WORKLIST)
        .PG_QC = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_PG_QC)
        .PG_WAITTIME = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_PG_WAITTIME)
    End With
End Sub

'''Public Sub Save_Raw_Data(ArgSQL As String)
'''    Dim FilNum
'''    Dim strFileName As String
'''
'''    FilNum = FreeFile
'''
'''    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
'''        MkDir (App.Path & "\Result")
'''    End If
'''
'''    strFileName = gtypREG_INFO.EQUIPCD & "_" & Format(Date, "yyyymmdd")
'''
'''    Open App.Path & "\Result\" & strFileName & ".txt" For Append As FilNum
'''    Print #FilNum, ArgSQL
'''    Close FilNum
'''End Sub

