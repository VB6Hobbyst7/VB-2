Attribute VB_Name = "modVPM_Main"
Option Explicit

'/���� ����
Public intX                 As Integer
Public intY                 As Integer
Public strTemp              As String

Public gstrREG_DB_CONSTR    As String '/����(����) �����ͺ��̽� ���Ṯ��

Public gstrHOS_CUSCD        As String

Type USER_INFO
    USERID          As String
    USERNM          As String
    USERPW          As String
End Type
Public gtypUSER     As USER_INFO

Public gstrUSERID           As String
Public gstrUSERNM           As String
Public gstrUSERPW           As String

Public gstrFTP_RH           As String
Public gstrFTP_RP           As String
Public gstrFTP_UN           As String
Public gstrFTP_PW           As String

Public gstrSTAUS_DB         As String '/DB �������(Y/N)
Public gstrSTAUS_FTP        As String '/FTP �������(Y/N)
Public gstrSTAUS_COM        As String '/COM Port �������(Y/N)
Public gstrSTAUS_PRN        As String '/���� ������ �������(Y/N)

Public sftp                 As New ChilkatSFtp

'/������Ʈ ���� ����-----------------------------------------------------------------------/
Global Const REG_MAKER      As String = "MEDIMATE"          '/���ۻ�
Global Const REG_PRODUCT    As String = "VPM"               '/��ǰ��

Global Const REG_DB_INFO    As String = "DB_INFO"
Global Const REG_DB_CONSTR  As String = "CONNECTSTRING"

Global Const REG_CLIENT_INFO            As String = "CLIENT_INFO"
Global Const REG_CLIENT_EQCD            As String = "EQUIPCD"
Global Const REG_CLIENT_EQNM            As String = "EQUIPNM"
Global Const REG_CLIENT_EQSEQ           As String = "EQUIPSEQ"
Global Const REG_CLIENT_EQPOS           As String = "EQUIPPOSITION"
Global Const REG_CLIENT_EQTYPE          As String = "EQUIPTYPE"
Global Const REG_CLIENT_SERIALYN        As String = "SERIALYN"
Global Const REG_CLIENT_SERIALPORT      As String = "SERIALPORT"
Global Const REG_CLIENT_SERIALBAUD      As String = "SERIALBAUD"
Global Const REG_CLIENT_SERIALDATABIT   As String = "SERIALDATABIT"
Global Const REG_CLIENT_SERIALSTARTBIT  As String = "SERIALSTARTBIT"
Global Const REG_CLIENT_SERIALSTOPBIT   As String = "SERIALSTOPBIT"
Global Const REG_CLIENT_SERIALPARITY    As String = "SERIALPARITY"
Global Const REG_CLIENT_SERIALRTS       As String = "SERIALRTS"
Global Const REG_CLIENT_SERIALDTR       As String = "SERIALDTR"
Global Const REG_CLIENT_RECEIVETYPE     As String = "RECEIVETYPE"
Global Const REG_CLIENT_EQUIPPORT       As String = "EQUIPPORT"
Global Const REG_CLIENT_ORDYN           As String = "ORDYN"
Global Const REG_CLIENT_QUERYTYPE       As String = "QUERYTYPE"
Global Const REG_CLIENT_ZIPYN           As String = "ZIPYN"
Global Const REG_CLIENT_ZIPNM           As String = "ZIPNM"
Global Const REG_CLIENT_EQIMGFILEPATH   As String = "EQIMGFILEPATH"
Global Const REG_CLIENT_FTPIMGFILEPATH  As String = "FTPIMGFILEPATH"

Type EQ_INFO
    EQUIPCODE       As String
    EQUIPSEQ        As Long
    EQUIPNM         As String
    EQUIPTYPE       As String
    SERIALYN        As String
    SERIALPORT      As String
    SERIALBAUD      As String
    SERIALDATABIT   As String
    SERIALSTARTBIT  As String
    SERIALSTOPBIT   As String
    SERIALPARITY    As String
    SERIALRTS       As String
    SERIALDTR       As String
    RECEIVETYPE     As String
    EQUIPPORT       As String
    ORDYN           As String
    DEPTCODE        As String
    QUERYTYPE       As String
    ZIPYN           As String
    ZIPNM           As String
    EQIMGFILEPATH   As String
    FTPIMGFILEPATH  As String
    REMARK          As String
End Type

Public gtypEQ_INFO  As EQ_INFO

''''/00008(�ΰ�����ü���ܱ�)-AL2000
'''Type AL2000
'''    PT      As String '/�˻��Ͻ�(mid(PT,13,14))
'''    BM      As String '/��������� �˸�
'''    HR      As String '/Right Header(Eye Type = mid(HR,3,14), Vavg = mid(HR,17,4), Vlens = mid(HR,21,4))
'''    VR      As String '/Right(Vacd = mid(HR,3,4))
'''    LR      As String '/Right(AXIAL, ACD, LENS) ���� �Ҽ��� 2�ڸ� ����
'''    KR      As String '/Right(K1, K2) ���� �Ҽ��� 2�ڸ� ����
'''    DR      As String '/Right(Desired Ref. = mid(HR,3,6))
'''    FR      As String '/Right(Formula = mid(HR,3,15))
'''    IR1     As String '/Right
'''    RR1     As String '/Right(mid(LENS const,3,19), mid(IOL Power,22,5), mid(Expected Ref., 27, 6))
'''    IR2     As String '/Right
'''    RR2     As String '/Right(mid(LENS const,3,19), mid(IOL Power,22,5), mid(Expected Ref., 27, 6))
'''    IR3     As String '/Right
'''    RR3     As String '/Right(mid(LENS const,3,19), mid(IOL Power,22,5), mid(Expected Ref., 27, 6))
'''    HL      As String '/Left Header(Eye Type, Vavg, Vlens)
'''    VL      As String '/Left(Vacd)
'''    LL      As String '/Left(AXIAL, ACD, LENS) ���� �Ҽ��� 2�ڸ� ����
'''    KL      As String '/Left(K1, K2) ���� �Ҽ��� 2�ڸ� ����
'''    DL      As String '/Left(Desired Ref. = mid(HR,3,6))
'''    FL      As String '/Left(Formula = mid(HR,3,15))
'''    IL1     As String '/Left
'''    RL1     As String '/Left(mid(LENS const,3,19), mid(IOL Power,22,5), mid(Expected Ref., 27, 6))
'''    IL2     As String '/Left
'''    RL2     As String '/Left(mid(LENS const,3,19), mid(IOL Power,22,5), mid(Expected Ref., 27, 6))
'''    IL3     As String '/Left
'''    RL3     As String '/Left(mid(LENS const,3,19), mid(IOL Power,22,5), mid(Expected Ref., 27, 6))
'''    WR      As String '/Right
'''    WL      As String '/Left
'''End Type
'''
'''Public gtyp0008  As AL2000

Public FtpScanFileName()                                        '/FTP Scan File Name
Public FtpScanFileDate()                                        '/FTP Scan File Date

Public FtpScanFileName_IMG()                                    '/FTP Scan File Name
Public FtpScanFileDate_IMG()                                    '/FTP Scan File Date

Public gstrMSCOMM_Buff          As String                       '/MSComm Input String

'/�⺻������ ����
Private Const HWND_BROADCAST    As Long = &HFFFF&
Private Const WM_WININICHANGE   As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Public Sub GET_EQUIPMENT_INFO(ArgEQCD As String, ArgEQSEQ As String)
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        gstrQuy = "SELECT A.*, B.EQUIPNM, B.EQUIPTYPE "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF A, MM_EMR_EQUIP B "
        gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQUIPCODE = B.EQUIPCODE "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPCODE = '" & ArgEQCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPSEQ  =  " & Val(ArgEQSEQ) & " "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            gtypEQ_INFO.EQUIPCODE = Trim(ADR!EQUIPCODE & "")
            gtypEQ_INFO.EQUIPSEQ = Trim(ADR!EQUIPSEQ & "")
            gtypEQ_INFO.EQUIPNM = Trim(ADR!EQUIPNM & "")
            gtypEQ_INFO.EQUIPTYPE = Trim(ADR!EQUIPTYPE & "")
            gtypEQ_INFO.EQIMGFILEPATH = Trim(ADR!EQIMGFILEPATH & "")
            gtypEQ_INFO.FTPIMGFILEPATH = Trim(ADR!FTPIMGFILEPATH & "")
            gtypEQ_INFO.SERIALYN = Trim(ADR!SERIALYN & "")
            gtypEQ_INFO.SERIALPORT = Trim(ADR!SERIALPORT & "")
            gtypEQ_INFO.SERIALBAUD = Trim(ADR!SERIALBAUD & "")
            gtypEQ_INFO.SERIALDATABIT = Trim(ADR!SERIALDATABIT & "")
            gtypEQ_INFO.SERIALSTARTBIT = Trim(ADR!SERIALSTARTBIT & "")
            gtypEQ_INFO.SERIALSTOPBIT = Trim(ADR!SERIALSTOPBIT & "")
            gtypEQ_INFO.SERIALPARITY = Trim(ADR!SERIALPARITY & "")
            gtypEQ_INFO.SERIALRTS = Trim(ADR!SERIALRTS & "")
            gtypEQ_INFO.SERIALDTR = Trim(ADR!SERIALDTR & "")
            gtypEQ_INFO.RECEIVETYPE = Trim(ADR!RECEIVETYPE & "")
            gtypEQ_INFO.EQUIPPORT = Trim(ADR!EQUIPPORT & "")
            gtypEQ_INFO.ORDYN = Trim(ADR!ORDYN & "")
            gtypEQ_INFO.DEPTCODE = Trim(ADR!DEPTCODE & "")
            gtypEQ_INFO.REMARK = Trim(ADR!REMARK & "")
            gtypEQ_INFO.QUERYTYPE = Trim(ADR!QUERYTYPE & "")
            gtypEQ_INFO.ZIPYN = Trim(ADR!ZIPYN & "")
            gtypEQ_INFO.ZIPNM = Trim(ADR!ZIPNM & "")

            ADR.Close: Set ADR = Nothing
        End If
        
        Call CloseDB
    End If
End Sub

Public Function SET_DEFAULT_FOLDER(ArgSection As String) As Boolean
'/ArgSection: "EQUIP" "FTP"
    SET_DEFAULT_FOLDER = False
    
On Error GoTo RTN_ERROR

    If Dir(App.Path & "\" & gtypEQ_INFO.EQUIPCODE, vbDirectory) = "" Then
        MkDir App.Path & "\" & gtypEQ_INFO.EQUIPCODE
    End If
    If Dir(App.Path & "\" & gtypEQ_INFO.EQUIPCODE & "\" & gtypEQ_INFO.EQUIPSEQ, vbDirectory) = "" Then
        MkDir App.Path & "\" & gtypEQ_INFO.EQUIPCODE & "\" & gtypEQ_INFO.EQUIPSEQ
    End If
    MkDir App.Path & "\" & gtypEQ_INFO.EQUIPCODE & "\" & gtypEQ_INFO.EQUIPSEQ & "\" & ArgSection

    SET_DEFAULT_FOLDER = True

RTN_ERROR:

End Function

Public Sub PrinterChange(strPrinterName As String)
    Dim strBuffer       As String
    Dim iRet            As Long
    Dim strPrinter()    As String

    If Len(Trim(strPrinterName)) > 0 Then
        strBuffer = Space(1024)
        
        iRet = GetProfileString("Devices", strPrinterName, "", strBuffer, Len(strBuffer))
        iRet = GetProfileString("PrinterPorts", strPrinterName, "", strBuffer, Len(strBuffer))

        strPrinter = Split(strBuffer, ",", -1, vbTextCompare)

        iRet = WriteProfileString("windows", "Device", strPrinterName & "," & strPrinter(0) & "," & strPrinter(1))
        iRet = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
    End If
End Sub

Public Sub Main()
    '/Main ȭ�� ����
    frmVPM_Main.Show
'''    If gtypEQ_INFO.EQUIPTYPE = "2" Then
'''        frmVPM_Main.Show
'''    Else
'''        MsgBox "�� ���α׷��� VPM ���� ���α׷��Դϴ�." & vbCrLf & vbCrLf & _
'''               "��� �´� ���α׷��� �̿��Ͻñ� �ٶ��ϴ�.", vbCritical, "���α׷� ����"
'''        End
'''    End If
End Sub

Public Function SAVE_00022() As Boolean
'''    Dim str����     As String
'''    Dim strü��     As String
'''    Dim strü������ As String
'''    Dim strBMI      As String
'''
'''    SAVE_00022 = False
'''
'''On Error GoTo RTN_ERR
'''
'''    If OpenDB1(gstrREG_DB_CONSTR_00022) = False Then End
'''
'''    ADC1.BeginTrans
'''
'''    gstrQuy = "SELECT * "
'''    gstrQuy = gstrQuy & vbCrLf & "  FROM OUTD_DATA "
'''    gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID  = '" & lbl���Ϲ�ȣ & "' "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND DATETIMES >= '" & Replace(lbló������, "-", "") & "000000' "
'''    gstrQuy = gstrQuy & vbCrLf & " ORDER BY DATETIMES "
'''    If ReadSQL1(gstrQuy, ADR1) = False Then Call CloseDB1: End
'''
'''    If Not ADR Is Nothing Then
'''        str���� = Trim(ADR1!Height & "")
'''        strü�� = Trim(ADR1!Weight & "")
'''        strü������ = Trim(ADR1!Height & "")
'''        strBMI = Trim(ADR1!BMI & "")
'''
'''
'''
'''        ADR.Close: Set ADR = Nothing
'''
'''        '/Server DB�� ����� �Է��� �Ǿ� ������ �˻����ڸ� Update ��.
'''        gstrQuy = "UPDATE MM_EMR_RES SET "
'''        gstrQuy = gstrQuy & vbCrLf & "       EXAMDATE  = TO_CHAR(TRUNC(SYSDATE),'YYYYMMDD') "
'''        gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & lbl���Ϲ�ȣ & "' "
'''        gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(lbló������, "-", "") & "' "
'''        gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(lbló��SEQ) & " "
'''        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
'''        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
'''        If RunSQL(gstrQuy) = False Then ADC1.RollbackTrans: Call CloseDB: End
'''    Else
'''        '/����ڵ庰 ó���ڵ� ��������
'''        gstrQuy = "INSERT INTO MM_EMR_RES "
'''        gstrQuy = gstrQuy & vbCrLf & " (PATNO,      ORDDATE,    ORDSEQ,     EXAMDATE,       DEPTCODE, "
'''        gstrQuy = gstrQuy & vbCrLf & "  PARTCODE,   EQUIPCODE,  EXAMCODE,   WORDNO,         ROOMNO, "
'''        gstrQuy = gstrQuy & vbCrLf & "  IOFLAG,     EXECID,     DRID,       IMGFILENAME,    IMGFILEPATH, "
'''        gstrQuy = gstrQuy & vbCrLf & "  RECEDATE,   RECESEQ,    EQUIPSEQ) "
'''        gstrQuy = gstrQuy & vbCrLf & " VALUES "
'''        gstrQuy = gstrQuy & vbCrLf & " ('" & lbl���Ϲ�ȣ & "', "                    '/PATNO(���Ϲ�ȣ)
'''        gstrQuy = gstrQuy & vbCrLf & "  '" & Replace(lbló������, "-", "") & "', "  '/ORDDATE(ó������)
'''        gstrQuy = gstrQuy & vbCrLf & "   " & Val(lbló��SEQ) & ", "                 '/ORDSEQ(ó��SEQ(�ǰ������� ��� ������ȣ))
'''        gstrQuy = gstrQuy & vbCrLf & "  TO_CHAR(TRUNC(SYSDATE),'YYYYMMDD'), "       '/EXAMDATE(����Է�����)
'''        gstrQuy = gstrQuy & vbCrLf & "  '" & lbl����� & "', "                      '/DEPTCODE(������ڵ�)
'''        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/PARTCODE(������ڵ�)
'''        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypEQ_INFO.EQUIPCODE & "', "          '/EQUIPCODE(����ڵ�)
'''        gstrQuy = gstrQuy & vbCrLf & "  '" & lbló���ڵ� & "', "                    '/EXAMCODE(�˻��ڵ�)
'''        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/WORDNO(�����ڵ�)
'''        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/ROOMNO(�����ڵ�)
'''        Select Case lbl�Կܱ���                                                     '/IOFLAG(�Կ�/�ܷ�/���� ����)
'''            Case "�Կ�": gstrQuy = gstrQuy & vbCrLf & "  'A', "
'''            Case "�ܷ�": gstrQuy = gstrQuy & vbCrLf & "  'O', "
'''            Case "����": gstrQuy = gstrQuy & vbCrLf & "  'M', "
'''            Case Else:   gstrQuy = gstrQuy & vbCrLf & "  '', "
'''        End Select
'''        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/EXECID(������ȣ)
'''        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/DRID(ó���ǹ�ȣ)
'''        gstrQuy = gstrQuy & vbCrLf & "  '" & strFileName & "', "                    '/IMGFILENAME(����̹������ϸ�)
'''        gstrQuy = gstrQuy & vbCrLf & "  '" & strIMGFILEPATH & "', "                 '/IMGFILEPATH(����̹������ϰ��)
'''        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/RECEDATE(��������)
'''        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/RECESEQ(����SEQ)
'''        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypEQ_INFO.EQUIPSEQ & "') "           '/EQUIPSEQ(���SEQ)
'''        If RunSQL(gstrQuy) = False Then ADC1.RollbackTrans: Call CloseDB: End
'''    End If
'''
'''    ADC1.CommitTrans
'''
'''    Call CloseDB
'''
'''    SAVE_00022 = True
'''
'''RTN_ERR:
'''
End Function
