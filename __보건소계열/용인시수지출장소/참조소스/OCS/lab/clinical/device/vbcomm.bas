Attribute VB_Name = "Communication"
Option Explicit
Dim chrs, chrsin, chrsout, idx

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Axsym
Global H_FRAME
Global P_FRAME
Global O_FRAME
Global L_FRAME
Global O_NO
Global P_NO
Global MOD_8

'Data Communication Setting
Global CfgPort              As String
Global CfgComm              As String
Global GnRCntNo             As Integer
Global GnPCntNo             As Integer
Global Errno

Global GPcCode              As String
Global PcSerial             As String
Global GGJCODE              As String
Global GGCODE               As String

Global ComPort              As String
Global Settings             As String

' --------------------------------------------- Byer KF3_
Public Const KF3_Receive_SampleNo = 0
Public Const KF3_Receive_WBC = 1
Public Const KF3_Receive_RBC = 2
Public Const KF3_Receive_HGB = 3
Public Const KF3_Receive_HCT = 4
Public Const KF3_Receive_MCV = 5
Public Const KF3_Receive_MCH = 6
Public Const KF3_Receive_MCHC = 7
Public Const KF3_Receive_RDW = 8
Public Const KF3_Receive_HDW = 9
Public Const KF3_Receive_PLT = 10
Public Const KF3_Receive_MPV = 11
Public Const KF3_Receive_PDW = 12

Public Const KF3_Receive_NEUT_CNT = 13
Public Const KF3_Receive_LYMP_CNT = 14
Public Const KF3_Receive_MONO_CNT = 15
Public Const KF3_Receive_EOS_CNT = 16
Public Const KF3_Receive_BASO_CNT = 17
Public Const KF3_Receive_LUC_CNT = 18

Public Const KF3_Receive_NEUT_PCT = 19
Public Const KF3_Receive_LYMP_PCT = 20
Public Const KF3_Receive_MONO_PCT = 21
Public Const KF3_Receive_EOS_PCT = 22
Public Const KF3_Receive_BASO_PCT = 23
Public Const KF3_Receive_LUC_PCT = 24

Public Const KF3_Receive_LI = 25
Public Const KF3_Receive_MPXI = 26
Public Const KF3_Receive_RBC_FLAGS = 27
Public Const KF3_Receive_WBC_FLAGS = 28

Public Const KF3_Receive_ANISO = 29
Public Const KF3_Receive_MICRO = 30
Public Const KF3_Receive_MACRO = 31
Public Const KF3_Receive_VAR = 32
Public Const KF3_Receive_HYPO = 33
Public Const KF3_Receive_HYPER = 34
Public Const KF3_Receive_L_SHIFT = 35
Public Const KF3_Receive_ATYP = 36
Public Const KF3_Receive_BLASTS = 37
Public Const KF3_Receive_OTHER1 = 38
Public Const KF3_Receive_OTHER2 = 39

Type KF3_ReceiveRecord
    strRecord(KF3_Receive_SampleNo To KF3_Receive_OTHER2) As String
    intLength(KF3_Receive_SampleNo To KF3_Receive_OTHER2) As Integer
    intPosition(KF3_Receive_SampleNo To KF3_Receive_OTHER2) As Integer
End Type

Type KF3_RcvData
    StrIdNumber             As String * 12
    WBC                     As Single
    RBC                     As Single
    HGB                     As Single
    HCT                     As Single
    MCV                     As Single
    MCH                     As Single
    MCHC                    As Single
    PLT                     As Single
    PctLYMPH                As Single
    PctMONO                 As Single
    PctNEUT                 As Single
    PctEO                   As Single
    PctBASO                 As Single
    uLLYMPH                 As Single
    uLMONO                  As Single
    uLNEUT                  As Single
    uLEO                    As Single
    uLBASO                  As Single
    RDW_CV                  As Single
    RDW_SD                  As Single
    PDW                     As Single
    MPV                     As Single
    P_LCR                   As Single
End Type

'---------------------------------------------------------------------- KF3_ Defiem end
'register 등록 조회용 function
Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Type testRecord ' 10 byte 사용자 정의 형식
'    s1 As String * 5
'    s2 As String * 10
'    s3 As String * 10
'End Type


Sub FormCenter(f)
    f.Left = (12120 - f.Width) / 2
    f.Top = (9120 - f.Height) / 3
    
End Sub

Public Function convLabnoToExpand(ByVal sComp5 As String) As String
    
    convLabnoToExpand = Format(DateAdd("d", Val(sComp5), "2000-10-01"), "YYYYMMDD")
    
End Function

Public Function convLabnoToComp(ByVal sYear8 As String) As String
    Dim sconvYear       As String
    
    sconvYear = Left(sYear8, 4) & "-" & Mid(sYear8, 5, 2) & "-" & Mid(sYear8, 7)
    
    convLabnoToComp = Format(DateDiff("d", "2000-10-01", sconvYear), "00000")


End Function


Public Sub SysDate_Get()
    
    strSQL = ""
    strSQL = "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD') SYS FROM DUAL"
    Result = adoSQL(strSQL)
    If Result = 0 Then GstrSysDate = AdoGetString(Rs, "Sys", 0)
    AdoCloseSet Rs

End Sub


Sub DELAY(Sec)
    Dim T1, T2, Dummy
    Dummy = DoEvents()
    T1 = Timer + Sec
    Do
        T2 = Timer
    Loop Until T2 > T1

End Sub


Sub SS_INIT(sctl As Control, C1, R1)
    Dim i, j

    For i = 1 To sctl.DataRowCnt
        For j = 1 To sctl.MaxCols
            sctl.Col = j
            sctl.Row = i
            sctl.ForeColor = RGB(0, 0, 0)
        Next j
    Next i

    For i = 1 To 500 'sctl.DataRowCnt
        For j = 1 To 6
            sctl.Col = j
            sctl.Row = i
            sctl.BackColor = RGB(217, 253, 196)
        Next j
    Next i
    
    sctl.BlockMode = True
    sctl.Col = C1: sctl.Col2 = sctl.MaxCols
    sctl.Row = R1: sctl.Row2 = sctl.DataRowCnt
    sctl.Action = SS_ACTION_CLEAR_TEXT
    sctl.BlockMode = False
    sctl.TopRow = 1
    sctl.Col = C1: sctl.Row = R1
    sctl.Action = SS_ACTION_ACTIVE_CELL

End Sub


'Function ChecksumTx(ByVal Cstr1 As String) As String
Function S_itemcd(ByVal Sitemcd As String) As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT CODEKY "
    strSQL = strSQL & "   FROM TWEXAM_ITEMML "                  ' 고객 MASTER
'    strSQL = strSQL & "  WHERE GEOMJAN1 = '10' "
    strSQL = strSQL & "  WHERE GEOMJAN1 = '" & GGCODE & "' "
    strSQL = strSQL & "    AND GEOMJAN3 = '" & Sitemcd & "' "
    
    Result = AdoOpenSet(Rs, strSQL)
    
    If Result Then
        Do While Not Rs.EOF
            S_itemcd = Trim$(Rs.Fields("CODEKY")) & ""
            Rs.MoveNext
        Loop
    End If

End Function


Function NameSearch(PtnoS As String)
        
    strSQL = ""
    strSQL = strSQL & " SELECT SNAME "
    strSQL = strSQL & "   FROM TWBAS_PATIENT "                  ' 고객 MASTER
    strSQL = strSQL & "  WHERE PTNO = '" & PtnoS & "'"
    
    Result = AdoOpenSet(Rs, strSQL)
    
    If Result Then
        Do While Not Rs.EOF
            NameSearch = Trim$(Rs.Fields("SNAME")) & ""
            Rs.MoveNext
        Loop
    End If

End Function

Function PTNOSearch(PtnoS As String)

    Dim Bdt
    Dim Bno1
    Dim Bno2
    
    Bdt = convLabnoToExpand(Mid(PtnoS, 1, 5))
    Bno1 = Mid(PtnoS, 6, 2)
    Bno2 = Mid(PtnoS, 8, 5)
    
    strSQL = ""
    strSQL = strSQL & " SELECT PTNO "
    strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB "                  ' 고객 MASTER
    strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bdt & "','YYYY-MM-DD')"
    strSQL = strSQL & "    AND SLIPNO1 =   '" & Bno1 & "'"        ' 일련번호
    strSQL = strSQL & "    AND SLIPNO2 =   '" & Bno2 & "'"        ' 일련번호
    
    Result = AdoOpenSet(Rs, strSQL)
    
    If Result Then
        Do While Not Rs.EOF
'            Rs.MoveLast
            PTNOSearch = Trim$(Rs.Fields("ptno")) & ""
            Rs.MoveLast
            Rs.MoveNext      'check   1건만 가져오게
            
        Loop
    End If
    
End Function

Function RSaveRecord(Rrecord As String)
    RSaveRecord = "Rx " & Format(Time, "hh:mm:ss") & " ]  " & Mid$(Rrecord, 1, (Len(Rrecord) - 2))
    
End Function

Function TSaveRecord(Trecord As String)
    If Len(Trecord) > 1 Then
        TSaveRecord = "Tx " & Format(Time, "hh:mm:ss") & " ]  " & Mid$(Trecord, 1, (Len(Trecord) - 2))
    Else
        TSaveRecord = "Tx " & Format(Time, "hh:mm:ss") & " ]  " & Trecord
    End If

End Function


Function ExtractTime(chrsin As String)
     
    If InStr(chrsin, "\") Then              'check to see if a forward slash exists
       For idx = Len(chrsin) To 1 Step -1   'step though until full name is extracted
           If Mid(chrsin, idx, 1) = "\" Then
              chrsout = Mid(chrsin, idx + 1)
              Exit For
           End If
       Next idx
    ElseIf InStr(chrsin, ":") = 2 Then      'otherwise, check to see if a colon exists
       chrsout = Mid(chrsin, 3)             'if so, return the filename
    Else
       chrsout = chrsin                     'otherwise, return the original string
    End If
    
    ExtractTime = chrsout & "  " & FileDateTime(chrsin)         'return the filename to the user
   
End Function


Public Function Checksum_Eci_Rx(ByVal strPrmValue As String)

    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    
    intValueLength = LenA(strPrmValue)
    
    For i = 2 To intValueLength - 4
        intCheck = intCheck + Asc(Mid(strPrmValue, i, 1))
    Next
    
    strCheck = Hex(intCheck)
    
    If Len(strCheck) = 1 Then
        Checksum_Eci_Rx = "0" & strCheck
    Else
        Checksum_Eci_Rx = Right(strCheck, 2)
    End If

End Function

Public Function CheckSum_ECi_Tx(ByVal strPrmValue As String)

    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    
    intValueLength = LenA(strPrmValue)
    
    For i = 1 To intValueLength
        intCheck = intCheck + Asc(Mid(strPrmValue, i, 1))
    Next
    
    strCheck = Hex(intCheck)
    
    If Len(strCheck) = 1 Then
        CheckSum_ECi_Tx = "0" & strCheck
    Else
        CheckSum_ECi_Tx = Right(strCheck, 2)
    End If

End Function



Public Function CheckSum_KF3_(ByVal strPrmValue As String)

    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    intValueLength = LenA(strPrmValue)
'    If intValueLength < 3 Then
'        KF3_VerifyCheckSum = False
'        Exit Function
'    End If
    
    For i = 2 To intValueLength - 6
        intCheck = intCheck Xor Asc(Mid(strPrmValue, i, 1))
    Next
    
    intCheck = intCheck And &HFF&
    If Hex(intCheck) = &H3 Then intCheck = &H7F
    
    CheckSum_KF3_ = intCheck

End Function


Public Function CheckSum_Ax2(ByVal strPrmValue As String)

    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    intValueLength = LenA(strPrmValue)
    
    For i = 2 To intValueLength - 6
        intCheck = intCheck + Asc(Mid(strPrmValue, i, 1))
    Next
    
    intCheck = intCheck And &HFF&

    CheckSum_Ax2 = Right("00" & Hex(intCheck), 2) 'intCheck

End Function




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   sta compact용
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ChecksumRx(CheckStr As String)
    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As String 'Integer
    Dim strCheck            As String

    intCheck = 0
    intValueLength = LenA(CheckStr)
    
    For i = 2 To intValueLength - 4
        intCheck = intCheck + Asc(Mid(CheckStr, i, 1))
    Next
    
    intCheck = Hex(intCheck)
    
    ChecksumRx = Right(intCheck, 2)


End Function


Function ChecksumTx(ByVal Cstr1 As String) As String
'STA COMPACT
    
    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As String 'Integer
    Dim strCheck            As String

    intCheck = 0
    intValueLength = LenA(Cstr1)
    
    For i = 1 To intValueLength '- 4
        intCheck = intCheck + Asc(Mid(Cstr1, i, 1))
    Next
    
    intCheck = Hex(intCheck)
    
    ChecksumTx = Right(intCheck, 2)

End Function


Public Function LenA(strPrmString As String) As Integer

    Dim i                   As Integer
    Dim intStrLen           As Integer
    Dim intAnsiStrLen       As Integer
    Dim strTemp             As String
    
    intStrLen = Len(strPrmString)
    For i = 1 To intStrLen
        strTemp = Mid(strPrmString, i, 1)
        
        Select Case AscW(strTemp)
        Case 0 To 255
            intAnsiStrLen = intAnsiStrLen + 1
        
        Case Else
            intAnsiStrLen = intAnsiStrLen + 2
        
        End Select
    Next
    
    LenA = intAnsiStrLen

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Axsym Data Control Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function RetPos_X(str As String, seq As Integer) As Integer

Dim i, cnt As Integer

    cnt = 0
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) = "^" Then cnt = cnt + 1
        If cnt = seq Then
            RetPos_X = i
            Exit For
        End If
    Next i

End Function

Function RetPos(str As String, seq As Integer) As Integer

Dim i, cnt As Integer

    cnt = 0
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) = "|" Then cnt = cnt + 1
        If cnt = seq Then
            RetPos = i
            Exit For
        End If
    Next i

End Function

Function GetData(str As String, seq As Integer) As String

Dim i, cnt As Integer

   GetData = Mid(str, RetPos(str, seq) + 1, RetPos(str, seq + 1) - RetPos(str, seq) - 1)
    
End Function

Function CheckSum(str As String) As String
   
     Dim i, CS As Integer
   
   
     CS = 0
     For i = 1 To Len(str)
       CS = CS + Asc(Mid(str, i, 1))
     Next i
     CheckSum = Right(Hex(CS + 16), 2)
End Function


Function MakeH() As String
'Header
Dim Part As String
    Part = H_FRAME
    
    MOD_8 = MOD_8 + 1
    Part = (MOD_8 Mod 8) & Part
    MakeH = Chr(2) & Part & Chr(13) & Chr(3) & CheckSum(Part) & vbCrLf
End Function
Function MakeP(SID As String) As String
'Patient Information
Dim Part As String
    
    Part = P_FRAME
    P_NO = P_NO + 1
    O_NO = 0
    MOD_8 = MOD_8 + 1
    Part = Mid(Part, 1, RetPos(Part, 3)) & SID & Mid(Part, RetPos(Part, 4), 100)
    Part = (MOD_8 Mod 8) & Mid(Part, 1, RetPos(Part, 1)) & P_NO & Mid(Part, RetPos(Part, 1), 100)
    
    MakeP = Chr(2) & Part & Chr(13) & Chr(3) & CheckSum(Part) & vbCrLf

End Function
Function MakeO(SID As String, ANO As String) As String
'Order
Dim Part As String

    Part = O_FRAME
    O_NO = O_NO + 1
    MOD_8 = MOD_8 + 1
    Part = Mid(Part, 1, RetPos(Part, 2)) & SID & Mid(Part, RetPos(Part, 3), 100)
    Part = Mid(Part, 1, RetPos(Part, 4) + 3) & ANO & Mid(Part, RetPos(Part, 4) + 4, 100)
    Part = (MOD_8 Mod 8) & Mid(Part, 1, RetPos(Part, 1)) & O_NO & Mid(Part, RetPos(Part, 1) + 1, 100)
    MakeO = Chr(2) & Part & Chr(13) & Chr(3) & CheckSum(Part) & vbCrLf
    
End Function


Function MakeL() As String
'Message Termination
Dim Part As String
    Part = L_FRAME
    
    MOD_8 = MOD_8 + 1
    Part = (MOD_8 Mod 8) & Part
    MakeL = Chr(2) & Part & Chr(13) & Chr(3) & CheckSum(Part) & vbCrLf

End Function



Function QGetSID(str As String) As String
'Query 접수번호
Dim Part As String
'<STX>2Q|1|^301*99090401317<CR><ETX>13<CR><LF>
  
'  Part = Mid(str, RetPos(str, 3) + 1, 100)
  
  Part = Mid(str, RetPos_X(str, 1) + 1, 100)
  QGetSID = Mid(Part, 1, InStr(1, Part, "|", 0) - 1)
  
  If Len(QGetSID) <> 12 Then QGetSID = ""

End Function

Function GetSID(str As String) As String
'Result 접수번호
Dim Part As String

  Part = Mid(str, RetPos(str, 3) + 1, 100)
  GetSID = Mid(Part, 1, InStr(1, Part, "^", 0) - 1)

  If Len(GetSID) <> 12 Then GetSID = ""

End Function
Function GetANO(str As String) As String
'검사번호
Dim Part As String
  Part = Mid(str, RetPos(str, 4) + 4, 100)
  GetANO = Mid(Part, 1, InStr(1, Part, "^", 0) - 1)

End Function
Function GetResult(str As String) As String
 
  GetResult = Mid(str, RetPos(str, 3) + 1, RetPos(str, 4) - RetPos(str, 3) - 1)

End Function



''
'' KF3_ section
''
''

Public Sub KF3_InitReceiveLength(udPrmReceive As KF3_ReceiveRecord)

    Const CNT_1st = 50
    Const CNT_2nd = 54
    Const CNT_3rd = 30
    Const CNT_4th = 43
    Const CNT_5th = 43
    Const CNT_6th = 25
    Const CNT_other = 57 + 22

    Dim i As Integer

    ' 1st line
    udPrmReceive.intLength(KF3_Receive_SampleNo) = 13
    udPrmReceive.intPosition(KF3_Receive_SampleNo) = 31
    
    ' 2nd line
    udPrmReceive.intLength(KF3_Receive_WBC) = 5
    udPrmReceive.intPosition(KF3_Receive_WBC) = CNT_1st
    
    For i = KF3_Receive_RBC To KF3_Receive_HDW
        udPrmReceive.intLength(i) = 5
        udPrmReceive.intPosition(i) = udPrmReceive.intPosition(i - 1) + udPrmReceive.intLength(i - 1) + 2
    Next
    
    ' 3rd line
    udPrmReceive.intLength(KF3_Receive_PLT) = 5
    udPrmReceive.intPosition(KF3_Receive_PLT) = 1 + CNT_1st + CNT_2nd + 9
    For i = KF3_Receive_MPV To KF3_Receive_PDW
        udPrmReceive.intLength(i) = 5
        udPrmReceive.intPosition(i) = udPrmReceive.intPosition(i - 1) + udPrmReceive.intLength(i - 1) + 2
    Next
    
    ' 4th line
    udPrmReceive.intLength(KF3_Receive_NEUT_CNT) = 5
    udPrmReceive.intPosition(KF3_Receive_NEUT_CNT) = 1 + CNT_1st + CNT_2nd + CNT_3rd + 8
    For i = KF3_Receive_LYMP_CNT To KF3_Receive_LUC_CNT
        udPrmReceive.intLength(i) = 5
        udPrmReceive.intPosition(i) = udPrmReceive.intPosition(i - 1) + udPrmReceive.intLength(i - 1) + 2
    Next
    
    ' 5th line
    udPrmReceive.intLength(KF3_Receive_NEUT_PCT) = 5
    udPrmReceive.intPosition(KF3_Receive_NEUT_PCT) = 1 + CNT_1st + CNT_2nd + CNT_3rd + CNT_4th + 8
    For i = KF3_Receive_LYMP_PCT To KF3_Receive_LUC_PCT
        udPrmReceive.intLength(i) = 5
        udPrmReceive.intPosition(i) = udPrmReceive.intPosition(i - 1) + udPrmReceive.intLength(i - 1) + 2
    Next
    
    ' 6th line
    udPrmReceive.intLength(KF3_Receive_LI) = 5
    udPrmReceive.intPosition(KF3_Receive_LI) = 1 + CNT_1st + CNT_2nd + CNT_3rd + CNT_4th + CNT_5th + 8
    For i = KF3_Receive_MPXI To KF3_Receive_MPXI
        udPrmReceive.intLength(i) = 5
        udPrmReceive.intPosition(i) = udPrmReceive.intPosition(i - 1) + udPrmReceive.intLength(i - 1) + 2
    Next
    
    udPrmReceive.intLength(KF3_Receive_RBC_FLAGS) = 4
    udPrmReceive.intPosition(KF3_Receive_RBC_FLAGS) = udPrmReceive.intPosition(KF3_Receive_RBC_FLAGS - 1) + udPrmReceive.intLength(KF3_Receive_RBC_FLAGS - 1) + 2
    
    udPrmReceive.intLength(KF3_Receive_WBC_FLAGS) = 4
    udPrmReceive.intPosition(KF3_Receive_WBC_FLAGS) = udPrmReceive.intPosition(KF3_Receive_WBC_FLAGS - 1) + udPrmReceive.intLength(KF3_Receive_WBC_FLAGS - 1) + 1
    
    ' 7th line
    udPrmReceive.intLength(KF3_Receive_ANISO) = 4
    udPrmReceive.intPosition(KF3_Receive_ANISO) = 1 + CNT_1st + CNT_2nd + CNT_3rd + CNT_4th + CNT_5th + CNT_6th + 8
    For i = KF3_Receive_MICRO To KF3_Receive_OTHER2
        udPrmReceive.intLength(i) = 4
        udPrmReceive.intPosition(i) = udPrmReceive.intPosition(i - 1) + udPrmReceive.intLength(i - 1) + 1
    Next
    
End Sub

Public Sub KF3_ReadReceiveBuffer(strPrmReceive As String, udPrmReceive As KF3_ReceiveRecord)
'   STX + KF3_ReceiveRecord + ETX

    Dim i As Integer, strBuffer As String

    ' KF3_ReceiveRecord
    strBuffer = MidA(strPrmReceive, 2, LenA(strPrmReceive) - 2)
    For i = KF3_Receive_SampleNo To KF3_Receive_OTHER2
        udPrmReceive.strRecord(i) = MidA(strBuffer, udPrmReceive.intPosition(i), udPrmReceive.intLength(i))
    Next

End Sub


Public Function MidA(strPrmString As String, intPrmStart As Integer, Optional intPrmLength As Integer) As String

    Dim i As Integer, intArrPos As Integer
    Dim intUniStrLen As Integer, intAnsiStrLen As Integer
    Dim intAnsiStart As Integer, intStrCatCnt As Integer
    Dim bytTempString() As Byte
    Dim strReturnValue As String, strTempChar As String
    Dim blnUnicode As Boolean
     
    If intPrmLength < 0 Then
        Exit Function
    End If
     
    intUniStrLen = Len(strPrmString)
    intAnsiStrLen = LenA(strPrmString)
    
    If intUniStrLen = intAnsiStrLen Then
        MidA = Mid(strPrmString, intPrmStart, intPrmLength)
        Exit Function
    End If
    
    If intPrmLength = Empty Then
        intPrmLength = intAnsiStrLen - intPrmStart
    End If
    
    bytTempString = strPrmString
    
    For i = 1 To intUniStrLen
        intArrPos = (i - 1) * 2
        
        If Hex(bytTempString(intArrPos + 1)) = "0" Then
            intAnsiStart = intAnsiStart + 1
            blnUnicode = False
        Else
            intAnsiStart = intAnsiStart + 2
            blnUnicode = True
        End If
        
        If intAnsiStart >= intPrmStart Then
            If intStrCatCnt >= intPrmLength Then Exit For
            
            intStrCatCnt = intStrCatCnt + 1
            
            If (strReturnValue = "" And intAnsiStart + 1 = intPrmStart) Or (intAnsiStart = intPrmStart) Then
                If blnUnicode Then
                    strTempChar = "_"
                Else
                    'strTempChar = ChrW("&H" & Hex(bytTempString(intArrPos + 1)) & Hex(bytTempString(intArrPos)))
                    strTempChar = Mid(strPrmString, i, 1)
                End If
            Else
                If blnUnicode Then
                    intStrCatCnt = intStrCatCnt + 1
                End If
                
                'strTempChar = ChrW("&H" & Hex(bytTempString(intArrPos + 1)) & Hex(bytTempString(intArrPos)))
                strTempChar = Mid(strPrmString, i, 1)
            End If
                
            strReturnValue = strReturnValue & strTempChar
        End If
    Next
        
    If intPrmLength < LenA(strReturnValue) Then
        MidA = Left(strReturnValue, Len(strReturnValue) - 1) & " "
    Else
        MidA = strReturnValue
    End If
    
End Function

