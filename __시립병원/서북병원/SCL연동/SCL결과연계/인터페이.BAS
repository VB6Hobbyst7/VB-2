Attribute VB_Name = "modInterface"
Option Explicit

Type HeaderRec
    HeadTyp As String
    HeadDel As String
    HeadSID As String
    HeadRev As String
    HeadPro As String
    HeadVer As String
End Type

Type PatientRec
    PatTypNam As String
    PatSeqNum As String
    PatPraSng As String
    PatPatNam As String
    PatBidDte As String
    PatSexNam As String
    PatPhyId As String
    PatLocNam As String
End Type

Type OrderRec
    OdrTypNam As String
    OdrSeqNum As String
    OdrSpcID As String
    OdrSamID As String
    OdrTryCup As String
    OdrUniTid As String
    OdrUniFir As String
    OdrUniSed As String
    OdrUniThr As String
    OdrUniFor As String
    OdrPriCod As String
    OdrActCod As String
    OdrRepTyp As String
End Type

Type ResultRec
    ResTypNam As String
    ResSeqNum As String
    ResUniTid As String
    ResDatVal As String
    ResDatUni As String
    ResAdnFlg As String
    ResRstSat As String
    ResRstDte As String
End Type

Type TerminationRec
    TerRecTyp As String
    TerSeqNum As String
    TerTerCod As String
End Type

Type RecordRec
    Header As HeaderRec
    Patient As PatientRec
    Order As OrderRec
    Result As ResultRec
    Termination As TerminationRec
End Type

Type SelMchType
    Buf         As String
    LabMchMst   As LabMchMstRec
    LabSelYon   As String
End Type

Public Const MAX_NUM_OF_LAB_MCH = 100

Public tGblSelMch(1 To MAX_NUM_OF_LAB_MCH) As SelMchType

Global LoopAgain As Integer
Global ComPortOpened As Integer
Global sGblTstDte As String
Global sSeqNum As String

Public iGblNumOfLabMchSelMnu As Integer
Public bGblBeInp As Boolean '선접수 인터페이스 여부

Global NumOfResultText As Integer
Global NumOfMultiResultText As Integer
Global NumOfResultCombo As Integer
Global NumOfTotalControls As Integer
Global NumOfLabel As Integer

Global Const UNIT_HEIGHT = 400
Global Const LABEL_WIDTH = 1800
Global Const TEXT_WIDTH = 1200
Global Const MARGIN_BETWEEN_COLUMNS = 400

Type ControlType
    RltTyp As String
    LabCod As String
    ResultControlIndex As Integer
    SclCodeComboIndex As Integer
End Type
Global Const MAX_TOTAL_CONTROLS = 9999
Global tGblControls(1 To MAX_TOTAL_CONTROLS) As ControlType

Public Sub CloseComPorts()

'    Dim i As Integer
'
'    On Error GoTo ErrorHandler2
'
'    For i = 1 To iGblNumOfLabMchSelMnu
'        If frmInterface.Com(i).PortOpen = True Then
'            frmInterface.Com(i).PortOpen = False
'        End If
'    Next i
'
'    If frmInterface.CommPrn.PortOpen = True Then
'        frmInterface.CommPrn.PortOpen = False
'    End If
'
'    Exit Sub
'
'ErrorHandler2:
'    Select Case err
'        Case 68
'            Resume Next
'        Case Else
'            Call MsgBox("[통신오류] " & err.Description & Chr(13) & Chr(13) & "▣ 포트 : " & frmInterface.Com(i).CommPort & Chr(13) & "▣ 설정 : " & frmInterface.Com(i).Settings, vbInformation)
'            'End
'            Resume Next
'    End Select

End Sub

Function checkchr(ByVal stest As String) As String
    Dim i As Integer
    Dim sSum As Integer
        For i = 1 To Len(stest)
        
            sSum = sSum + Asc(Mid(stest, i, 1))
            
        Next i
    checkchr = Right(Hex(Str(sSum) + 16), 2)
End Function

Public Sub StartCom()

'2004/04/06 이상은 - 주석처리
'    Dim i As Integer
'    Dim ComIndex As Integer
'    Dim Dummy As Integer
'    Dim OldBackColor As Long
'    Dim sBuf As String
'    Dim Buf As String
'
'    '2002.01.17 학부접수에 사용.
'    Dim sMchCod As String
'    Dim sAcpCod As String
'    Dim sAnswer As String
'    Dim i1 As Integer
'    Dim exec As String
'    Dim ststBuf As String
'    Dim ststBuf1 As String
'    Dim ststBuf2 As String
'    Dim ststBuf3 As String
'    Dim stest As String
'    Dim sC1C2 As String
'    Dim iRow As Integer
''    Dim sss As String
''    Dim s7170OrderString  As String
'
'    Do
'        Dummy = DoEvents()
'        If (LoopAgain = False) Then
'            Exit Do
'        End If
'
'        For i = 1 To iGblNumOfLabMchSelMnu
'
'            sGblPartErr = "N"                                                               '학부접수에서 사용하는 파라메터값.
'
'            If (frmInterface.Com(i).InBufferCount > 0) Then                                '테스트중 2002.01.16 sebal
'
'
'                sBuf = frmInterface.Com(i).Input
'
'                tGblSelMch(i).Buf = tGblSelMch(i).Buf & sBuf
'
'                Select Case tGblSelMch(i).LabMchMst.LabMchCod
'
'                Case "CD1700"
'                    Call CD_1700(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'                Case "MIDM"
'                    frmInterface.Com(i).Output = Chr(&H2) & Chr(&H3E) & Chr(&H3) & "  " & Chr(CR)
'                    Call Miditron_M(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'                Case "ELECSYS"
'                    frmInterface.Com(i).Output = Chr(ACK)
'                    Call ELECSYS_1010(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'                Case "PENTRA"
'                    Call PENTRA60(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'                Case "HP800"
'                    frmInterface.Com(i).Output = Chr(ACK)
'                    Call Roche_HITACHI_800(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'                End Select
'
'
'                '2002/11/06 EverSky  양방향 인터페이스
''                Case "7170A"
''
''                    Call GetBetween(tGblSelMch(i).Buf, ststBuf, Chr(2), Chr(3))
''
''                    Call H7170A(ststBuf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''
''                '2002/11/06 EverSky  양방향 인터페이스
''                Case "STAGO"
''
''                    sBuf = frmInterface.Com(i).Input                                           '테스트중 2002.01.16 sebal
''
''
''                    tGblSelMch(i).Buf = tGblSelMch(i).Buf & sBuf
''
''                    Call GetBetween(tGblSelMch(i).Buf, ststBuf, Chr(2), Chr(10))
''
''                    If Left(ststBuf, 2) = "1H" Then
''                        ststBuf1 = Chr(2) & ststBuf & Chr(10)
''                    End If
''
''                    If Left(ststBuf, 2) = "3L" Then
''                        ststBuf3 = Chr(2) & "4" & Left(ststBuf, 2) & Chr(10)
''                    End If
''
''                    frmInterface.Com(i).Output = Chr(ACK)
'''                    Debug.Print ststBuf
''
''                    If Left(ststBuf, 2) = "3L" Then
''                        frmInterface.Com(i).Output = ststBuf1
''                        stest = "2P|1|||^^^"
''                        sC1C2 = checkchr(stest)
''                        frmInterface.Com(i).Output = Chr(2) & stest & vbCr & Chr(3) & sC1C2 & vbCrLf
''                        '오더부분
''                        stest = "3O|1|200200000158||^^^1" & Chr(92) & "^^^2|R"
''                        sC1C2 = checkchr(stest)
''                        frmInterface.Com(i).Output = Chr(2) & stest & vbCr & Chr(3) & sC1C2 & vbCrLf
''
''                        frmInterface.Com(i).Output = ststBuf
''                    End If
'
'
''                Case "CD1700"   '20021107 lek Add For CEL-DYN 1700
''
''                    'Debug.Print "at StartCom Case CD1700"
''                    Call CD_1700(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
''                Case "CD3700"
''                    'Debug.Print "at StartCom Case CD3700"
''                    Call CD_3700(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''
''                Case "NOVA-CRT" '20021109 lek Add For NOVA Electrolyte
''
''                    'Debug.Print "at StartCom Case NOVA-CRT"
''                    Call NOVA_ELEC(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''
''                Case "NOVA-SPM" '20021109 lek Add
''
''                    'Debug.Print "at StartCom Case NOVA-SPM"
''                    Debug.Print tGblSelMch(i).Buf
''
''                    'Call Nova_SPM(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CLINITEK"
''
''                    sBuf = frmInterface.Com(i).Input                                           '테스트중 2002.01.16 sebal
''
''
''                    tGblSelMch(i).Buf = tGblSelMch(i).Buf & sBuf
''
''                    Call GetBetween(tGblSelMch(i).Buf, ststBuf, Chr(2), Chr(3))
''
''                    If ststBuf <> "" Then
''                        Call CLINITEK(ststBuf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                    End If
'
'                '2003.01.27 대구성삼병원
'
'            End If
'        Next i
'
'    Loop
'                    'Call H7170A(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'
'
''                Case "RL9"
''                    Call RL9(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "EPX"
''                    Call EPX(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "C100"
''                    Call Clinitek100(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                'Case "C200"
''                '    Call Clinitek200(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "H7150"
''                    Call Hitachi7150(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "H7050"
''                    Call Hitachi7050(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "K1000"
''                    Call K1000(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "K4500"
''                    Call K4500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "NOVA-SP3"
''                    Call NovaSP3(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CC288"
''                    Call CibaCorning288(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "F800"
''                    Call F800(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "MINOS"
''                    Call Minos(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "ATB"
''                    Call Atb(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "DADE"
''                    Call Dade_Dimension(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "US2100"
''                    Call US_2100(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "GAS348"
''                    Call GAS_348(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "S300"
''                    Call Scan_300(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "AVL9130"
''                    Call AVL_9130(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "JUNIOR"
''                    Call Miditron_Junior(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "IMX"
''                    Call Imx(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CCX"
''                    Call CCX(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CA500"
''                    Call CA500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "SE9000"
''                    Call SE9000(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "VITROS"
''                    Call VITROS(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "RAPIDO"
''                    Call RAPIDO(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "TBA20R"
''                    frmInterface.CommPrn.Output = sBuf
''                    Call TBA_20R(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "MIDITRON"
''                    Call Miditron_M(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "VITEK"
''                    Call VITEK(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "NOVA-E"
''                    Call NOVA_ELEC(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                 '시화병원
''                Case "VITROS_ECI"
''                    Call VITROS_ECI(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "TBA20FR"
''                    Call TBA_20FR(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "ELECSYS"
''                    Call ELECSYS_1010(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CHIRON248"
''                    Call CHIRON_248(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CHIRON644"
''                    Call CHIRON_644(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "MEDONIC"
''                    Call MEDONIC_CA530(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "Lab"
''                    Call Lab(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                '김해중앙
''                Case "CC644"
''                    Call CibaCorning644(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                'Case "AXSYM"
''                '    Call AxSym(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CELL"
''                    Call Cell_Dyn_1600(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "ACS180"
''                    Call Acs(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
''                Case "AVL"
''                    Call AVL(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''
''                '2001년 03월 14일 개발과 오성연
''                '처음으로 모듈을 만들다....(김해에서)
''                Case "COULTER"
''                    Call COULTER(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''
''                '###############################################################################
''                '상주적십자
''                '###############################################################################
''                Case "AXSYM"
''                    Call AxSym(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "NOVA-SPM"
''                    Call Nova_SPM(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CELLDYN"
''                    Call Cell_Dyn3500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "URISCAN-PRO"
''                    Call URISCAN_PRO(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "REFLOTRON"
''                    Call REFLOTRON(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "TBA-80FR"
''                    Call TBA_80FR(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                '###############################################################################
''
''                Case "EASYBLOOD"
''                    Call EasyBld(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'
'
'''
'''                'Bestian 병원
'''                '02.04.21 일요일 sebal
'''                Case "NOVA5"
'''                    Call NOVA_ELEC(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''
'''
'''                '02.04.24 sebal
'''                Case "CD1700"
'''                    Call CD_1700(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''
'''                '의왕성모병원
'''                Case "CLINITEK500"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call Clinitek_500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "VITROS_ECI"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call VITROS_ECI(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "VITROS_750"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call VITROS_750(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "ACL100"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call ACL_100(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "ADVIA120"
'''                    'DoEvents
'''                    'DoEvents
'''                    'DoEvents
'''                    Call PAUSE(1)
'''                    Call ADVIA_120(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    Call PAUSE(1)
'''                    'DoEvents
'''                    'DoEvents
'''                    'DoEvents
'''
'''                Case "RAPID_LAB"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call RAPIDO(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                End Select
'''
'''                If frmInterface.lvwMch.ListItems(i).Selected = True Then
'''                    frmInterface.lvwMch.ListItems(i).Icon = "RUN"
'''                Else
'''                    frmInterface.lvwMch.ListItems(i).Icon = "MCH" & CStr(i)
'''                End If
'''
'''            End If                                         '테스트중 2002.01.16 sebal
'''
'''            '2002.01.17 sebal
'''            '학부접수장비중, 학부접수 하지 않고, 바로 검사를 하면, RsbInf.RsbIntYon = "N" 이기
'''            '때문에 ResInf에 자동으로 저장 할때 접수번호를 가져 오지 못한다. (학부접수를 하면
'''            'RsbInf.RsbIntYon이 "P"로 저장 된다) 그리하야, 학부접수를 한다고 setting 해 놓고
'''            '학부 접수하지 않고 바로 검사하는 장비에 대해서 아래의 메세지 뿌려준다.
'''            'sGblPartErr = "N" : P로 저장 되어정상.
'''            'sGblPartErr = "Y" : 학부접수 하지않아 접수번호 가져 오지 못함.
'''
'''            If sGblPartErr = "Y" Then
'''                sMchCod = tGblSelMch(i).LabMchMst.LabMchCod
'''                sAcpCod = MasterHelpDetail("LabMchMst", sMchCod & Chr(5), sMchCod & Chr(5), 12)
'''                sAnswer = MsgBox(sMchCod & " 장비의 " & sAcpCod & " 학부접수 후 재 검사 하세요")
'''                sGblPartErr = "N"
'''            End If
'''
'''        Next i
'''
'''    Loop
'
'    Close #1
    
End Sub

''corebrain 2003-05-23
''성삼병원 Interface작업으로 인해서 첨단병원 Interface 모듈주석처리함.

'Public Sub StartCom()
'
'    Dim i As Integer
'    Dim ComIndex As Integer
'    Dim Dummy As Integer
'    Dim OldBackColor As Long
'    Dim sBuf As String, Buf As String
'
'    '2002.01.17 학부접수에 사용.
'    Dim sMchCod As String
'    Dim sAcpCod As String
'    Dim sAnswer As String
'    Dim i1 As Integer
'    Dim exec As String
'    Dim ststBuf As String
'    Dim ststBuf1 As String
'    Dim ststBuf2 As String
'    Dim ststBuf3 As String
'    Dim stest As String
'    Dim sC1C2 As String
'
'
'
'
'    Dim sBarCode As String
'    Dim sOrderRes As String
'    Dim iTmp As Integer
''    Dim sss As String
''    Dim s7170OrderString  As String
'
'    Do
'        Dummy = DoEvents()
'
'        If (LoopAgain = False) Then
'            Exit Do
'        End If
'
'        For i = 1 To iGblNumOfLabMchSelMnu
'
'            sGblPartErr = "N"                                                               '학부접수에서 사용하는 파라메터값.
'
'            If (frmInterface.Com(i).InBufferCount > 0) Then                                '테스트중 2002.01.16 sebal
'
'                'frmInterface.lvwMch.ListItems(i).Icon = "COMM"
'
'                sBuf = frmInterface.Com(i).Input                                           '테스트중 2002.01.16 sebal
'                tGblSelMch(i).Buf = tGblSelMch(i).Buf & sBuf
''                Debug.Print "1===" & tGblSelMch(i).Buf
'
'                Select Case tGblSelMch(i).LabMchMst.LabMchCod
'
'                '2002/11/06 EverSky  양방향 인터페이스
'                Case "7170A"
'
'                    Call GetBetween(tGblSelMch(i).Buf, ststBuf, Chr(2), Chr(3))
'
'                    Call H7170A(ststBuf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'                '2002/11/06 EverSky  양방향 인터페이스
'                Case "STAGO"
'                    Select Case tGblSelMch(i).Buf
'
'                        Case Chr(EOT)    '통신 종료
'
'                            Select Case sOrderRes
'
'                                Case "Result"
'                                    Call LabResultSave
'                                    sBuf = ""
'                                    tGblSelMch(i).Buf = ""
'                                    sOrderRes = ""   ' 오더 부분 초기화
'                                    iTmp = 0
'                                Case "Order"
'                                    sBuf = ""
'                                    tGblSelMch(i).Buf = ""
'
'                                Case Else
'
'                                    sBuf = ""
'                                    tGblSelMch(i).Buf = ""
'                                    sOrderRes = ""   ' 오더 부분 초기화
'
'
'                            End Select
'
'
'                        Case Chr(ENQ)    '통신 시작
'                            frmInterface.Com(i).Output = Chr(ACK)
'                            tGblSelMch(i).Buf = ""
'                            sBuf = ""
'
'                        Case Else
'
''                            Debug.Print tGblSelMch(i).Buf & sOrderRes
'
'                            Call GetBetween(tGblSelMch(i).Buf, ststBuf, Chr(2), Chr(10))
'                                '오더 전달부분
'                            If tGblSelMch(i).Buf = Chr(ACK) And sOrderRes = "Order" Then
'                                iTmp = iTmp + 1
'                                Call OrderInput(iTmp, ststBuf1, sBarCode, i)
'                                If iTmp = 5 Then
'                                    sOrderRes = ""
'                                End If
'                            End If
'
'                            If sOrderRes = "Result" And Mid(ststBuf, 2, 1) = "R" Then
'                                '결과 전달부분
'                                iTmp = iTmp + 1
''                                Debug.Print sBarCode & "  " & iTmp & "    " & ststBuf
'                                Call ResultInput(iTmp, sBarCode, ststBuf)
'                                frmInterface.Com(i).Output = Chr(ACK)
'                            Else
'                                '결과 및 환자 바코드 인식부분
'                                Call OrderRes(ststBuf, sOrderRes, ststBuf1, sBarCode, iTmp, i)
'                            End If
'
'                    End Select
'
'
'                Case "CD1700"   '20021107 lek Add For CEL-DYN 1700
'
'                    'Debug.Print "at StartCom Case CD1700"
'                    Call CD_1700(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'                Case "CD3700"
'                    'Debug.Print "at StartCom Case CD3700"
'                    Call CD_3700(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'                Case "NOVA-CRT" '20021109 lek Add For NOVA Electrolyte
'
'                    'Debug.Print "at StartCom Case NOVA-CRT"
'                    Call NOVA_ELEC(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'                Case "NOVA-SPM" '20021109 lek Add
'
'                    'Debug.Print "at StartCom Case NOVA-SPM"
'
'                    frmInterface.Com(i).Output = Chr(ACK)
'
'                    Call Nova_SPM(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'                Case "CLINITEK"
'
'                    sBuf = frmInterface.Com(i).Input                                           '테스트중 2002.01.16 sebal
'                    tGblSelMch(i).Buf = tGblSelMch(i).Buf & sBuf
'
'                    Call GetBetween(tGblSelMch(i).Buf, ststBuf, Chr(2), Chr(3))
'
'                    If ststBuf <> "" Then
'                        Call CLINITEK(ststBuf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'                    End If
'                End Select
'            End If
'        Next i
'
'    Loop
'                    'Call H7170A(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'
'
''                Case "RL9"
''                    Call RL9(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "EPX"
''                    Call EPX(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "C100"
''                    Call Clinitek100(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                'Case "C200"
''                '    Call Clinitek200(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "H7150"
''                    Call Hitachi7150(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "H7050"
''                    Call Hitachi7050(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "K1000"
''                    Call K1000(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "K4500"
''                    Call K4500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "NOVA-SP3"
''                    Call NovaSP3(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CC288"
''                    Call CibaCorning288(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "F800"
''                    Call F800(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "MINOS"
''                    Call Minos(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "ATB"
''                    Call Atb(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "DADE"
''                    Call Dade_Dimension(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "US2100"
''                    Call US_2100(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "GAS348"
''                    Call GAS_348(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "S300"
''                    Call Scan_300(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "AVL9130"
''                    Call AVL_9130(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "JUNIOR"
''                    Call Miditron_Junior(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "IMX"
''                    Call Imx(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CCX"
''                    Call CCX(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CA500"
''                    Call CA500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "SE9000"
''                    Call SE9000(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "VITROS"
''                    Call VITROS(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "RAPIDO"
''                    Call RAPIDO(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "TBA20R"
''                    frmInterface.CommPrn.Output = sBuf
''                    Call TBA_20R(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "MIDITRON"
''                    Call Miditron_M(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "VITEK"
''                    Call VITEK(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "NOVA-E"
''                    Call NOVA_ELEC(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                 '시화병원
''                Case "VITROS_ECI"
''                    Call VITROS_ECI(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "TBA20FR"
''                    Call TBA_20FR(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "ELECSYS"
''                    Call ELECSYS_1010(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CHIRON248"
''                    Call CHIRON_248(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CHIRON644"
''                    Call CHIRON_644(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "MEDONIC"
''                    Call MEDONIC_CA530(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "Lab"
''                    Call Lab(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                '김해중앙
''                Case "CC644"
''                    Call CibaCorning644(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                'Case "AXSYM"
''                '    Call AxSym(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CELL"
''                    Call Cell_Dyn_1600(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "ACS180"
''                    Call Acs(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "H7060"
''                    Call Hitachi_7060(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "AVL"
''                    Call AVL(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''
''                '2001년 03월 14일 개발과 오성연
''                '처음으로 모듈을 만들다....(김해에서)
''                Case "COULTER"
''                    Call COULTER(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''
''                '###############################################################################
''                '상주적십자
''                '###############################################################################
''                Case "AXSYM"
''                    Call AxSym(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "NOVA-SPM"
''                    Call Nova_SPM(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "CELLDYN"
''                    Call Cell_Dyn3500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "URISCAN-PRO"
''                    Call URISCAN_PRO(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "REFLOTRON"
''                    Call REFLOTRON(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                Case "TBA-80FR"
''                    Call TBA_80FR(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
''                '###############################################################################
''
''                Case "EASYBLOOD"
''                    Call EasyBld(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'
'
'
'''
'''                'Bestian 병원
'''                '02.04.21 일요일 sebal
'''                Case "NOVA5"
'''                    Call NOVA_ELEC(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''
'''
'''                '02.04.24 sebal
'''                Case "CD1700"
'''                    Call CD_1700(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''
'''                '의왕성모병원
'''                Case "CLINITEK500"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call Clinitek_500(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "VITROS_ECI"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call VITROS_ECI(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "VITROS_750"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call VITROS_750(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "ACL100"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call ACL_100(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                Case "ADVIA120"
'''                    'DoEvents
'''                    'DoEvents
'''                    'DoEvents
'''                    Call PAUSE(1)
'''                    Call ADVIA_120(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    Call PAUSE(1)
'''                    'DoEvents
'''                    'DoEvents
'''                    'DoEvents
'''
'''                Case "RAPID_LAB"
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''                    Call RAPIDO(tGblSelMch(i).Buf, tGblSelMch(i).LabMchMst.LabMchCod, i)
'''                    DoEvents
'''                    DoEvents
'''                    DoEvents
'''
'''                End Select
'''
'''                If frmInterface.lvwMch.ListItems(i).Selected = True Then
'''                    frmInterface.lvwMch.ListItems(i).Icon = "RUN"
'''                Else
'''                    frmInterface.lvwMch.ListItems(i).Icon = "MCH" & CStr(i)
'''                End If
'''
'''            End If                                         '테스트중 2002.01.16 sebal
'''
'''            '2002.01.17 sebal
'''            '학부접수장비중, 학부접수 하지 않고, 바로 검사를 하면, RsbInf.RsbIntYon = "N" 이기
'''            '때문에 ResInf에 자동으로 저장 할때 접수번호를 가져 오지 못한다. (학부접수를 하면
'''            'RsbInf.RsbIntYon이 "P"로 저장 된다) 그리하야, 학부접수를 한다고 setting 해 놓고
'''            '학부 접수하지 않고 바로 검사하는 장비에 대해서 아래의 메세지 뿌려준다.
'''            'sGblPartErr = "N" : P로 저장 되어정상.
'''            'sGblPartErr = "Y" : 학부접수 하지않아 접수번호 가져 오지 못함.
'''
'''            If sGblPartErr = "Y" Then
'''                sMchCod = tGblSelMch(i).LabMchMst.LabMchCod
'''                sAcpCod = MasterHelpDetail("LabMchMst", sMchCod & Chr(5), sMchCod & Chr(5), 12)
'''                sAnswer = MsgBox(sMchCod & " 장비의 " & sAcpCod & " 학부접수 후 재 검사 하세요")
'''                sGblPartErr = "N"
'''            End If
'''
'''        Next i
'''
'''    Loop
'
'    Close #1
'
'End Sub

Public Sub CreateControls(ByVal MchCod As String)

'2004/04/06 이상은 - 주석 처리
'    Dim sCurKey As String
'    Dim sCmpKey As String
'    Dim sValue As String
'    Dim LABMST As NewLabMstRec
'    Dim MchCodMst As MchCodMstRec
'    Dim sDBName As String
'    Dim iTop As Integer
'    Dim iLeft As Integer
'    Dim iHeight As Integer
'    Dim iSky As Integer
'    Dim iEarth As Integer
'    Dim i As Integer
'    Dim s As String
'
'    iSky = 60
'
'    'iSky = frmInterface.Panel3D1.Top + frmInterface.Panel3D1.Height + 30
'    'TestNameLabel(0).Top
'    iEarth = frmInterface.pnlBack.Height - 60  'frmInterface.Height - 400
'
'    iTop = iSky
'    iLeft = frmInterface.TestNameLabel(0).Left + 60
'
'    sDBName = "MchCodMstCodPos"
'
'    sCmpKey = MchCod & Chr(5)
'    sCurKey = sCmpKey
'    sCurKey = mSetNext(sDBName, sCurKey)
'
'    Do
'        sCurKey = mReadNext(sDBName, sCurKey, sCmpKey, sValue)
'        If sCurKey = "" Then Exit Do
'
'        Call MchCodMstLoad(sValue, MchCodMst)
'        Call NewLabMstRead(MchCodMst.MchLabCod, LABMST)
'
'        '먼저, 카운터 체크해서 Control을 눈에 보이게 한다
'        NumOfTotalControls = NumOfTotalControls + 1
'
'        '라벨을 Enable 시킨다
'        Select Case UCase(LABMST.LabRltTyp)
'            Case "S", "M", "C"
'                Load frmInterface.TestNameLabel(NumOfTotalControls)
'                Call EnableControl(frmInterface.TestNameLabel(NumOfTotalControls))
'            Case "L"
'            Case Else
'                'Disabled Label
'                Load frmInterface.TestNameLabel(NumOfTotalControls)
'        End Select
'
'        '결과 넣을 자리를 enable 시킨다
'        Select Case UCase(LABMST.LabRltTyp)
'            Case "S", "s"
'                NumOfResultText = NumOfResultText + 1
'                Load frmInterface.ResultText(NumOfResultText)
'                Call EnableControl(frmInterface.ResultText(NumOfResultText))
'                frmInterface.TestNameLabel(NumOfTotalControls).Caption = LABMST.LabCodNam
'                frmInterface.ResultText(NumOfResultText).Tag = LABMST.LabCod
'            Case "M", "m"
'                NumOfMultiResultText = NumOfMultiResultText + 1
'                Load frmInterface.MultiResultText(NumOfMultiResultText)
'                Call EnableControl(frmInterface.MultiResultText(NumOfMultiResultText))
'                frmInterface.TestNameLabel(NumOfTotalControls).Caption = LABMST.LabCodNam
'                frmInterface.MultiResultText(NumOfMultiResultText).Tag = LABMST.LabCod
'            Case "C", "c"
'                NumOfResultCombo = NumOfResultCombo + 1
'                Load frmInterface.ResultCombo(NumOfResultCombo)
'                Call EnableControl(frmInterface.ResultCombo(NumOfResultCombo))
'                frmInterface.TestNameLabel(NumOfTotalControls).Caption = LABMST.LabCodNam
'                frmInterface.ResultCombo(NumOfResultCombo).Tag = LABMST.LabCod
'            Case "L", "l"
'            Case Else
'                NumOfLabel = NumOfLabel + 1
'                Load frmInterface.TestTitleLabel(NumOfLabel)
'                Call EnableControl(frmInterface.TestTitleLabel(NumOfLabel))
'                frmInterface.TestTitleLabel(NumOfLabel).Caption = LABMST.LabCodNam
'        End Select
'
'        'TabStop 셋팅
'        Select Case LABMST.LabRltTyp
'            Case "S", "s"
'                frmInterface.ResultText(NumOfResultText).TabIndex = NumOfTotalControls + 100
'            Case "M", "m"
'                frmInterface.MultiResultText(NumOfMultiResultText).TabIndex = NumOfTotalControls + 100
'            Case "C", "c"
'                frmInterface.ResultCombo(NumOfResultCombo).TabIndex = NumOfTotalControls + 100
'            Case "L", "l"
'            Case Else
'                '
'        End Select
'
'        'Global 변수에 저장
'        tGblControls(NumOfTotalControls).RltTyp = LABMST.LabRltTyp
'        tGblControls(NumOfTotalControls).LabCod = LABMST.LabCod
'        Select Case LABMST.LabRltTyp
'            Case "S", "s"
'                tGblControls(NumOfTotalControls).ResultControlIndex = NumOfResultText
'            Case "M", "m"
'                tGblControls(NumOfTotalControls).ResultControlIndex = NumOfMultiResultText
'            Case "C", "c"
'                tGblControls(NumOfTotalControls).ResultControlIndex = NumOfResultCombo
'            Case "L", "l"
'            Case Else
'                tGblControls(NumOfTotalControls).ResultControlIndex = NumOfLabel
'        End Select
'
'        '높이를 계산해서 너무 높으면 오른쪽 공간에서 다시 시작한다
'        Select Case LABMST.LabRltTyp
'            Case "M", "m"
'                iHeight = UNIT_HEIGHT * CInteger(LABMST.LabMaxLin)
'            Case Else
'                iHeight = UNIT_HEIGHT
'        End Select
'
'        If ((iTop + iHeight) > iEarth) Then
'            iLeft = iLeft + LABEL_WIDTH + TEXT_WIDTH + MARGIN_BETWEEN_COLUMNS
'            iTop = iSky
'        End If
'
'        '자기 자리를 잡는다 (라벨부터 !)
'        Select Case LABMST.LabRltTyp
'            Case "S", "s", "M", "m", "C", "c"
'                frmInterface.TestNameLabel(NumOfTotalControls).Left = iLeft
'                frmInterface.TestNameLabel(NumOfTotalControls).Top = iTop
'                frmInterface.TestNameLabel(NumOfTotalControls).Width = LABEL_WIDTH
'                frmInterface.TestNameLabel(NumOfTotalControls).Height = iHeight
'            Case Else
'                '
'        End Select
'
'        '다음은 결과 넣을 자리 ! (또는 노란 타이틀)
'        Select Case LABMST.LabRltTyp
'            Case "S", "s"
'                frmInterface.ResultText(NumOfResultText).Left = iLeft + LABEL_WIDTH
'                frmInterface.ResultText(NumOfResultText).Top = iTop
'                frmInterface.ResultText(NumOfResultText).Width = TEXT_WIDTH
'            Case "M", "m"
'                frmInterface.MultiResultText(NumOfMultiResultText).Left = iLeft + LABEL_WIDTH
'                frmInterface.MultiResultText(NumOfMultiResultText).Top = iTop
'                frmInterface.MultiResultText(NumOfMultiResultText).Width = TEXT_WIDTH
'                frmInterface.MultiResultText(NumOfMultiResultText).Height = iHeight
'            Case "C", "c"
'                frmInterface.ResultCombo(NumOfResultCombo).Left = iLeft + LABEL_WIDTH
'                frmInterface.ResultCombo(NumOfResultCombo).Top = iTop
'                frmInterface.ResultCombo(NumOfResultCombo).Width = TEXT_WIDTH
'            Case "L", "l"
'            Case Else
'                frmInterface.TestTitleLabel(NumOfLabel).Left = iLeft
'                frmInterface.TestTitleLabel(NumOfLabel).Top = iTop
'                frmInterface.TestTitleLabel(NumOfLabel).Width = LABEL_WIDTH + TEXT_WIDTH
'        End Select
'
'        '선택가능 검사 결과 셋팅
'        Select Case LABMST.LabRltTyp
'            Case "C", "c"
'                frmInterface.ResultCombo(NumOfResultCombo).Clear
'                i = 0
'                Do
'                    i = i + 1
'                    s = piece(LABMST.LabRltOpt, "|", i)
'                    If (s = "") Then
'                        Exit Do
'                    End If
'                    frmInterface.ResultCombo(NumOfResultCombo).AddItem (s)
'                Loop
'            Case Else
'                '
'        End Select
'
'        iTop = iTop + iHeight
'
'        If (NumOfTotalControls = MAX_TOTAL_CONTROLS) Then
'            Exit Do
'        End If
'    Loop

End Sub
'02.04.24 sebal 저장후 Text clear
Public Sub ClearIntTxt()

'2004/04/06 이상은 - 주석처리

'    frmInterface.ChtNumTxt.Text = ""
'    frmInterface.PatNamTxt.Text = ""
'    frmInterface.AgeSexTxt.Text = ""
'    frmInterface.AcpDteTxt.Text = ""
'    frmInterface.AcpNumTxt.Text = ""
'    frmInterface.txtBarCode.Text = ""
'    frmInterface.DepTxt.Text = ""
'    frmInterface.EntDtmTxt.Text = ""
'    frmInterface.SpmCodTxt.Text = ""

End Sub
Public Sub DeleteIntInf()

'2004/04/06 이상은 - 주석처리
'    Dim CookedTstDte As String
'    Dim TstDte As String
'    Dim SeqNum As String
'    Dim MchCod As String
'    Dim sCurKey As String
'    Dim sCmpKey As String
'    Dim sValue As String
'    Dim IntInf As IntInfRec
'    Dim sDbName As String
'
'    If (frmInterface.SeqNumList.Text = "") Then
'        Exit Sub
'    End If
'
'    MchCod = frmInterface.Tag
'
'    CookedTstDte = Left$(frmInterface.SeqNumList.Text, 10)
'    TstDte = Left$(CookedTstDte, 4) & Mid$(CookedTstDte, 6, 2) & Mid$(CookedTstDte, 9, 2)
'
'    SeqNum = Format(Trim(Mid$(frmInterface.SeqNumList.Text, 13)), "@@@@@@@@@@@@")
'
'    sDbName = "IntInf"
'    sCmpKey = MchCod & Chr(5) & TstDte & Chr(5) & SeqNum & Chr(5)
'    sCurKey = MchCod & Chr(5) & TstDte & Chr(5) & SeqNum & Chr(5)
'    sCurKey = mSetNext(sDbName, sCurKey)
'    Do
'        sCurKey = mReadNext(sDbName, sCurKey, sCmpKey, sValue)
'        If (sCurKey = "") Then
'
'
'            'Corebrain 2003-08-12
'            '성삼병원엔 PENTRA에서 빈값이 저장되는경우가 있다 삭제할경우
'            'SeqNum의 Format형식이 달라서 삭제가 안되는경우가 있었슴
'            '길이가 11인 경우이면 정상적으로 데이터가 들어간 경우임.
'            '비정상적인 Null값이 들어간 경우는 10으로 들어가있음
'            If Len(Mid$(frmInterface.SeqNumList.Text, 13)) = "11" Then
'                Exit Do
'            Else
'
'                SeqNum = Format(Trim(Mid$(frmInterface.SeqNumList.Text, 13)), "@@@@@@@@@@@")
'                sDbName = "IntInf"
'                sCmpKey = MchCod & Chr(5) & TstDte & Chr(5) & SeqNum & Chr(5)
'                sCurKey = MchCod & Chr(5) & TstDte & Chr(5) & SeqNum & Chr(5)
'                sCurKey = mSetNext(sDbName, sCurKey)
'                Do
'                    sCurKey = mReadNext(sDbName, sCurKey, sCmpKey, sValue)
'                    If (sCurKey = "") Then
'                        Exit Sub
'                    End If
'
'                    If (mDelete(sDbName, sCurKey) = False) Then
'                        MsgBox ("삭제 오류")
'                        Exit Sub
'                    End If
'                Loop
'            End If
'
'        End If
'
'        If (mDelete(sDbName, sCurKey) = False) Then
'            MsgBox ("삭제 오류")
'            Exit Sub
'        End If
'    Loop
'
'    'MultiSelect Delete와 같이사용하므로 DeleteIntInf에서 빼낸다.
'    'frmInterface.SeqNumList.RemoveItem (frmInterface.SeqNumList.ListIndex)
'
End Sub

'Public Sub DisplayInfPanel(ResInf As ResInfRec)
'
'    Dim Dte As String
'    Dim tim As String
'
'    '인적사항표시
'    Call DisplayPbsInf(ResInf.ResChtNum)
'
''    '접수일시 표시
''    Dte = Mid$(ResInf.ResAcpDtm, 1, 8)
''    tim = Mid$(ResInf.ResAcpDtm, 9, 4)
''    frmInterface.AcpDtmTxt.Text = CookDate(Dte) & " " & CookTime(tim)
''
''    '의뢰일시 표시
''    Dte = Mid$(ResInf.ResOdrDtm, 1, 8)
''    tim = Mid$(ResInf.ResOdrDtm, 9, 4)
''    frmInterface.EntDtmTxt.Text = CookDate(Dte) & " " & CookTime(tim)
''
''    'OcmInf 표시
''    Call DisplayOcmInf(ResInf.ResOcmNum, ResInf.ResOspIsp, Dte)
'
'    '검체종류 표시
'    frmAxsym.SpmCodTxt.Text = ResInf.ResSpmCod
'    SetText frmAxsym.vasID, frmAxsym.SpmCodTxt.Text, frmAxsym.vasID.ActiveRow, 16
'End Sub

Public Sub DisplayIntInf()

'    Dim i As Integer
'    Dim CookedTstDte As String
'    Dim TstDte As String
'    Dim SeqNum As String
'    Dim MchCod As String
'    Dim sCurKey As String
'    Dim sCmpKey As String
'    Dim sValue As String
'    Dim IntInf As IntInfRec
'    Dim sDbName As String
'
'    MchCod = frmInterface.Tag
'
'    CookedTstDte = Left$(frmInterface.SeqNumList.Text, 10)
'    TstDte = Left$(CookedTstDte, 4) & Mid$(CookedTstDte, 6, 2) & Mid$(CookedTstDte, 9, 2)
'
'    SeqNum = Format(Trim(Mid$(frmInterface.SeqNumList.Text, 13)), "@@@@@@@@@@@@")
'
'    sDbName = "IntInf"
'    sCmpKey = MchCod & Chr(5) & TstDte & Chr(5) & SeqNum & Chr(5)
'    sCurKey = MchCod & Chr(5) & TstDte & Chr(5) & SeqNum & Chr(5)
'    sCurKey = mSetNext(sDbName, sCurKey)
'    Do
'        sCurKey = mReadNext(sDbName, sCurKey, sCmpKey, sValue)
'        If (sCurKey = "") Then
'            Exit Do
'        End If
'
'        Call IntInfLoad(sValue, IntInf)
'
'        For i = 1 To NumOfTotalControls
'            If (tGblControls(i).LabCod = IntInf.ResLabCod) Then
'                Select Case tGblControls(i).RltTyp
'                    Case "S", "s"
'                        frmInterface.ResultText(tGblControls(i).ResultControlIndex).BackColor = &HC0E0FF
'                        frmInterface.ResultText(tGblControls(i).ResultControlIndex).Text = IntInf.ResLabRlt
'                    Case "M", "m"
'                        frmInterface.MultiResultText(tGblControls(i).ResultControlIndex).BackColor = &HC0E0FF
'                        frmInterface.MultiResultText(tGblControls(i).ResultControlIndex).Text = IntInf.ResLabRlt
'                    Case "C", "c"
'                        frmInterface.ResultCombo(tGblControls(i).ResultControlIndex).BackColor = &HC0E0FF
'                        frmInterface.ResultCombo(tGblControls(i).ResultControlIndex).Text = IntInf.ResLabRlt
'                    'case "L", "l"
'                    Case Else
'                        '
'                End Select
'            End If
'        Next i
'    Loop

End Sub

Public Sub DisplayOcmInf(OcmNum As String, IspOrOsp As String, Dte As String)

'    Dim sCurKey As String
'    Dim sCmpKey As String
'    Dim sValue As String
'    Dim OcmInf As OcmInfRec
'    Dim IcmInf As IcmInfRec
'    Dim DepCod As String
'    Dim DepMst As DepMstRec
'
'    If (IspOrOsp = "ISP") Then
'        'Call IcmInfRead(OcmNum, IcmInf, "IcmInf")
'        Call IcmInfRead(OcmNum, IcmInf)
'        DepCod = IcmInf.IcmDepCod
'    ElseIf (IspOrOsp = "OSP") Then
'        Call OcmInfRead(OcmNum, OcmInf)
'        DepCod = OcmInf.OcmDepCod
'    Else
'        Exit Sub
'    End If
'
'    Call DepMstRead(DepCod, Dte, DepMst)
'    frmInterface.DepTxt.Text = DepMst.DepKorNam

End Sub

'Public Sub DisplayPbsInf(sChtNum As String)
'
'    Dim PbsInf As PbsInfRec
'    Dim Age As String
'    Dim Sex As String
'
'    Call PbsInfRead(Format$(sChtNum, "@@@@@@@@"), PbsInf)
'
''    Form1.ChtNumTxt.Text = PbsInf.PbsChtNum
''    Form1.PatNamTxt.Text = PbsInf.PbsPatNam
''    Call GetAgeSex(Age, Sex, PbsInf.PbsResNum)
''    Form1.AgeSexTxt.Text = Sex & "/" & Age
'
'    SetText frmAxsym.vasID, PbsInf.PbsChtNum, frmAxsym.vasID.ActiveRow, 6        '챠트번호
'
'    SetText frmAxsym.vasID, PbsInf.PbsPatNam, frmAxsym.vasID.ActiveRow, 7 '이름
'
'    Call GetAgeSex(Age, Sex, PbsInf.PbsResNum)
''    SetText frmCD3200.vasID, Sex, frmCD3200.vasID.ActiveRow, 8       '성별
'    Select Case Sex
'    Case "여"
'        SetText frmAxsym.vasID, "F", frmAxsym.vasID.ActiveRow, 8       '성별
'    Case "남"
'        SetText frmAxsym.vasID, "M", frmAxsym.vasID.ActiveRow, 8       '성별
'    End Select
'    SetText frmAxsym.vasID, Age, frmAxsym.vasID.ActiveRow, 9       '나이
'
'    'sGblChtNum = PbsInf.PbsChtNum
'End Sub

Public Sub DisplaySeqNum(ByVal MchCod As String)

'    Dim sDbName As String
'    Dim sCurKey As String
'    Dim sCmpKey As String
'    Dim sValue As String
'    Dim IntInf As IntInfRec
'    Dim Buf As String
'    Dim OldTstDte As String
'    Dim OldSeqNum As String
'    Dim sDate As String
'
'    OldTstDte = ""
'    OldSeqNum = ""
'    frmInterface.SeqNumList.Clear
'
'    'sDBName = "IntInfCfmMchTstSeq"
'    'sCmpKey = "N" & Chr(5) & MchCod & Chr(5)
'    sDbName = "IntInf"
'
'    If frmInterface.dtpDate.Tag = "ALL" Then
'        sCmpKey = MchCod & Chr(5)
'    Else
'        sDate = ""
'        Call Setting_DateTimePicker(frmInterface.dtpDate, sDate)
'        sCmpKey = MchCod & Chr(5) & sDate & Chr(5)
'    End If
'
'    sCurKey = sCmpKey
'    sCurKey = mSetNext(sDbName, sCurKey)
'    Do
'        sCurKey = mReadNext(sDbName, sCurKey, sCmpKey, sValue)
'        If (sCurKey = "") Then
'            Exit Do
'        End If
'
'        Call IntInfLoad(sValue, IntInf)
'
'        If ((OldTstDte <> IntInf.ResTstDte) Or (OldSeqNum <> IntInf.ResSeqNum)) Then
'            'If IntInf.ResTstDte <> "20030718" Then Stop
'            Buf = CookDate(IntInf.ResTstDte) & " " & IntInf.ResSeqNum
'            frmInterface.SeqNumList.AddItem Buf
'
'            OldTstDte = IntInf.ResTstDte
'            OldSeqNum = IntInf.ResSeqNum
'        End If
'
'        'IntInf.ResSeqNum = IntInf.ResSeqNum & Chr(1)
'        'Call IntInfStore(sCurKey, sValue, IntInf)
'    Loop
'
'    If (frmInterface.SeqNumList.ListCount > 0) Then
'        'frmInterface.SeqNumList.TopIndex = frmInterface.SeqNumList.ListCount - 1
'        frmInterface.SeqNumList.TopIndex = 0
'    End If

End Sub

Public Sub GetFirstRecord(SpmDte As String, AcpCod As String, SpmNum As String, ResInf As ResInfRec, Found As Integer)

    Dim sCurKey As String
    Dim sCmpKey As String
    Dim sValue As String
    Dim sDBName As String

    Found = False
    sDBName = "ResInf"
    sCmpKey = SpmDte & Chr(5) & AcpCod & Chr(5) & SpmNum & Chr(5)
    sCurKey = sCmpKey
    sCurKey = mSetNext(sDBName, sCurKey)
    Do
        sCurKey = mReadNext(sDBName, sCurKey, sCmpKey, sValue)

        If sCurKey = "" Then Exit Do

        Call ResInfLoad(sValue, ResInf)

        '통계여부에 "OC"인 취소된 처방은 제외시킨다.
        If ResInf.ResStaYon <> "OC" Then
            Found = True
            Exit Sub
        End If
    Loop

End Sub

Public Sub InitfrmInterface()

'    Dim i As Integer
'    Dim FocusMoved As Integer   '화면상 제일 처음 결과가 입력
'                                '된 곳으로 커서를 옮기기 위함
'
'    FocusMoved = False
'
'    For i = 1 To NumOfTotalControls
'        Select Case tGblControls(i).RltTyp
'            Case "S", "s"
'                frmInterface.ResultText(tGblControls(i).ResultControlIndex).BackColor = WHITE
'                frmInterface.ResultText(tGblControls(i).ResultControlIndex).Text = ""
'                If (FocusMoved = False) Then
'                    FocusMoved = True
'                    frmInterface.ResultText(tGblControls(i).ResultControlIndex).SetFocus
'                End If
'
'            Case "M", "m"
'                frmInterface.MultiResultText(tGblControls(i).ResultControlIndex).BackColor = WHITE
'                frmInterface.MultiResultText(tGblControls(i).ResultControlIndex).Text = ""
'                If (FocusMoved = False) Then
'                    FocusMoved = True
'                    frmInterface.MultiResultText(tGblControls(i).ResultControlIndex).SetFocus
'                End If
'
'            Case "C", "c"
'                frmInterface.ResultCombo(tGblControls(i).ResultControlIndex).BackColor = WHITE
'                frmInterface.ResultCombo(tGblControls(i).ResultControlIndex).Text = ""
'                If (FocusMoved = False) Then
'                    FocusMoved = True
'                    frmInterface.ResultCombo(tGblControls(i).ResultControlIndex).SetFocus
'                End If
'
'            'case "L", "l"
'            Case Else
'                '
'        End Select
'    Next i

End Sub

Public Sub SaveChangedData(ByVal LabCod As String, ByVal Result As String)

'    Dim CookedTstDte As String
'    Dim sCurKey As String
'    Dim sValue As String
'    Dim IntInf As IntInfRec
'    Dim sDbName As String
'
'    sDbName = "IntInf"
'
'    CookedTstDte = Left$(frmInterface.SeqNumList.Text, 10)
'
'    IntInf.ResMchCod = frmInterface.Tag
'    IntInf.ResTstDte = Left$(CookedTstDte, 4) & Mid$(CookedTstDte, 6, 2) & Mid$(CookedTstDte, 9, 2)
'    IntInf.ResSeqNum = Mid$(frmInterface.SeqNumList.Text, 13)
'    IntInf.ResLabCod = LabCod
'
'    Call IntInfStore(sCurKey, sValue, IntInf)
'    sCurKey = mSetReadEqual(sDbName, sCurKey, sValue)
'    If (sCurKey = "") Then
'        MsgBox ("저장에러")
'    End If
'
'    Call IntInfLoad(sValue, IntInf)
'    IntInf.ResLabRlt = Result
'
'    Call IntInfStore(sCurKey, sValue, IntInf)
'    If (mWrite(sDbName, sCurKey, sValue) = False) Then
'        Call IntInfStore(sCurKey, sValue, IntInf)
'        If (mUpdate(sDbName, sCurKey, sValue) = False) Then
'            MsgBox ("저장에러")
'        End If
'    End If

End Sub
Public Sub SaveToIntInf(MchCod As String, TstDte As String, SeqNum As String, LabCod As String, LabRlt As String)

'    Dim IntInf      As IntInfRec
'    Dim sCurKey     As String
'    Dim sValue      As String
'    Dim sDtlPartYn   As String
'
'    IntInf.ResMchCod = MchCod
'    IntInf.ResTstDte = TstDte
'    IntInf.ResSeqNum = Format(Trim(SeqNum), "@@@@@@@@@@@@")
'    IntInf.ResLabCod = LabCod
'    IntInf.ResLabRlt = LabRlt
'    IntInf.ResAcpDte = ""
'    IntInf.ResAcpCod = ""
'    IntInf.ResAcpNum = ""
'    IntInf.ResCfmYon = "N"
'
'    If (LabCod = "") Then
'        Exit Sub
'    End If
'
'    Call IntInfStore(sCurKey, sValue, IntInf)
'    If (mWrite("IntInf", sCurKey, sValue) = False) Then
'        If (mUpdate("IntInf", sCurKey, sValue) = False) Then
'            Beep
'        End If
'        'MsgBox ("IntInf mWrite Error")
'
'    End If
'
'
'    '2002.01.16 sebal
'    '트레이를 사용하여 파트접수를 하는 장비는 자동으로 ResInf로 저장....
'    'LabMst.LabTryYon = Y 이고, 상세정보의 LABMST중, PARYON = Y일 경우만.
'
'    'EverSky 나중에 풀어서 저장한다. . . .. . . .  .바코드를 사용 하여.....
'
'
'
''    sDtlPartYn = piece(MasterHelpDetail("DtlMst", "LABMST" & Chr(5) & "PARTYN" & Chr(5), "LABMST" & Chr(5) & "PARTYN" & Chr(5), 3), "=", 1)
''
''    If MasterHelpDetail("LabMchMst", MchCod & Chr(5), MchCod & Chr(5), 13) = "Y" And sDtlPartYn = "Y" Then
''        Call Int2Res4PartReg(MchCod, TstDte, SeqNum, LabCod, LabRlt)
''    End If

End Sub

'Public Sub SaveToResInf_Interface(tPrmResInf As ResInfRec, Optional bPrmConfirmSetting As Boolean = False)

'2004/04/06 이상은 수정
'============================================================================================================
'    Dim sDbName As String
'    Dim sCurKey As String
'    Dim sValue As String
'    Dim sCmpKey As String
'    Dim ResInf As ResInfRec
'    Dim sNow As String
'    Dim Result As String
'    Dim Msg As String
'    Dim Reply As Integer
'    Dim i As Integer
'    Dim Ok As Integer
'    Dim Mst As String
'
''    If (Form1.SeqNumList.Text = "") Then
''        MsgBox ("검사 결과를 선택하세요 !")
''        Exit Sub
''    End If
'
''    Reply = MsgBox("저장하시겠습니까 ?", MB_YESNO + MB_ICONQUESTION, "저장하시겠습니까?")
''    If (Reply = IDNO) Then
''        Exit Sub
''    End If
'
'    Ok = True
'    sDbName = "ResInf"
'
''    For i = 1 To NumOfTotalControls
''        Select Case tGblControls(i).RltTyp
''            Case "S", "s"
''                Result = Trim(frmInterface.ResultText(tGblControls(i).ResultControlIndex).Text)
''            Case "M", "m"
''                Result = Trim(frmInterface.MultiResultText(tGblControls(i).ResultControlIndex).Text)
''            Case "C", "c"
''                Result = Trim(frmInterface.ResultCombo(tGblControls(i).ResultControlIndex).Text)
''            'case "L", "l"
''            Case Else
''                GoTo NextI
''        End Select
''        If (Result = "") Then
''            GoTo NextI
''        End If
'
''        ResInf.ResAcpDte = tPrmResInf.ResAcpDte
''        ResInf.ResAcpCod = tPrmResInf.ResAcpCod
''        ResInf.ResAcpNum = tPrmResInf.ResAcpNum
'
'        ResInf.ResAcpDte = tPrmResInf.ResAcpDte
'        ResInf.ResAcpCod = tPrmResInf.ResAcpCod
'        ResInf.ResAcpNum = tPrmResInf.ResAcpNum
'        ResInf.ResSpmCod = tPrmResInf.ResSpmCod
'
'
'        ResInf.ResLabCod = "B1050"
'        Call ResInfStore(sCurKey, sValue, ResInf)
'        sCurKey = mSetReadEqual(sDbName, sCurKey, sValue)
'        If (sCurKey = "") Then
'            Ok = False
'            GoTo NextI
'        End If
'        Call ResInfLoad(sValue, ResInf)
'
'        '담당의사 조회 여부 체크
'        If ((ResInf.ResSeeYon = "Y") And (Result <> ResInf.ResMzhMnt)) Then
'            Msg = "담당의사가 " & ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과를 조회하였으므로"
'            Msg = Msg & Chr(13) & "검사결과를 수정할 수 없습니다."
'            MsgBox (Msg)
'            GoTo NextI
'        End If
'
'        If ((ResInf.ResConLvl = "2") And (Result <> ResInf.ResMzhMnt)) Then
'            Msg = ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과는 확정되었으므로"
'            Msg = Msg & Chr(13) & "검사결과를 수정할 수 없습니다."
'            MsgBox (Msg)
'            GoTo NextI
'        End If
'
'        sNow = SystemTime()
'        ResInf.ResTstDtm = sGblTstDte & sNow
'        ResInf.ResUpdDtm = sGblTstDte & sNow
'        ResInf.ResUpdUid = tGblUidData.UidCod
'        ResInf.ResMzhMnt = Result
'
'        'If (ResInf.ResTstDtm = "") Then
'        ResInf.ResTstDtm = "" 'sGblTstDte & sNow
'        'End If
'        If (ResInf.ResTstUid = "") Then
'            ResInf.ResTstUid = tGblUidData.UidCod
'        End If
'
'        If (ResInf.ResConLvl <> "2") Then
'            ResInf.ResConLvl = "0"
'        End If
'
'        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'        ' 2001/12/18 james
'        ' Optional로 인터페이스에서 바로 검사승인을 할수 있도록...
'        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'        If bPrmConfirmSetting Then
'
'            ResInf.ResTstDtm = sGblTstDte & sNow
'            ResInf.ResTstUid = tGblUidData.UidCod
'            ResInf.ResSclCod = ""   '=> 검사승인에서도 수탁부분이 주석처리되어 있어서 일단은 Null로
'            ResInf.ResConLvl = "1"
'            ResInf.ResOkYon = "Y"
'            ResInf.ResOkUid = tGblUidData.UidCod
'            ResInf.ResOkDtm = sGblTstDte & sNow
'
'            Select Case GetRtnCode(ResInf.ResLabCod)
'                Case "B" '손코드
'                    If Result <> "" Then
'                        ResInf.ResStaYon = "B"
'                        Call SetLabCodtoN(ResInf, sGblTstDte & sNow)
'                    Else
'                        ResInf.ResStaYon = ""
'                    End If
'                Case "N" '자코드
'                    If Result <> "" Then
'                        ResInf.ResStaYon = "N"
'                        Call SetLabCodtoY(ResInf)
'                    Else
'                        ResInf.ResStaYon = ""
'                    End If
'                Case "Y"  '모코드
'                    If Result <> "" Then
'                        ResInf.ResStaYon = "Y"
'                    Else
'                        ResInf.ResStaYon = ""
'                    End If
'            End Select
'
'            If sFnlYon = "Y" Then
'                ResInf.ResRepTyp = "F" 'Final Report
'            Else
'                ResInf.ResRepTyp = "I" 'Interim Report
'            End If
'
''        End If
'        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'        '드디어 저장
'        Call ResInfStore(sCurKey, sValue, ResInf)
'        If (mUpdate(sDbName, sCurKey, sValue) = False) Then
'            MsgBox ("ResInf mUpdate Error")
'        End If
'NextI:
'    Next i
'
'    If (Ok = False) Then
'        Msg = "저장하지 못한 결과가 있습니다." & Chr(KEY_RETURN)
'        Msg = Msg & "인터페이스된 자료를 삭제할까요 ?"
'        Reply = MsgBox(Msg, MB_YESNO + MB_ICONEXCLAMATION, "삭제할까요?")
'        If (Reply = IDNO) Then
'            Exit Sub
'        End If
'    End If
'
'    Call DeleteIntInf
'    Call ClearIntTxt            '02.04.24 sebal 저장후 Text clear
'
'    'MultiSelect Delete와 같이 사용하므로 DeleteIntInf에서 빼낸다.
'    frmInterface.SeqNumList.RemoveItem (frmInterface.SeqNumList.ListIndex)
'============================================================================================================

'End Sub

'Public Sub SaveToResInf_Interface(tPrmResInf As ResInfRec, argRow As Long, Optional bPrmConfirmSetting As Boolean = False)
'
'    Dim sDBName As String
'    Dim sCurKey As String
'    Dim sValue As String
'    Dim sCmpKey As String
'    Dim ResInf As ResInfRec
'    Dim sNow As String
'    Dim Result As String
'    Dim Msg As String
'    Dim Reply As Integer
'    Dim i As Integer
'    Dim Ok As Integer
'    Dim Mst As String
'
''    If (Form1.SeqNumList.Text = "") Then
''        MsgBox ("검사 결과를 선택하세요 !")
''        Exit Sub
''    End If
'
''    Reply = MsgBox("저장하시겠습니까 ?", MB_YESNO + MB_ICONQUESTION, "저장하시겠습니까?")
''    If (Reply = IDNO) Then
''        Exit Sub
''    End If
'
'    Ok = True
'    sDBName = "ResInf"
'
''    For i = 1 To NumOfTotalControls
''        Select Case tGblControls(i).RltTyp
''            Case "S", "s"
''                Result = Trim(frmInterface.ResultText(tGblControls(i).ResultControlIndex).Text)
''            Case "M", "m"
''                Result = Trim(frmInterface.MultiResultText(tGblControls(i).ResultControlIndex).Text)
''            Case "C", "c"
''                Result = Trim(frmInterface.ResultCombo(tGblControls(i).ResultControlIndex).Text)
''            'case "L", "l"
''            Case Else
''                GoTo NextI
''        End Select
''        If (Result = "") Then
''            GoTo NextI
''        End If
'
''        ResInf.ResAcpDte = tPrmResInf.ResAcpDte
''        ResInf.ResAcpCod = tPrmResInf.ResAcpCod
''        ResInf.ResAcpNum = tPrmResInf.ResAcpNum
'
'        ResInf.ResAcpDte = Format(frmAxsym.Text_Today.Text, "yyyymmdd")        '접수일자
'        ResInf.ResAcpCod = gPart        '접수코드 = 파트
'        ResInf.ResAcpNum = Format(frmAxsym.txtBarcode.Text, "@@@@@@@@@@")  '접수번호
'        ResInf.ResSpmCod = frmAxsym.SpmCodTxt.Text     '검체코드
'
'
'        ResInf.ResLabCod = Trim(GetText(frmAxsym.vasRes, argRow, 2))    '검사코드
'        Call ResInfStore(sCurKey, sValue, ResInf)
'        Debug.Print sCurKey
'
'        sCurKey = mSetReadEqual(sDBName, sCurKey, sValue)
''        If (sCurKey = "") Then
''            Ok = False
''            GoTo NextI
''        End If
'        Call ResInfLoad(sValue, ResInf)
'
'        '담당의사 조회 여부 체크
''        If ((ResInf.ResSeeYon = "Y") And (Result <> ResInf.ResMzhMnt)) Then
''            Msg = "담당의사가 " & ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과를 조회하였으므로"
''            Msg = Msg & Chr(13) & "검사결과를 수정할 수 없습니다."
''            MsgBox (Msg)
''            GoTo NextI
''        End If
'
'        If ((ResInf.ResConLvl = "2") And (Result <> ResInf.ResMzhMnt)) Then
'            Msg = ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과는 확정되었으므로"
'            Msg = Msg & Chr(13) & "검사결과를 수정할 수 없습니다."
'            MsgBox (Msg)
''            GoTo NextI
'        End If
'
'        sNow = SystemTime()
'        ResInf.ResTstDtm = sGblTstDte & sNow
'        ResInf.ResUpdDtm = sGblTstDte & sNow
'        ResInf.ResUpdUid = tGblUidData.UidCod
'        ResInf.ResMzhMnt = Trim(GetText(frmAxsym.vasRes, argRow, 4))       '검사결과
'
'        'If (ResInf.ResTstDtm = "") Then
'        ResInf.ResTstDtm = "" 'sGblTstDte & sNow
'        'End If
'        If (ResInf.ResTstUid = "") Then
'            ResInf.ResTstUid = tGblUidData.UidCod
'        End If
'
'        If (ResInf.ResConLvl <> "2") Then
'            ResInf.ResConLvl = "0"
'        End If
'
'        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'        ' 2001/12/18 james
'        ' Optional로 인터페이스에서 바로 검사승인을 할수 있도록...
'        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''        If bPrmConfirmSetting Then
''
''            ResInf.ResTstDtm = sGblTstDte & sNow
''            ResInf.ResTstUid = tGblUidData.UidCod
''            ResInf.ResSclCod = ""   '=> 검사승인에서도 수탁부분이 주석처리되어 있어서 일단은 Null로
''            ResInf.ResConLvl = "1"
''            ResInf.ResOkYon = "Y"
''            ResInf.ResOkUid = tGblUidData.UidCod
''            ResInf.ResOkDtm = sGblTstDte & sNow
''
''            Select Case GetRtnCode(ResInf.ResLabCod)
''                Case "B" '손코드
''                    If Result <> "" Then
''                        ResInf.ResStaYon = "B"
''                        Call SetLabCodtoN(ResInf, sGblTstDte & sNow)
''                    Else
''                        ResInf.ResStaYon = ""
''                    End If
''                Case "N" '자코드
''                    If Result <> "" Then
''                        ResInf.ResStaYon = "N"
''                        Call SetLabCodtoY(ResInf)
''                    Else
''                        ResInf.ResStaYon = ""
''                    End If
''                Case "Y"  '모코드
''                    If Result <> "" Then
''                        ResInf.ResStaYon = "Y"
''                    Else
''                        ResInf.ResStaYon = ""
''                    End If
''            End Select
''
''            If sFnlYon = "Y" Then
''                ResInf.ResRepTyp = "F" 'Final Report
''            Else
''                ResInf.ResRepTyp = "I" 'Interim Report
''            End If
''
'''        End If
'        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'        '드디어 저장
'        Call ResInfStore(sCurKey, sValue, ResInf)
'        If (mUpdate(sDBName, sCurKey, sValue) = False) Then
'            MsgBox ("ResInf mUpdate Error")
'        End If
''NextI:
''    Next i
'
''    If (Ok = False) Then
''        Msg = "저장하지 못한 결과가 있습니다." & Chr(KEY_RETURN)
''        Msg = Msg & "인터페이스된 자료를 삭제할까요 ?"
''        Reply = MsgBox(Msg, MB_YESNO + MB_ICONEXCLAMATION, "삭제할까요?")
''        If (Reply = IDNO) Then
''            Exit Sub
''        End If
''    End If
''
''    Call DeleteIntInf
''    Call ClearIntTxt            '02.04.24 sebal 저장후 Text clear
''
''    'MultiSelect Delete와 같이 사용하므로 DeleteIntInf에서 빼낸다.
''    frmInterface.SeqNumList.RemoveItem (frmInterface.SeqNumList.ListIndex)
'
'End Sub

Public Sub SetDefaultSpmNumDte()

'    frmInterface.AcpDteTxt.Text = AddCentury(SystemDate())

End Sub

Public Sub UnloadControls()

'On Error Resume Next
'
'    Dim i As Integer
'
'    For i = 1 To NumOfResultText
'        Unload frmInterface.ResultText(i)
'    Next i
'    For i = 1 To NumOfMultiResultText
'        Unload frmInterface.MultiResultText(i)
'    Next i
'    For i = 1 To NumOfResultCombo
'        Unload frmInterface.ResultCombo(i)
'    Next i
'    For i = 1 To NumOfLabel
'        Unload frmInterface.TestTitleLabel(i)
'    Next i
'
'    For i = 1 To NumOfTotalControls
'        Unload frmInterface.TestNameLabel(i)
'    Next i
'
'    NumOfResultText = 0
'    NumOfMultiResultText = 0
'    NumOfResultCombo = 0
'    NumOfLabel = 0
'    NumOfTotalControls = 0

End Sub

