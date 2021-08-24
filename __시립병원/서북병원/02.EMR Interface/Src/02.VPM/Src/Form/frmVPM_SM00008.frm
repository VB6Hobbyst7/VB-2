VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmVPM_SM00008 
   Caption         =   "00008-인공수정체진단기/Tomey/AL-2000"
   ClientHeight    =   10665
   ClientLeft      =   10215
   ClientTop       =   4680
   ClientWidth     =   9945
   Icon            =   "frmVPM_SM00008.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   9945
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   7995
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   9615
      _Version        =   393216
      _ExtentX        =   16960
      _ExtentY        =   14102
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   32
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmVPM_SM00008.frx":9F8A
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtBuff 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "txtBuff"
      Top             =   900
      Width           =   4935
   End
   Begin VB.TextBox txtInput 
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Text            =   "txtInput"
      Top             =   60
      Width           =   4935
   End
   Begin FPSpread.vaSpread sprPrint 
      Height          =   9015
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
      _Version        =   393216
      _ExtentX        =   7858
      _ExtentY        =   15901
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   28
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmVPM_SM00008.frx":B8F8
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "txtBuff"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "MSComm1.Input"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1410
   End
End
Attribute VB_Name = "frmVPM_SM00008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HWND_BROADCAST    As Long = &HFFFF&
Private Const WM_WININICHANGE   As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Public Sub FUNC_MM_PRINT(argData As String)
    Dim strSubData()
    Dim strResult()
    Dim lngCnt                  As Long
    Dim intCnt                  As Integer
    Dim intResultCntR           As Integer
    Dim intResultCntL           As Integer
    
    On Error Resume Next
    
    '/Step1.자료 변환
    intCnt = 0
    For lngCnt = 1 To Len(argData)
        If Mid(argData, lngCnt, 1) = Chr(1) Or Mid(argData, lngCnt, 1) = Chr(2) Then
            intCnt = intCnt + 1
            ReDim Preserve strSubData(intCnt)
            strSubData(intCnt) = ""
        ElseIf Mid(argData, lngCnt, 1) = Chr(13) Then
'            sprPrint.SetText 1, intCnt, strSubData(intCnt)
        Else
            strSubData(intCnt) = strSubData(intCnt) & Mid(argData, lngCnt, 1)
        End If
    Next lngCnt

    For intX = 1 To UBound(strSubData)
        Select Case Trim(Left(strSubData(intX), 2))
            Case ""
                strTemp = "Date : " & Format(Trim(Mid(strSubData(intX), 13, 14)), "@@@@-@@-@@ @@:@@:@@")
                Call SET_CELL(vaSpread1, 1, 2, strTemp)
            Case "BM"
            Case "HR"
                Call SET_CELL(vaSpread1, 7, 4, Trim(Mid(strSubData(intX), 3, 14)))
            Case "VR"
            Case "LR"
                Call SET_CELL(vaSpread1, 7, 5, Trim(Mid(strSubData(intX), 8, 4)))   '/ACD
                Call SET_CELL(vaSpread1, 7, 6, Trim(Mid(strSubData(intX), 12, 4)))  '/LENS
                Call SET_CELL(vaSpread1, 7, 7, Trim(Mid(strSubData(intX), 3, 5)))   '/AXIAL
            Case "KR"
                Call SET_CELL(vaSpread1, 7, 8, Trim(Mid(strSubData(intX), 3, 5)))
                Call SET_CELL(vaSpread1, 7, 9, Trim(Mid(strSubData(intX), 8, 5)))
            Case "DR"
                Call SET_CELL(vaSpread1, 7, 10, Trim(Mid(strSubData(intX), 3, 6)))
            Case "FR"
                Call SET_CELL(vaSpread1, 9, 4, Trim(Mid(strSubData(intX), 3, 15)))
            Case "IR"
            Case "RR"
                intResultCntR = intResultCntR + 1
                Select Case intResultCntR
                    Case 1
                        Call SET_CELL(vaSpread1, 7, 12, Trim(Mid(strSubData(intX), 3, 19)))
                        Call SET_CELL(vaSpread1, 7, 13, Trim(Mid(strSubData(intX), 22, 5)))
                        strTemp = Trim(Mid(strSubData(intX), 22, 5))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 6, 20 + intY, Format(Val(strTemp) + (intY * 0.5), "#0.00")
                            Next intY
                        End If
                        Call SET_CELL(vaSpread1, 7, 14, Trim(Mid(strSubData(intX), 27, 6)))
                        strTemp = Trim(Replace(Mid(strSubData(intX), 27, 6), " ", ""))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 7, 20 + intY, Format(Val(strTemp) - (intY * 0.4), "#0.00")
                            Next intY
                        End If
                    Case 2
                        Call SET_CELL(vaSpread1, 8, 12, Trim(Mid(strSubData(intX), 3, 19)))
                        Call SET_CELL(vaSpread1, 8, 13, Trim(Mid(strSubData(intX), 22, 5)))
                        strTemp = Trim(Mid(strSubData(intX), 22, 5))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 8, 20 + intY, Format(Val(strTemp) + (intY * 0.5), "#0.00")
                            Next intY
                        End If
                        Call SET_CELL(vaSpread1, 8, 14, Trim(Mid(strSubData(intX), 27, 6)))
                        strTemp = Trim(Replace(Mid(strSubData(intX), 27, 6), " ", ""))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 9, 20 + intY, Format(Val(strTemp) - (intY * 0.4), "#0.00")
                            Next intY
                        End If
                    Case 3
                        Call SET_CELL(vaSpread1, 9, 12, Trim(Mid(strSubData(intX), 3, 19)))
                        Call SET_CELL(vaSpread1, 9, 13, Trim(Mid(strSubData(intX), 22, 5)))
                        strTemp = Trim(Mid(strSubData(intX), 22, 5))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 6, 29 + intY, Format(Val(strTemp) + (intY * 0.5), "#0.00")
                            Next intY
                        End If
                        Call SET_CELL(vaSpread1, 9, 14, Trim(Mid(strSubData(intX), 27, 6)))
                        strTemp = Trim(Replace(Mid(strSubData(intX), 27, 6), " ", ""))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 7, 29 + intY, Format(Val(strTemp) - (intY * 0.4), "#0.00")
                            Next intY
                        End If
                End Select
            Case "HL"
                Call SET_CELL(vaSpread1, 2, 4, Trim(Mid(strSubData(intX), 3, 14)))
            Case "VL"
            Case "LL"
                Call SET_CELL(vaSpread1, 2, 5, Trim(Mid(strSubData(intX), 8, 4)))   '/ACD
                Call SET_CELL(vaSpread1, 2, 6, Trim(Mid(strSubData(intX), 12, 4)))  '/LENS
                Call SET_CELL(vaSpread1, 2, 7, Trim(Mid(strSubData(intX), 3, 5)))   '/AXIAL
            Case "KL"
                Call SET_CELL(vaSpread1, 2, 8, Trim(Mid(strSubData(intX), 3, 5)))
                Call SET_CELL(vaSpread1, 2, 9, Trim(Mid(strSubData(intX), 8, 5)))
            Case "DL"
                Call SET_CELL(vaSpread1, 2, 10, Trim(Mid(strSubData(intX), 3, 6)))
            Case "FL"
                Call SET_CELL(vaSpread1, 4, 4, Trim(Mid(strSubData(intX), 3, 15)))
            Case "IL"
            Case "RL":
                intResultCntL = intResultCntL + 1
                Select Case intResultCntL
                    Case 1
                        Call SET_CELL(vaSpread1, 2, 12, Trim(Mid(strSubData(intX), 3, 19)))
                        Call SET_CELL(vaSpread1, 2, 13, Trim(Mid(strSubData(intX), 22, 5)))
                        strTemp = Trim(Mid(strSubData(intX), 22, 5))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 1, 20 + intY, Format(Val(strTemp) + (intY * 0.5), "#0.00")
                            Next intY
                        End If
                        Call SET_CELL(vaSpread1, 2, 14, Trim(Mid(strSubData(intX), 27, 6)))
                        strTemp = Trim(Replace(Mid(strSubData(intX), 27, 6), " ", ""))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 2, 20 + intY, Format(Val(strTemp) - (intY * 0.4), "#0.00")
                            Next intY
                        End If
                    
                    Case 2
                        Call SET_CELL(vaSpread1, 3, 12, Trim(Mid(strSubData(intX), 3, 19)))
                        Call SET_CELL(vaSpread1, 3, 13, Trim(Mid(strSubData(intX), 22, 5)))
                        strTemp = Trim(Mid(strSubData(intX), 22, 5))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 3, 20 + intY, Format(Val(strTemp) + (intY * 0.5), "#0.00")
                            Next intY
                        End If
                        Call SET_CELL(vaSpread1, 3, 14, Trim(Mid(strSubData(intX), 27, 6)))
                        strTemp = Trim(Replace(Mid(strSubData(intX), 27, 6), " ", ""))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 4, 20 + intY, Format(Val(strTemp) - (intY * 0.4), "#0.00")
                            Next intY
                        End If
                    
                    Case 3
                        Call SET_CELL(vaSpread1, 4, 12, Trim(Mid(strSubData(intX), 3, 19)))
                        Call SET_CELL(vaSpread1, 4, 13, Trim(Mid(strSubData(intX), 22, 5)))
                        strTemp = Trim(Mid(strSubData(intX), 22, 5))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 1, 29 + intY, Format(Val(strTemp) + (intY * 0.5), "#0.00")
                            Next intY
                        End If
                        Call SET_CELL(vaSpread1, 4, 14, Trim(Mid(strSubData(intX), 27, 6)))
                        strTemp = Trim(Replace(Mid(strSubData(intX), 27, 6), " ", ""))
                        If IsNumeric(strTemp) = True Then
                            For intY = -3 To 3
                                vaSpread1.SetText 2, 29 + intY, Format(Val(strTemp) - (intY * 0.4), "#0.00")
                            Next intY
                        End If
                End Select
            
            Case "WR"
            Case "WL"
        End Select
    Next intX
    
    With vaSpread1
        '/Step3.Zan Image Printer(color)
        strTemp = GET_DEFAULT_PRINTER '/기존 설정된 기본프린터 찾기
        Call SET_DEFAULT_PRINTER(gtypEQ_INFO.ZIPNM)  '/임시로 가상프린터를 기본프린터로 지정
'''        Call ZanPrinterSetting '/장비Image기본폴더 및 파일명 설정.

        '/Step4.출력(이미지생성)
        Dim strFont1  As String
        Dim strFont2  As String
        Dim strHead1  As String

        strFont1 = "/fn""굴림체""/fz""15""/fb1/fi0/fu1/fk0/fs1"
        strFont2 = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs2"

        strHead1 = "/f1/c" & "인공수정체진단기 결과" & "/n/n/n"

        .PrintAbortMsg = "인공수정체진단기 결과 이미지 출력 중..."
        .PrintHeader = strFont1 + strHead1 + strFont2
        '''.PrintFooter = "/c" & "PAGE : " & "/P"
        .PrintBorder = True
        .PrintGrid = True
        .PrintColHeaders = False
        .PrintRowHeaders = False
        .PrintColor = True
        .PrintMarginTop = 500
        .PrintMarginBottom = 500
        .PrintMarginLeft = 500
        .PrintMarginRight = 0
        .PrintType = PrintTypeAll
        .PrintShadows = False
        .PrintUseDataMax = False
        .Action = ActionSmartPrint

        Call SET_DEFAULT_PRINTER(strTemp) '/기존 설정된 기본프린터 지정
    End With

    On Error GoTo 0
    
''''/00008(인공수정체진단기)-AL2000
'''Type AL2000
'''    PT      As String '/검사일시(mid(PT,23,14))
'''    BM      As String '/결과시작점 알림
'''    HR      As String '/Right Header(Eye Type = mid(HR,3,14), Vavg = mid(HR,17,4), Vlens = mid(HR,21,4))
'''    VR      As String '/Right(Vacd = mid(HR,3,4))
'''    LR      As String '/Right(AXIAL, ACD, LENS) 각각 소수점 2자리 고정
'''    KR      As String '/Right(K1, K2) 각각 소수점 2자리 고정
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
'''    LL      As String '/Left(AXIAL, ACD, LENS) 각각 소수점 2자리 고정
'''    KL      As String '/Left(K1, K2) 각각 소수점 2자리 고정
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
    






'''    Dim strData             As String
'''    Dim strSubData(1 To 23) As String
'''    Dim lngCnt                As Long
'''    Dim intCnt              As Integer
'''    Dim ii                  As Integer
'''    Dim strAXIAL            As String
'''    Dim strK1               As String
'''    Dim strK2               As String
'''    Dim strDR               As String
'''    Dim strAconst1          As String
'''    Dim strAconst2          As String
'''    Dim strAconst3          As String
'''    Dim strPower1           As String
'''    Dim strPower2           As String
'''    Dim strPower3           As String
'''    Dim strIOL1             As String
'''    Dim strIOL2             As String
'''    Dim strIOL3             As String
'''    Dim strRef1             As String
'''    Dim strRef2             As String
'''    Dim strRef3             As String
'''    Dim iRow                As Integer
'''
'''    On Error Resume Next
'''
'''    '/Step1.자료 변환
'''    strData = argData
'''    intCnt = 0
'''    For lngCnt = 1 To Len(strData)
'''        If Mid(strData, lngCnt, 1) = Chr(1) Or Mid(strData, lngCnt, 1) = Chr(2) Then
'''            intCnt = intCnt + 1
'''            If intCnt > 23 Then
'''                Exit For
'''            End If
'''
'''            strSubData(intCnt) = ""
'''        ElseIf Mid(strData, lngCnt, 1) = Chr(13) Then
''''            sprPrint.SetText 1, intCnt, strSubData(intCnt)
'''        Else
'''            strSubData(intCnt) = strSubData(intCnt) & Mid(strData, lngCnt, 1)
'''        End If
'''    Next
'''
'''    strAXIAL = Replace(Trim(Mid(strSubData(5), 3)), " ", "")
'''    strK1 = Replace(Trim(Mid(strSubData(6), 3, 5)), " ", "")
'''    strK2 = Replace(Trim(Mid(strSubData(6), 8, 5)), " ", "")
'''    strDR = Replace(Trim(Mid(strSubData(7), 3)), " ", "")
'''
'''    strAconst1 = Replace(Trim(Mid(strSubData(10), 3, 10)), " ", "")
'''    strAconst2 = Replace(Trim(Mid(strSubData(12), 3, 10)), " ", "")
'''    strAconst3 = Replace(Trim(Mid(strSubData(14), 3, 10)), " ", "")
'''
'''    strPower1 = Replace(Trim(Mid(strSubData(10), 22, 5)), " ", "")
'''    strPower2 = Replace(Trim(Mid(strSubData(12), 22, 5)), " ", "")
'''    strPower3 = Replace(Trim(Mid(strSubData(14), 22, 5)), " ", "")
'''
'''    strRef1 = Replace(Trim(Mid(strSubData(10), 27, 6)), " ", "")
'''    strRef2 = Replace(Trim(Mid(strSubData(12), 27, 6)), " ", "")
'''    strRef3 = Replace(Trim(Mid(strSubData(14), 27, 6)), " ", "")
'''
'''    With sprPrint
'''        '/Step2.자료 Setting
'''        sprPrint.SetText 2, 4, strAXIAL
'''        sprPrint.SetText 2, 5, strK1
'''        sprPrint.SetText 2, 6, strK2
'''        sprPrint.SetText 2, 7, strDR
'''
'''        sprPrint.SetText 2, 9, strAconst1
'''        sprPrint.SetText 3, 9, strAconst2
'''        sprPrint.SetText 4, 9, strAconst3
'''
'''        sprPrint.SetText 2, 10, strPower1
'''        sprPrint.SetText 3, 10, strPower2
'''        sprPrint.SetText 4, 10, strPower3
'''
'''        If IsNumeric(strPower1) = True Then
'''            iRow = 16
'''            For ii = -3 To 3
'''                sprPrint.SetText 1, iRow + ii, Format(CStr(CCur(strPower1) + (ii * 0.5)), "#0.00")
'''            Next ii
'''        End If
'''        If IsNumeric(strPower2) = True Then
'''            iRow = 16
'''            For ii = -3 To 3
'''                sprPrint.SetText 3, iRow + ii, Format(CStr(CCur(strPower2) + (ii * 0.5)), "#0.00")
'''            Next ii
'''        End If
'''        If IsNumeric(strPower3) = True Then
'''            iRow = 25
'''            For ii = -3 To 3
'''                sprPrint.SetText 1, iRow + ii, Format(CStr(CCur(strPower3) + (ii * 0.5)), "#0.00")
'''            Next ii
'''        End If
'''
'''        If IsNumeric(strRef1) = True Then
'''            iRow = 16
'''            For ii = -3 To 3
'''                sprPrint.SetText 2, iRow + ii, Format(CStr(CCur(strRef1) - (ii * 0.4)), "#0.00")
'''            Next ii
'''
'''        End If
'''        If IsNumeric(strRef2) = True Then
'''            iRow = 16
'''            For ii = -3 To 3
'''                sprPrint.SetText 4, iRow + ii, Format(CStr(CCur(strRef2) - (ii * 0.4)), "#0.00")
'''            Next ii
'''        End If
'''        If IsNumeric(strRef3) = True Then
'''            iRow = 25
'''            For ii = -3 To 3
'''                sprPrint.SetText 2, iRow + ii, Format(CStr(CCur(strRef3) - (ii * 0.4)), "#0.00")
'''            Next ii
'''        End If
'''
'''        '/Step3.Zan Image Printer(color)
'''        strTemp = GET_DEFAULT_PRINTER '/기존 설정된 기본프린터 찾기
'''        Call SET_DEFAULT_PRINTER(gtypEQ_INFO.ZIPNM)  '/임시로 가상프린터를 기본프린터로 지정
'''        Call ZanPrinterSetting '/장비Image기본폴더 및 파일명 설정.
'''
'''        '/Step4.출력(이미지생성)
'''        Dim strFont1  As String
'''        Dim strFont2  As String
'''        Dim strHead1  As String
'''
'''        strFont1 = "/fn""굴림체""/fz""15""/fb1/fi0/fu1/fk0/fs1"
'''        strFont2 = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs2"
'''
'''        strHead1 = "/f1/c" & "인공수정체진단기 결과" & "/n/n/n"
'''
'''        .PrintAbortMsg = "인공수정체진단기 결과 이미지 출력 중..."
'''        .PrintHeader = strFont1 + strHead1 + strFont2
'''        '''.PrintFooter = "/c" & "PAGE : " & "/P"
'''        .PrintBorder = True
'''        .PrintGrid = True
'''        .PrintColHeaders = False
'''        .PrintRowHeaders = False
'''        .PrintColor = True
'''        .PrintMarginTop = 500
'''        .PrintMarginBottom = 500
'''        .PrintMarginLeft = 500
'''        .PrintMarginRight = 0
'''        .PrintType = PrintTypeAll
'''        .PrintShadows = False
'''        .PrintUseDataMax = False
'''        .Action = ActionSmartPrint
'''
'''        Call SET_DEFAULT_PRINTER(strTemp) '/기존 설정된 기본프린터 지정
'''    End With
'''
'''    On Error GoTo 0
End Sub

Private Sub cmdPrint_Click()
    For intX = 1 To Len(txtInput)
        strTemp = Mid(txtInput, intX, 1)
        
        Select Case strTemp
            Case Chr(1) 'SOH
                txtBuff = ""
                txtBuff = strTemp
            Case Chr(4) 'EOT
                txtBuff = txtBuff & strTemp
                Call FUNC_MM_PRINT(txtBuff)
                txtBuff = ""
            Case Else
                txtBuff = txtBuff & strTemp
        End Select
    Next intX
End Sub

