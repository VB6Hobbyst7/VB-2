VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmVPM_SM00025 
   Caption         =   "00025-자동굴절계/Topcon/KR7100"
   ClientHeight    =   10665
   ClientLeft      =   2970
   ClientTop       =   2400
   ClientWidth     =   9945
   Icon            =   "frmVPM_SM00025.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   9945
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   8655
      Left            =   60
      TabIndex        =   5
      Top             =   1320
      Width           =   8715
      _Version        =   393216
      _ExtentX        =   15372
      _ExtentY        =   15266
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   16
      SpreadDesigner  =   "frmVPM_SM00025.frx":9F8A
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
      Top             =   900
      Width           =   4935
   End
   Begin VB.TextBox txtInput 
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "txtBuff"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   120
      Width           =   1410
   End
End
Attribute VB_Name = "frmVPM_SM00025"
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
'''    MsgBox "공사 중...", vbInformation, "확인"

    Dim strSubData()
    Dim strResult()
    Dim lngCnt                  As Long
    Dim intCnt                  As Integer
    Dim intResultCntR           As Integer
    Dim intResultCntL           As Integer

    On Error Resume Next

    '/Step1.자료 변환
    intCnt = 0
    ReDim Preserve strSubData(intCnt)
    For lngCnt = 1 To Len(argData)
        If Mid(argData, lngCnt, 1) = Chr(1) Or Mid(argData, lngCnt, 1) = Chr(2) Then
            intCnt = intCnt + 1
            ReDim Preserve strSubData(intCnt)
            strSubData(intCnt) = ""
        ElseIf Mid(argData, lngCnt, 1) = Chr(13) Then
            intCnt = intCnt + 1
            ReDim Preserve strSubData(intCnt)
        ElseIf Mid(argData, lngCnt, 1) = Chr(10) Then
        
        Else
            strSubData(intCnt) = strSubData(intCnt) & Mid(argData, lngCnt, 1)
        End If
    Next lngCnt

    For intX = 1 To UBound(strSubData)
        Select Case intX
            Case 2: Call SET_CELL(vaSpread1, 7, 4, strSubData(intX))
            Case 3: Call SET_CELL(vaSpread1, 8, 4, strSubData(intX))
            Case 4: Call SET_CELL(vaSpread1, 9, 4, strSubData(intX))
            
            Case 7: Call SET_CELL(vaSpread1, 2, 4, strSubData(intX))
            Case 8: Call SET_CELL(vaSpread1, 3, 4, strSubData(intX))
            Case 9: Call SET_CELL(vaSpread1, 4, 4, strSubData(intX))
        
            Case 14: Call SET_CELL(vaSpread1, 7, 7, strSubData(intX))
            Case 15: Call SET_CELL(vaSpread1, 8, 7, strSubData(intX))
            Case 16: Call SET_CELL(vaSpread1, 9, 7, strSubData(intX))
        
            Case 17: Call SET_CELL(vaSpread1, 7, 8, strSubData(intX))
            Case 18: Call SET_CELL(vaSpread1, 8, 8, strSubData(intX))
            Case 19: Call SET_CELL(vaSpread1, 9, 8, strSubData(intX))
        
            Case 20: Call SET_CELL(vaSpread1, 7, 10, strSubData(intX))
            Case 21: Call SET_CELL(vaSpread1, 8, 10, strSubData(intX))
            
            Case 22: Call SET_CELL(vaSpread1, 8, 12, strSubData(intX))
            Case 23: Call SET_CELL(vaSpread1, 9, 12, strSubData(intX))
            
            Case 25: Call SET_CELL(vaSpread1, 2, 7, strSubData(intX))
            Case 26: Call SET_CELL(vaSpread1, 3, 7, strSubData(intX))
            Case 27: Call SET_CELL(vaSpread1, 4, 7, strSubData(intX))
        
            Case 28: Call SET_CELL(vaSpread1, 2, 8, strSubData(intX))
            Case 29: Call SET_CELL(vaSpread1, 3, 8, strSubData(intX))
            Case 30: Call SET_CELL(vaSpread1, 4, 8, strSubData(intX))
        
            Case 31: Call SET_CELL(vaSpread1, 2, 10, strSubData(intX))
            Case 32: Call SET_CELL(vaSpread1, 3, 10, strSubData(intX))
            
            Case 33: Call SET_CELL(vaSpread1, 3, 12, strSubData(intX))
            Case 34: Call SET_CELL(vaSpread1, 4, 12, strSubData(intX))
        
            Case 11: Call SET_CELL(vaSpread1, 1, 15, "PD = " & strSubData(intX) & "mm")
        
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

