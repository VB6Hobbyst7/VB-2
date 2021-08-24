VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmVPM_SM00016 
   Caption         =   "00016-자동안압계/Topcon/CT-80"
   ClientHeight    =   6975
   ClientLeft      =   8580
   ClientTop       =   5835
   ClientWidth     =   6675
   Icon            =   "frmVPM_SM00016.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtBuff 
      Height          =   330
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
      Height          =   5595
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
      _Version        =   393216
      _ExtentX        =   8705
      _ExtentY        =   9869
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   30
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmVPM_SM00016.frx":9F8A
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "txtBuff"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1020
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
Attribute VB_Name = "frmVPM_SM00016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FUNC_MM_PRINT(asData As String)
    Dim strData            As String
    Dim strSubData(1 To 5) As String
    
    With sprPrint
        '/Step1.자료 Setting
        strData = asData
        strData = Replace(strData, Chr(1), "")
        strData = Replace(strData, Chr(2), "")
        strData = Replace(strData, Chr(3), "")
        strData = Replace(strData, Chr(4), "")
        
        intY = 1
        
        For intX = 1 To Len(strData)
            Select Case Mid(strData, intX, 1)
                Case Chr(13)
                    .SetText 1, intY, strSubData(intY)
                    
                    intY = intY + 1
                    
                    If intY > 5 Then Exit For
                    
                    strSubData(intY) = ""
            
                Case Chr(10)
                
                Case Else
                    strSubData(intY) = strSubData(intY) & Mid(strData, intX, 1)
            End Select
        Next
        
        '/Step3.Zan Image Printer(color)
        strTemp = GET_DEFAULT_PRINTER '/기존 설정된 기본프린터 찾기
        Call SET_DEFAULT_PRINTER(gtypEQ_INFO.ZIPNM) '/임시로 가상프린터를 기본프린터로 지정
'''        Call ZanPrinterSetting

        '/Step3.출력(이미지생성)
        Dim strFont1  As String
        Dim strFont2  As String
        Dim strHead1  As String
    
        strFont1 = "/fn""굴림체""/fz""15""/fb1/fi0/fu1/fk0/fs1"
        strFont2 = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs2"
    
        strHead1 = "/f1/c" & "자동안압계 결과" & "/n/n/n"
    
        .PrintAbortMsg = "자동안압계 결과 이미지 출력 중..."
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
End Sub

Private Sub cmdPrint_Click()
    For intX = 1 To Len(txtInput)
        strTemp = Mid(txtInput, intX, 1)
            
        Select Case strTemp
            Case Chr(1)
                txtBuff = strTemp
            
            Case Chr(4)
                txtBuff = txtBuff & strTemp
                Call FUNC_MM_PRINT(txtBuff)
                
                txtBuff = ""
            
            Case Else
                txtBuff = txtBuff & strTemp
        End Select
    Next intX
End Sub

