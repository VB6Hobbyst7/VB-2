VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmVPM_SM00014 
   Caption         =   "00014-렌즈미터//CHAROPS-LM"
   ClientHeight    =   6120
   ClientLeft      =   2580
   ClientTop       =   3795
   ClientWidth     =   6675
   Icon            =   "frmVPM_SM00014.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   540
      Width           =   4935
   End
   Begin VB.TextBox txtInput 
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox txtBuff 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin FPSpread.vaSpread sprPrint 
      Height          =   3315
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   3795
      _Version        =   393216
      _ExtentX        =   6694
      _ExtentY        =   5847
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   2
      MaxRows         =   8
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "frmVPM_SM00014.frx":9F8A
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
      Top             =   180
      Width           =   1410
   End
End
Attribute VB_Name = "frmVPM_SM00014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FUNC_MM_PRINT(argData As String)
    With sprPrint
        '/Step1.자료 Setting
        .ClearRange -1, -1, -1, -1, -1
        
        .SetText 1, 1, Mid(argData, 9, 8) '/NO:
        
        '/RIGHT
        .SetText 1, 2, Mid(argData, 67, 8)  '/PD
        .SetText 1, 3, "<RIGHT>"            '/<RIGHT>
        .SetText 1, 4, Mid(argData, 23, 9)  '/S
        .SetText 1, 5, Mid(argData, 32, 9)  '/C
        .SetText 1, 6, Mid(argData, 41, 6)  '/A
        .SetText 1, 7, Mid(argData, 47, 10) '/PX
        .SetText 1, 8, Mid(argData, 57, 10) '/PY
        
        '/LEFT
        .SetText 2, 2, Mid(argData, 141, 8)  '/PD
        .SetText 2, 3, "<LEFT>"              '/<LEFT>
        .SetText 2, 4, Mid(argData, 97, 9)   '/S
        .SetText 2, 5, Mid(argData, 106, 9)  '/C
        .SetText 2, 6, Mid(argData, 115, 6)  '/A
        .SetText 2, 7, Mid(argData, 121, 10) '/PX
        .SetText 2, 8, Mid(argData, 131, 10) '/PY
        
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
    
        strHead1 = "/f1/c" & "LensMeter 결과" & "/n/n/n"
    
        .PrintAbortMsg = "LensMeter 결과 이미지 출력 중..."
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
            Case vbCr
            
            Case vbLf
                If Mid(txtBuff, 1, 5) = "LM2RK" Then
                    Call FUNC_MM_PRINT(txtBuff)
                End If
                txtBuff = ""
            Case Else
                txtBuff = txtBuff & strTemp
        End Select
    Next intX
End Sub

