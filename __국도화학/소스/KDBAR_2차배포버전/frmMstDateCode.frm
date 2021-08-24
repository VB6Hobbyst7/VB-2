VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMstDateCode 
   BackColor       =   &H00FFFFFF&
   Caption         =   "날짜코드 범례"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   15615
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1065
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "테스트"
      Height          =   315
      Left            =   1950
      TabIndex        =   4
      Top             =   240
      Width           =   1005
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   810
      TabIndex        =   3
      Top             =   240
      Width           =   1065
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   240
      Width           =   465
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "닫기"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14310
      Style           =   1  '그래픽
      TabIndex        =   0
      ToolTipText     =   "현재화면을 모두 지웁니다"
      Top             =   150
      Width           =   1095
   End
   Begin FPSpread.vaSpread spdDate 
      Height          =   6645
      Left            =   240
      TabIndex        =   1
      Top             =   630
      Width           =   15165
      _Version        =   393216
      _ExtentX        =   26749
      _ExtentY        =   11721
      _StockProps     =   64
      ColsFrozen      =   8
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   9
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16774636
      SpreadDesigner  =   "frmMstDateCode.frx":0000
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmMstDateCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdTest_Click()
    Dim strTmp
    
    strTmp = Get_YMD(txtCode.Text, txtInput.Text)
    
    txtOutput.Text = strTmp
    
End Sub

'-----------------------------------------------------------------------------'
'   파일명  : frmMstDateCode.frm
'   작성자  : 오세원
'   내  용  : 날짜코드 범례
'   작성일  : 2020-02-23
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub Form_Load()
    
    Call CtlInitializing
    
    Call SetValue

End Sub

Private Sub CtlInitializing()

    With spdDate
        .MaxRows = 18
        .MaxCols = 5
        .FontName = "돋움체"
        .FontBold = True
        
        Call SetText(spdDate, "구분", 0, 1):    .ColWidth(1) = 10
        Call SetText(spdDate, "코드", 0, 2):    .ColWidth(2) = 10
        Call SetText(spdDate, "타입", 0, 3):    .ColWidth(3) = 30
        Call SetText(spdDate, "표기", 0, 4):    .ColWidth(4) = 40
        Call SetText(spdDate, "설명", 0, 5):    .ColWidth(5) = 30
        
    
    End With

End Sub

Private Sub SetValue()
    Dim i   As Integer
    
    i = 1
    
    With spdDate
        Call SetText(spdDate, "년", i, 1):  Call SetText(spdDate, "Y1", i, 2):  Call SetText(spdDate, "4자리 숫자", i, 3):  Call SetText(spdDate, "2020=2020, 2021=2021, 2022=2022 ...", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "Y2", i, 2):  Call SetText(spdDate, "2자리 숫자", i, 3):  Call SetText(spdDate, "2020=20,   2021=21,   2022=22 ...", i, 4):   Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "Y3", i, 2):  Call SetText(spdDate, "1자리 숫자", i, 3):  Call SetText(spdDate, "2020=0,    2021=1,    2022=2 ...", i, 4):    Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "Y4", i, 2):  Call SetText(spdDate, "1자리 영문", i, 3):  Call SetText(spdDate, "2010=A,    2021=B,    2022=C ...", i, 4):    Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "Y5", i, 2):  Call SetText(spdDate, "1자리 영문", i, 3):  Call SetText(spdDate, "2011=A,    2012=B,    2013=C ...", i, 4):    Call SetText(spdDate, "I,O,U,V 제외됨", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "Y6", i, 2):  Call SetText(spdDate, "1자리 영문", i, 3):  Call SetText(spdDate, "2010=A,    2011=B,    2012=C ...", i, 4):    Call SetText(spdDate, "N,O 제외됨", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "Y7", i, 2):  Call SetText(spdDate, "1자리 영문", i, 3):  Call SetText(spdDate, "2011=A,    2012=B,    2013=C ...", i, 4):    Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "", i, 2):    Call SetText(spdDate, "", i, 3):            Call SetText(spdDate, "", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1

        Call SetText(spdDate, "월", i, 1):  Call SetText(spdDate, "M1", i, 2):  Call SetText(spdDate, "2자리 숫자", i, 3):                      Call SetText(spdDate, "01=01, 02=02 ... 10=10, 11=1, 12=12", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "M2", i, 2):  Call SetText(spdDate, "1자리 숫자 & 영문", i, 3):               Call SetText(spdDate, "01=1,  02=2  ... 10=A,  11=B, 12=C", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "M3", i, 2):  Call SetText(spdDate, "1자리 숫자 & 영문(월) 첫글자", i, 3):    Call SetText(spdDate, "01=1,  02=2  ... 10=O,  11=N, 12=D", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "", i, 2):    Call SetText(spdDate, "", i, 3): Call SetText(spdDate, "", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1

        Call SetText(spdDate, "일", i, 1):  Call SetText(spdDate, "D1", i, 2):  Call SetText(spdDate, "2자리 숫자", i, 3):          Call SetText(spdDate, "01=01,02=02 ... 30=30, 31=31", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "D2", i, 2):  Call SetText(spdDate, "1자리 숫자 & 영문", i, 3):   Call SetText(spdDate, "01=1, 02=2  ... 10=A ... 35=Z", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "D3", i, 2):  Call SetText(spdDate, "1자리 숫자 & 영문", i, 3):   Call SetText(spdDate, "01=1, 02=2  ... 10=A ... 33=Z", i, 4): Call SetText(spdDate, "I,O 제외됨", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "D4", i, 2):  Call SetText(spdDate, "1자리 숫자 & 영문", i, 3):   Call SetText(spdDate, "01=1, 02=2  ... 10=A ... 31=Z", i, 4): Call SetText(spdDate, "I,O,U,V 제외됨", i, 5): i = i + 1
        Call SetText(spdDate, "", i, 1):    Call SetText(spdDate, "", i, 2):    Call SetText(spdDate, "", i, 3): Call SetText(spdDate, "", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1

        Call SetText(spdDate, "제품길이", i, 1): Call SetText(spdDate, "L1", i, 2): Call SetText(spdDate, "1자리 영문", i, 3): Call SetText(spdDate, "A=10, B=100, C=1000, D=10000", i, 4): Call SetText(spdDate, "", i, 5): i = i + 1

        For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            .TypeHAlign = TypeHAlignLeft
            .TypeVAlign = TypeVAlignCenter
        Next
    
    End With
    
End Sub

Private Sub spdDate_Click(ByVal Col As Long, ByVal Row As Long)
    
    With spdDate
        
        txtCode.Text = GetText(spdDate, Row, 2)
        
    End With

End Sub
