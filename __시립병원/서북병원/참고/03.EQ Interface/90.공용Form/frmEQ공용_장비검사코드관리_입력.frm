VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEQ공용_장비검사코드관리_입력 
   BorderStyle     =   1  '단일 고정
   Caption         =   "장비검사코드관리"
   ClientHeight    =   7275
   ClientLeft      =   6540
   ClientTop       =   2715
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_장비검사코드관리_입력.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9315
   Begin VB.Frame fra상세내역 
      Caption         =   "[상세내역]"
      Height          =   6075
      Left            =   60
      TabIndex        =   30
      Top             =   1140
      Width           =   9195
      Begin VB.ComboBox cboEQORDYN 
         Height          =   300
         Left            =   1680
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   600
         Width           =   1515
      End
      Begin FPSpread.vaSpread sprEXCD 
         Height          =   5415
         Left            =   6420
         TabIndex        =   26
         Top             =   540
         Width           =   2655
         _Version        =   393216
         _ExtentX        =   4683
         _ExtentY        =   9551
         _StockProps     =   64
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         SpreadDesigner  =   "frmEQ공용_장비검사코드관리_입력.frx":263A
      End
      Begin VB.Frame Frame3 
         Caption         =   "[CutOff] 장비결과 문자변환"
         Height          =   2115
         Left            =   120
         TabIndex        =   47
         Top             =   3840
         Width           =   6135
         Begin VB.ComboBox cboEQCUTRTYPE 
            Height          =   300
            Left            =   1020
            Style           =   2  '드롭다운 목록
            TabIndex        =   25
            Top             =   1680
            Width           =   4995
         End
         Begin VB.ComboBox cboEQCUTOFFNM 
            Height          =   300
            Left            =   1020
            Style           =   2  '드롭다운 목록
            TabIndex        =   24
            Top             =   1320
            Width           =   4995
         End
         Begin VB.ComboBox cboEQCUTMNM 
            Height          =   300
            Left            =   1020
            Style           =   2  '드롭다운 목록
            TabIndex        =   23
            Top             =   960
            Width           =   4995
         End
         Begin VB.TextBox txtEQCUTLVAL 
            Height          =   300
            Left            =   4020
            TabIndex        =   21
            Text            =   "txtEQCUTLVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cboEQCUTLREF 
            Height          =   300
            Left            =   4920
            Style           =   2  '드롭다운 목록
            TabIndex        =   22
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtEQCUTHVAL 
            Height          =   300
            Left            =   1020
            TabIndex        =   19
            Text            =   "txtEQCUTHVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cboEQCUTHREF 
            Height          =   300
            Left            =   1980
            Style           =   2  '드롭다운 목록
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cboEQCUTOFFGB 
            Height          =   300
            Left            =   1020
            Style           =   2  '드롭다운 목록
            TabIndex        =   18
            Top             =   240
            Width           =   4995
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "표시형태"
            Height          =   180
            Index           =   19
            Left            =   180
            TabIndex        =   53
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "값 형 태"
            Height          =   180
            Index           =   18
            Left            =   180
            TabIndex        =   52
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "중 간 값"
            Height          =   180
            Index           =   17
            Left            =   180
            TabIndex        =   51
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "하 한 값"
            Height          =   180
            Index           =   16
            Left            =   3240
            TabIndex        =   50
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "상 한 값"
            Height          =   180
            Index           =   15
            Left            =   180
            TabIndex        =   49
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "적용구분"
            Height          =   180
            Index           =   14
            Left            =   180
            TabIndex        =   48
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "[Limit] 장비결과 수치제한"
         Height          =   675
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   6135
         Begin VB.ComboBox cboEQLIMITFLAG2 
            Height          =   300
            Left            =   4020
            Style           =   2  '드롭다운 목록
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtEQLIMITVALUE2 
            Height          =   300
            Left            =   5160
            TabIndex        =   17
            Text            =   "txtEQLIMITVALUE2"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtEQLIMITVALUE1 
            Height          =   300
            Left            =   1980
            TabIndex        =   15
            Text            =   "txtEQLIMITVALUE1"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboEQLIMITFLAG1 
            Height          =   300
            Left            =   840
            Style           =   2  '드롭다운 목록
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "상한값"
            Height          =   180
            Index           =   13
            Left            =   3300
            TabIndex        =   46
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "하한값"
            Height          =   180
            Index           =   8
            Left            =   180
            TabIndex        =   45
            Top             =   300
            Width           =   540
         End
      End
      Begin MSMask.MaskEdBox mskEQSEQ 
         Height          =   195
         Left            =   4980
         TabIndex        =   4
         Top             =   1020
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame1 
         Caption         =   "[정상참고치]"
         Height          =   1035
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   6135
         Begin VB.ComboBox cboEQRFLREF 
            Height          =   300
            Left            =   1740
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtEQRFLVAL 
            Height          =   300
            Left            =   840
            TabIndex        =   10
            Text            =   "txtEQAFLVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtEQRFHVAL 
            Height          =   300
            Left            =   4020
            TabIndex        =   12
            Text            =   "txtEQAFHVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cboEQRFHREF 
            Height          =   300
            Left            =   4920
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cboEQRMLREF 
            Height          =   300
            Left            =   1740
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtEQRMLVAL 
            Height          =   300
            Left            =   840
            TabIndex        =   6
            Text            =   "txtEQAMLVAL"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtEQRMHVAL 
            Height          =   300
            Left            =   4020
            TabIndex        =   8
            Text            =   "txtEQAMHVAL"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboEQRMHREF 
            Height          =   300
            Left            =   4920
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "여 Low"
            Height          =   180
            Index           =   12
            Left            =   180
            TabIndex        =   40
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "여 High"
            Height          =   180
            Index           =   10
            Left            =   3300
            TabIndex        =   39
            Top             =   660
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "남 Low"
            Height          =   180
            Index           =   11
            Left            =   180
            TabIndex        =   38
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "남 High"
            Height          =   180
            Index           =   5
            Left            =   3300
            TabIndex        =   37
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.TextBox txtEQUNIT 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   195
         IMEMode         =   8  '영문
         Left            =   1680
         TabIndex        =   3
         Text            =   "1234567890"
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtEQNM 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   195
         IMEMode         =   8  '영문
         Left            =   1680
         TabIndex        =   1
         Text            =   "12345678901234567890123456789012345678901234567890"
         Top             =   300
         Width           =   4575
      End
      Begin MSMask.MaskEdBox mskEQRSTRANGE 
         Height          =   195
         Left            =   2220
         TabIndex        =   5
         Top             =   1380
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   120
         X2              =   6240
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "장비전송여부"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   55
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "처방코드"
         Height          =   195
         Index           =   6
         Left            =   6420
         TabIndex        =   54
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "※ 장비수신결과가 문자일 경우 0으로 표시"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "(0: 전체표시, 1 이상 : 숫자만큼 표시)"
         Height          =   180
         Index           =   2
         Left            =   2700
         TabIndex        =   42
         Top             =   1380
         Width           =   3330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   120
         X2              =   6240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "소수점이하 표시 자리수"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   41
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정렬순서"
         Height          =   195
         Index           =   4
         Left            =   3420
         TabIndex        =   35
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   120
         X2              =   6240
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "검사결과단위"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "장비검사명"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   1515
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   120
         X2              =   6240
         Y1              =   540
         Y2              =   540
      End
   End
   Begin VB.TextBox txtEQCD 
      Appearance      =   0  '평면
      BorderStyle     =   0  '없음
      Height          =   195
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8340
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7380
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   60
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   32
      Top             =   600
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbl장비명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장비검사코드입력"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   33
      Top             =   60
      Width           =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   7
      X1              =   120
      X2              =   2700
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "장비검사코드"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   780
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmEQ공용_장비검사코드관리_입력"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SUB_MM_CANCEL() As Boolean
    barStatus.Max = 100
    barStatus.Value = 100
    
    txtEQCD = ""
    
    Call SUB_MM_KEY_CLEAR
End Function

Public Function MM_DELETE() As Boolean

End Function

Private Sub SUB_MM_INITIAL()
    Me.Height = 7755
    Me.Width = 9435
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    
    Call SUB_MM_CANCEL

    If gstrInputUpdate = "2" Then '/1.Input, 2.Update
        txtEQCD = gstrArgTemp1
        txtEQCD.BackColor = RGB(255, 255, 240)
        txtEQCD.Enabled = False
        Call FUNC_MM_VIEW
    End If
    
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:
    '/장비검사코드 장비전송여부(Y.전송, N.미전송)
    cboEQORDYN.AddItem "전송" & Space(100) & "Y"
    cboEQORDYN.AddItem "미전송" & Space(100) & "N"
    
    '/Reference
    cboEQRMHREF.AddItem ""
    cboEQRMHREF.AddItem "이상" & Space(100) & "1"
    cboEQRMHREF.AddItem "초과" & Space(100) & "2"
    cboEQRMHREF.AddItem "이하" & Space(100) & "3"
    cboEQRMHREF.AddItem "미만" & Space(100) & "4"

    cboEQRMLREF.AddItem ""
    cboEQRMLREF.AddItem "이상" & Space(100) & "1"
    cboEQRMLREF.AddItem "초과" & Space(100) & "2"
    cboEQRMLREF.AddItem "이하" & Space(100) & "3"
    cboEQRMLREF.AddItem "미만" & Space(100) & "4"

    cboEQRFHREF.AddItem ""
    cboEQRFHREF.AddItem "이상" & Space(100) & "1"
    cboEQRFHREF.AddItem "초과" & Space(100) & "2"
    cboEQRFHREF.AddItem "이하" & Space(100) & "3"
    cboEQRFHREF.AddItem "미만" & Space(100) & "4"

    cboEQRFLREF.AddItem ""
    cboEQRFLREF.AddItem "이상" & Space(100) & "1"
    cboEQRFLREF.AddItem "초과" & Space(100) & "2"
    cboEQRFLREF.AddItem "이하" & Space(100) & "3"
    cboEQRFLREF.AddItem "미만" & Space(100) & "4"

    '/LIMIT 하한값 구분(0.미적용, 1.이하, 2.미만)
    cboEQLIMITFLAG1.AddItem "미적용" & Space(100) & "0"
    cboEQLIMITFLAG1.AddItem "이하" & Space(100) & "1"
    cboEQLIMITFLAG1.AddItem "미만" & Space(100) & "2"

    '/LIMIT 상한값 구분(0.미적용, 1.이상, 2.초과)
    cboEQLIMITFLAG2.AddItem "미적용" & Space(100) & "0"
    cboEQLIMITFLAG2.AddItem "이상" & Space(100) & "1"
    cboEQLIMITFLAG2.AddItem "초과" & Space(100) & "2"

    '/CUTOFF 적용구분(0.적용안함, 1.상한 Positive, 2.하한 Positive, 3.장비결과값적용(수치와 CutOff값이 동시에 나올경우)
    cboEQCUTOFFGB.AddItem "적용안함" & Space(100) & "0"
    cboEQCUTOFFGB.AddItem "상한 Positive" & Space(100) & "1"
    cboEQCUTOFFGB.AddItem "하한 Positive" & Space(100) & "2"
    cboEQCUTOFFGB.AddItem "장비결과값적용" & Space(100) & "3"
        
    '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
    cboEQCUTHREF.AddItem ""
    cboEQCUTHREF.AddItem "이상" & Space(100) & "1"
    cboEQCUTHREF.AddItem "초과" & Space(100) & "2"
            
    '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
    cboEQCUTLREF.AddItem ""
    cboEQCUTLREF.AddItem "이하" & Space(100) & "1"
    cboEQCUTLREF.AddItem "미만" & Space(100) & "2"
            
    '/CUTOFF 중간값(CUTOFF 상,하한값이 다를 경우 적용 1.Grayzone, 2.Weakly positive, 3.Low Titer) 코드값은 사용자 요구로 변할 수 있음
    cboEQCUTMNM.AddItem ""
    cboEQCUTMNM.AddItem "Grayzone" & Space(100) & "1"
    cboEQCUTMNM.AddItem "Weakly positive" & Space(100) & "2"
    cboEQCUTMNM.AddItem "Low Titer" & Space(100) & "3"
            
    '/CUTOFF 값형태(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) 코드값은 사용자 요구로 변할 수 있음
    cboEQCUTOFFNM.AddItem ""
    cboEQCUTOFFNM.AddItem "Negative/Positive" & Space(100) & "1"
    cboEQCUTOFFNM.AddItem "Neg/Pos" & Space(100) & "2"
    cboEQCUTOFFNM.AddItem "Nonreactive/Reactive" & Space(100) & "3"
    cboEQCUTOFFNM.AddItem "NEGATIVE/POSITIVE" & Space(100) & "4"
    cboEQCUTOFFNM.AddItem "NEG/POS" & Space(100) & "5"
            
    '/CUTOFF 표시형태(1.Negative/Positive, 2.Negative/Positive(수치), 3.Negative/Grayzone(수치)/Positive(수치), 4.Negative(수치)/Grayzone(수치)/Positive(수치)
    cboEQCUTRTYPE.AddItem ""
    cboEQCUTRTYPE.AddItem "Negative/Positive" & Space(100) & "1"
    cboEQCUTRTYPE.AddItem "Negative/Positive(수치)" & Space(100) & "2"
    cboEQCUTRTYPE.AddItem "Negative/Grayzone(수치)/Positive(수치)" & Space(100) & "3"
    cboEQCUTRTYPE.AddItem "Negative(수치)/Grayzone(수치)/Positive(수치)" & Space(100) & "4"
Return
End Sub

Public Function MM_INPUT() As Boolean

End Function

Private Sub SUB_MM_KEY_CLEAR()
    fra상세내역.Enabled = False
    
    txtEQNM = ""
    cboEQORDYN.ListIndex = 0
    txtEQUNIT = ""
    mskEQRSTRANGE = "0"
    mskEQSEQ = ""
    
    txtEQRMHVAL = ""
    txtEQRMLVAL = ""
    txtEQRFHVAL = ""
    txtEQRFLVAL = ""

    cboEQRMHREF.ListIndex = -1      '/Reference 남자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
    cboEQRMLREF.ListIndex = -1      '/Reference 남자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
    cboEQRFHREF.ListIndex = -1      '/Reference 여자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
    cboEQRFLREF.ListIndex = -1      '/Reference 여자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
    
    cboEQLIMITFLAG1.ListIndex = 0   '/LIMIT 하한값 구분(0.미적용, 1.이하, 2.미만)
    cboEQLIMITFLAG2.ListIndex = 0   '/LIMIT 상한값 구분(0.미적용, 1.이상, 2.초과)
    
    cboEQCUTOFFGB.ListIndex = 0     '/CUTOFF 적용구분
    cboEQCUTHREF.ListIndex = -1     '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
    cboEQCUTLREF.ListIndex = -1     '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
    cboEQCUTMNM.ListIndex = -1      '/CUTOFF 중간값(CUTOFF 상,하한값이 다를 경우 적용 1.Grayzone, 2.Weakly positive, 3.Low Titer) 코드값은 사용자 요구로 변할 수 있음
    cboEQCUTOFFNM.ListIndex = -1    '/CUTOFF 값형태(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) 코드값은 사용자 요구로 변할 수 있음
    cboEQCUTRTYPE.ListIndex = -1    '/CUTOFF 표시형태(1.Negative/Positive, 2.Negative/Positive(수치), 3.Negative/Grayzone(수치)/Positive(수치), 4.Negative(수치)/Grayzone(수치)/Positive(수치)
    
    If sprEXCD.MaxRows > 0 Then sprEXCD.MaxRows = 0: sprEXCD.MaxRows = 1
End Sub

Public Function MM_PRINT() As Boolean

End Function

Public Function FUNC_MM_SAVE() As Boolean
    FUNC_MM_SAVE = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    ADC_LOC.BeginTrans
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        ADR_LOC.Close: Set ADR_LOC = Nothing
    
        gstrQuy = "UPDATE EQ_MST SET "
        gstrQuy = gstrQuy & vbCrLf & "       EQNM           = '" & Trim(TEXT_LSET(txtEQNM, 50)) & "', "     '/장비검사명
        gstrQuy = gstrQuy & vbCrLf & "       EQUNIT         = '" & Trim(TEXT_LSET(txtEQUNIT, 10)) & "', "   '/결과단위
        gstrQuy = gstrQuy & vbCrLf & "       EQSEQ          =  " & Val(mskEQSEQ) & ",  "                    '/화면정렬순서
        gstrQuy = gstrQuy & vbCrLf & "       EQRSTRANGE     =  " & Val(mskEQRSTRANGE) & ", "                '/장비검사결과 소수점표시( 0: 전체표시, >1 : 숫자만큼 표시)
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITFLAG1   = '" & Trim(Right(cboEQLIMITFLAG1, 10)) & "', " '/LIMIT 하한값 구분(0.미적용, 1.이하, 2.미만)
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITVALUE1  = '" & Trim(txtEQLIMITVALUE1) & "', "           '/LIMIT 하한값
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITFLAG2   = '" & Trim(Right(cboEQLIMITFLAG2, 10)) & "', " '/LIMIT 상한값 구분(0.미적용, 1.이상, 2.초과)
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITVALUE2  = '" & Trim(txtEQLIMITVALUE2) & "', "           '/LIMIT 상한값
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTOFFGB     = '" & Trim(Right(cboEQCUTOFFGB, 10)) & "', "   '/CUTOFF 적용구분(0.적용안함, 1.상한 Positive, 2.하한 Positive, 3.장비결과값적용(수치와 CutOff값이 동시에 나올경우)
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTOFFNM     = '" & Trim(Right(cboEQCUTOFFNM, 10)) & "', "   '/CUTOFF 값형태(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) 코드값은 사용자 요구로 변할 수 있음
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTHVAL      = '" & Trim(txtEQCUTHVAL) & "', "               '/CUTOFF 상한값
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTHREF      = '" & Trim(Right(cboEQCUTHREF, 10)) & "', "    '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTLVAL      = '" & Trim(txtEQCUTLVAL) & "', "               '/CUTOFF 하한값
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTLREF      = '" & Trim(Right(cboEQCUTLREF, 10)) & "', "    '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTMNM       = '" & Trim(Right(cboEQCUTMNM, 10)) & "', "     '/CUTOFF 중간값(CUTOFF 상,하한값이 다를 경우 적용 1.Grayzone, 2.Weakly positive, 3.Low Titer) 코드값은 사용자 요구로 변할 수 있음
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTRTYPE     = '" & Trim(Right(cboEQCUTRTYPE, 10)) & "', "   '/CUTOFF 표시형태(1.Negative/Positive, 2.Negative/Positive(수치), 3.Negative/Grayzone(수치)/Positive(수치), 4.Negative(수치)/Grayzone(수치)/Positive(수치)
        gstrQuy = gstrQuy & vbCrLf & "       EQRMHVAL       = '" & Trim(txtEQRMHVAL) & "', "                '/Reference 남자 High Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRMHREF       = '" & Trim(Right(cboEQRMHREF, 10)) & "', "     '/Reference 남자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
        gstrQuy = gstrQuy & vbCrLf & "       EQRMLVAL       = '" & Trim(txtEQRMLVAL) & "', "                '/Reference 남자 Low Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRMLREF       = '" & Trim(Right(cboEQRMLREF, 10)) & "', "     '/Reference 남자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
        gstrQuy = gstrQuy & vbCrLf & "       EQRFHVAL       = '" & Trim(txtEQRFHVAL) & "', "                '/Reference 여자 High Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRFHREF       = '" & Trim(Right(cboEQRFHREF, 10)) & "', "     '/Reference 여자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
        gstrQuy = gstrQuy & vbCrLf & "       EQRFLVAL       = '" & Trim(txtEQRFLVAL) & "', "                '/Reference 여자 Low Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRFLREF       = '" & Trim(Right(cboEQRFLREF, 10)) & "' "      '/Reference 여자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD           = '" & Trim(txtEQCD) & "' "
    Else
        gstrQuy = "INSERT INTO EQ_MST "
        gstrQuy = gstrQuy & vbCrLf & " (EQCD,           EQNM,           EQUNIT,         EQSEQ,          EQRSTRANGE, "
        gstrQuy = gstrQuy & vbCrLf & "  EQLIMITFLAG1,   EQLIMITVALUE1,  EQLIMITFLAG2,   EQLIMITVALUE2,  EQCUTOFFGB, "
        gstrQuy = gstrQuy & vbCrLf & "  EQCUTOFFNM,     EQCUTHVAL,      EQCUTHREF,      EQCUTLVAL,      EQCUTLREF, "
        gstrQuy = gstrQuy & vbCrLf & "  EQCUTMNM,       EQCUTRTYPE,     EQRMHVAL,       EQRMHREF,       EQRMLVAL, "
        gstrQuy = gstrQuy & vbCrLf & "  EQRMLREF,       EQRFHVAL,       EQRFHREF,       EQRFLVAL,       EQRFLREF) "
        gstrQuy = gstrQuy & vbCrLf & " VALUES "
        gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(TEXT_LSET(txtEQCD, 10)) & "', "     '/장비검사코드
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(TEXT_LSET(txtEQNM, 50)) & "', "     '/장비검사명
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(TEXT_LSET(txtEQUNIT, 10)) & "', "   '/결과단위
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(mskEQSEQ) & ",  "                    '/화면정렬순서
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(mskEQRSTRANGE) & ", "                '/장비검사결과 소수점표시( 0: 전체표시, >1 : 숫자만큼 표시)
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQLIMITFLAG1, 10)) & "', " '/LIMIT 하한값 구분(0.미적용, 1.이하, 2.미만)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQLIMITVALUE1) & "', "           '/LIMIT 하한값
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQLIMITFLAG2, 10)) & "', " '/LIMIT 상한값 구분(0.미적용, 1.이상, 2.초과)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQLIMITVALUE2) & "', "           '/LIMIT 상한값
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTOFFGB, 10)) & "', "   '/CUTOFF 적용구분(0.적용안함, 1.상한 Positive, 2.하한 Positive, 3.장비결과값적용(수치와 CutOff값이 동시에 나올경우)
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTOFFNM, 10)) & "', "   '/CUTOFF 값형태(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) 코드값은 사용자 요구로 변할 수 있음
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQCUTHVAL) & "', "               '/CUTOFF 상한값
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTHREF, 10)) & "', "    '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQCUTLVAL) & "', "               '/CUTOFF 하한값
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTLREF, 10)) & "', "    '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTMNM, 10)) & "', "     '/CUTOFF 중간값(CUTOFF 상,하한값이 다를 경우 적용 1.Grayzone, 2.Weakly positive, 3.Low Titer) 코드값은 사용자 요구로 변할 수 있음
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTRTYPE, 10)) & "', "   '/CUTOFF 표시형태(1.Negative/Positive, 2.Negative/Positive(수치), 3.Negative/Grayzone(수치)/Positive(수치), 4.Negative(수치)/Grayzone(수치)/Positive(수치)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRMHVAL) & "', "                '/Reference 남자 High Value
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRMHREF, 10)) & "', "     '/Reference 남자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRMLVAL) & "', "                '/Reference 남자 Low Value
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRMLREF, 10)) & "', "     '/Reference 남자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRFHVAL) & "', "                '/Reference 여자 High Value
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRFHREF, 10)) & "', "     '/Reference 여자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRFLVAL) & "', "                '/Reference 여자 Low Value
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRFLREF, 10)) & "') "     '/Reference 여자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
    End If
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    gstrQuy = "DELETE FROM EX_MST "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    For intX = 1 To sprEXCD.DataRowCnt
        If Trim(GET_CELL(sprEXCD, 1, intX)) <> "" Then
            gstrQuy = "SELECT * "
            gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EXCD = '" & Trim(GET_CELL(sprEXCD, 1, intX)) & "' "
            If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
            
            If Not ADR_LOC Is Nothing Then
                ADR_LOC.Close: Set ADR_LOC = Nothing
            Else
                gstrQuy = "INSERT INTO EX_MST (EQCD, EXCD, EQORDREADYN, EQRESSENDYN) "
                gstrQuy = gstrQuy & vbCrLf & " VALUES "
                gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(txtEQCD) & "',"                                           '/장비검사코드
                gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(GET_CELL(sprEXCD, 1, intX)) & "', "                       '/처방코드
                gstrQuy = gstrQuy & vbCrLf & "  '" & IIf(Trim(GET_CELL(sprEXCD, 2, intX)) = "1", "Y", "N") & "', "  '/HIS처방읽기여부(Y.대상, N.비대상)
                gstrQuy = gstrQuy & vbCrLf & "  '" & IIf(Trim(GET_CELL(sprEXCD, 3, intX)) = "1", "Y", "N") & "') "   '/장비검사결과 HIS전송여부(Y.대상, N.비대상)
                If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
            End If
        End If
    Next intX
    
    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_MM_SAVE = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_MM_VIEW() As Boolean
    FUNC_MM_VIEW = False
    
On Error GoTo RTN_ERR

    If Trim(txtEQCD) = "" Then Exit Function
    
    Call SUB_MM_KEY_CLEAR
    
    If ConnDB_LOC = False Then Exit Function
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        If gstrInputUpdate = "2" Then '/1.Input, 2.Update
            fra상세내역.Enabled = True
            
            txtEQNM = Trim(ADR_LOC!EQNM & "")               '/장비검사명
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQORDYN & ""), cboEQORDYN) '/장비검사코드 장비전송여부(Y.전송, N.미전송)
            txtEQUNIT = Trim(ADR_LOC!EQUNIT & "")           '/검사결과단위
            mskEQSEQ = Trim(ADR_LOC!EQSEQ & "")             '/화면정렬순서
            mskEQRSTRANGE = Trim(ADR_LOC!EQRSTRANGE & "")   '/소수점자릿수
            
            txtEQRMLVAL = Trim(ADR_LOC!EQRMLVAL & "")                   '/Reference 남자 Low Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRMLREF & ""), cboEQRMLREF) '/Reference 남자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
            txtEQRMHVAL = Trim(ADR_LOC!EQRMHVAL & "")                   '/Reference 남자 High Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRMHREF & ""), cboEQRMHREF) '/Reference 남자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
            txtEQRFLVAL = Trim(ADR_LOC!EQRFLVAL & "")                   '/Reference 여자 Low Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRFLREF & ""), cboEQRFLREF) '/Reference 여자 Low Reference(1.이상, 2.초과, 3.이하, 4.미만)
            txtEQRFHVAL = Trim(ADR_LOC!EQRFHVAL & "")                   '/Reference 여자 High Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRFHREF & ""), cboEQRFHREF) '/Reference 여자 High Reference(1.이상, 2.초과, 3.이하, 4.미만)
        
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQLIMITFLAG1 & ""), cboEQLIMITFLAG1) '/LIMIT 하한값 구분(0.미적용, 1.이하, 2.미만)
            txtEQLIMITVALUE1 = Trim(ADR_LOC!EQLIMITVALUE1 & "")                 '/LIMIT 하한값
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQLIMITFLAG2 & ""), cboEQLIMITFLAG2) '/LIMIT 상한값 구분(0.미적용, 1.이상, 2.초과)
            txtEQLIMITVALUE2 = Trim(ADR_LOC!EQLIMITVALUE2 & "")                 '/LIMIT 상한값
        
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTOFFGB & ""), cboEQCUTOFFGB) '/CUTOFF 구분(0.적용안함, 1.상한 Positive, 2.하한 Positive, 3.장비결과값적용(수치와 CutOff값이 동시에 나올경우)
            txtEQCUTHVAL = Trim(ADR_LOC!EQCUTHVAL & "")                     '/CUTOFF 상한값
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTHREF & ""), cboEQCUTHREF)   '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
            txtEQCUTLVAL = Trim(ADR_LOC!EQCUTLVAL & "")                     '/CUTOFF 하한값
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTLREF & ""), cboEQCUTLREF)   '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTMNM & ""), cboEQCUTMNM)     '/CUTOFF 중간값(CUTOFF 상,하한값이 다를 경우 적용 1.Grayzone, 2.Weakly positive, 3.Low Titer) 코드값은 사용자 요구로 변할 수 있음
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTOFFNM & ""), cboEQCUTOFFNM) '/CUTOFF 값형태(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) 코드값은 사용자 요구로 변할 수 있음
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTRTYPE & ""), cboEQCUTRTYPE) '/CUTOFF 표시형태(1.Negative/Positive, 2.Negative/Positive(수치), 3.Negative/Grayzone(수치)/Positive(수치), 4.Negative(수치)/Grayzone(수치)/Positive(수치)
        Else
            MsgBox "기존자료가 있습니다!", vbInformation, "확인"
        End If
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        With sprEXCD
            If .MaxRows > 0 Then .MaxRows = 0
            
            gstrQuy = "SELECT * "
            gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
            gstrQuy = gstrQuy & vbCrLf & " ORDER BY EXCD "
            If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
            
            If Not ADR_LOC Is Nothing Then
                Do Until ADR_LOC.EOF
                    .MaxRows = .MaxRows + 1: .Row = .MaxRows
                
                    .Col = 1: .Text = Trim(ADR_LOC!EXCD & "")
                    .Col = 2: .Text = IIf(Trim(ADR_LOC!EQORDREADYN & "") = "Y", "1", "0")
                    .Col = 3: .Text = IIf(Trim(ADR_LOC!EQRESSENDYN & "") = "Y", "1", "0")
                    
                    ADR_LOC.MoveNext
                Loop
            End If
            
            .MaxRows = .MaxRows + 1
        End With
    Else
        If gstrInputUpdate = "1" Then '/1.Input, 2.Update
            fra상세내역.Enabled = True
            
            txtEQNM.SetFocus
        End If
    End If

    Call CloseDB_LOC
    
    FUNC_MM_VIEW = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Private Sub cboEQCUTHREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTLREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTMNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTOFFGB_Click()
    '/CUTOFF 적용구분(0.적용안함, 1.상한 Positive, 2.하한 Positive, 3.장비결과값적용(수치와 CutOff값이 동시에 나올경우)
    If Trim(Right(cboEQCUTOFFGB, 10)) = "0" Then
        txtEQCUTHVAL = ""               '/CUTOFF 상한값
        cboEQCUTHREF.ListIndex = -1     '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
        txtEQCUTLVAL = ""               '/CUTOFF 하한값
        cboEQCUTLREF.ListIndex = -1     '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
        cboEQCUTMNM.ListIndex = -1      '/CUTOFF 중간값(CUTOFF 상,하한값이 다를 경우 적용 1.Grayzone, 2.Weakly positive, 3.Low Titer) 코드값은 사용자 요구로 변할 수 있음
        cboEQCUTOFFNM.ListIndex = -1    '/CUTOFF 값형태(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) 코드값은 사용자 요구로 변할 수 있음
        cboEQCUTRTYPE.ListIndex = -1    '/CUTOFF 표시형태(1.Negative/Positive, 2.Negative/Positive(수치), 3.Negative/Grayzone(수치)/Positive(수치), 4.Negative(수치)/Grayzone(수치)/Positive(수치)
    
        txtEQCUTHVAL.Enabled = False
        cboEQCUTHREF.Enabled = False
        txtEQCUTLVAL.Enabled = False
        cboEQCUTLREF.Enabled = False
        cboEQCUTMNM.Enabled = False
        cboEQCUTOFFNM.Enabled = False
        cboEQCUTRTYPE.Enabled = False
    Else
        txtEQCUTHVAL.Enabled = True
        cboEQCUTHREF.Enabled = True
        txtEQCUTLVAL.Enabled = True
        cboEQCUTLREF.Enabled = True
        cboEQCUTMNM.Enabled = True
        cboEQCUTOFFNM.Enabled = True
        cboEQCUTRTYPE.Enabled = True
    End If
End Sub

Private Sub cboEQCUTOFFGB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTOFFNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTRTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQLIMITFLAG1_Click()
    '/LIMIT 하한값 구분(0.미적용, 1.이하, 2.미만)
    If Trim(Right(cboEQLIMITFLAG1, 10)) = "0" Then
        txtEQLIMITVALUE1 = ""
        txtEQLIMITVALUE1.Enabled = False
    Else
        txtEQLIMITVALUE1.Enabled = True
    End If
End Sub

Private Sub cboEQLIMITFLAG1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQLIMITFLAG2_Click()
    '/LIMIT 상한값 구분(0.미적용, 1.이상, 2.초과)
    If Trim(Right(cboEQLIMITFLAG2, 10)) = "0" Then
        txtEQLIMITVALUE2 = ""
        txtEQLIMITVALUE2.Enabled = False
    Else
        txtEQLIMITVALUE2.Enabled = True
    End If
End Sub

Private Sub cboEQLIMITFLAG2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQORDYN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRFHREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRFLREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRMHREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRMLREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtEQCD) = "" Then MsgBox "장비검사코드를 (재)입력하십시오!", vbInformation, "확인": txtEQCD.SetFocus: Exit Sub
    
    If IsNumeric(mskEQSEQ) = False Then
        MsgBox "화면정렬순서를 (재)입력하십시오!" & vbCrLf & _
               "입력형식은 숫자타입입니다.", vbInformation, "확인": mskEQSEQ.SetFocus: Exit Sub
    End If
    
    If IsNumeric(mskEQRSTRANGE) = False Then
        MsgBox "소수점이하 표시 자리수를 (재)입력하십시오!" & vbCrLf & _
               "입력형식은 숫자타입입니다.", vbInformation, "확인": mskEQRSTRANGE.SetFocus: Exit Sub
    End If
    
    '/정상참고치(남Low)
    If Trim(txtEQRMLVAL) = "" Or Trim(cboEQRMLREF) = "" Then '/둘중 하나가 공백일때
        If Not (Trim(txtEQRMLVAL) = "" And Trim(cboEQRMLREF) = "") Then '/둘다 공백이 아닐때
            If Trim(txtEQRMLVAL) = "" Then MsgBox "[정상참고치] 남 Low 값을 (재)입력하십시오!", vbInformation, "확인": txtEQRMLVAL.SetFocus: Exit Sub
            If Trim(cboEQRMLREF) = "" Then MsgBox "[정상참고치] 남 Low 기준을 (재)입력하십시오!", vbInformation, "확인": cboEQRMLREF.SetFocus: Exit Sub
        End If
    End If
    
    '/정상참고치(남High)
    If Trim(txtEQRMHVAL) = "" Or Trim(cboEQRMHREF) = "" Then '/둘중 하나가 공백일때
        If Not (Trim(txtEQRMHVAL) = "" And Trim(cboEQRMHREF) = "") Then '/둘다 공백이 아닐때
            If Trim(txtEQRMHVAL) = "" Then MsgBox "[정상참고치] 남 High 값을 (재)입력하십시오!", vbInformation, "확인": txtEQRMHVAL.SetFocus: Exit Sub
            If Trim(cboEQRMHREF) = "" Then MsgBox "[정상참고치] 남 High 기준을 (재)입력하십시오!", vbInformation, "확인": cboEQRMHREF.SetFocus: Exit Sub
        End If
    End If
    
    '/정상참고치(여Low)
    If Trim(txtEQRFLVAL) = "" Or Trim(cboEQRFLREF) = "" Then '/둘중 하나가 공백일때
        If Not (Trim(txtEQRFLVAL) = "" And Trim(cboEQRFLREF) = "") Then '/둘다 공백이 아닐때
            If Trim(txtEQRFLVAL) = "" Then MsgBox "[정상참고치] 여 Low 값을 (재)입력하십시오!", vbInformation, "확인": txtEQRFLVAL.SetFocus: Exit Sub
            If Trim(cboEQRFLREF) = "" Then MsgBox "[정상참고치] 여 Low 기준을 (재)입력하십시오!", vbInformation, "확인": cboEQRFLREF.SetFocus: Exit Sub
        End If
    End If
    
    '/정상참고치(여High)
    If Trim(txtEQRFHVAL) = "" Or Trim(cboEQRFHREF) = "" Then '/둘중 하나가 공백일때
        If Not (Trim(txtEQRFHVAL) = "" And Trim(cboEQRFHREF) = "") Then '/둘다 공백이 아닐때
            If Trim(txtEQRFHVAL) = "" Then MsgBox "[정상참고치] 여 High 값을 (재)입력하십시오!", vbInformation, "확인": txtEQRFHVAL.SetFocus: Exit Sub
            If Trim(cboEQRFHREF) = "" Then MsgBox "[정상참고치] 여 High 기준을 (재)입력하십시오!", vbInformation, "확인": cboEQRFHREF.SetFocus: Exit Sub
        End If
    End If
    
    '/CUTOFF 상한값 Equal 포함 여부(1.이상, 2.초과)
    If Trim(txtEQCUTHVAL) = "" Or Trim(cboEQCUTHREF) = "" Then '/둘중 하나가 공백일때
        If Not (Trim(txtEQCUTHVAL) = "" And Trim(cboEQCUTHREF) = "") Then '/둘다 공백이 아닐때
            If Trim(txtEQCUTHVAL) = "" Then MsgBox "CUTOFF 상한값을 (재)입력하십시오!", vbInformation, "확인": txtEQCUTHVAL.SetFocus: Exit Sub
            If Trim(cboEQCUTHREF) = "" Then MsgBox "CUTOFF 상한 기준을 (재)입력하십시오!", vbInformation, "확인": cboEQCUTHREF.SetFocus: Exit Sub
        End If
    End If
    
    '/CUTOFF 하한값 Equal 포함 여부(1.이하, 2.미만)
    If Trim(txtEQCUTLVAL) = "" Or Trim(cboEQCUTLREF) = "" Then '/둘중 하나가 공백일때
        If Not (Trim(txtEQCUTLVAL) = "" And Trim(cboEQCUTLREF) = "") Then '/둘다 공백이 아닐때
            If Trim(txtEQCUTLVAL) = "" Then MsgBox "CUTOFF 하한값을 (재)입력하십시오!", vbInformation, "확인": txtEQCUTLVAL.SetFocus: Exit Sub
            If Trim(cboEQCUTLREF) = "" Then MsgBox "CUTOFF 하한 기준을 (재)입력하십시오!", vbInformation, "확인": cboEQCUTLREF.SetFocus: Exit Sub
        End If
    End If
    
    If MsgBox("저장하겠습니까?", vbQuestion + vbOKCancel, "저장질의") = vbCancel Then Exit Sub
    
    If FUNC_MM_SAVE = True Then
        gstrInputUpdateYN = True '/저장 성공여부 Set(광역변수)
        
        MsgBox "저장되었습니다!", vbInformation, "확인"
        
        Call SUB_MM_CANCEL
        
        If gstrInputUpdate = "1" Then '/1.Input, 2.Update(신규 및 수정 시 화면처리)
            txtEQCD.SetFocus
        Else
            Unload Me
        End If
    Else
        gstrInputUpdateYN = False '/저장 성공여부 Set(광역변수)
        MsgBox "저장오류!", vbCritical, "확인"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Set frmEQ공용_장비검사코드관리_입력 = Nothing
End Sub

Private Sub sprEXCD_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If ChangeMade = False Then Exit Sub
    If Col <> 1 Then Exit Sub
    
    If sprEXCD.DataRowCnt = sprEXCD.MaxRows Then
        sprEXCD.MaxRows = sprEXCD.MaxRows + 1
    End If
End Sub

Private Sub txtEQCD_Change()
    Call SUB_MM_KEY_CLEAR
End Sub

Private Sub txtEQCD_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call FUNC_MM_VIEW
End Sub

Private Sub txtEQCUTHVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQCUTHVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQCUTLVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQCUTLVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQLIMITVALUE1_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQLIMITVALUE1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQLIMITVALUE2_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQLIMITVALUE2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRFHVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRFHVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRFLVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRFLVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRMHVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRMHVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRMLVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRMLVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQNM_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub mskEQSEQ_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub mskEQSEQ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQUNIT_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub mskEQRSTRANGE_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub mskEQRSTRANGE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
