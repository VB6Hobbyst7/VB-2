VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm108AccCancel 
   BackColor       =   &H00DBE6E6&
   Caption         =   "채혈/접수 취소"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00800000&
      Caption         =   "전체선택(&A)"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   7815
      TabIndex        =   42
      Top             =   2715
      Width           =   1560
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   9450
      TabIndex        =   41
      Top             =   2685
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "접수취소 사유"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   5910
      TabIndex        =   39
      Top             =   45
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "환자기본정보"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   315
      Left            =   75
      TabIndex        =   38
      Top             =   45
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "접수번호"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   5340
      Left            =   75
      TabIndex        =   36
      Tag             =   "10114"
      Top             =   3030
      Width           =   9345
      _Version        =   196608
      _ExtentX        =   16484
      _ExtentY        =   9419
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      MaxCols         =   8
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis108.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   75
      TabIndex        =   13
      Top             =   285
      Width           =   5820
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1470
         MaxLength       =   12
         TabIndex        =   58
         Top             =   360
         Width           =   2300
      End
      Begin VB.OptionButton seLOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Tag             =   "125"
         Top             =   405
         Width           =   1155
      End
      Begin VB.OptionButton seLOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4050
         TabIndex        =   56
         Tag             =   "125"
         Top             =   405
         Width           =   1155
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DBE6E6&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1425
         TabIndex        =   17
         Top             =   705
         Width           =   3825
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "D/C"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   165
            TabIndex        =   37
            Tag             =   "125"
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "처방상태로.."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1005
            TabIndex        =   3
            Tag             =   "125"
            Top             =   180
            Width           =   1395
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈상태로.."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2415
            TabIndex        =   4
            Tag             =   "123"
            Top             =   180
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1425
         ScaleHeight     =   300
         ScaleWidth      =   2355
         TabIndex        =   14
         Top             =   315
         Width           =   2415
         Begin VB.TextBox txtAccSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1725
            TabIndex        =   2
            Top             =   30
            Width           =   600
         End
         Begin VB.TextBox txtAccDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            MaxLength       =   6
            TabIndex        =   1
            Top             =   30
            Width           =   810
         End
         Begin VB.TextBox txtWorkArea 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   45
            MaxLength       =   2
            TabIndex        =   0
            Top             =   30
            Width           =   465
         End
         Begin VB.Label Label17 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1530
            TabIndex        =   16
            Top             =   -45
            Width           =   210
         End
         Begin VB.Label Label2 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   525
            TabIndex        =   15
            Top             =   -30
            Width           =   210
         End
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "상태"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   570
         TabIndex        =   35
         Top             =   900
         Width           =   450
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00DF6A3E&
         FillStyle       =   0  '단색
         Height          =   360
         Left            =   240
         Shape           =   4  '둥근 사각형
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Option"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   405
         TabIndex        =   28
         Tag             =   "151"
         Top             =   885
         Width           =   540
      End
   End
   Begin VB.Frame fraPtInfo 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   5910
      TabIndex        =   12
      Tag             =   "104"
      Top             =   285
      Width           =   8550
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   330
         Left            =   3900
         TabIndex        =   18
         Top             =   585
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDob 
         Height          =   330
         Left            =   1050
         TabIndex        =   19
         Top             =   585
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   3900
         TabIndex        =   20
         Top             =   165
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtId 
         Height          =   330
         Left            =   1035
         TabIndex        =   25
         Top             =   165
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   330
         Left            =   6660
         TabIndex        =   26
         Top             =   585
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   330
         Left            =   1050
         TabIndex        =   27
         Top             =   1020
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   90
         TabIndex        =   49
         Top             =   165
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "환자 ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   2925
         TabIndex        =   50
         Top             =   165
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "성 명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   5715
         TabIndex        =   51
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   90
         TabIndex        =   52
         Top             =   585
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "생년월일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   10
         Left            =   2925
         TabIndex        =   53
         Top             =   585
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   11
         Left            =   5715
         TabIndex        =   54
         Top             =   585
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "병실"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   90
         TabIndex        =   55
         Top             =   1020
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7590
         TabIndex        =   23
         Top             =   225
         Width           =   60
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   7905
         TabIndex        =   22
         Top             =   255
         Width           =   345
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6705
         TabIndex        =   21
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Caption         =   "             /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6660
         TabIndex        =   24
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1035
      Left            =   75
      TabIndex        =   11
      Top             =   1650
      Width           =   14385
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   330
         Left            =   180
         TabIndex        =   29
         Top             =   555
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblColDtTm 
         Height          =   330
         Left            =   2070
         TabIndex        =   30
         Top             =   555
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   4710
         TabIndex        =   31
         Top             =   555
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblRcvDtTm 
         Height          =   330
         Left            =   7785
         TabIndex        =   32
         Top             =   555
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblRcvNm 
         Height          =   330
         Left            =   10500
         TabIndex        =   33
         Top             =   555
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   44
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검체"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   45
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   4710
         TabIndex        =   46
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   7785
         TabIndex        =   47
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   10500
         TabIndex        =   48
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "접수자"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   5460
      Left            =   9465
      TabIndex        =   8
      Top             =   2925
      Width           =   4995
      Begin VB.ComboBox cboReason 
         BackColor       =   &H00FCEFE9&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1065
         Style           =   2  '드롭다운 목록
         TabIndex        =   34
         Top             =   225
         Width           =   3720
      End
      Begin VB.TextBox txtReason 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   615
         Width           =   4620
      End
      Begin VB.CheckBox chkRefund 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Refund to Patient"
         Height          =   240
         Left            =   150
         TabIndex        =   9
         Tag             =   "10404"
         Top             =   5130
         Visible         =   0   'False
         Width           =   2385
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   135
         TabIndex        =   43
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "취소사유"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "접수취소(&S)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "123"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료 (&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   40
      Top             =   2685
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "처방 내역"
      LeftGab         =   100
   End
End
Attribute VB_Name = "frm108AccCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tmpAccDt As String
Private MySql As New clsLISSqlStatement

Private MultiFg As String
Private MultiLabNo()
Private ClearFg As Boolean

Private objCanAcc As New clsLISAccCancel
Private objAccData As New clsDictionary

Private Sub cboReason_Click()
    txtReason.Text = medGetP(cboReason.Text, 2, ":")
End Sub

Private Sub chkAll_Click()
    With tblOrdSheet
        .Col = 8: .COL2 = 8
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .value = chkAll.value
        .BlockMode = False
    End With
End Sub

Private Sub cmdCancel_Click()

    Dim i As Integer
    Dim strStsCd As String
    Dim blnCancel As Boolean
    Dim lngCnt As Long
    Dim sFlag   As Long
    
    If Trim(txtReason.Text) = "" Then
        MsgBox "접수취소 사유를 반드시 입력하십시오..", vbCritical + vbOKOnly, "메세지"
        txtReason.SetFocus
        Exit Sub
    End If
    
    lngCnt = 0
    For i = 1 To tblOrdSheet.DataRowCnt
        tblOrdSheet.Row = i
        tblOrdSheet.Col = 8
        If tblOrdSheet.value = 1 Then lngCnt = lngCnt + 1
    Next
    If lngCnt <= 0 Then
        MsgBox "접수취소할 항목을 선택하십시오..", vbExclamation, "메세지"
        Exit Sub
    End If
    
    On Error GoTo Err_Trap
    
    '백업처방인지 OCS처방인지를 구분한다. OCS처방일 경우 D/C를 할 수 없다.
    Dim RS          As New Recordset
    Dim strSQL      As String
    
    If optOption(2).value = True Then
        With objAccData
            .MoveFirst
            strSQL = " SELECT * FROM " & T_LAB102 & _
                     "  WHERE " & DBW("workarea", .Fields("workarea"), 2) & _
                     "    AND " & DBW("accdt", .Fields("accdt"), 2) & _
                     "    AND " & DBW("accseq", .Fields("accseq"), 2)
            RS.Open strSQL, DBConn
            
            If RS.RecordCount > 0 Then
                MsgBox "OCS처방은 D/C 할 수 없습니다.처방상태로 하신 뒤 오더 D/C를 요청하세요", vbCritical, "조회오류"
                Call cmdClear_Click
                Set RS = Nothing
                Exit Sub
            End If
            Set RS = Nothing
        End With
    End If
    
    If optOption(2).value = True Then 'DC 처방
        sFlag = 3
        strStsCd = "D"
    Else    '처방상태, 채혈상태
        strStsCd = IIf(optOption(0).value, enStsCd.StsCd_LIS_Cancel, enStsCd.StsCd_LIS_Collection)
        sFlag = IIf(optOption(0).value, 1, 2)
    End If
    
    If lngCnt = tblOrdSheet.DataRowCnt Then
        '## 전체취소
        DBConn.BeginTrans
        With objAccData
            .MoveFirst
            blnCancel = objCanAcc.DoCancelAccession(.Fields("ptid"), .Fields("workarea"), .Fields("accdt"), .Fields("accseq"), _
                                             strStsCd, ObjSysInfo.EmpId, txtReason.Text, tblOrdSheet, sFlag)
            If Not blnCancel Then GoTo Err_Trap
        End With
'        blnCancel = OCSActingCheck
'        If Not blnCancel Then GoTo Err_Trap
        
        DBConn.CommitTrans
    Else
        '## 부분취소
        DBConn.BeginTrans
        With objAccData
            .MoveFirst
            blnCancel = objCanAcc.DoCancelPart(.Fields("ptid"), .Fields("workarea"), .Fields("accdt"), .Fields("accseq"), _
                                             strStsCd, ObjSysInfo.EmpId, txtReason.Text, tblOrdSheet, sFlag)
            If Not blnCancel Then GoTo Err_Trap
        End With
'        blnCancel = OCSActingCheck
'        If Not blnCancel Then GoTo Err_Trap
        
        DBConn.CommitTrans
    End If
    
    Call ClearRtn
    DoEvents
    txtWorkArea.SetFocus
    Exit Sub
    
Err_Trap:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Function OCSActingCheck() As Boolean
    Dim RS          As Recordset
    Dim SqlStmt     As String
    Dim strOcsOrdNo As String
    Dim strBussdiv  As String
    Dim strPtid     As String
    Dim strOrdNo    As String
    Dim strOrdDt    As String
    Dim strOrdSeq   As String
    
    Dim ii          As Integer
    
On Error GoTo Errors

    strPtid = lblPtId.Caption
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 8
            If .value = "1" Then
                .Col = 1: strOrdDt = Replace(Trim(.value), "-", "")
                .Col = 2: strOrdNo = Trim(.value)
                .Col = 7: strOrdSeq = Trim(.value)
                '접수시 OCS 관련 Table 에 Acting_Check를 해준다.
                
                SqlStmt = " SELECT a.ocsordno,b.bussdiv " & _
                          " FROM " & T_LAB101 & " b," & T_LAB102 & " a" & _
                          " WHERE " & DBW("a.ptid =", strPtid) & _
                          " AND " & DBW("a.orddt=", strOrdDt) & _
                          " AND " & DBW("a.ordno=", strOrdNo) & _
                          " AND " & DBW("a.ordseq=", strOrdSeq) & _
                          " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno"
                          
                Set RS = Nothing
                Set RS = New Recordset
                RS.Open SqlStmt, DBConn
                
                If Not RS.EOF Then
                    strOcsOrdNo = Val(Trim(RS.Fields("ocsordno").value & ""))
                    strBussdiv = Trim(RS.Fields("bussdiv").value & "")
                    If strOcsOrdNo <> "" Then
                        '병동은 ipd_order_dmc,ipd_order_update_dmc 업데이트
                        '외래는 opd_order_dmc 업데이트
                        If strBussdiv = enBussDiv.BussDiv_InPatient Then
                            SqlStmt = " UPDATE med_ocs.ipd_order_dmc SET acting_check='0' where order_key=" & strOcsOrdNo
                            DBConn.Execute SqlStmt
                            SqlStmt = " UPDATE med_ocs.ipd_order_update_dmc SET acting_check='0' where order_key=" & strOcsOrdNo
                        Else
                            SqlStmt = " UPDATE med_ocs.opd_order_dmc SET acting_check='0' where order_key=" & strOcsOrdNo
                            DBConn.Execute SqlStmt
                        End If
                    End If
                End If
                
                Set RS = Nothing
            End If
        Next
    End With
    Set RS = Nothing
    OCSActingCheck = True
    Exit Function
    
Errors:
    Set RS = Nothing
    OCSActingCheck = False
End Function

Private Sub cmdClear_Click()
    Call ClearRtn
    DoEvents
    txtWorkArea.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
'    Set frm104DelOrder = Nothing
End Sub

Private Sub Form_Load()
    
    optOption(0).value = True
    seLOption(0).value = True
    
    ClearFg = True
    
    Call objCanAcc.LoadReasonTemplate(cboReason)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objCanAcc = Nothing
    Set objAccData = Nothing
End Sub

Private Sub seLOption_Click(Index As Integer)
    Select Case Index
        Case 0
            txtBarcode.Visible = False
            txtBarcode.Enabled = False
        Case 1
            txtBarcode.Enabled = True
            txtBarcode.Visible = True
    End Select
End Sub

Private Sub txtBarCode_Change()
    If Len(txtBarcode.Text) = txtBarcode.MaxLength Then
        If txtAccDt.Enabled Then txtAccDt.SetFocus
    End If
End Sub

Private Sub txtBarCode_GotFocus()
    
    With txtBarcode
       .SelStart = 0
       .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
   
    Dim strStsNm As String, strMultiFg As String
    Dim tmpStatus As String
    Dim AccFg As Boolean
    Dim Resp As VbMsgBoxResult
    Dim objPatient As New clsPatient
   
    Dim tmpStr As String
    Dim tmpRs As Recordset
    Dim strSpcYY As String, strSpcNo As String
'    Dim strWorkarea As String, strAccDT As String, strAccSeq As String
    
    If KeyAscii = vbKeyReturn Then

        If txtBarcode.Text = "" Then Exit Sub

        strSpcYY = Mid(txtBarcode.Text, 1, 2)
        strSpcNo = Mid(txtBarcode.Text, 3, 11)
        tmpStr = objCanAcc.SqlAccOrderByBar(strSpcYY, strSpcNo)
        Set tmpRs = New Recordset
        tmpRs.Open tmpStr, DBConn
    
        If tmpRs.EOF Then GoTo NoData
        
        txtWorkArea.Text = "" & tmpRs.Fields("workarea").value
        tmpAccDt = "" & tmpRs.Fields("accdt").value
        txtAccDt.Text = Mid(tmpAccDt, 3)
        txtAccSeq.Text = "" & tmpRs.Fields("accseq").value
        
        Set tmpRs = Nothing

        objAccData.Clear
        
        objAccData.FieldInialize "workarea,accdt,accseq", "stscd,ptid,orddoct,deptcd,wardid,roomid,bedid,hosilid," & _
                                 "spccd,coldt,coltm,colid,rcvdt,rcvtm,rcvid,multifg,spcnm"
      
        tmpStatus = objCanAcc.CheckStatus(txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, objAccData)
        
        If tmpStatus = "0" Then
            
            Resp = MsgBox("입력하신 접수번호는 정상적인 데이타가 아닙니다. 그래도 취소하시겠습니까?", vbYesNo + vbExclamation, "경고")
            If Resp = vbNo Then
                Call ClearRtn
                txtWorkArea.SetFocus
                Exit Sub
            End If
            optOption(1).value = True
            optOption(1).Enabled = False
            
        ElseIf Val(objAccData.Fields("stscd")) = enStsCd.StsCd_LIS_Cancel Then
            
            MsgBox "입력하신 접수번호는 이미 취소되었습니다.", vbOKOnly + vbCritical, "메세지"
            Call ClearRtn
            txtWorkArea.SetFocus
            Exit Sub
            
        ElseIf Val(objAccData.Fields("stscd")) > enStsCd.StsCd_LIS_Accession Then
            
            Resp = MsgBox("이미 검사가 수행되었습니다. 그래도 취소하시겠습니까?", vbYesNo + vbExclamation, "경고")
            If Resp = vbNo Then
                Call ClearRtn
                txtWorkArea.SetFocus
                Exit Sub
            End If
            
        End If
      
        optOption(0).value = True
        optOption(1).Enabled = IIf((Val(objAccData.Fields("stscd")) = enStsCd.StsCd_LIS_Collection), False, True)
        '복수검체는 처방상태로만...
        If objAccData.Fields("multifg") <> "" Then optOption(1).Enabled = False
        
        lblStatus.Caption = tmpStatus
        lblStatus.tag = objAccData.Fields("stscd")
        
        '처방의
        lblDoctNm.Caption = GetEmpNm(objAccData.Fields("orddoct"))
        
        '진료과
'        objLisComCode.DeptCd.KeyChange objAccData.Fields("deptcd")
        lblDeptNm.Caption = GetDeptNm(objAccData.Fields("deptcd")) 'objLisComCode.DeptCd.Fields("deptnm")
        
        '환자정보
        lblPtId.Caption = objAccData.Fields("PtId")
        With objPatient
            If .GETPatient(objAccData.Fields("PtId")) Then
                lblPtNm.Caption = .ptnm
                lblSex.Caption = .SEXNM
                lblAge.Caption = .Age
                lblAgeDiv.Caption = .AGEDIV
            End If
        End With
        Call ICSPatientMark(lblPtId.Caption, enICSNum.LIS_ALL)
        
        lblLocation.Caption = objAccData.Fields("WardId") & "-" & objAccData.Fields("RoomId")
        lblSpcNm.Caption = objAccData.Fields("SpcNm")
        lblColDtTm.Caption = Format(objAccData.Fields("ColDt"), CS_DateMask) & "  " & _
                             Format(objAccData.Fields("ColTm"), CS_TimeLongMask)
        
        
        lblColNm.Caption = GetEmpNm(objAccData.Fields("ColId"))
        
        If Val(objAccData.Fields("stscd")) > enStsCd.StsCd_LIS_Collection Then
            lblRcvDtTm.Caption = Format(objAccData.Fields("RcvDt"), CS_DateMask) & "  " & _
                                 Format(objAccData.Fields("RcvTm"), CS_TimeLongMask)
            lblRcvNm.Caption = GetEmpNm(objAccData.Fields("RcvId"))
        End If
        
        AccFg = objCanAcc.DisplayOrder(tblOrdSheet, txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, tmpStatus)
      
        txtWorkArea.Enabled = False
        txtAccDt.Enabled = False
        txtAccSeq.Enabled = False
        
        cmdCancel.Enabled = True
        chkAll.value = 1
        Call chkAll_Click
    End If
 

NoData:
   Set tmpRs = Nothing

 End Sub

Private Sub txtWorkArea_Change()
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then
        If txtAccDt.Enabled Then txtAccDt.SetFocus
    End If
End Sub

Private Sub txtWorkArea_GotFocus()
    
    With txtWorkArea
       .SelStart = 0
       .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtWorkArea_KeyPress(KeyAscii As Integer)
    'Call ClearRtn
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If txtWorkArea = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccDt.SetFocus
 End Sub

Private Sub txtAccDt_Change()
    
    If Mid(txtAccDt.Text, 1, 1) = "9" Then
        tmpAccDt = "19" & txtAccDt.Text
    ElseIf Mid(txtAccDt.Text, 1, 1) = "0" Then
        tmpAccDt = "20" & txtAccDt.Text
    ElseIf Mid(txtAccDt.Text, 1, 1) = "1" Then
        tmpAccDt = "20" & txtAccDt.Text
    Else
        tmpAccDt = ""
    End If
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then
       If txtAccSeq.Enabled Then txtAccSeq.SetFocus
    End If

End Sub

Private Sub txtAccDt_GotFocus()
    With txtAccDt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    If txtAccDt.Text = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccSeq.SetFocus
End Sub

Private Sub txtAccSeq_GotFocus()
    
    With txtAccSeq
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'% 접수번호를 입력했을 경우
Private Sub txtAccSeq_KeyPress(KeyAscii As Integer)
   
    Dim strStsNm As String, strMultiFg As String
    Dim tmpStatus As String
    Dim AccFg As Boolean
    Dim Resp As VbMsgBoxResult
    Dim objPatient As New clsPatient
   
    If KeyAscii = vbKeyReturn Then

        If txtAccSeq.Text = "" Then Exit Sub
        
        objAccData.Clear
        objAccData.FieldInialize "workarea,accdt,accseq", "stscd,ptid,orddoct,deptcd,wardid,roomid,bedid,hosilid," & _
                                 "spccd,coldt,coltm,colid,rcvdt,rcvtm,rcvid,multifg,spcnm"
      
        tmpStatus = objCanAcc.CheckStatus(txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, objAccData)
        
        If tmpStatus = "0" Then
            
            Resp = MsgBox("입력하신 접수번호는 정상적인 데이타가 아닙니다. 그래도 취소하시겠습니까?", vbYesNo + vbExclamation, "경고")
            If Resp = vbNo Then
                Call ClearRtn
                txtWorkArea.SetFocus
                Exit Sub
            End If
            optOption(1).value = True
            optOption(1).Enabled = False
            
        ElseIf Val(objAccData.Fields("stscd")) = enStsCd.StsCd_LIS_Cancel Then
            
            MsgBox "입력하신 접수번호는 이미 취소되었습니다.", vbOKOnly + vbCritical, "메세지"
            Call ClearRtn
            txtWorkArea.SetFocus
            Exit Sub
            
        ElseIf Val(objAccData.Fields("stscd")) > enStsCd.StsCd_LIS_Accession Then
            
            Resp = MsgBox("이미 검사가 수행되었습니다. 그래도 취소하시겠습니까?", vbYesNo + vbExclamation, "경고")
            If Resp = vbNo Then
                Call ClearRtn
                txtWorkArea.SetFocus
                Exit Sub
            End If
            
        End If
      
        optOption(0).value = True
        optOption(1).Enabled = IIf((Val(objAccData.Fields("stscd")) = enStsCd.StsCd_LIS_Collection), False, True)
        '복수검체는 처방상태로만...
        If objAccData.Fields("multifg") <> "" Then optOption(1).Enabled = False
        
        lblStatus.Caption = tmpStatus
        lblStatus.tag = objAccData.Fields("stscd")
        
        '처방의
        lblDoctNm.Caption = GetEmpNm(objAccData.Fields("orddoct"))
        
        '진료과
'        objLisComCode.DeptCd.KeyChange objAccData.Fields("deptcd")
        lblDeptNm.Caption = GetDeptNm(objAccData.Fields("deptcd")) 'objLisComCode.DeptCd.Fields("deptnm")
        
        '환자정보
        lblPtId.Caption = objAccData.Fields("PtId")
        With objPatient
            If .GETPatient(objAccData.Fields("PtId")) Then
                lblPtNm.Caption = .ptnm
                lblSex.Caption = .SEXNM
                lblAge.Caption = .Age
                lblAgeDiv.Caption = .AGEDIV
            End If
        End With
        Call ICSPatientMark(lblPtId.Caption, enICSNum.LIS_ALL)
        
        lblLocation.Caption = objAccData.Fields("WardId") & "-" & objAccData.Fields("RoomId")
        lblSpcNm.Caption = objAccData.Fields("SpcNm")
        lblColDtTm.Caption = Format(objAccData.Fields("ColDt"), CS_DateMask) & "  " & _
                             Format(objAccData.Fields("ColTm"), CS_TimeLongMask)
        
        
        lblColNm.Caption = GetEmpNm(objAccData.Fields("ColId"))
        
        If Val(objAccData.Fields("stscd")) > enStsCd.StsCd_LIS_Collection Then
            lblRcvDtTm.Caption = Format(objAccData.Fields("RcvDt"), CS_DateMask) & "  " & _
                                 Format(objAccData.Fields("RcvTm"), CS_TimeLongMask)
            lblRcvNm.Caption = GetEmpNm(objAccData.Fields("RcvId"))
        End If
        
        AccFg = objCanAcc.DisplayOrder(tblOrdSheet, txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, tmpStatus)
      
        txtWorkArea.Enabled = False
        txtAccDt.Enabled = False
        txtAccSeq.Enabled = False
        
        cmdCancel.Enabled = True
   
    End If

End Sub

'Private Function GetEmpNm(ByVal vEmpID As String) As String
'    Dim objData As New clsBasisData
'
'    GetEmpNm = objData.GetEmpNm(vEmpID)
'    Set objData = Nothing
'End Function

'Private Function GetDeptNm(ByVal vDeptCd As String) As String
'    Dim objData As New clsBasisData
'
'    GetDeptNm = objData.GetDeptNm(vDeptCd)
'    Set objData = Nothing
'End Function

Private Sub ClearRtn()
    
    txtWorkArea.Enabled = True
    txtAccDt.Enabled = True
    txtAccSeq.Enabled = True
    
    txtWorkArea.Text = ""
    txtAccDt.Text = ""
    txtAccSeq.Text = ""
    
    optOption(0).value = True
    optOption(1).Enabled = True
    lblStatus.Caption = ""
   
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDoctNm.Caption = ""
    lblDeptNm.Caption = ""
    lblLocation.Caption = ""
    
    lblSpcNm.Caption = ""
    lblColDtTm.Caption = ""
    lblColNm.Caption = ""
    lblRcvDtTm.Caption = ""
    lblRcvNm.Caption = ""
    
    tblOrdSheet.MaxRows = 0
    tblOrdSheet.MaxRows = 16
    txtReason.Text = ""
    chkAll.value = 0
    
    cmdCancel.Enabled = False
    ClearFg = True
    Call ICSPatientMark
End Sub
