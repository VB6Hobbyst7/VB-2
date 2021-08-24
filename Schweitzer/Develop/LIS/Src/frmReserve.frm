VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmReserve 
   BackColor       =   &H00DBE6E6&
   Caption         =   "검사예약 및 취소"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14430
   WindowState     =   2  '최대화
   Begin DRcontrol1.DrFrame fraChange 
      Height          =   2745
      Left            =   4620
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4842
      Title           =   "= 검체 채혈일시를 변경합니다. ="
      BackColor       =   16776439
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFCF7&
         Caption         =   "닫기"
         Height          =   510
         Left            =   4725
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   1380
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   165
         TabIndex        =   23
         Top             =   420
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
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
         Caption         =   "환자정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtid 
         Height          =   360
         Left            =   1470
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   420
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblsPtnm 
         Height          =   360
         Left            =   2505
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   360
         Index           =   0
         Left            =   165
         TabIndex        =   26
         Top             =   1920
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
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
         Caption         =   "희망채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReqdt 
         Height          =   360
         Left            =   1470
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel10 
         Height          =   375
         Left            =   165
         TabIndex        =   28
         Top             =   2295
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
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
         Caption         =   "변경일시"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpReqdate 
         Height          =   390
         Left            =   1470
         TabIndex        =   29
         Top             =   2295
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   688
         _Version        =   393216
         CustomFormat    =   "yyy-MM-dd HH:mm:ss"
         Format          =   73924611
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel lblChangeReqdate 
         Height          =   360
         Left            =   3555
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2295
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   360
         Index           =   1
         Left            =   165
         TabIndex        =   31
         Top             =   795
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
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
         Caption         =   "처방정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrddt 
         Height          =   360
         Left            =   1470
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   795
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOrdNo 
         Height          =   360
         Left            =   2970
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   795
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestcd 
         Height          =   360
         Left            =   1470
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   360
         Left            =   2505
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1170
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcCd 
         Height          =   360
         Left            =   1470
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   360
         Left            =   2505
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BackColor       =   15857140
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "01"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   165
         TabIndex        =   38
         Top             =   1170
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
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
         Caption         =   "검사정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   360
         Index           =   2
         Left            =   165
         TabIndex        =   39
         Top             =   1545
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
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
         Caption         =   "검체정보"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00FFFCF7&
         Caption         =   "확인"
         Height          =   510
         Left            =   4725
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   870
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00DBE6E6&
      Caption         =   "예약조회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3495
      Style           =   1  '그래픽
      TabIndex        =   19
      Top             =   840
      Width           =   1320
   End
   Begin VB.OptionButton optDiv 
      BackColor       =   &H00EFFFEE&
      Caption         =   "예약리스트"
      Height          =   225
      Index           =   1
      Left            =   2085
      TabIndex        =   16
      Top             =   915
      Width           =   1260
   End
   Begin VB.OptionButton optDiv 
      BackColor       =   &H00EFFFEE&
      Caption         =   "예약대상"
      Height          =   225
      Index           =   0
      Left            =   945
      TabIndex        =   15
      Top             =   990
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00DBE6E6&
      Caption         =   "예약대상"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3495
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   375
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel7 
      Height          =   300
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   45
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   529
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
      Alignment       =   1
      Caption         =   "예약대상항목 리스트"
      Appearance      =   0
      LeftGab         =   200
   End
   Begin FPSpread.vaSpread tblTest 
      Height          =   7020
      Left            =   75
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1335
      Width           =   3405
      _Version        =   196608
      _ExtentX        =   6006
      _ExtentY        =   12382
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      DisplayRowHeaders=   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16777215
      MaxCols         =   5
      MaxRows         =   29
      OperationMode   =   1
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   15663103
      SpreadDesigner  =   "frmReserve.frx":0000
   End
   Begin MedControls1.LisLabel LisLabel7 
      Height          =   300
      Index           =   1
      Left            =   3510
      TabIndex        =   3
      Top             =   45
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   529
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
      Caption         =   "예약대상항목 리스트"
      Appearance      =   0
      LeftGab         =   200
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
      Height          =   1035
      Left            =   4845
      TabIndex        =   5
      Top             =   285
      Width           =   9600
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         MaxLength       =   10
         TabIndex        =   6
         Top             =   165
         Width           =   1425
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   4590
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "김미경"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   7920
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   165
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "김미경"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   300
         Left            =   1155
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   540
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "김미경"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   4590
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   525
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "김미경"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   315
         Left            =   7920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   525
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "김미경"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   40
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "환자   ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReceptNo 
         Height          =   315
         Left            =   90
         TabIndex        =   41
         Top             =   540
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "처 방 의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   3510
         TabIndex        =   42
         Top             =   165
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "성      명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   3510
         TabIndex        =   43
         Top             =   525
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "진 료 과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   11
         Left            =   6855
         TabIndex        =   44
         Top             =   165
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "성 / 나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   6855
         TabIndex        =   45
         Top             =   525
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "병      실"
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   7035
      Left            =   3510
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1335
      Width           =   10935
      _Version        =   196608
      _ExtentX        =   19288
      _ExtentY        =   12409
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16777215
      MaxCols         =   16
      MaxRows         =   29
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   15663103
      SpreadDesigner  =   "frmReserve.frx":0508
      TextTip         =   1
   End
   Begin MSComCtl2.DTPicker dtpOrdDt 
      Height          =   315
      Left            =   945
      TabIndex        =   17
      Top             =   465
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
      Format          =   73924608
      CurrentDate     =   36342.5951388889
   End
   Begin VB.Label lblDt 
      BackColor       =   &H00EFFFEE&
      Caption         =   "조회구분"
      Height          =   225
      Index           =   1
      Left            =   165
      TabIndex        =   18
      Tag             =   "15104"
      Top             =   1005
      Width           =   735
   End
   Begin VB.Label lblDt 
      BackColor       =   &H00EFFFEE&
      Caption         =   "처방일"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Tag             =   "15104"
      Top             =   615
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EFFFEE&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   930
      Index           =   1
      Left            =   75
      Top             =   375
      Width           =   3390
   End
End
Attribute VB_Name = "frmReserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents mnuPopup     As menu
'Private WithEvents mnuChange    As menu
Private WithEvents objPop As clsPopupMenu
Private Const MENU_RESERVE& = 1

Private lngSelRow               As Long

Private Enum TblCol
    enPtid = 1
    enPtNm
    enOrdDate
    enTestNm
    enSpcNm
    
    enReqdate
    enChangeDate
    enStatus
    enOrdDt
    enOrdNo
    
    enOrdSeq
    enOrdDoct
    enDeptcd
    enTestCd
    enSpccd

    enSSN
End Enum


Private Sub cmdClose_Click()
    fraChange.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CommAND1_Click()
    optDiv(0).Value = False
    optDiv(0).Value = True
    dtpOrdDt.Value = GetSystemdate
End Sub

Private Sub dtpReqdate_LostFocus()
    lblChangeReqdate.Caption = Format(dtpReqdate.Value, "YYYY-MM-DD   HH:MM:SS")
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_RESERVE
            Call ClearChange
            If lngSelRow = 0 Then Exit Sub
            
            With tblList
                .Row = lngSelRow
                .Col = TblCol.enSpccd:
                If .Value = "" Then Exit Sub
                .Col = TblCol.enPtid:    lblPtid.Caption = .Value
                .Col = TblCol.enPtNm:    lblsPtnm.Caption = .Value
                .Col = TblCol.enSpccd:   lblSpcCd.Caption = .Value
                .Col = TblCol.enSpcNm:   lblSpcNm.Caption = .Value
                .Col = TblCol.enTestCd:  lblTestcd.Caption = .Value
                .Col = TblCol.enTestNm:  lblTestNm.Caption = .Value
                .Col = TblCol.enOrdDt:   lblOrddt.Caption = Format(.Value, "####-##-##")
                .Col = TblCol.enOrdNo:   lblOrdNo.Caption = .Value
                .Col = TblCol.enReqdate: lblReqdt.Caption = .Value
            End With
            fraChange.Visible = True
    End Select
End Sub

Private Sub optDiv_Click(Index As Integer)
    Call medClearTable(tblList)
    cmdQuery.Enabled = False
    cmdList.Enabled = False
    If Index = 0 Then
        cmdQuery.Enabled = True
    Else
        cmdList.Enabled = True
    End If
    Call tblList_Click(1, 1)
End Sub

Private Sub ClearData(Optional ByVal FirstClear As Boolean = True)
    dtpOrdDt.Value = GetSystemdate
    optDiv(0).Value = True
    txtPtId.Text = ""
    lblPtNm.Caption = "": lblSexAge.Caption = "": lblDeptNm.Caption = "": lblDoctNm.Caption = ""
    lblLocation.Caption = ""
    With tblTest
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .CellType = CellTypeStaticText
        .Value = ""
        .BlockMode = False
    End With
    Call medClearTable(tblList)
End Sub

' 예약가능 검사항목 조회하기
Private Sub GetReserveTestItem()
    Dim SSQL    As String
    Dim RS      As Recordset
        
    Call ClearData
    SSQL = GetReserveTestItemSQL
    
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    
    If Not RS.EOF Then
        With tblTest
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter
                .Col = 2: .Value = RS.Fields("field1").Value & ""
                .Col = 3: .Value = RS.Fields("field2").Value & ""
                .Col = 4: .Value = RS.Fields("cdval1").Value & ""
                .Col = 5: .Value = RS.Fields("cdval2").Value & ""
                RS.MoveNext
            Loop
        End With
    End If
    
    Set RS = Nothing
End Sub


Private Sub cmdQuery_Click()
    Dim sOrdDt  As String
    Dim sOrdcd  As String
    Dim sSpcCd  As String
    Dim ii      As Integer
    
    
    Call medClearTable(tblList)
    
    sOrdDt = Format(dtpOrdDt.Value, "YYYYMMDD")
    
    With tblTest
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1
            If .CellType = CellTypeCheckBox And .Value = 1 Then
                .Col = 4: sOrdcd = Trim(.Value)
                .Col = 5: sSpcCd = Trim(.Value)
            End If
        Next
    End With
    Call GetReserveTestListQuery(sOrdDt, , sOrdcd, sSpcCd)
End Sub

Private Sub Form_Activate()
    Call GetReserveTestItem
End Sub

Private Sub GetReserveTestListQuery(ByVal sOrdDt As String, Optional ByVal sPtid As String = "", _
                                    Optional ByVal sOrdcd As String = "", Optional ByVal sSpcCd As String = "", _
                                    Optional ByVal ReserveOK As Boolean = False)
    Dim RS      As Recordset
    Dim SSQL    As String
    

    If ReserveOK = False Then
        SSQL = GetReserveTestListSQL(sOrdDt, sPtid, sOrdcd, sSpcCd)
    Else
        SSQL = GetReserveListSQL(sOrdDt)
    End If
    
    Set RS = New Recordset
    
    RS.Open SSQL, dbconn
    
    If Not RS.EOF Then
        With tblList
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1: lngSelRow = .Row
                .Col = TblCol.enPtid:    .Value = RS.Fields("ptid").Value & ""
                .Col = TblCol.enPtNm:    .Value = RS.Fields("ptnm").Value & ""
                .Col = TblCol.enDeptcd:  .Value = RS.Fields("deptcd").Value & ""
                .Col = TblCol.enOrdDate: .Value = Format(RS.Fields("orddt").Value & "", "####-##-##") & " " & _
                                                  Format(RS.Fields("ordtm").Value & "", "00:00:00")
                .Col = TblCol.enOrdDoct: .Value = RS.Fields("orddoct").Value & ""
                .Col = TblCol.enOrdDt:   .Value = RS.Fields("orddt").Value & ""
                .Col = TblCol.enOrdNo:   .Value = RS.Fields("ordno").Value & ""
                .Col = TblCol.enOrdSeq:  .Value = RS.Fields("ordseq").Value & ""
                .Col = TblCol.enReqdate: .Value = Format(RS.Fields("reqdt").Value & "", "####-##-##") & " " & _
                                                  Format(RS.Fields("reqtm").Value & "", "00:00:00")
                .Col = TblCol.enTestNm:  .Value = RS.Fields("testnm").Value & ""
                .Col = TblCol.enSpcNm:   .Value = RS.Fields("spcnm").Value & ""
                .Col = TblCol.enTestCd:  .Value = RS.Fields("ordcd").Value & ""
                .Col = TblCol.enSpccd:   .Value = RS.Fields("spccd").Value & ""
                .Col = TblCol.enSSN:     .Value = GetSexAGE(RS.Fields("ssn").Value & "")
                Call GetReserveDataDSP(RS.Fields("ptid").Value & "", RS.Fields("orddt").Value & "", RS.Fields("ordno").Value & "", _
                                       RS.Fields("ordseq").Value & "")
                If ReserveOK = True Then
                    .Col = TblCol.enStatus
                    If RS.Fields("stscd") >= enStsCd.StsCd_LIS_Collection Then
                        .Value = "처리완료"
                    End If
                End If
                RS.MoveNext
            Loop
            Call tblList_Click(1, 1)
        End With
    End If
    Set RS = Nothing
End Sub


Private Function GetReserveTestItemSQL() As String
    Dim SSQL As String
    
    SSQL = " SELECT * FROM " & T_LAB031 & _
           " WHERE  " & DBW("cdindex=", LC4_TestItemComment) & _
           " AND    " & DBW("text2=", "1") & _
           " ORDER BY cdval1"
    GetReserveTestItemSQL = SSQL
End Function

Private Function GetReserveTestListSQL(ByVal sOrdDt As String, Optional ByVal sPtid As String = "", _
                                       Optional ByVal sTestcd As String = "", Optional ByVal sSpcCd As String = "") As String
    Dim SSQL    As String
        
    SSQL = " SELECT b.ptid ,b.orddt,b.ordno,b.ordseq,a.ordtm,b.ordcd,b.spccd,a.wardid,a.deptcd,a.reqdt,a.reqtm, " & _
           "        a.majdoct,a.orddoct,d." & F_PTNM & " as ptnm,c.field1 as testnm,c.field2 as spcnm," & _
           "    " & F_SSN2("d") & " as ssn" & _
           " FROM " & T_HIS001 & " d," & T_LAB102 & " b," & T_LAB031 & " c," & T_LAB101 & " a " & _
           " WHERE "
           
    If sPtid <> "" Then
        SSQL = SSQL & DBW("a.ptid=", sPtid) & " AND " & DBW("a.orddt=", sOrdDt)
    Else
        SSQL = SSQL & DBW("a.orddt=", sOrdDt)
    End If
    SSQL = SSQL & " AND " & DBW("a.orddiv=", lis_orddiv) & _
                  " AND " & DBW("a.donefg=", enStsCd.StsCd_LIS_Order) & _
                  " AND " & DBW("a.bussdiv=", enBussDiv.BussDiv_OutPatient) & _
                  " AND a.ptid=b.ptid" & _
                  " AND a.orddt=b.orddt" & _
                  " AND a.ordno=b.ordno" & _
                  " AND " & DBW("c.cdindex=", LC4_TestItemComment)
    If sTestcd <> "" Then
        SSQL = SSQL & " AND " & DBW("b.ordcd=", sTestcd)
    End If
    If sSpcCd <> "" Then
        SSQL = SSQL & " AND " & DBW("b.spccd=", sSpcCd)
    End If
    SSQL = SSQL & " AND c.cdval1=b.ordcd" & _
                  " AND c.cdval2=b.spccd" & _
                  " AND " & DBW("c.text2=", "1") & _
                  " AND a.ptid=d." & F_PTID
    GetReserveTestListSQL = SSQL
End Function

Private Sub ClearChange()
    lblPtid.Caption = "":   lblsPtnm.Caption = "": lblReqdt.Caption = "": lblTestcd.Caption = ""
    lblTestNm.Caption = "": lblSpcCd.Caption = "": lblSpcNm.Caption = ""
    lblOrddt.Caption = "":  lblOrdNo.Caption = "": lblChangeReqdate.Caption = ""
    
    dtpReqdate.Value = GetSystemdate
    
End Sub
'Private Sub mnuChange_Click()
'
'    Call ClearChange
'    If lngSelRow = 0 Then Exit Sub
'
'    With tblList
'        .Row = lngSelRow
'        .Col = TblCol.enSpccd:
'        If .Value = "" Then Exit Sub
'        .Col = TblCol.enPtid:    lblPtid.Caption = .Value
'        .Col = TblCol.enPtNm:    lblsPtnm.Caption = .Value
'        .Col = TblCol.enSpccd:   lblSpcCd.Caption = .Value
'        .Col = TblCol.enSpcNm:   lblSpcNm.Caption = .Value
'        .Col = TblCol.enTestCd:  lblTestcd.Caption = .Value
'        .Col = TblCol.enTestNm:  lblTestNm.Caption = .Value
'        .Col = TblCol.enOrdDt:   lblOrddt.Caption = Format(.Value, "####-##-##")
'        .Col = TblCol.enOrdNo:   lblOrdNo.Caption = .Value
'        .Col = TblCol.enReqdate: lblReqdt.Caption = .Value
'    End With
'    fraChange.Visible = True
'End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim objData As clsBasisData
    Dim strData As String
    
    txtPtId.Text = ""
    lblDeptNm.Caption = "": lblPtNm.Caption = "": lblSexAge.Caption = "": lblLocation.Caption = ""
    lblDoctNm.Caption = ""
    
    If Row > tblList.DataRowCnt Or Row < 1 Then Exit Sub
    
'    Set objData = New clsBasisData
    
    
    With tblList
        .Row = Row
        .Col = TblCol.enPtid:       txtPtId.Text = .Value
        .Col = TblCol.enPtNm:       lblPtNm.Caption = .Value
        .Col = TblCol.enDeptcd:     'Call objLisComCode.DeptCd.KeyChange(.Value)
                                    lblDeptNm.Caption = GetDeptNm(.Value) 'objLisComCode.DeptCd.Fields("deptnm")
        .Col = TblCol.enOrdDoct:
                                    lblDoctNm.Caption = GetEmpNm(.Value) 'GetEmpName("doctnm")
        .Col = TblCol.enSSN:        lblSexAge.Caption = .Value
    End With
    
'    Set objData = Nothing
End Sub

Private Sub tblList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    lngSelRow = Row
    If Row > tblList.DataRowCnt Or Row < 1 Then Exit Sub
    
    With tblList
        .Row = Row
        .Col = TblCol.enStatus
        If .Value <> "" Then Exit Sub
    End With
    
    Set objPop = Nothing
    Set objPop = New clsPopupMenu
    
    With objPop
        .AddMenu MENU_RESERVE, "검사예약"
        
        .PopupMenus Me.hwnd
    End With
    
    Set objPop = Nothing
    
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuChange = frmControls.mnuSub
'
'    mnuChange.Caption = "검사예약"
'    frmControls.mnuSub1.Visible = False
'    frmControls.mnuSub2.Visible = False
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuChange = Nothing
End Sub

Private Function GetSexAGE(ByVal ssn As String) As String
    Dim strTmp As String
    Dim strSEX As String
    Dim strAge As String
    Dim strDOB As String
    
    Dim strYY  As String
    Dim strMM  As String
    Dim strDD  As String
    
    strYY = Trim(Mid(ssn, 1, 2))
    strMM = Trim(Mid(ssn, 3, 2))
    strDD = Trim(Mid(ssn, 5, 2))
    
    If Val(strMM) < 1 Then strMM = "01"
    If Val(strMM) > 12 Then strMM = "12"
    If Val(strDD) < 1 Then strDD = "01"
    If Val(strDD) > 31 Then strDD = "31"
    
    
    On Error Resume Next
    
    If IsDate(strYY & "-" & strMM & "-" & strDD) = False Then
        strDD = "01"
    End If
    
    strSEX = "기타": strAge = "": strDOB = ""
    
    If ssn <> "" Then
        strTmp = Mid(ssn, 7, 1)
        Select Case strTmp
            Case "0": strSEX = "여": strDOB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1": strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2": strSEX = "여": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3": strSEX = "남": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4": strSEX = "여": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        
        If Len(ssn) = 13 Then
            strAge = medFindAge(Replace(strDOB, "-", ""), "Y")
        Else
            strAge = ""
        End If
        GetSexAGE = strSEX & "/" & strAge
    Else
        GetSexAGE = ""
    End If
End Function



Private Sub tblTest_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or Row > tblTest.DataRowCnt Then Exit Sub
    If Col <> 1 Then Exit Sub
    Call medClearTable(tblList)
    Call tblList_Click(1, 1)
    With tblTest
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .COL2 = 1
        .BlockMode = True
        .Value = "0"
        .BlockMode = False
        .Row = Row: .Col = Col: .Value = IIf(.Value = "1", "0", "1")
        
    End With
End Sub

Private Sub txtPtId_GotFocus()
    txtPtId.SelStart = 0
    txtPtId.SelLength = Len(txtPtId.Text)
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call txtPtId_LostFocus
    End If
End Sub

Private Sub txtPtId_LostFocus()
    Dim sOrdDt  As String
    
    lblDeptNm.Caption = "": lblPtNm.Caption = "": lblSexAge.Caption = "": lblLocation.Caption = ""
    lblDoctNm.Caption = ""
    
    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    Call medClearTable(tblList)
    
    sOrdDt = Format(dtpOrdDt.Value, "YYYYMMDD")
    Call GetReserveTestListQuery(sOrdDt, txtPtId.Text)
    
    If tblList.DataRowCnt = 0 Then
        txtPtId.SelStart = 0
        txtPtId.SelLength = Len(txtPtId.Text)
    End If
End Sub

Private Sub cmdOk_Click()
    Dim sPtid   As String
    Dim sOrdDt  As String
    Dim sOrdNo  As String
    Dim sOrdSeq As String
    Dim sReqdt  As String
    Dim sReqtm  As String
    Dim sSysDt  As String
    Dim sSysTm  As String
    Dim sEntDt  As String
    Dim sEntTm  As String
    Dim sEntId  As String
    
    Dim SSQL    As String
    Dim strTmp  As String
    
    If lblChangeReqdate.Caption = "" Then
        MsgBox "예약일시를 선택하세요", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    With tblList
        .Row = lngSelRow
        .Col = TblCol.enPtid:       sPtid = .Value
'        .Col = TblCol.enPtnm:       sPtNm = .Value
        .Col = TblCol.enOrdDt:      sOrdDt = .Value
        .Col = TblCol.enOrdNo:      sOrdNo = .Value
        .Col = TblCol.enOrdSeq:     sOrdSeq = .Value
        .Col = TblCol.enReqdate:    sReqdt = medGetP(Replace(.Value, "-", ""), 1, " ")
    End With
    
    sSysDt = Format(dtpReqdate.Value, "YYYYMMDD")
    sSysTm = Format(dtpReqdate.Value, "HHMMSS")
    
    sEntDt = Format(GetSystemdate, "YYYYMMDD")
    sEntTm = Format(GetSystemdate, "HHMMSS")
    sEntId = ObjSysInfo.EmpId
    
    If GetOrderStatusChk(sPtid, sOrdDt, sOrdNo) = False Then Exit Sub
    
    If GetOrderInfoNotDupChk(sPtid, sOrdDt, sOrdNo, sOrdSeq) = False Then
        strTmp = MsgBox("예약항목이외의 검사들이 있습니다." & vbCRLF & _
                        "예약일시 변경시 모든 항목들에 대해서 적용됩니다." & vbCRLF & _
                        "변경하시겠습니까?", vbYesNo + vbInformation, "info")
        If strTmp = vbNo Then Exit Sub
    End If
    
    '101의 희망채혈일시 변경
    's2reserve 삭제후 Insert
    On Error GoTo SAVE_ERROR
    dbconn.BeginTrans
    
    SSQL = " update " & T_LAB101 & " set " & _
                        DBW("reqdt=", sSysDt, 1) & _
                        DBW("reqtm=", sSysTm) & _
           " WHERE  " & DBW("ptid=", sPtid) & _
           " AND    " & DBW("orddt=", sOrdDt) & _
           " AND    " & DBW("ordno=", sOrdNo)
    dbconn.Execute SSQL
    
    SSQL = " delete " & t_lab904 & " " & _
           " WHERE  " & _
                        DBW("ptid=", sPtid) & _
           " AND    " & DBW("orddt=", sOrdDt) & _
           " AND    " & DBW("ordno=", sOrdNo) & _
           " AND    " & DBW("ordseq=", sOrdSeq)
    
    dbconn.Execute SSQL
    
    SSQL = " insert into " & t_lab904 & " (ptid,orddt,ordno,ordseq,reqdt,reqtm,sysdt,systm,entdt,enttm,entid) " & _
           " values( " & _
                    DBV("ptid  ", sPtid, 1) & DBV("orddt", sOrdDt, 1) & DBV("ordno", sOrdNo, 1) & DBV("ordseq", sOrdSeq, 1) & _
                    DBV("reqdt", sReqdt, 1) & DBV("reqtm", sReqtm, 1) & DBV("sysdt", sSysDt, 1) & _
                    DBV("systm", sSysTm, 1) & DBV("entdt", sEntDt, 1) & DBV("enttm", sEntTm, 1) & _
                    DBV("entid", sEntId) & _
           ")"
    dbconn.Execute SSQL
    dbconn.CommitTrans
    
    With tblList
        .Row = lngSelRow
        .Col = TblCol.enChangeDate: .Value = Format(sSysDt, "####-##-##") & " " & _
                                             Format(sSysTm, "0#:##:##")
                                    .ForeColor = DCM_LightRed
        .Col = TblCol.enStatus:     .Value = "예약"
    End With
    
    Exit Sub
SAVE_ERROR:
    dbconn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Function GetOrderStatusChk(ByVal sPtid As String, ByVal sOrdDt As String, _
                                   ByVal sOrdNo As String) As Boolean
    Dim RS  As Recordset
    Dim SSQL As String
    
    SSQL = " SELECT stscd FROM " & T_LAB102 & _
           " WHERE " & _
                        DBW("ptid=", sPtid) & _
           " AND    " & DBW("orddt=", sOrdDt) & _
           " AND    " & DBW("ordno=", sOrdNo)
    GetOrderStatusChk = True
    
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    If Not RS.EOF Then
        Do Until RS.EOF
        
            If Val(RS.Fields("stscd").Value & "") >= Val(enStsCd.StsCd_LIS_Collection) Then
                MsgBox "이미 채혈된항목이 있어 일시를 변경할수 없습니다.", vbInformation + vbOKOnly, "Info"
                GetOrderStatusChk = False
                Exit Do
            End If
            RS.MoveNext
        Loop
    End If
    
    Set RS = Nothing
End Function


Private Function GetOrderInfoNotDupChk(ByVal sPtid As String, ByVal sOrdDt As String, _
                                       ByVal sOrdNo As String, ByVal sOrdSeq As String) As Boolean
    Dim SSQL As String
    Dim RS   As Recordset
    
    GetOrderInfoNotDupChk = True
    
    SSQL = " SELECT ptid,orddt,ordno,ordseq FROM " & T_LAB102 & _
           " WHERE " & _
                     DBW("ptid=", sPtid) & _
           " AND " & DBW("orddt=", sOrdDt) & _
           " AND " & DBW("ordno=", sOrdNo)
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            If RS.Fields("ordseq").Value & "" <> sOrdSeq Then
                GetOrderInfoNotDupChk = False
                Exit Do
            End If
            RS.MoveNext
        Loop
    End If
    
    Set RS = Nothing
End Function

Private Sub GetReserveDataDSP(ByVal sPtid As String, ByVal sOrdDt As String, _
                              ByVal sOrdNo As String, ByVal sOrdSeq As String)
    Dim RS      As Recordset
    Dim SSQL    As String
    
    
    SSQL = " SELECT sysdt,systm FROM " & t_lab904 & " " & _
           " WHERE  " & _
                        DBW("ptid=", sPtid) & _
           " AND    " & DBW("orddt=", sOrdDt) & _
           " AND    " & DBW("ordno=", sOrdNo) & _
           " AND    " & DBW("ordseq=", sOrdSeq)
    
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    If Not RS.EOF Then
        With tblList
            .Row = lngSelRow
            .Col = TblCol.enChangeDate: .Value = Format(RS.Fields("sysdt").Value & "", "####-##-##") & " " & _
                                                 Format(RS.Fields("systm").Value & "", "0#:##:##")
                                        .ForeColor = DCM_LightRed
            .Col = TblCol.enStatus:     .Value = "예약"
        End With
    End If
    Set RS = Nothing

End Sub
Private Function GetReserveListSQL(ByVal sReqdt As String) As String
    Dim SSQL    As String
    
    SSQL = " SELECT b.ptid ,b.orddt,b.ordno,b.ordseq,a.ordtm,b.ordcd,b.spccd,a.wardid,a.deptcd,a.reqdt,a.reqtm, " & _
           "        a.majdoct,a.orddoct,d." & F_PTNM & " as ptnm,c.field1 as testnm,c.field2 as spcnm," & _
           "    " & F_SSN2("d") & " as ssn,b.stscd" & _
           " FROM " & T_HIS001 & " d," & T_LAB102 & " b," & T_LAB031 & " c," & T_LAB101 & " a, " & t_lab904 & " z" & _
           " WHERE " & _
                    DBW("z.sysdt=", sReqdt) & _
           " AND z.ptid=b.ptid AND z.orddt=b.orddt AND z.ordno=b.ordno AND z.ordseq=b.ordseq " & _
           " AND z.ptid=a.ptid AND z.orddt=a.orddt AND z.ordno=a.ordno" & _
           " AND " & DBW("c.cdindex=", LC4_TestItemComment) & _
           " AND c.cdval1=b.ordcd AND c.cdval2=b.spccd" & _
           " AND z.ptid=d." & F_PTID
           
    GetReserveListSQL = SSQL
End Function
Private Sub cmdList_Click()
    Dim sReqdt  As String
    
    Call medClearTable(tblList)
    
    sReqdt = Format(dtpOrdDt.Value, "YYYYMMDD")
    
    Call GetReserveTestListQuery(sReqdt, , , , True)
    
End Sub

Private Sub tblList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim RS          As Recordset
    Dim tmpToolTip  As String
    Dim SSQL        As String
    Dim sPtid       As String
    Dim sOrdDt      As String
    Dim sOrdNo      As String
    
    If Row = 0 Then Exit Sub
    If Row > tblList.DataRowCnt Then Exit Sub
    
    With tblList
        .Row = Row
        .Col = TblCol.enPtid: sPtid = .Value
        .Col = TblCol.enOrdDt: sOrdDt = .Value
        .Col = TblCol.enOrdNo: sOrdNo = .Value
    
        
        SSQL = " SELECT a.orddt,a.ordno,a.ordseq,a.ordcd,a.spccd,b.abbrnm10 as testnm,c.field3 as spcnm " & _
               " FROM " & T_LAB032 & " c," & T_LAB001 & " b," & T_LAB102 & " a" & _
               " WHERE " & _
                         DBW("a.ptid=", sPtid) & _
               " AND " & DBW("a.orddt=", sOrdDt) & _
               " AND " & DBW("a.ordno=", sOrdNo) & _
               " AND a.ordcd=b.testcd" & _
               " AND " & DBW("c.cdindex=", LC3_Specimen) & _
               " AND a.spccd=c.cdval1" & _
               " ORDER BY ordseq"
             
        tmpToolTip = vbCRLF & " ♣ 동일발생 처방정보" & vbCRLF & vbCRLF
        Set RS = New Recordset
        RS.Open SSQL, dbconn
        If Not RS.EOF Then
            tmpToolTip = tmpToolTip & "  처방일" & Space(5) & "처방번호" & Space(8) & " 검사명" & Space(12) & "검체명" & vbCRLF
            
            Do Until RS.EOF
            
            
                
                tmpToolTip = tmpToolTip & " " & Format(RS.Fields("orddt").Value & "", "####-##-##") & Space(5) & _
                                         " " & RS.Fields("ordseq").Value & "" & Space(10 - Len(RS.Fields("ordseq").Value & "")) & _
                                         " " & RS.Fields("testnm").Value & "" & "(" & RS.Fields("ordcd").Value & "" & ")" & Space(17 - (Len(RS.Fields("testnm").Value & "" & RS.Fields("ordcd").Value & "") + 2)) & _
                                         " " & RS.Fields("spcnm").Value & "" & "(" & RS.Fields("spccd").Value & "" & ")" & vbCRLF
                RS.MoveNext
            Loop
        
            MultiLine = 1
            TipText = tmpToolTip
            TipWidth = 5500
            .TextTipDelay = 1000
            Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
            ShowTip = True
        End If
    End With
    Set RS = Nothing


End Sub
