VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWRD001 
   BackColor       =   &H00FFFFFF&
   Caption         =   "병동일괄채혈"
   ClientHeight    =   9240
   ClientLeft      =   285
   ClientTop       =   465
   ClientWidth     =   14055
   FillColor       =   &H00808080&
   FillStyle       =   0  '단색
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWRD001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   14055
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면닫기(&O)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11745
      Style           =   1  '그래픽
      TabIndex        =   71
      Tag             =   "0"
      Top             =   8520
      Width           =   1320
   End
   Begin VB.Frame fraQuery 
      BackColor       =   &H00DBE6E6&
      Height          =   8370
      Left            =   60
      TabIndex        =   4
      Top             =   -75
      Width           =   7275
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Left            =   4290
         TabIndex        =   6
         Top             =   135
         Width           =   2925
         _ExtentX        =   5159
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
         Caption         =   "채취 일시"
         LeftGab         =   100
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   1230
         Left            =   4275
         TabIndex        =   8
         Top             =   360
         Width           =   2955
         Begin VB.OptionButton optApplyColTm 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체적용"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   300
            Width           =   1035
         End
         Begin VB.OptionButton optApplyColTm 
            BackColor       =   &H00DBE6E6&
            Caption         =   "현재 Row만 적용"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1215
            TabIndex        =   9
            Top             =   300
            Width           =   1710
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   11
            Top             =   765
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   344
            BackColor       =   14411494
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            AutoSize        =   -1  'True
            Caption         =   "채취일시"
            Appearance      =   0
         End
         Begin MSComCtl2.DTPicker dtpColDtTm 
            Height          =   315
            Left            =   930
            TabIndex        =   12
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
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
            CustomFormat    =   "yyy-MM-dd  HH:mm"
            Format          =   61734915
            UpDown          =   -1  'True
            CurrentDate     =   36328.5416666667
         End
      End
      Begin VB.CheckBox ChkMornFg 
         BackColor       =   &H00800000&
         Caption         =   "채혈간호사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   990
         TabIndex        =   5
         Top             =   165
         Visible         =   0   'False
         Width           =   1530
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   300
         Left            =   0
         TabIndex        =   7
         Top             =   150
         Width           =   4245
         _ExtentX        =   7488
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
         Caption         =   "병동 선택"
         LeftGab         =   100
      End
      Begin FPSpread.vaSpread tblPtList 
         Height          =   6300
         Left            =   15
         TabIndex        =   21
         Top             =   1980
         Width           =   7215
         _Version        =   196608
         _ExtentX        =   12726
         _ExtentY        =   11113
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   24
         MaxRows         =   50
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frmWRD001.frx":08CA
         TextTip         =   4
         ScrollBarTrack  =   3
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   300
         Left            =   0
         TabIndex        =   22
         Top             =   1605
         Width           =   7215
         _ExtentX        =   12726
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
         Caption         =   "검체 채취 리스트"
         LeftGab         =   100
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   1215
         Left            =   0
         TabIndex        =   13
         Top             =   375
         Width           =   4275
         Begin VB.TextBox txtWardID 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   915
            MaxLength       =   9
            TabIndex        =   16
            Top             =   315
            Width           =   1395
         End
         Begin VB.CommandButton cmdWardList 
            BackColor       =   &H0098A7A5&
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2310
            Style           =   1  '그래픽
            TabIndex        =   15
            Tag             =   "WardID"
            Top             =   315
            Width           =   360
         End
         Begin VB.CommandButton cmdGetOrders 
            BackColor       =   &H00E0E0E0&
            Caption         =   "조회(&F)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3270
            Style           =   1  '그래픽
            TabIndex        =   14
            Tag             =   "0"
            Top             =   690
            Width           =   930
         End
         Begin MedControls1.LisLabel LisLabel3 
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   315
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   450
            BackColor       =   14411494
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
            Alignment       =   1
            Caption         =   "병동 ID"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   795
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   344
            BackColor       =   14411494
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
            Alignment       =   1
            Caption         =   "처방일"
            Appearance      =   0
         End
         Begin MSComCtl2.DTPicker dtpToTime 
            Height          =   315
            Left            =   900
            TabIndex        =   19
            Top             =   750
            Width           =   2310
            _ExtentX        =   4075
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
            CustomFormat    =   "yyy-MM-dd    HH:mm:ss"
            Format          =   61734912
            CurrentDate     =   36328
         End
         Begin MedControls1.LisLabel lblWardNm 
            Height          =   315
            Left            =   2685
            TabIndex        =   20
            Top             =   330
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
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
            BorderStyle     =   0
            Caption         =   ""
            Appearance      =   0
            LeftGab         =   100
         End
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "일괄채혈(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3015
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "15101"
      Top             =   8520
      Width           =   1320
   End
   Begin VB.CommandButton cmdSaveNurse 
      BackColor       =   &H00E0E0E0&
      Caption         =   "개별채혈(&P)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10395
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "0"
      Top             =   8520
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13350
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "0"
      Top             =   8520
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4365
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "0"
      Top             =   8520
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblCollect 
      Height          =   885
      Left            =   -90
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8475
      Visible         =   0   'False
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   1561
      _StockProps     =   64
      BackColorStyle  =   3
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   10
      MaxRows         =   50
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   15463405
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmWRD001.frx":146E
      Appearance      =   1
   End
   Begin MedControls1.LisLabel lblErrString 
      Height          =   300
      Left            =   5865
      TabIndex        =   70
      Top             =   8535
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
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
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin VB.Frame fraWard 
      BackColor       =   &H00DBE6E6&
      Height          =   8370
      Left            =   7290
      TabIndex        =   23
      Top             =   -75
      Width           =   7365
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   300
         Left            =   45
         TabIndex        =   24
         Top             =   135
         Width           =   7275
         _ExtentX        =   12832
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
         Caption         =   "출력 옵션"
         LeftGab         =   100
      End
      Begin VB.Frame fraPrtOption 
         BackColor       =   &H00DBE6E6&
         Height          =   1245
         Left            =   45
         TabIndex        =   32
         Top             =   345
         Width           =   7290
         Begin VB.TextBox txtCopy 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   345
            Left            =   2760
            TabIndex        =   37
            Top             =   810
            Width           =   750
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "바코드 Only"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   600
            TabIndex        =   36
            Top             =   465
            Width           =   1365
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "바코드Lable And 채혈 리스트"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2070
            TabIndex        =   35
            Top             =   465
            Value           =   -1  'True
            Width           =   2745
         End
         Begin VB.CheckBox chkPrintFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "출력안함"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   615
            TabIndex        =   34
            Top             =   150
            Width           =   1305
         End
         Begin VB.CheckBox chkTestdiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "검사코드출력"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2025
            TabIndex        =   33
            Top             =   150
            Width           =   1425
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   360
            Left            =   3495
            TabIndex        =   38
            Top             =   780
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MedControls1.LisLabel lblColList 
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   840
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   450
            BackColor       =   14411494
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
            Alignment       =   1
            Caption         =   "채혈리스트 출력장수"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblPage 
            Height          =   255
            Left            =   3780
            TabIndex        =   40
            Top             =   855
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   450
            BackColor       =   14411494
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
            Alignment       =   1
            Caption         =   "부"
            Appearance      =   0
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00DBE6E6&
         Height          =   5895
         Left            =   45
         ScaleHeight     =   5835
         ScaleWidth      =   7230
         TabIndex        =   25
         Top             =   2265
         Width           =   7290
         Begin MedControls1.LisLabel lblColNm 
            Height          =   330
            Left            =   345
            TabIndex        =   26
            Top             =   555
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   582
            BackColor       =   13752531
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
            Appearance      =   0
            LeftGab         =   100
         End
         Begin MedControls1.LisLabel lblPtCount 
            Height          =   330
            Left            =   345
            TabIndex        =   27
            Top             =   1440
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   582
            BackColor       =   13752531
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
            Appearance      =   0
            LeftGab         =   100
         End
         Begin FPSpread.vaSpread tblCount 
            Height          =   4530
            Left            =   2415
            TabIndex        =   28
            Tag             =   "15109"
            Top             =   0
            Width           =   2970
            _Version        =   196608
            _ExtentX        =   5239
            _ExtentY        =   7990
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            BackColorStyle  =   1
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   14737632
            MaxCols         =   3
            MaxRows         =   30
            Protect         =   0   'False
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            ShadowText      =   0
            SpreadDesigner  =   "frmWRD001.frx":1CAB
            VisibleCols     =   3
            VisibleRows     =   15
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   2400
            X2              =   2400
            Y1              =   0
            Y2              =   4770
         End
         Begin VB.Label Label6 
            BackColor       =   &H00DBE6E6&
            Caption         =   "환자수"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   345
            TabIndex        =   31
            Tag             =   "20104"
            Top             =   1170
            Width           =   765
         End
         Begin VB.Label lblBuildCnt 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈자"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   345
            TabIndex        =   30
            Tag             =   "20104"
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label4 
            BackColor       =   &H00DBE6E6&
            Caption         =   "명"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1620
            TabIndex        =   29
            Tag             =   "20104"
            Top             =   1515
            Width           =   270
         End
      End
      Begin MSComctlLib.ProgressBar pbrPtCnt 
         Height          =   150
         Left            =   210
         TabIndex        =   41
         Top             =   2025
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   300
         Left            =   45
         TabIndex        =   42
         Top             =   1620
         Width           =   7275
         _ExtentX        =   12832
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
         Caption         =   "진행 상황"
         LeftGab         =   100
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00D8DEDA&
         FillStyle       =   0  '단색
         Height          =   330
         Index           =   1
         Left            =   45
         Shape           =   4  '둥근 사각형
         Top             =   1935
         Width           =   7290
      End
   End
   Begin VB.Frame fraNurse 
      BackColor       =   &H00DBE6E6&
      Height          =   8370
      Left            =   5730
      TabIndex        =   43
      Top             =   -75
      Width           =   8940
      Begin VB.CheckBox chkChangeColTm 
         BackColor       =   &H00800000&
         Caption         =   "채취시간변경 : "
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   255
         Left            =   5610
         TabIndex        =   45
         Top             =   1620
         Width           =   1500
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   285
         Left            =   75
         TabIndex        =   52
         Top             =   135
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   503
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
         Caption         =   "환자 기본정보"
         LeftGab         =   100
      End
      Begin VB.Frame Frame3 
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
         Height          =   1245
         Left            =   60
         TabIndex        =   56
         Top             =   345
         Width           =   8805
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
            Left            =   990
            MaxLength       =   10
            TabIndex        =   57
            Top             =   315
            Width           =   1425
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   195
            Index           =   1
            Left            =   3225
            TabIndex        =   58
            Top             =   360
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   344
            BackColor       =   14411494
            ForeColor       =   4210752
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
            Caption         =   "성     명"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel3 
            Height          =   255
            Index           =   5
            Left            =   6180
            TabIndex        =   59
            Top             =   360
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   450
            BackColor       =   14411494
            ForeColor       =   4210752
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
            Caption         =   "성 / 나이"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   344
            BackColor       =   14411494
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
            Caption         =   "환자 ID"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   195
            Index           =   4
            Left            =   3225
            TabIndex        =   61
            Top             =   780
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   344
            BackColor       =   14411494
            ForeColor       =   4210752
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
            Caption         =   "진 료 과"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel3 
            Height          =   255
            Index           =   7
            Left            =   6180
            TabIndex        =   62
            Top             =   780
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   450
            BackColor       =   14411494
            ForeColor       =   4210752
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
            Caption         =   "병      실"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   63
            Top             =   765
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   344
            BackColor       =   14411494
            ForeColor       =   4210752
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
            Caption         =   "처 방 의"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblPtNm 
            Height          =   300
            Left            =   4095
            TabIndex        =   64
            Top             =   300
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
         Begin MedControls1.LisLabel lblSexAge 
            Height          =   300
            Left            =   7125
            TabIndex        =   65
            Top             =   315
            Width           =   1395
            _ExtentX        =   2461
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
         Begin MedControls1.LisLabel lblDoctNm 
            Height          =   300
            Left            =   1005
            TabIndex        =   66
            Top             =   720
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
            Height          =   300
            Left            =   4110
            TabIndex        =   67
            Top             =   735
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
         Begin MedControls1.LisLabel lblLocation 
            Height          =   300
            Left            =   7155
            TabIndex        =   68
            Top             =   765
            Width           =   1380
            _ExtentX        =   2434
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
      End
      Begin VB.TextBox txtMesg 
         BackColor       =   &H00F7FDF8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   1245
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   51
         ToolTipText     =   "검사 리마크를 입력하세요."
         Top             =   7350
         Width           =   7245
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00800000&
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1290
         TabIndex        =   46
         Top             =   1605
         Width           =   4890
         Begin VB.CheckBox chkSelAll 
            BackColor       =   &H00800000&
            Caption         =   "전체(&A)"
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H004A4189&
            Height          =   240
            Left            =   45
            TabIndex        =   47
            Top             =   30
            Width           =   1050
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00553755&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   1
            Left            =   1320
            Shape           =   3  '원형
            Top             =   60
            Width           =   330
         End
         Begin VB.Label lblLIS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "임상병리"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00553755&
            Height          =   180
            Left            =   1680
            TabIndex        =   49
            Top             =   45
            Width           =   795
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00496835&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   2
            Left            =   2670
            Shape           =   3  '원형
            Top             =   45
            Width           =   330
         End
         Begin VB.Label lblBBS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "혈액은행"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00496835&
            Height          =   180
            Left            =   3030
            TabIndex        =   48
            Top             =   60
            Width           =   795
         End
      End
      Begin MSComCtl2.DTPicker DTPNurse 
         Height          =   300
         Left            =   7080
         TabIndex        =   44
         Top             =   1605
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         CustomFormat    =   "yyyy-MM-dd H:mm"
         Format          =   61734915
         UpDown          =   -1  'True
         CurrentDate     =   36851.6291666667
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   285
         Left            =   180
         TabIndex        =   55
         Top             =   7365
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   503
         BackColor       =   15728622
         ForeColor       =   0
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
         Caption         =   "◈ Remark"
      End
      Begin MedControls1.LisLabel lblBar 
         Height          =   285
         Left            =   60
         TabIndex        =   50
         Top             =   1605
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   503
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
         Caption         =   "처방 리스트"
         LeftGab         =   100
      End
      Begin VB.Frame fraOrder 
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
         Height          =   5415
         Left            =   60
         TabIndex        =   53
         Top             =   1815
         Width           =   8835
         Begin FPSpread.vaSpread tblOrdSheet 
            Height          =   5025
            Left            =   90
            TabIndex        =   54
            Tag             =   "10114"
            Top             =   195
            Width           =   8610
            _Version        =   196608
            _ExtentX        =   15187
            _ExtentY        =   8864
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
            GridColor       =   14737632
            MaxCols         =   36
            MaxRows         =   19
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   14737632
            ShadowText      =   0
            SpreadDesigner  =   "frmWRD001.frx":20D3
            StartingColNumber=   2
            VirtualRows     =   24
            VisibleCols     =   5
            VisibleRows     =   19
         End
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00EFFFEE&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   930
         Index           =   0
         Left            =   75
         Shape           =   4  '둥근 사각형
         Top             =   7290
         Width           =   8505
      End
   End
End
Attribute VB_Name = "frmWRD001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsFirst             As Boolean

Private objLISCollect       As New clsLIS_Collection
Private MyPatient           As New clsPatient

Private blnCleanFg          As Boolean
Private blnCollectFg        As Boolean

Private blnCleared          As Boolean
Private PtFg                As Boolean
Private MsgFg               As Boolean
Private OrdFg               As Boolean
Private SelAllFg            As Boolean
Private blnBBS              As Boolean

Private intPtCount          As Long
Private intErrCount         As Long

Private sWorkDt             As String
Private sWorkTm             As String
Private strBBS_BldCd        As String
Private strBlgCd            As String
Private strErBldCd          As String
Private strGBldCd           As String

Private mvarWardID          As String
Private mvarWardNm          As String
Private mvarDeptCd          As String
Private mvarHosilID         As String
Private mvarRoomID          As String

Private Const lngMaxRows = 19
Private Const lngRowHeight = 12

Private lngNo As Long
Private lngWorkNo As Long
Private lngPtNm   As Long
Private lngPtid   As Long
Private lngSex    As Long
Private lngHos    As Long
Private lngColdt  As Long
Private lngTest   As Long
Private lngSpc    As Long
Private SortTF    As Boolean

Private Sub cmdClear_Click()
    Call ClearRtn(1)
    Call ClearRtn(2)
    If txtPtId.Enabled = True Then txtPtId.Text = ""
    fraQuery.ZOrder 0
    fraWard.ZOrder 0
    txtWardID.SetFocus
End Sub

Private Sub tblordersheet()
    With tblOrdSheet
        .SortBy = SortByRow
        .SortKey(1) = enCOLLIST.tcORDDIV
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Action = ActionSort
    End With
End Sub

Private Sub cmdClose_Click()
    Call ClearRtn(2)
    
    With tblPtList
        If .DataRowCnt > 0 Then
            cmdSave.Enabled = True
        End If
    End With
    
    cmdSaveNurse.Enabled = False
    fraWard.ZOrder 0
    fraQuery.ZOrder 0
End Sub

Private Sub cmdSaveNurse_Click()
     Dim objPrgBar      As clsProgress
     Dim APSColSuccess  As Boolean
     Dim BBSColSuccess  As Boolean
     Dim LISColSuccess  As Boolean
     
     Dim iCheckOrder    As Long
     Dim lngBarCnt      As Long
     Dim lngSelCnt      As Long
     Dim BarCount       As Long
     Dim SelCount       As Long
     
     Dim strColdt       As String
     Dim strcoltm       As String
     Dim strColId       As String
     
     Dim ii As Long
     
     If CollectionTargetChk = False Then
        MsgBox "채취할 항목을 선택하세요..", vbInformation, "항목선택"
        tblOrdSheet.SetFocus
        Exit Sub
     End If
   
    iCheckOrder = objLISCollect.CheckSameOrder(tblOrdSheet, 1)
    If iCheckOrder > 0 Then GoTo OrdCheck1
    MouseRunning
    Set objPrgBar = New clsProgress
    With objPrgBar
        .Container = Me
        
        .Left = fraNurse.Left + lblBar.Left
        .Top = lblBar.Top - 70
        .Width = lblBar.Width - 10
        .ForeColor = &HFA8B10
        .Appearance = ccFlat
        .BorderStyle = ccNone
        .Height = lblBar.Height - 10
        .Message = "선택된 검사항목에 대해 채취처리중입니다."
        .Max = 90
        .Min = 0
        .Value = 10
        DoEvents
    End With

    DoEvents
    
    Call tblordersheet
    
    Dim objDic As New clsDictionary
    
    objDic.Clear
    objDic.FieldInialize "orddiv", "first,last,coldt,coltm"
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = enCOLLIST.tcORDDIV
            Select Case .Value
                Case BBS_ORDDIV
                    If objDic.Exists(.Value) Then
                        objDic.KeyChange BBS_ORDDIV
                        objDic.Fields("last") = .Row
                    Else
                        .Col = enCOLLIST.tcREQDTTM
                        objDic.AddNew BBS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & _
                                      Format(.Text, "yyyymmdd") & COL_DIV & Format(.Text, "HHmm")
                    End If
                Case LIS_ORDDIV
                    If objDic.Exists(.Value) Then
                        objDic.KeyChange LIS_ORDDIV
                        objDic.Fields("last") = .Row
                    Else
                        objDic.AddNew LIS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & "" & COL_DIV & ""
                    End If
            End Select
        Next
        objDic.MoveFirst
        Do Until objDic.EOF
            If objDic.Fields("last") = "" Then objDic.Fields("last") = objDic.Fields("first")
            objDic.MoveNext
        Loop
    End With
   
    With objDic
        .MoveFirst
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case APS_ORDDIV: iCheckOrder = objLISCollect.ChkSpcnm(tblOrdSheet, .Fields("first"), .Fields("last"))
                Case LIS_ORDDIV: iCheckOrder = objLISCollect.ChkSpcnm(tblOrdSheet, .Fields("first"), .Fields("last"))
            End Select
            If iCheckOrder > 0 Then GoTo OrdCheck2
            .MoveNext
        Loop
    End With
  
    With objDic
        .MoveFirst
        BBSColSuccess = True: APSColSuccess = True: LISColSuccess = True
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case BBS_ORDDIV: BBSColSuccess = CollectForBBS_NEW(.Fields("first"), .Fields("last"), _
                                                                    Format(GetSystemDate, "yyyymmdd"), _
                                                                    Format(GetSystemDate, "HHmmss"), objPrgBar)
                Case LIS_ORDDIV: LISColSuccess = CollectForLIS_New(.Fields("first"), .Fields("last"), objPrgBar)
            End Select
            .MoveNext
        Loop
    End With
    Date = Format(GetSystemDate, CS_DateLongFormat)
    Time = Format(GetSystemDate, CS_TimeLongFormat)
     
    SelCount = 0: lngBarCnt = 0
    
    If BBSColSuccess = False Or LISColSuccess = False Then
        Set objPrgBar = Nothing
        If lblErrString.Caption = "" Then
            MsgBox "채혈처리중 오류가 발생했습니다 !!" & vbCrLf & _
                   "재실행하신 후 오류가 계속되면 전산실 혹은 임상병리과로 연락바랍니다.", _
                   vbCritical, "오류"
        End If
    End If
    
    MouseDefault
    Set objPrgBar = Nothing
    Set objDic = Nothing
    
ExitPos:
    Call cmdGetOrders_Click
    Set objDic = Nothing
    GoTo GO_DBCLOSE

OrdCheck1:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    tblOrdSheet.SetFocus
    GoTo GO_DBCLOSE
    
OrdCheck2:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "지정검체 정보가 없습니다. 전산실 혹은 임상병리과로 연락하세요.", vbInformation + vbOKOnly, "오류"
    tblOrdSheet.SetFocus
    Set objDic = Nothing
    GoTo GO_DBCLOSE
    
GO_DBCLOSE:
    If IsDBOpen = True Then
        Call DBClose
    End If
    
End Sub

Private Sub dtpColDtTm_Change()
    Dim Resp As VbMsgBoxResult

    If blnCleanFg Then Exit Sub
    If dtpColDtTm.Value < GetSystemDate Then
        If dtpColDtTm.Tag = "1" Then
            dtpColDtTm.Tag = "0"
        Else
            Resp = MsgBox("채혈시간이 현재시간보다 이전입니다. 적용하시겠습니까?", _
                   vbQuestion + vbYesNo, "채혈시간적용")
            If Resp = vbYes Then
                dtpColDtTm.Tag = "1"
            Else
                dtpColDtTm.Tag = "0"
                dtpColDtTm.Value = GetSystemDate
            End If
        End If
    End If

    With tblPtList
        If optApplyColTm(0).Value Then
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 19: .Col2 = 19
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .BlockMode = False
            .Col = 20: .Col2 = 20
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "HH:MM")
            .BlockMode = False
        Else
            If .ActiveRow > .DataRowCnt Then Exit Sub
            .Row = .ActiveRow
            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")
        End If
    End With
End Sub

Private Sub dtpToTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdGetOrders.SetFocus
End Sub

Private Sub Form_Activate()
    If Not IsFirst Then Exit Sub
    IsFirst = False
    
    txtCopy.Text = 1
    dtpToTime.Value = Format(GetSystemDate, "YYYY-MM-DD HH:MM:SS")
    dtpColDtTm.Value = GetSystemDate
    blnCleanFg = True
    intErrCount = 0
    
    txtWardID.Text = gWardId: mvarWardID = gWardId
    lblWardNm.Caption = gWardNm: mvarWardNm = gWardNm
    
GO_DBCLOSE:
    If IsDBOpen = True Then
        Call DBClose
    End If
    
On Error GoTo Err_Trap
  
    chkPrintFg.Value = 0
    optOption(0).Value = True
    Exit Sub
    
Err_Trap:
    Resume Next
End Sub

Private Sub Form_Load()
    IsFirst = True
    With tblPtList
'        ChkMornFg.Visible = False
        LisLabel7.Visible = False
        Frame1.Visible = False
        LisLabel1.Width = 5595
        Frame2.Width = 5595
        cmdGetOrders.Left = cmdGetOrders.Left + 100
    End With
End Sub

Private Sub chkPrintFg_Click()
    If chkPrintFg.Value = 1 Then
        optOption(0).Value = False
        optOption(1).Value = False
    Else
        optOption(1).Value = True
    End If
End Sub

Private Sub cmdExit_Click()
    Set objLISCollect = Nothing
    Set MyPatient = Nothing
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim Resp        As VbMsgBoxResult
    Dim intSelCount As Long
    Dim sBuildCd    As String
    Dim sBuildNm    As String
    Dim strSavePtId As String
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long

    Date = Format(GetSystemDate, "YYYY-MM-DD")
    Time = Format(GetSystemDate, "HH:mm:SS")
    blnCollectFg = False
    Set objLISCollect = New clsLIS_Collection

    sWorkDt = Format(GetSystemDate, CS_DateDbFormat)
    sWorkTm = Format(GetSystemDate, CS_TimeDbFormat)

    tblCount.Row = 0
    intErrCount = 0
    intSelCount = 0
    strSavePtId = ""

    Call SetLock(True)

    Me.MousePointer = 11

    With tblPtList
        pbrPtCnt.Visible = True
        If .DataRowCnt <> 0 Then
            pbrPtCnt.Max = .DataRowCnt * 3 * 101
        End If
        pbrPtCnt.Min = 0
        lblPtCount.Caption = ""

        For i = 1 To .DataRowCnt
            .Row = i

            .Col = 1: If .Value = 1 Then GoTo Skip

            intSelCount = intSelCount + 1
            
            .Col = 17
            If Trim(.Value) <> "" Then Call DoCollectionForBBS(i)
            
            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
            pbrPtCnt.Value = pbrPtCnt.Value + 100
            DoEvents
            
            .Col = 15
            If Trim(.Value) <> "" Then Call DoCollectionForLIS(i)
            
            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
            pbrPtCnt.Value = pbrPtCnt.Value + 100
            DoEvents




            .Row = i: .Col = 3
            If strSavePtId <> Trim(.Text) Then
               lblPtCount.Caption = Val(lblPtCount.Caption) + 1
               strSavePtId = .Text
            End If

            objLISCollect.InitRtn
            DoEvents
Skip:
        Next

        lblColNm.Caption = ObjSysInfo.EmpId
    End With

    If intSelCount = 0 Then
         Screen.MousePointer = vbDefault
         Call cmdClear_Click
         MsgBox "처리된 데이타가 없습니다..", vbInformation, "Message"
         GoTo GO_DBCLOSE
    End If
    
    If blnCollectFg = True Then
        pbrPtCnt.Value = pbrPtCnt.Max
        DoEvents
        MouseDefault
        If intErrCount > 0 Then
            MsgBox CStr(intErrCount) & "건의 오류가 발생했습니다.."
        Else
            If optOption(0).Value Then
                Call medClearTable(tblPtList)
                Resp = MsgBox("모두 정상적으로 채취처리 되었습니다.." & vbCrLf & _
                              "채취리스트를 지금 출력하시겠습니까 ? ", vbInformation + vbYesNo, "채취리스트 출력")
                If Resp = vbYes Then
                    For i = 1 To tblCount.DataRowCnt
                        tblCount.Row = i
                        tblCount.Col = 3:  sBuildCd = tblCount.Value
                        tblCount.Col = 1:  sBuildNm = tblCount.Value
                        Call CollectListPrint(txtWardID.Text, sWorkDt, sWorkTm, sBuildCd)
                    Next
                End If
            Else
                Call MsgBox("모두 정상적으로 채취처리 되었습니다..", vbInformation, "메세지")
            End If
    
            Call ClearRtn(1)
            
            '-- 일단 제외 함 By M.G.Choi 2005.06.30 =========
'            Call cmdGetOrders_Click
            '================================================
            
            On Error GoTo Err_Trap
            
        End If
    Else
        Call ClearRtn(0)
        On Error GoTo Err_Trap
        txtWardID.SetFocus
    End If
    pbrPtCnt.Visible = False
    
Err_Trap:
    Me.MousePointer = 0

GO_DBCLOSE:
    If IsDBOpen = True Then
        Call DBClose
    End If
    
End Sub

Private Sub SetLock(ByVal blnLock As Boolean)
    txtWardID.Enabled = Not blnLock
    txtWardID.BackColor = IIf(blnLock, &H8000000F, vbWhite)
    cmdWardList.Enabled = Not blnLock
    dtpToTime.Enabled = Not blnLock
    cmdGetOrders.Enabled = Not blnLock
End Sub

Private Sub DoCollectionForBBS(ByVal Row As Long) '
    Dim strPtID     As String
    Dim strPtNm     As String
    Dim strColId    As String
    Dim strColdt    As String
    Dim strcoltm    As String
    Dim strBuildcd  As String
    Dim strHosilId  As String
    Dim strStatFg   As String
    Dim blnCollect  As Boolean
    
    Dim lngErCnt    As Long
    Dim lngGcnt     As Long
    Dim lngBldRow   As Long
    Dim j           As Long

    If P_IncludeBBSSystem Then
    
        Dim objDic          As New clsDictionary
        Dim objBBSCollect   As New clsBBS_Collection
        
        Call objBBSCollect.SETWardColList(txtWardID.Text, sWorkDt, sWorkTm)
        strColId = ObjSysInfo.EmpId
        With tblPtList
            .Row = Row
            .Col = 3: strPtID = Trim(.Value)
            .Col = 4: strPtNm = .Value
            .Col = 5
            If .Value = "※" Then
                lngErCnt = lngErCnt + 1
            Else
                lngGcnt = lngGcnt + 1
            End If
            
            .Col = 12:  strHosilId = Trim(.Value)
            .Col = 19:  strColdt = Format(.Text, "YYYYMMDD")
            .Col = 20:  strcoltm = Format(.Text, "HHMMss")
            .Col = 23:  strStatFg = IIf(.Value = "1", "1", "")
            .Col = 18
            
            If .Value = "존재" Then
                If objBBSCollect.SetAccessCheck(strPtID) = True Then
                    '검체를 채취(채혈내역작성) 할필요 없이 바로 접수 하여도 된다.
                    blnCollect = objBBSCollect.SetWardAccess(strPtID, enBussDiv.BussDiv_InPatient, strColdt, strcoltm, ObjSysInfo.EmpId)
                    If blnCollect Then blnCollectFg = True
                Else
                    GoTo BBSCollect
                End If
                
            Else
BBSCollect:
                objDic.Clear
                objDic.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"
        
                objDic.AddNew strPtID, Join(Array(strPtNm, strColdt, strcoltm, strColId, _
                                            enBussDiv.BussDiv_InPatient, strBBS_BldCd, strHosilId, strStatFg), COL_DIV)
                
                If P_ApplyBuildingInfo Then
                    strBuildcd = ObjSysInfo.BuildingCd
                Else
                    strBuildcd = "10"
                End If
                    
                If objDic.RecordCount > 0 Then
                    If objBBSCollect.SET_Collect(objDic, strBuildcd, , True) Then
'                        Call objComBuilding.KeyChange(strBuildcd)
                        lngBldRow = 0
                        For j = 1 To tblCount.DataRowCnt
                            tblCount.Row = j: tblCount.Col = 3
                            If tblCount.Value = strBuildcd Then
                                lngBldRow = j
                                Exit For
                            End If
                        Next
        
                        If lngBldRow = 0 Then lngBldRow = tblCount.DataRowCnt + 1
                        tblCount.Row = lngBldRow
                        tblCount.Col = 1: tblCount.Text = "본원" 'objComBuilding.Fields("buildnm")
                        tblCount.Col = 2: tblCount.Text = Val(tblCount.Text) + 1
                        tblCount.Col = 3: tblCount.Text = strBuildcd
        
                        Dim objBAR As New clsDictionary
        
                        Set objBAR = objBBSCollect.BldDic
                        If objBAR.RecordCount > 0 Then
                            BarCode_Print objBAR
                            blnCollectFg = True
                        End If
                    End If
                End If
            End If
        End With
    
        Set objBBSCollect = Nothing
        Set objDic = Nothing
        Set objBAR = Nothing
    End If
End Sub

Private Sub BarCode_Print(objDic As clsDictionary)
    
If P_IncludeBBSSystem Then
    Dim strBuildNm  As String
    Dim strPtID     As String
    Dim strPtNm     As String
    Dim strColdt    As String
    Dim strcoltm    As String
    Dim strSpcNo    As String
    Dim strAccSeq   As String
    Dim HosilID     As String
    Dim strStatFg   As String
    Dim strBarW_H   As String
    Dim strColId As String
    Dim objPt As clsPatient
    Dim strSexAge As String
    
    If P_ApplyBuildingInfo Then
        strBuildNm = ObjSysInfo.BuildingNm
    Else
        strBuildNm = BBSName
    End If

    objDic.MoveFirst

    Do Until objDic.EOF
        strPtID = medGetP(objDic.GetString, 1, COL_DIV)
        strPtNm = medGetP(objDic.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDic.GetString, 3, COL_DIV)
        strColdt = medGetP(objDic.GetString, 4, COL_DIV)
        strColdt = Format(Mid(strColdt, 5, 4), "##/##")
        strcoltm = Mid(medGetP(objDic.GetString, 5, COL_DIV), 1, 4)
        strcoltm = Format(strcoltm, "##:##")
        HosilID = medGetP(objDic.GetString, 6, COL_DIV)
        strStatFg = medGetP(objDic.GetString, 7, COL_DIV)
        Call GetColInfo("B", strSpcNo, "", "", strColId)
        
        If HosilID <> "" Then
            strBarW_H = txtWardID & "/" & HosilID
        Else
            strBarW_H = txtWardID
        End If
        
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        
        Set objPt = New clsPatient
        objPt.GETPatient (strPtID)
        strSexAge = objPt.SEXAGE
        Set objPt = Nothing
        
        Call PrintOutBarcode("혈액", "", strColId, "", strSpcNo, strPtID, _
                         strPtNm, strSexAge, "", strStatFg, strBarW_H, _
                          strColdt, strcoltm, "", Val(txtCopy.Text))
'        BarCodePrint _
                        "XM", strBuildNm, "", strAccSeq, strSpcNo, strPtID, _
                        strPtNm, "", "", strStatFg, strBarW_H, _
                        strColdt, strcoltm, "", Val(txtCopy)
        objDic.MoveNext
    Loop

End If

End Sub

Private Sub DoCollectionForLIS(ByVal Row As Long)
    Dim tmpRs       As New Recordset
    Dim MySql       As New clsLIS_SQL
    
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpTestFg   As String
    Dim SqlStmt     As String
    Dim tmpData()   As String
    
    Dim tmpDeptCd   As String
    Dim tmpOrdDoct  As String
    Dim tmpMajDoct  As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sBuildCd    As String
    Dim sBuildNm    As String
    Dim blnMornCol  As Boolean
    Dim blnSuccess  As Boolean
    
    Dim lngBldRow   As Long
    Dim iAccseq     As Integer
    Dim i           As Long
    Dim j           As Long

    Call objLISCollect.SetWardCol(sWorkDt, sWorkTm, Trim(txtWardID))
    objLISCollect.MornFg = ChkMornFg.Value

    ReDim tmpData(0 To 16)
    
    With tblPtList
        .Row = Row
                    tmpData(0) = Mid(Format(Now, "YYYY"), 4)
        .Col = 3:   tmpData(1) = Trim(.Value)
        .Col = 4:   tmpData(2) = .Value
        .Col = 14:  tmpData(3) = .Value
        .Col = 7:
                    If IsDate(Format(.Value, CS_DateMask)) Then
                        tmpData(4) = DateDiff("y", Format(.Value, CS_DateMask), GetSystemDate)
                    Else
                        tmpData(4) = Mid(.Value, 1, 4) & "-01-01"
                        If IsDate(tmpData(4)) Then
                            tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
                        Else
                            tmpData(4) = 0
                        End If
                    End If
        .Col = 8:   tmpData(5) = .Value
                    tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)
                    tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)
                    tmpData(8) = ObjSysInfo.EmpId
                    tmpData(9) = ""
                    tmpData(10) = Format(GetSystemDate, CS_DateDbFormat)
                    objLISCollect.ColTm = Format(GetSystemDate, "HHMMSS")
                    tmpData(11) = ObjSysInfo.EmpId
        .Col = 2:   tmpData(12) = .Value
        .Col = 12:  tmpData(13) = .Value
        .Col = 13:  tmpData(14) = .Value
                    tmpData(15) = ""
                    If P_ApplyBuildingInfo Then
                        tmpData(16) = ObjSysInfo.BuildingCd
                    Else
                        tmpData(16) = "10"
                    End If
        
        Call objLISCollect.SetColData(tmpData)
        
        .Col = 22:  blnMornCol = Choose(Val(.Text) + 1, False, True)
        .Col = 5:   tmpStatFg = .Value
        .Col = 9:   tmpDeptCd = .Value
        .Col = 10:  tmpOrdDoct = .Value
        .Col = 11:  tmpMajDoct = .Value
    End With

    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)

    If blnMornCol Then
        SqlStmt = MySql.SqlReadOrderForMornCol(objLISCollect.Ptid, tmpDate, tmpTime)
    Else
    
        SqlStmt = MySql.SqlReadWardCollect(objLISCollect.Ptid, tmpDate, tmpTime, , _
                                            enBussDiv.BussDiv_InPatient, , LIS_ORDDIV)
    End If
    
    
    Set MySql = Nothing
    tmpRs.CursorLocation = adUseClient
    tmpRs.Open SqlStmt, DBConn, adOpenStatic, adLockReadOnly, adCmdText

    ReDim tmpData(0 To 20)
    With tmpRs
        For i = 1 To .RecordCount
            MsgBox "갯수:" & .RecordCount
            tmpStatFg = medGetP("" & .Fields("StatFlags").Value, 1, ";")
            tmpTestFg = medGetP("" & .Fields("StatFlags").Value, 2, ";")
    
            If P_ApplyBuildingInfo Then
                If Trim(.Fields("StatFg").Value) = "1" Then
                   If Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                      If ObjSysInfo.BuildingCd = CentralLab Or _
                         ObjSysInfo.BuildingCd = AneLab Then
                         tmpData(0) = EmergencyLab
                      Else
                         tmpData(0) = ObjSysInfo.BuildingCd
                      End If
                      tmpData(4) = "1"
                      GoTo DataSet
                   Else
                      If ObjSysInfo.BuildingCd = WomLab Or _
                         ObjSysInfo.BuildingCd = HrtLab Then
                           If Mid(tmpStatFg, EmergencyNo, 1) = "1" Then
                             tmpData(0) = EmergencyLab
                             tmpData(4) = "1"
                             GoTo DataSet
                           End If
                      End If
                   End If
                End If
                tmpData(4) = "0"
                If Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                   tmpData(0) = ObjSysInfo.BuildingCd
                   sBuildCd = tmpData(0)
                Else
                   tmpData(0) = CentralLab
                End If
            Else
                tmpData(0) = "10"
                sBuildCd = tmpData(0)
                tmpData(4) = Trim(.Fields("StatFg").Value)
            End If
                
DataSet:
            tmpData(1) = Trim("" & .Fields("WorkArea").Value)
            tmpData(2) = Trim("" & .Fields("SpcCd").Value)
            tmpData(3) = Trim("" & .Fields("StoreCd").Value)
            tmpData(5) = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)
            tmpData(6) = Trim("" & .Fields("TestDiv").Value)
            tmpData(7) = Trim("" & .Fields("MultiFg").Value)
            tmpData(8) = Trim("" & .Fields("SpcGrp").Value)
            tmpData(9) = Trim("" & .Fields("OrdDt").Value)
            tmpData(10) = Trim("" & .Fields("OrdNo").Value)
            tmpData(11) = Trim("" & .Fields("OrdSeq").Value)
            tmpData(12) = Trim("" & .Fields("OrdCd").Value)
            tmpData(13) = tmpDeptCd
            tmpData(14) = tmpOrdDoct
            tmpData(15) = tmpMajDoct
            tmpData(16) = Trim("" & .Fields("AbbrNm5").Value)
            tmpData(17) = Trim("" & .Fields("LabelCnt").Value)
            tmpData(18) = GetLabDiv(tmpData(12))
'            Call objComLisSpc.KeyChange(tmpData(2))
'            tmpData(19) = Trim("" & .Fields("SpcNm5").Value)
'            tmpData(20) = objComLisSpc.Fields("labrange")
            
            Dim strSpcAbbr As String
            Dim strLabRng As String
            Call GetSpcInfo(tmpData(2), strSpcAbbr, strLabRng)
            tmpData(19) = Trim("" & .Fields("SpcNm5").Value)
            tmpData(20) = strLabRng
            
            Call objLISCollect.SetAddLabCollect(tmpData)
            .MoveNext
            
            MsgBox "STEP:" & i
        Next
    End With

    If tmpRs.RecordCount > 0 Then
        objLISCollect.SetTrans = True
        blnSuccess = objLISCollect.DoCollection(pbrPtCnt)
        blnCollectFg = True
    Else
        GoTo Skip
    End If

Err_Trap:
    If Not blnSuccess Then
        tblPtList.Row = Row
        tblPtList.Col = -1
        tblPtList.ForeColor = vbRed
        intErrCount = intErrCount + 1
        tblPtList.Row = Row
        tblPtList.Col = 24: tblPtList.Value = objLISCollect.ErrString
    Else
        DoEvents

         For i = 1 To objLISCollect.ColCount
            Call objLISCollect.GetLabNumbers(i, sWorkArea, sAccDt, iAccseq, sBuildCd)
'            Call objComBuilding.KeyChange(sBuildCd)
           
            lngBldRow = 0
            For j = 1 To tblCount.DataRowCnt
                tblCount.Row = j: tblCount.Col = 3
                If tblCount.Value = sBuildCd Then
                    lngBldRow = j
                    Exit For
                End If
            Next
            If lngBldRow = 0 Then lngBldRow = tblCount.DataRowCnt + 1
            tblCount.Row = lngBldRow
            tblCount.Col = 1: tblCount.Text = "본원" 'objComBuilding.Fields("buildnm")
            tblCount.Col = 2: tblCount.Text = Val(tblCount.Text) + 1
            tblCount.Col = 3: tblCount.Text = sBuildCd
        Next

    End If
Skip:
    tmpRs.Close
    Set tmpRs = Nothing
End Sub

Private Sub cmdGetOrders_Click()
    Dim obj001      As clsLIS_SQL
    Dim tmpRs       As New Recordset
    Dim objProgress As clsProgress
    Dim Resp        As VbMsgBoxResult
    Dim i           As Long
    Dim SqlStmt     As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String

    If Trim(txtWardID.Text) = "" Then
        MsgBox "병동ID를 입력하세요.", vbInformation, "병동선택"
        txtWardID.SetFocus
        Exit Sub
    End If
    
    Set obj001 = New clsLIS_SQL
    Call ClearRtn(2): txtPtId.Text = ""
    
    fraWard.ZOrder 0: fraQuery.ZOrder 0
    cmdSave.Enabled = True: cmdSaveNurse.Enabled = False
    
    If Not obj001.Archive_WardColData(txtWardID.Text) Then
        MsgBox "병동일괄채취 내역 Archive중 오류가 발생했습니다." & vbCrLf & _
                "전산실 혹은 임상병리과로 연락바랍니다. (☎" & ObjSysInfo.HelpLine & ")", vbCritical, "오류발생"
    End If
    
    If ChkMornFg.Value = 1 Then
        Resp = MsgBox("채혈간호사 작업을 시작하시겠습니까?", vbQuestion + vbYesNo, "아침채혈")
        If Resp = vbNo Then
            Set obj001 = Nothing
            GoTo GO_DBCLOSE
        End If
    End If
    Me.MousePointer = 11
    
    Set objProgress = New clsProgress
    Call TableClear(1)

    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)
    MouseRunning
    
    With objProgress
        .Container = MainFrm.stsBar
        .Message = Trim(txtWardID.Text) & " 병동 환자들의 처방을 검색중입니다..."
'        .Caption = "병동일괄채취"
'        .Msg = Trim(txtWardID.Text) & " 병동 환자들의 처방을 검색중입니다.."
'        .Mode = 1
    End With
    
    '** 추가 ===================================================
    ' * 병동 Parameter (WardID, OrdDt(ReqDt))
    Dim objORDER    As New S2CON_HOS.clsPatient
            
    Call objORDER.CreateOrder("", Trim(txtWardID.Text), tmpDate)
    
    Set objORDER = Nothing
    '===========================================================
    
    If ChkMornFg.Value = 1 Then
        SqlStmt = obj001.SqlOrderForMornCol(tmpDate, txtWardID.Text)
    Else
        SqlStmt = obj001.SqlWardOrder(tmpDate, tmpTime, txtWardID.Text)
    End If

    tmpRs.CursorLocation = adUseClient
    tmpRs.Open SqlStmt, DBConn, adOpenStatic, adLockReadOnly, adCmdText
   
    If Not tmpRs.EOF Then Call DisplayOrders(tmpRs, objProgress)
        
    If Get_SpcAdd(Format(GetSystemDate, "yyyymmdd"), Trim(txtWardID.Text)) = False Then
        If tblPtList.DataRowCnt = 0 Then
            MsgBox "채취대상이 없습니다..", vbInformation, "병동채혈"
            cmdSave.Enabled = False
            GoTo Err_Trap
        End If
    End If

    cmdSave.Enabled = True
    blnCleanFg = False
    tmpRs.Close
    DoEvents

    tblPtList.SetFocus

Err_Trap:
    Set tmpRs = Nothing
    Set obj001 = Nothing
    Set objProgress = Nothing
    Me.MousePointer = 0
    MouseDefault
    
GO_DBCLOSE:
    If IsDBOpen = True Then
        Call DBClose
    End If

End Sub

Private Sub DisplayOrders(ByVal objRs As Recordset, Optional ByRef objPrgBar As Object = Nothing)
    Dim objSQL      As New clsBBS_Collection
    Dim MySql       As New clsLIS_SQL
    Dim TmpPTID     As String
    Dim tmpStatFg   As String
    Dim tmpSpcCd    As String
    Dim tmpOrdDiv   As String
    Dim i           As Long

    With tblPtList
        If P_ApplyBuildingInfo Then
            strBBS_BldCd = Get_BuildingCd(txtWardID.Text)
        Else
            strBBS_BldCd = "10"
        End If

        If Not objPrgBar Is Nothing Then
            objPrgBar.Min = 0
            objPrgBar.Max = objRs.RecordCount * 100 + 1
            objPrgBar.Value = 0
'            objPrgBar.Visible = True
            DoEvents
        End If

        .MaxRows = 0
        .MaxRows = IIf(objRs.RecordCount < 29, 29, objRs.RecordCount)
        .Row = 1

        intPtCount = 0
        For i = 1 To objRs.RecordCount
            If TmpPTID <> Trim(objRs.Fields("PtId").Value & "") Then
                If Not objPrgBar Is Nothing Then objPrgBar.Value = objPrgBar.Value + 50
                DoEvents

                intPtCount = intPtCount + 1
                .Row = intPtCount
                .Col = 2: .Text = "" & objRs.Fields("WardId").Value
                .Col = 3: .Text = Trim("" & objRs.Fields("PtId").Value)
                .Col = 4: .Text = "" & objRs.Fields("PtNm").Value
                .Col = 7: .Text = "" & objRs.Fields("DOB").Value
                .Col = 8: .Text = "" & objRs.Fields("BedInDt").Value
                .Col = 14:
                .Text = Trim("" & objRs.Fields("Sex").Value)
                If IsNumeric(.Text) Then .Text = Choose((Val(.Text) Mod 2) + 1, "F", "M")
                TmpPTID = "" & objRs.Fields("PtId").Value
            End If

            .Col = 9:  .Text = "" & objRs.Fields("DeptCd").Value
            .Col = 10: .Text = "" & objRs.Fields("OrdDoct").Value
            .Col = 11: .Text = "" & objRs.Fields("MajDoct").Value
            .Col = 12: .Text = "" & objRs.Fields("HosilId").Value
            .Col = 13: .Text = "" & objRs.Fields("RoomId").Value

            tmpStatFg = "" & objRs.Fields("StatFg").Value
            tmpOrdDiv = "" & objRs.Fields("orddiv").Value
            tmpSpcCd = "" & objRs.Fields("SpcCd").Value
            
            If tmpOrdDiv = BBS_ORDDIV Then .Col = 23: .Value = tmpStatFg
            
            If chkTestdiv.Value = 1 Then
                If tmpSpcCd = "" And tmpOrdDiv = BBS_ORDDIV Then tmpSpcCd = "혈액"
            Else
                If tmpOrdDiv = LIS_ORDDIV Then
'                    If objComLisSpc.Exists(tmpSpcCd) Then
'                        objComLisSpc.KeyChange (tmpSpcCd)
'                        tmpSpcCd = objComLisSpc.Fields("spcbarnm")
'                    Else
                        tmpSpcCd = MySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
'                    End If
                Else
                    tmpSpcCd = MySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
                End If
                If tmpSpcCd = "" And tmpOrdDiv = BBS_ORDDIV Then
                    tmpSpcCd = "혈액"
                End If
            End If
            
            If tmpStatFg = "1" Then
                .Col = 5
                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpSpcCd & ", "
                End If
                .Col = 22: .Text = "0"
            Else
                .Col = 6
                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpSpcCd & ", "
                End If
            End If
            
            If ChkMornFg.Value = 1 Then
                .Col = 22: .Text = "1"
            Else
                .Col = 22: .Text = "0"
            End If
            
            Select Case tmpOrdDiv
                Case LIS_ORDDIV:
                    .Col = 15: .ForeColor = vbRed: .Text = "√"
                Case BBS_ORDDIV:
                    .Col = 17: .ForeColor = vbRed: .Text = "√"
                    If objSQL.Blood_Existence(TmpPTID, Format(GetSystemDate, "yyyyMMdd"), _
                                                Format(GetSystemDate, "HHmm")) = True Then
                        .Col = 18: .ForeColor = vbBlue: .Value = "신규"
                    Else
                        .Col = 18: .ForeColor = DCM_Gray: .Value = "존재"
                    End If
            End Select
            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")
            objRs.MoveNext
        Next

        If Not objPrgBar Is Nothing Then objPrgBar.Value = objPrgBar.Max
        DoEvents

        pbrPtCnt.Min = 0
        pbrPtCnt.Max = .DataRowCnt * 10
        pbrPtCnt.Value = 0

        dtpColDtTm.Value = GetSystemDate
    End With

    Set MySql = Nothing
    Set objSQL = Nothing
End Sub

Private Function Get_BuildingCd(ByVal DeptCd As String) As String
    Dim objSQL  As New clsLIS_SQL
    Dim rs      As New Recordset
    
    rs.Open objSQL.Get_BuildingCd(DeptCd), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Get_BuildingCd = IIf(IsNull(rs.Fields("bldgb").Value) = True, "", rs.Fields("bldgb").Value)
    Else
        Get_BuildingCd = ""
    End If
    rs.Close
    Set rs = Nothing
    Set objSQL = Nothing
End Function

Private Function Get_SpcAdd(ByVal OrdDt As String, Wardid As String) As Boolean
    Dim objSQL      As New clsLIS_SQL
    Dim DrRS        As New Recordset
    Dim strErChk    As String
    Dim strPtID     As String
    Dim strColdt    As String
    Dim strcoltm    As String
    Dim cnt         As Long
    Dim lngRow      As Long

    Get_SpcAdd = True
    strColdt = Format(GetSystemDate, "yyyy-mm-dd")
    strcoltm = Format(GetSystemDate, "HH:mm")
    
    If P_ApplyBuildingInfo Then
        strBBS_BldCd = Get_BuildingCd(txtWardID.Text)
    Else
        strBBS_BldCd = "10"
    End If

    DrRS.Open objSQL.Get_SpcAdd(UCase(Wardid)), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not DrRS.EOF Then
        With tblPtList
            Do Until DrRS.EOF
                If DupCheck("" & DrRS.Fields("ptid").Value, lngRow) = False Then
                    If .DataRowCnt <= .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .ForeColor = vbBlue
                    .Col = 2: .Value = Wardid
                    .Col = 3: .Value = Trim("" & DrRS.Fields("ptid").Value): strPtID = Trim("" & DrRS.Fields("ptid").Value)
                    .Col = 4: .Value = "" & DrRS.Fields("ptnm").Value
                    strErChk = objSQL.ER_Chk(strPtID, "" & DrRS.Fields("orddt").Value)
                    .Col = 5: .Value = IIf(strErChk = "1", "혈액,", "")
                    .Col = 6: .Value = IIf(strErChk = "0", "혈액,", "")
                    .Col = 7: .Value = "" & DrRS.Fields("dob").Value
                    .Col = 8: .Value = "" & DrRS.Fields("bedindt").Value
                    .Col = 14:
                        .Text = Trim("" & DrRS.Fields("Sex").Value)
                        If IsNumeric(.Text) Then .Text = Choose((Val(.Text) Mod 2) + 1, "F", "M")
                    Select Case "" & DrRS.Fields("orddiv").Value
                        Case "L":
                            .Col = 15: .ForeColor = vbRed: .Text = "√"
                        Case "B":
                            .Col = 17: .ForeColor = vbRed: .Text = "√"
                    End Select
                    .Col = 18: .Value = "추가"
                    .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
                    .Col = 20: .Value = Format(dtpColDtTm.Value, "HH:MM:SS")

                    .Col = 9: .Text = "" & DrRS.Fields("DeptCd").Value
                    .Col = 10: .Text = "" & DrRS.Fields("OrdDoct").Value
                    .Col = 11: .Text = "" & DrRS.Fields("MajDoct").Value
                    .Col = 12: .Text = "" & DrRS.Fields("RoomId").Value
                    .Col = 13: .Text = "" & DrRS.Fields("HosilId").Value
                    cnt = cnt + 1
                Else
                    .Row = lngRow
                    .ForeColor = vbBlue
                    .Col = 17: .ForeColor = vbRed: .Text = "√"
                    .Col = 18: .Value = "추가"
                    .Col = 21: .Value = "*"
                End If
                DrRS.MoveNext
            Loop
        End With
    Else
        Get_SpcAdd = False
    End If

    If cnt = 0 Then Get_SpcAdd = False

    DrRS.Close
    Set DrRS = Nothing
    Set objSQL = Nothing
End Function

Private Function DupCheck(ByVal pPtID As String, ByRef plngRow As Long) As Boolean
    Dim ii      As Long
    
    DupCheck = False
    With tblPtList
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 3
            If .Value = pPtID Then
                DupCheck = True
                plngRow = ii
                Exit For
            End If
        Next
    End With
End Function

Private Sub dtpToTime_Change()
    If Not blnCleanFg Then Call TableClear(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objLISCollect = Nothing
    Set MyPatient = Nothing
End Sub

Private Sub optApplyColTm_Click(Index As Integer)
    Dim Resp As VbMsgBoxResult

    If dtpColDtTm.Value < Now Then
        If dtpColDtTm.Tag = "1" Then
            dtpColDtTm.Tag = "0"
        Else
            Resp = MsgBox("채취시간이 현재시간보다 이전입니다. 적용하시겠습니까?", _
                   vbQuestion + vbYesNo, "채취시간적용")
            If Resp = vbYes Then
                dtpColDtTm.Tag = "1"
            Else
                dtpColDtTm.Tag = "0"
                dtpColDtTm.Value = Format(GetSystemDate, "YY-MM-DD HH:MM")
            End If
        End If
    End If

    With tblPtList
        If optApplyColTm(0).Value Then
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 19: .Col2 = 19
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .BlockMode = False
            .Col = 20: .Col2 = 20
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "HH:MM")
            .BlockMode = False
        Else
            If .ActiveRow > .DataRowCnt Then Exit Sub
            .Row = .ActiveRow
            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")
            optApplyColTm(1).Value = False
        End If
    End With
End Sub

Private Sub optOption_Click(Index As Integer)
    Select Case Index
        Case 0, 2: txtCopy.Text = 1
                txtCopy.Enabled = True
        Case 1: txtCopy.Text = 0
                txtCopy.Enabled = False
    End Select
End Sub

Private Sub cmdWardList_Click()
    Dim objMySQL    As New clsLIS_SQL
    Dim objPop      As clsPopUpList
    
    Set objPop = New clsPopUpList
    With objPop
        .Connection = DBConn
        .FormCaption = "병동 리스트"
        .ColumnHeaderText = "병동코드;병동명"
        .ColumnHeaderWidth = "1260.284;1620.284"
        .FormWidth = 3360
        .FormLeft = 990
        .FormTop = 2775
        .LoadPopUp GetSQLWardList '(objMySQL.LoadWardId)
        txtWardID.Text = medGetP(.SelectedString, 1, .Delimiter)
        lblWardNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
    End With
    
    Set objPop = Nothing
    Set objMySQL = Nothing
    
GO_DBCLOSE:
    If IsDBOpen = True Then
        Call DBClose
    End If
    
End Sub

Private Sub tblOrdSheet_DblClick(ByVal Col As Long, ByVal Row As Long)
    cmdSave.Enabled = True
    cmdSaveNurse.Enabled = False
    fraWard.ZOrder 0
    fraQuery.ZOrder 0
End Sub


Private Sub SPreadSort(ByVal Col As Integer)
    If Col < 3 Or Col > 4 Then Exit Sub
    With tblPtList
        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = Col
        If SortTF = True Then
            .SortKeyOrder(1) = SortKeyOrderAscending
            SortTF = False
        Else
            SortTF = True
            .SortKeyOrder(1) = SortKeyOrderDescending
        End If
        .Col = 1:  .Col2 = .MaxCols
        .Row = 1:  .Row2 = .DataRowCnt
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
    End With
End Sub

Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        Call SPreadSort(Col)
        Exit Sub
    End If
    If Col < 2 Then Exit Sub
    tblPtList.Row = Row
    tblPtList.Col = 3
    If tblPtList.Value = "" Then Exit Sub
    
    txtPtId.Text = tblPtList.Value
    cmdSave.Enabled = False
    Call GetPtCollection(txtPtId.Text)
    Call DisplayOrder
    If blnBBS = False Then Call Get_PtSpcAdd
    fraNurse.ZOrder 0
End Sub

Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim objSQL      As New clsLIS_SQL
    Dim tmpRs       As New Recordset
    
    Dim tmpToolTip  As String
    Dim strSql      As String
    Dim strPtID     As String
    Dim strOrdDate  As String
    Dim strOrdDiv   As String
    Dim strWardID   As String
    Dim strAPSORDCd As String
    Dim strBBSORDCd As String
    Dim strLISORDCd As String

    If Row = 0 Then Exit Sub

    tmpToolTip = vbCrLf

    With tblPtList
        .Row = Row

        .Col = 2: If Trim(.Value) = "" Then Exit Sub

        .Col = 4: tmpToolTip = tmpToolTip & "  " & .Value & vbCrLf & vbCrLf
        .Col = 5: tmpToolTip = tmpToolTip & "  응급검체 : " & .Value & vbCrLf
        .Col = 6: tmpToolTip = tmpToolTip & "  일반검체 : " & .Value & vbCrLf
        
        .Col = 3: strPtID = Trim(.Value)
        strOrdDate = Format(dtpToTime.Value, CS_DateDbFormat)
        strWardID = Trim(txtWardID.Text)
        
        strSql = objSQL.WardMn_ORDCD(strPtID, strOrdDate, strWardID)
        tmpRs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If tmpRs.BOF = False Then
            Do Until tmpRs.EOF = True
                strOrdDiv = Trim(tmpRs.Fields("orddiv").Value & "")
                
                Select Case strOrdDiv
                    Case "B"
                        strBBSORDCd = strBBSORDCd & tmpRs.Fields("abbrnm5").Value & ","
                        
                    Case "L"
                        strLISORDCd = strLISORDCd & tmpRs.Fields("abbrnm5").Value & ","
                End Select
                
                tmpRs.MoveNext
            Loop
        End If

        If strBBSORDCd <> "" Then
                tmpToolTip = tmpToolTip & "  혈액은행 : " & strBBSORDCd & vbCrLf
        ElseIf strLISORDCd <> "" Then
                tmpToolTip = tmpToolTip & "  임상병리 : " & strLISORDCd & vbCrLf
        End If
        
        With tblPtList
            .Row = Row
            .Col = 24
            If Trim(.Value) <> "" Then tmpToolTip = tmpToolTip & "  오    류 : " & .Value
        End With
        
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 5000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
    tmpRs.Close
    Set tmpRs = Nothing
    Set objSQL = Nothing
End Sub

Private Sub txtWardId_Change()
    If Not blnCleanFg Then Call TableClear(1)
End Sub

Private Sub ClearRtn(ByVal intOpt As Long)
    Select Case intOpt
        Case 1
            txtWardID.Enabled = True
            txtWardID.BackColor = &H80000005
            cmdWardList.Enabled = True
            dtpToTime.Enabled = True
            cmdGetOrders.Enabled = True
            cmdSave.Enabled = False
        
            sWorkDt = "": sWorkTm = ""
            dtpToTime.Value = Format(GetSystemDate, "YYYY-MM-DD hh:mm:ss")
            dtpColDtTm.Value = GetSystemDate
            dtpColDtTm.Tag = "0"
            pbrPtCnt.Value = 0
            chkPrintFg = 0
            optOption(1).Value = True
            optApplyColTm(0).Value = True
            intErrCount = 0
            Call TableClear(intOpt)
        Case Else
             lblPtNm.Caption = ""
             lblSexAge.Caption = ""
             lblDeptNm.Caption = ""
             lblLocation.Caption = ""
             lblDoctNm.Caption = ""
             txtMesg.Text = ""
             chkSelAll.Value = 0
             chkChangeColTm.Value = 0
             dtpColDtTm.Value = GetSystemDate
             dtpColDtTm.Enabled = False
             lblErrString.Caption = ""
             With tblOrdSheet
                 .Row = -1
                 .Col = -1
                 .BlockMode = True
                 .Action = ActionClearText
                 .BlockMode = False
             End With
             
             cmdSaveNurse.Enabled = True
    End Select
    
GO_DBCLOSE:
    If IsDBOpen = True Then
        Call DBClose
    End If
    
End Sub

Private Sub TableClear(ByVal intOpt As Long)
    tblPtList.MaxRows = 0
    tblPtList.MaxRows = 50
    If intOpt = 1 Then
        lblColNm.Caption = ""
        lblPtCount.Caption = ""
        tblCount.MaxRows = 0
        tblCount.MaxRows = 50
        blnCleanFg = True
    End If
End Sub

Private Sub txtWardId_GotFocus()
    With txtWardID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWardID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Trim(txtWardID.Text) = "" Then Call cmdWardList_Click
    End If
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_Trap

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = vbKeyReturn Then
        If txtWardID.Text = "" Then
            lblWardNm.Caption = ""
            Exit Sub
        Else
'            Dim objWard As clsBasisData
            Dim rs As Recordset
            Dim strWard As String
            
'            Set objWard = New clsBasisData
            Set rs = New Recordset
            
            strWard = GetSQLWard(txtWardID.Text)
            
            rs.Open strWard, DBConn
            
            If rs.EOF = False Then
                ObjSysInfo.BuildingCd = rs.Fields("bldgb").Value & ""
                ObjSysInfo.BuildingNm = rs.Fields("bldnm").Value & ""
                ObjSysInfo.BuildingNo = rs.Fields("bldno").Value & ""
                txtWardID.Tag = txtWardID.Text
            Else
                MsgBox "병동 코드를 확인하세요.", vbInformation
                txtWardID.Text = ""
                lblWardNm.Caption = ""
                txtWardID.SetFocus
                Call txtWardID_KeyDown(vbKeyDown, 0)
            End If
            Set rs = Nothing
'            Set objWard = Nothing

'            With ObjLISComCode.WardId
'                If .Exists(txtWardID.Text) Then
'                    Call .KeyChange(txtWardID.Text)
'                    lblWardNm.Caption = .Fields("WardNm")
'                    objsysinfo.BuildingCd = .Tags("bldgb")
'                    objsysinfo.BuildingNm = .Tags("bldnm")
'                    objsysinfo.BuildingNo = .Tags("bldno")
'                    dtpToTime.SetFocus
'                Else
'                    MsgBox "병동 코드를 확인하세요..", vbInformation, "코드입력오류"
'                    txtWardID.Text = ""
'                    lblWardNm.Caption = ""
'                    txtWardID.SetFocus
'                    Call txtWardId_KeyDown(vbKeyDown, 0)
'                    Exit Sub
'                End If
'            End With
        End If
    End If
    
GO_DBCLOSE:
    If IsDBOpen = True Then
        Call DBClose
    End If
    
    Exit Sub

Err_Trap:
    Resume Next

End Sub

Private Sub CollectListPrint(ByVal pWardID As String, _
                             ByVal pWorkDt As String, ByVal pWorkTm As String, _
                             ByVal pBuildCd As String)
    Dim ii          As Long
    Dim objMySQL    As New clsLIS_WardColList
    
    objMySQL.objDictionary.DELETEALL

    Me.MousePointer = 11
    
    If objMySQL.CollectQueryTF(pWorkDt, pWardID, pWorkTm, pBuildCd, chkTestdiv.Value) = True Then
        ii = 1
        With tblCollect
            .MaxRows = 0
            .MaxRows = objMySQL.objDictionary.RecordCount
            
            Call medClearTable(tblCollect)
            objMySQL.objDictionary.MoveFirst
            Do Until objMySQL.objDictionary.EOF
                .Row = ii
                .Col = 2:       .Value = objMySQL.objDictionary.Fields("seq")
                .Col = 3:       .Value = objMySQL.objDictionary.Fields("workno")
                .Col = 4:       .Value = objMySQL.objDictionary.Fields("ptnm")
                .Col = 5:       .Value = objMySQL.objDictionary.Fields("ptid")
                .Col = 6:       .Value = objMySQL.objDictionary.Fields("sexage")
                .Col = 7:       .Value = objMySQL.objDictionary.Fields("hosilid")
                .Col = 8:       .Value = objMySQL.objDictionary.Fields("collectdt")
                .Col = 9:       .Value = objMySQL.objDictionary.Fields("testlist")
                .Col = 10:      .Value = objMySQL.objDictionary.Fields("spcnm")
                ii = ii + 1
                objMySQL.objDictionary.MoveNext
            Loop
        End With
        
        Call VBPrint
    End If
    Set objMySQL = Nothing
    Me.MousePointer = 0
End Sub

Public Sub Call_WardId_KeyPress()
    Call txtWardId_KeyPress(vbKeyReturn)
End Sub

Public Sub Call_cmdGetOrders_click()
     Call cmdGetOrders_Click
End Sub

Private Sub txtWardId_LostFocus()
On Error GoTo Err_Trap
    If ActiveControl.Name = cmdWardList.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If txtWardID.Text = "" Then
        lblWardNm.Caption = ""
        Exit Sub
    Else
        Call txtWardId_KeyPress(vbKeyReturn)
    End If
    Exit Sub
Err_Trap:
    Resume Next
End Sub

Private Sub GetPtCollection(ByVal sPtid As String)

    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    If Not blnCleared Then Call ClearRtn(2)
    DoEvents
    
'    Call MyPatient.ClearData
    If MyPatient.GETPatient(Trim(txtPtId.Text)) Then
        lblPtNm.Caption = MyPatient.PtNm
        lblSexAge.Caption = MyPatient.SEXNM & " / " & MyPatient.AGE & " " & MyPatient.AGEDIV
        lblDeptNm.Caption = MyPatient.DeptNm
        lblLocation.Caption = MyPatient.Wardid & "-" & MyPatient.RoomId & "-" & MyPatient.BedID
        mvarWardID = MyPatient.Wardid
        DoEvents
        PtFg = True
        MouseRunning
        MouseDefault
    Else
        txtPtId.Text = ""
        MsgFg = True
        MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
        MsgFg = False
        txtPtId.SetFocus
        PtFg = False
        Call txtPtId_GotFocus
        Exit Sub
    End If
    If OrdFg Then
    Else
        cmdSave.Enabled = False
        txtPtId.SetFocus
        Call txtPtId_GotFocus
    End If
End Sub

Private Sub txtPtId_Change()
    If Not blnCleared Then
       Call ClearRtn(2)
    End If
End Sub

Private Sub txtPtId_GotFocus()
   With txtPtId
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)

    If Trim(txtPtId.Text) = "" Then Exit Sub

    If KeyAscii = vbKeyReturn Then
        Call GetPtCollection(Trim(txtPtId.Text))
        Call DisplayOrder
    End If
End Sub

Private Sub DisplayOrder()
    Dim i           As Long
    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpTestFg   As String
    Dim strErChk    As String
    Dim strOrdDiv   As String
    
    Dim sTestCD     As String
    
    Dim objProInSts As clsProgress
    Dim MySql       As New clsLIS_SQL
    Dim objGetSql   As New clsBBS_Collection
    Dim tmpRs       As New Recordset
    Dim strDoctnm As String
   
On Error GoTo NoData
    TestBuilding_Search
   
    Set objProInSts = New clsProgress
    With objProInSts
        .Container = Me
        
        .Left = fraNurse.Left + lblBar.Left
        .Top = lblBar.Top - 70
        .Width = lblBar.Width
        .ForeColor = &HFA8B10
        .Appearance = ccFlat
        .BorderStyle = ccNone
        .Height = lblBar.Height '
        .Message = "해당환자의 처방 내역을 검색 중입니다...."
        .Max = 90
        .Min = 0
        .Value = 10
        DoEvents
    End With

    DoEvents
    txtMesg.Text = ""
    
    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)

    strOrdDiv = Mid(ObjSysInfo.ProjectId, 1, 1)
    
    blnBBS = False

    SqlStmt = MySql.SqlReadWardOrder(Trim(txtPtId.Text), tmpDate, tmpTime, , enBussDiv.BussDiv_InPatient, , strOrdDiv)

    tmpRs.CursorLocation = adUseClient
    tmpRs.Open SqlStmt, DBConn, adOpenStatic, adLockReadOnly, adCmdText
    If tmpRs.EOF Then
        tmpRs.Close
        Set tmpRs = Nothing
        Set MySql = Nothing
        Set objProInSts = Nothing
        blnBBS = True
        Call Get_PtSpcAdd
        If tblOrdSheet.DataRowCnt = 0 Then
            MsgBox MyPatient.PtNm & " 님의 처방내역이 없습니다", vbInformation, "환자별 채혈"
            If Not blnCleared Then Call ClearRtn(2): txtPtId.Text = ""
        Else
            With tblOrdSheet
                .Row = -1
                .Col = 2: .Col2 = .MaxCols
                .BlockMode = True
                .Lock = True
                .Protect = True
                .BlockMode = False
            End With
            OrdFg = True
            fraOrder.Enabled = True
            blnCleared = False
        End If
        Exit Sub
    End If
    
    With tblOrdSheet
        .ReDraw = False
        .MaxRows = 0
        If tmpRs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
            .Row = tmpRs.RecordCount + 1
            .Row2 = lngMaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        Else
            .MaxRows = tmpRs.RecordCount
        End If
       
        objProInSts.Max = tmpRs.RecordCount

        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
        
        SelAllFg = True
        
        For i = 1 To tmpRs.RecordCount
            objProInSts.Value = i
            .Row = i
            .Col = enCOLLIST.tcCHECK: .Value = 1
            strDoctnm = GetDoctNm(tmpRs.Fields("orddoct").Value & "")
            
            If SvOrdDt <> Trim("" & tmpRs.Fields("OrdDt").Value) Then
                .Col = enCOLLIST.tcOrddt:   .Text = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateShortMask)
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)
                .Col = enCOLLIST.tcSpcNm:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)
                .Col = enCOLLIST.tcDOCTNM:  .Text = strDoctnm 'Trim("" & tmpRs.Fields("DoctNm").Value)
                SvOrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvOrdDoct = strDoctnm 'Trim("" & tmpRs.Fields("DoctNm").Value)
            End If
            If SvOrdNo <> Trim("" & tmpRs.Fields("OrdNo").Value) Then
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)
                .Col = enCOLLIST.tcSpcNm:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)
                .Col = enCOLLIST.tcDOCTNM:  .Text = strDoctnm 'Trim("" & tmpRs.Fields("DoctNm").Value)
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvOrdDoct = strDoctnm 'Trim("" & tmpRs.Fields("DoctNm").Value)
            End If
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                .Col = enCOLLIST.tcSpcNm:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            End If
            If SvOrdDoct <> Trim(strDoctnm) Then
                .Col = enCOLLIST.tcDOCTNM: .Text = strDoctnm
                SvOrdDoct = Trim(strDoctnm)
            End If

            tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")
            tmpTestFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 2, ";")

            Select Case tmpRs.Fields("orddiv")
            Case APS_ORDDIV:
                .Col = enCOLLIST.tcStatfg:  .Text = Trim("" & tmpRs.Fields("StatFg").Value)      '응급여부  --> 위에서 처리...
                .Col = enCOLLIST.tcBUILDCD: .Text = CentralLab
                .Col = enCOLLIST.tcBUILDNM: .Text = CentralLabNm
            
            Case BBS_ORDDIV:
                strErChk = MySql.ER_Chk(txtPtId.Text, SvOrdDt)
                .Col = enCOLLIST.tcStatfg: .Value = Trim("" & tmpRs.Fields("StatFg").Value)     '응급여부  --> 위에서 처리...
                .Col = enCOLLIST.tcBUILDCD: .Value = IIf(strErChk = "1", strErBldCd, strGBldCd)
                Dim strBuildcd As String
                strBuildcd = .Value
'                If objComBuilding.Exists(.Value) Then
'                    objComBuilding.KeyChange (.Value)
'                End If
                .Col = enCOLLIST.tcBUILDNM: .Value = GetBuildNm(strBuildcd) 'objComBuilding.Fields("buildnm")
            
            Case LIS_ORDDIV:
                If P_ApplyBuildingInfo Then
                   If Trim(tmpRs.Fields("StatFg").Value) = "1" Then
                       If Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                           If ObjSysInfo.BuildingCd = CentralLab Or _
                              ObjSysInfo.BuildingCd = AneLab Then
                               .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
                               .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
                           Else
                               .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
                               .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
                           End If
                           .Col = enCOLLIST.tcSTATFLAG: .Text = "1"
                           GoTo DataSet
                       Else
                           If ObjSysInfo.BuildingCd = WomLab Or ObjSysInfo.BuildingCd = HrtLab Then
                               If Mid(tmpStatFg, EmergencyNo, 1) = "1" Then
                                   .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
                                   .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
                                   .Col = enCOLLIST.tcSTATFLAG:   .Text = "1"
                                   GoTo DataSet
                               End If
                           End If
                       End If
                   End If
    
                   .Col = enCOLLIST.tcSTATFLAG: .Text = "0"
                   If Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                       .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
                       .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
                   Else
                       .Col = enCOLLIST.tcBUILDCD: .Text = CentralLab
                       .Col = enCOLLIST.tcBUILDNM: .Text = CentralLabNm
                   End If
                Else
                    .Col = enCOLLIST.tcBUILDCD:  .Text = "10"
                    .Col = enCOLLIST.tcBUILDNM:  .Text = "본원"
                    .Col = enCOLLIST.tcSTATFLAG: .Text = Trim(tmpRs.Fields("StatFg").Value)
                End If
            End Select
DataSet:
            .Col = enCOLLIST.tcTestnm:  .Text = Trim("" & tmpRs.Fields("TestNm").Value)
            
            Select Case tmpRs.Fields("orddiv")
                Case APS_ORDDIV: .ForeColor = &H5E3F00
                Case BBS_ORDDIV: .ForeColor = &H496835
                Case LIS_ORDDIV: .ForeColor = &H553755
            End Select
            
            If Trim("" & tmpRs.Fields("WorkArea").Value) = RI_WORKAREA Or Trim("" & tmpRs.Fields("WorkArea").Value) = RI_WORKAREA Then
                .ForeColor = DCM_LightRed
            End If
            .Col = enCOLLIST.tcStatfg:  .Text = IIf("" & tmpRs.Fields("StatFg").Value = "0", "", "Y")
                                        .ForeColor = DCM_Red
            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)
            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & tmpRs.Fields("OrdDt").Value)
            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & tmpRs.Fields("OrdNo").Value)
            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & tmpRs.Fields("OrdSeq").Value)
            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & tmpRs.Fields("OrdCd").Value)

            sTestCD = Trim("" & tmpRs.Fields("OrdCd").Value)
            .Col = enCOLLIST.tcLABDIV:  .Text = GetLabDiv(sTestCD)
            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & tmpRs.Fields("SpcCd").Value)
            
            Dim tmpSpcNm As String
            Dim tmpLabRng As String
            
            Call GetSpcInfo("" & tmpRs.Fields("SpcCd").Value, tmpSpcNm, tmpLabRng)

'            Call objComLisSpc.KeyChange(.Text)
            .Col = enCOLLIST.tcSPCABBR:  .Text = Trim("" & tmpRs.Fields("spcnm5").Value)
            .Col = enCOLLIST.tcLABRANGE: .Text = tmpLabRng 'objComLisSpc.Fields("labrange")
            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & tmpRs.Fields("WorkArea").Value)
            .Col = enCOLLIST.tcSTORECD:  .Text = Trim("" & tmpRs.Fields("StoreCd").Value)
            .Col = enCOLLIST.tcTESTDIV:  .Text = Trim("" & tmpRs.Fields("TestDiv").Value)
            .Col = enCOLLIST.tcMULTIFG:  .Text = Trim("" & tmpRs.Fields("MultiFg").Value)
            .Col = enCOLLIST.tcSPCGRP:   .Text = Trim("" & tmpRs.Fields("SpcGrp").Value)
            .Col = enCOLLIST.tcORDDOCT:  .Text = Trim("" & tmpRs.Fields("OrdDoct").Value)
                                         If .Text <> "" And lblDoctNm.Caption = "" Then lblDoctNm.Caption = strDoctnm ' Trim("" & tmpRs.Fields("DoctNm").Value)
            .Col = enCOLLIST.tcMAJDODT:  .Text = Trim("" & tmpRs.Fields("MajDoct").Value)
            .Col = enCOLLIST.tcDeptCD:   .Text = Trim("" & tmpRs.Fields("DeptCd").Value)
                                         If .Text <> "" And lblDeptNm.Caption = "" Then
                                            lblDeptNm.Caption = GetDeptNm(.Text)
'                                            If objComDeptCD.Exists(.Text) Then
'                                                objComDeptCD.KeyChange (.Text)
'                                                lblDeptNm.Caption = objComDeptCD.Fields("deptnm")
'                                            End If
                                         End If
            .Col = enCOLLIST.tcABBRNM:  .Text = Trim("" & tmpRs.Fields("AbbrNm5").Value)
            .Col = enCOLLIST.tcBARCNT:  .Text = Trim("" & tmpRs.Fields("LabelCnt").Value)
            .Col = enCOLLIST.tcPAYDT:   .Text = Trim("" & tmpRs.Fields("ReceptNo").Value)
                                        .ForeColor = vbRed

            .Col = enCOLLIST.tcWardID:  .Text = Trim("" & tmpRs.Fields("WardId").Value)
                                        mvarWardID = .Text
            .Col = enCOLLIST.tcROOMID:  .Text = Trim("" & tmpRs.Fields("hosilid").Value)
                                        mvarHosilID = .Text
            .Col = enCOLLIST.tcBEDID:   .Text = Trim("" & tmpRs.Fields("roomid").Value)
                                        mvarRoomID = .Text
            .Col = enCOLLIST.tcFRZFG:   .Text = Trim("" & tmpRs.Fields("FzFg").Value)
            .Col = enCOLLIST.tcORDDIV:  .Text = Trim("" & tmpRs.Fields("OrdDiv").Value)
            
            If mvarWardID <> "" Then lblLocation.Caption = mvarWardID & "-" & mvarHosilID & "-" & mvarRoomID
            
            If Trim("" & tmpRs.Fields("Mesg").Value) <> "" Then
                txtMesg.Text = txtMesg.Text & "# " & Format(Trim("" & tmpRs.Fields("OrdNo").Value), "##") & " - "
                txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("TestNm").Value) & vbCrLf
                txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("Mesg").Value) & vbCrLf
            End If
            tmpRs.MoveNext
        Next
        .RowHeight(-1) = lngRowHeight
        .ReDraw = True
    End With
    OrdFg = True
    fraOrder.Enabled = True
    blnCleared = False
    Set objProInSts = Nothing
    
NoData:
'    tmpRs.Close
    Set tmpRs = Nothing
    Set MySql = Nothing
End Sub

Private Sub TestBuilding_Search()
    Dim objSQL As New clsBBS_SQL
    Dim strTmp As String
    
    With objSQL
        If P_ApplyBuildingInfo Then
            If mvarWardID = "" Then
                strBlgCd = ObjSysInfo.BuildingCd
            Else
                strBlgCd = Get_BuildingCd(gWardId)
            End If
        Else
            strBlgCd = "10"
        End If
        strTmp = .TestBuildCd(strBlgCd)
        strErBldCd = medGetP(strTmp, 1, COL_DIV)
        strGBldCd = medGetP(strTmp, 2, COL_DIV)
    End With

    Set objSQL = Nothing
End Sub

Private Function CollectForLIS_New(ByVal FRowCnt As Long, _
                               ByVal LRowCnt As Long, _
                               ByRef objProgress As Object) As Boolean
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim SqlStmt     As String
    Dim tmpData()   As String
    Dim ColSuccess  As Boolean
    Dim i           As Long
    Dim SelCount    As Long
    Dim strTmp1     As String
    Dim strReqDt    As String
    Dim strReqtm    As String
    Dim strReqTm1   As String
    Dim strLastTm   As String
    Dim CollectCnt  As Long
    
    strLastTm = ""

    Date = GetSystemDate
    Time = GetSystemDate

    CollectCnt = 0
    Call objLISCollect.InitRtn

    With tblOrdSheet

        ReDim tmpData(0 To 20)
        .Row = FRowCnt: .Col = enCOLLIST.tcWardID: mvarWardID = .Value
                        .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
                        .Col = enCOLLIST.tcDeptCD: mvarDeptCd = .Value
        For i = FRowCnt To LRowCnt
            
            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            
            .Row = i
            
            .Col = enCOLLIST.tcCHECK
            If .Value <> 1 Then GoTo Skip
            
            CollectCnt = CollectCnt + 1
            .Col = 36: strTmp1 = .Value
            .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .Value
            .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .Value
            .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .Value
            .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .Value
            .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .Value
            .Col = enCOLLIST.tcREQDTTM:
                If strTmp1 = "1" Then
                    strReqDt = medGetP(.Value, 1, " ")
                    If strLastTm = "" Then
                        strReqtm = Val(Mid(medGetP(.Value, 2, " "), 1, 2)) + 1
                        strLastTm = strReqtm
                    Else
                        strReqtm = Val(strLastTm) + 1
                    End If
                    strReqTm1 = Mid(medGetP(.Value, 2, " "), 3)
                    strReqtm = strReqtm & strReqTm1
                    strReqDt = strReqDt & " " & strReqtm
                    tmpData(5) = strReqDt
                Else
                    tmpData(5) = .Value
                End If
            .Col = enCOLLIST.tcTESTDIV:  tmpData(6) = .Value
            .Col = enCOLLIST.tcMULTIFG:  tmpData(7) = .Value
            .Col = enCOLLIST.tcSPCGRP:   tmpData(8) = .Value
            .Col = enCOLLIST.tcORDDATE:  tmpData(9) = .Value
            .Col = enCOLLIST.tcORDNUM:   tmpData(10) = .Value
            .Col = enCOLLIST.tcORDSEQ:   tmpData(11) = .Value
            .Col = enCOLLIST.tcTESTCD:   tmpData(12) = .Value
            .Col = enCOLLIST.tcDeptCD:   tmpData(13) = .Value
            .Col = enCOLLIST.tcORDDOCT:  tmpData(14) = .Value
            .Col = enCOLLIST.tcMAJDODT:  tmpData(15) = .Value
            .Col = enCOLLIST.tcABBRNM:   tmpData(16) = .Value
            .Col = enCOLLIST.tcBARCNT:   tmpData(17) = .Value
            .Col = enCOLLIST.tcLABDIV:   tmpData(18) = .Value
            .Col = enCOLLIST.tcSPCABBR:  tmpData(19) = .Value
            .Col = enCOLLIST.tcLABRANGE: tmpData(20) = .Value
            
            Call objLISCollect.AddLabCollect(tmpData)
Skip:
        Next
    End With

    If CollectCnt = 0 Then
        CollectForLIS_New = True
        Exit Function
    End If

    With objLISCollect

        ReDim tmpData(0 To 16)

        tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)
        tmpData(1) = MyPatient.Ptid
        tmpData(2) = MyPatient.PtNm
        tmpData(3) = MyPatient.Sex
        If IsDate(Format(MyPatient.DOB, CS_DateLongMask)) Then
            tmpData(4) = DateDiff("y", Format(MyPatient.DOB, CS_DateLongMask), GetSystemDate)
        Else
            tmpData(4) = Mid(MyPatient.DOB, 1, 4) & "-01-01"
            If IsDate(tmpData(4)) Then
                tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
            Else
                tmpData(4) = 0
            End If
        End If
        tmpData(5) = MyPatient.BedIndt
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)
        tmpData(8) = ObjSysInfo.EmpId
        tmpData(9) = ""
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat)
        tmpData(11) = ObjSysInfo.EmpId
        tmpData(12) = mvarWardID
        tmpData(13) = mvarHosilID
        tmpData(14) = ""
        tmpData(15) = ""
        If P_ApplyBuildingInfo Then
            tmpData(16) = ObjSysInfo.BuildingCd
        Else
            tmpData(16) = "10"
        End If
        Call .SetColData(tmpData)
        
        If chkChangeColTm.Value = 1 Then
            .ColDt = Format(dtpColDtTm.Value, CS_DateDbFormat)
            .ColTm = Format(dtpColDtTm.Value, "HHMMSS")
        Else
            .ColDt = Format(GetSystemDate, CS_DateDbFormat)
            .ColTm = Format(GetSystemDate, "HHMMSS")
        End If
    End With
    
    objLISCollect.SetTrans = True
    ColSuccess = objLISCollect.DoCollection(objProgress)
    If Not ColSuccess Then
        Set objProgress = Nothing
        lblErrString.Caption = objLISCollect.ErrString
        MsgBox lblErrString.Caption
        MouseDefault
        CollectForLIS_New = False
        Exit Function
    End If
    CollectForLIS_New = True
End Function

Private Function Get_PtSpcAdd() As Boolean
    Dim objSQL      As New clsLIS_SQL
    Dim DrRS        As New Recordset
    Dim strErChk    As String
    Dim strPtID     As String
    Dim strColdt    As String
    Dim strcoltm    As String
    Dim cnt         As Long

    Get_PtSpcAdd = True
        
    DrRS.Open objSQL.Get_PtSpcAdd(Trim(txtPtId.Text)), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not DrRS.EOF Then
        With tblOrdSheet
            Do Until DrRS.EOF
                
                If .DataRowCnt <= .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = 1
                .ForeColor = vbBlue
                .Col = 2: .Value = Format(Trim(DrRS.Fields("reqdt").Value & ""), CS_DateShortMask)
                .Col = 3: .Value = Trim("" & DrRS.Fields("seq").Value & "")
                .Col = 4: .Value = "추가요청"
                .Col = 5: .Value = "혈액"
                .Col = 6: .Value = lblDoctNm.Caption
                .Col = 31: .Value = "" & DrRS.Fields("wardid").Value
                .Col = 32: .Value = "" & DrRS.Fields("hosilid").Value
                .Col = 34: .Value = "" & DrRS.Fields("orddiv").Value
                cnt = cnt + 1
                
                DrRS.MoveNext
            Loop
            .Row = 1
            .Row2 = .DataRowCnt
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Lock = False
            .Protect = True
            .BlockMode = False
        End With
    Else
        Get_PtSpcAdd = False
    End If

    If cnt = 0 Then Get_PtSpcAdd = False
    
    DrRS.Close
    Set DrRS = Nothing
    Set objSQL = Nothing
End Function

Private Function CollectForBBS_NEW(ByVal FRowCnt As Long, ByVal LRowCnt As Long, _
                                   ByVal ColDt As String, ByVal ColTm As String, _
                                   ByRef objProgress As Object) As Boolean

    
    Dim dicBBS      As clsDictionary
    Dim objCollect  As clsBBS_Collection
    Dim objBAR      As clsDictionary
    
    Dim tmpClipData As String
    
    Dim tmpTotData  As Variant
    Dim tmpRowData  As Variant
    
    Dim strColdt    As String
    Dim strcoltm    As String
    
    Dim i As Long
    Dim lngColCnt   As Long
    Dim strStatFg   As String
    
    lngColCnt = 0
    mvarHosilID = medGetP(lblLocation.Caption, 2, "-")
    
    With tblOrdSheet
        For i = FRowCnt To LRowCnt
            .Row = i
            .Col = 1
            If .Value = 1 Then GoTo BBS_Save
        Next i
    End With
    CollectForBBS_NEW = True
    Exit Function
    
BBS_Save:
    Set objCollect = New clsBBS_Collection
    
    If objCollect.Blood_Existence(txtPtId.Text, Format(GetSystemDate, "yyyymmdd"), Format(GetSystemDate, "hhmmss")) = False Then
        If objCollect.SetAccessCheck(txtPtId.Text) = True Then
           '검체가 이미 존재하는 경우
           CollectForBBS_NEW = objCollect.SetWardAccess(txtPtId.Text, enBussDiv.BussDiv_InPatient, Format(GetSystemDate, "yyyymmdd"), _
                                    Format(GetSystemDate, "hhmmss"), ObjSysInfo.EmpId)
                
            Set objCollect = Nothing
            Exit Function
        Else
            GoTo BBSCollect
        End If
    End If
BBSCollect:
    Set dicBBS = New clsDictionary
    Set objBAR = New clsDictionary
    
    With tblOrdSheet
        .Row = FRowCnt: .Col = enCOLLIST.tcWardID: mvarWardID = .Value
                        .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
                        .Col = enCOLLIST.tcDeptCD: mvarDeptCd = .Value

        strColdt = ColDt
        strcoltm = ColTm
        lngColCnt = 1 'lngColCnt + 1
        
        dicBBS.Clear
        dicBBS.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"
        dicBBS.AddNew txtPtId.Text, Join(Array(lblPtNm.Caption, strColdt, strcoltm, _
                      ObjSysInfo.EmpId, enBussDiv.BussDiv_InPatient, strBlgCd, mvarHosilID, strStatFg), COL_DIV)

'Skip:
'       Next
    
    End With
    
    If lngColCnt = 0 Then
        CollectForBBS_NEW = True
        Set objCollect = Nothing
        Set objBAR = Nothing
        Set dicBBS = Nothing
        Exit Function
    End If
    
    CollectForBBS_NEW = objCollect.SET_Collect(dicBBS, , objProgress)
    
    If CollectForBBS_NEW Then
        Set objBAR = objCollect.BldDic
        If objBAR.RecordCount > 0 Then
            BarCodePrintForBBS objBAR
        Else
            Set objProgress = Nothing
            MsgBox "검체가 이미 존재하므로 바코드가 출력되지 않습니다.", vbInformation + vbOKOnly, "바코드출력"
        End If
        If objCollect.Spc72Chk Then
            MsgBox "해당 환자는 72시간내에 채혈한 검체가 존재합니다.", vbInformation + vbOKOnly, "바코드출력"
        End If
    End If
    
    Set objCollect = Nothing
    Set objBAR = Nothing
    Set dicBBS = Nothing
End Function

Private Sub BarCodePrintForBBS(objDic As clsDictionary)
    Dim strPtID     As String
    Dim strPtNm     As String
    Dim strColdt    As String
    Dim strcoltm    As String
    Dim strSpcNo    As String
    Dim strW_Dept   As String
    Dim strBuildNm  As String
    Dim strAccSeq   As String
    Dim strHosilId  As String
    Dim strStatFg   As String
    Dim strColId As String
    Dim objPt As clsPatient
    Dim strSexAge As String
    
    strW_Dept = mvarWardID
    
    If strW_Dept = "" Then
        strW_Dept = mvarDeptCd
    End If
    
    If lblLocation.Caption <> "" Then
        If lblLocation.Caption <> "--" Then strW_Dept = strW_Dept & "/" & mvarHosilID
    End If
    
    If P_ApplyBuildingInfo Then
        strBuildNm = ObjSysInfo.BuildingNm
    Else
        strBuildNm = BBSName
    End If
    
    objDic.MoveFirst
    Do Until objDic.EOF
        strPtID = medGetP(objDic.GetString, 1, COL_DIV)
        strPtNm = medGetP(objDic.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDic.GetString, 3, COL_DIV)
        strColdt = Mid(medGetP(objDic.GetString, 4, COL_DIV), 5)
        strcoltm = Mid(medGetP(objDic.GetString, 5, COL_DIV), 1, 4)
        strStatFg = medGetP(objDic.GetString, 7, COL_DIV)
        strColdt = Format(strColdt, "00/00")
        strcoltm = Format(strcoltm, "0#:##")
        Call GetColInfo("B", strSpcNo, "", "", strColId)
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '
        Set objPt = New clsPatient
        objPt.GETPatient (strPtID)
        strSexAge = objPt.SEXAGE
        Set objPt = Nothing
        
        PrintOutBarcode "혈액", "", strColId, "", strSpcNo, strPtID, _
                     strPtNm, strSexAge, "", strStatFg, strW_Dept, strColdt, strcoltm, _
                    "", 1
'        BarCodePrint "XM", strBuildNm, "", strAccSeq, strSpcNo, strPtID, _
                                            strPtNm, "", "", strStatFg, strW_Dept, strColdt, strcoltm, _
                                            "", 1
        objDic.MoveNext
    Loop
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim i           As Long
    Dim ButtonValue As Variant
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String

    If SelAllFg Then Exit Sub
    
    With tblOrdSheet
       .Row = Row
       .Col = Col:   ButtonValue = .Value
       
       If .Value = 0 Then Exit Sub
       
       .Col = 9:      SvOrdDt = .Value
       .Col = 10:     SvOrdNo = .Value
       
       For i = 1 To .MaxRows
          If i <> Row Then
             .Row = i
             .Col = 9
             If .Value = SvOrdDt Then
                .Col = 10
                If .Value = SvOrdNo Then
                   .Col = 1
                   If .Value <> ButtonValue Then .Value = ButtonValue
                End If
             End If
          End If
       Next
    End With
End Sub

Private Function CollectionTargetChk() As Boolean
    Dim ii As Long
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = enCOLLIST.tcCHECK
            If .Value = 1 Then
                CollectionTargetChk = True
                Exit For
            End If
        Next
    End With
End Function

Private Sub chkSelAll_Click()
    SelAllFg = True
    With tblOrdSheet
        .Col = 1: .Col2 = 1
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .Value = chkSelAll.Value
        .BlockMode = False
    End With
    SelAllFg = False
End Sub

Private Sub PrintIntionlize()
    PrtLeft = 0
    LineSpace = 6
    
    lngNo = PrtLeft
    lngWorkNo = 10
    lngPtNm = 40
    lngPtid = 60
    lngSex = 80
    lngHos = 92
    lngColdt = 104
    lngTest = 129
    lngSpc = 180
    
    Printer.Font = "굴림체"
    Printer.FontSize = 9
    Printer.Orientation = vbPRORPortrait
    Printer.ScaleMode = vbMillimeters

    Twidth = Printer.ScaleWidth

    LastLinetop = Printer.ScaleHeight
End Sub

Private Sub PrintHeader(ByVal vData As Long)
    Dim strBase As String
    Dim spcCnt  As Long
    
    spcCnt = vData
    lngCurtop = 5
    
    Printer.FontBold = True
    Printer.FontSize = 16
    Printer.FontUnderline = True
    Call Print_Setting("병동 채혈리스트", PrtLeft, LineSpace, Twidth, "C", "C", False)
    Printer.FontBold = False: Printer.FontSize = 9: Printer.FontUnderline = False
    
    lngCurtop = lngCurtop + 20
    
    strBase = "채혈장소: " & txtWardID.Text & _
              "  작업일시 : " & Format(GetSystemDate, "yyyy-mm-dd") & "  " & _
                                Format(GetSystemDate, "HH:MM") & _
              "     채혈자 : " & Trim(ObjSysInfo.EmpNm)
    
    Call Print_Setting(strBase, PrtLeft, LineSpace, Twidth, , "C", False)
    Call Print_Setting("검체수 : " & spcCnt, lngSpc, LineSpace, Twidth, "L", "C")
   
    Printer.Line (PrtLeft, lngCurtop)-(Twidth - PrtLeft, lngCurtop)
    
    Call Print_Setting("NO", lngNo, LineSpace, , "L", "C", False)
    Call Print_Setting("WorkNo", lngWorkNo, LineSpace, , "L", "C", False)
    Call Print_Setting("환자명", lngPtNm, LineSpace, , "L", "C", False)
    Call Print_Setting("환자ID", lngPtid, LineSpace, , "L", "C", False)
    Call Print_Setting("S/A", lngSex, LineSpace, , "L", "C", False)
    Call Print_Setting("호실", lngHos, LineSpace, , "L", "C", False)
    Call Print_Setting("채혈일시", lngColdt, LineSpace, , "L", "C", False)
    Call Print_Setting("검사종목", lngTest, LineSpace, , "L", "C", False)
    Call Print_Setting("검체", lngSpc, LineSpace, , "L", "C")
    
    Printer.Line (PrtLeft, lngCurtop)-(Twidth - PrtLeft, lngCurtop)
    Printer.Line (PrtLeft, LastLinetop)-(Twidth - PrtLeft, LastLinetop)
End Sub

Private Sub VBPrint()
    Dim strNo      As String
    Dim strWorkNo  As String
    Dim strPtNm    As String
    Dim strPtID    As String
    Dim strSa      As String
    Dim strHosilId As String
    Dim strColdt   As String
    Dim strTest    As String
    Dim strSpcnm   As String
    
    Dim ii         As Long
    Dim spcCnt     As Long
    
    With tblCollect
        spcCnt = .DataRowCnt
        If spcCnt < 1 Then Exit Sub
        Call PrintIntionlize
        Call PrintHeader(spcCnt)
        Me.MousePointer = 11
            
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 2:      strNo = .Value
            .Col = 3:  strWorkNo = .Value
            .Col = 4:    strPtNm = .Value
            .Col = 5:    strPtID = .Value
            .Col = 6:      strSa = .Value
            .Col = 7: strHosilId = .Value
            .Col = 8:   strColdt = .Value
            .Col = 9:    strTest = .Value
            .Col = 10:   strSpcnm = .Value
            
            If lngCurtop > Printer.ScaleHeight - LineSpace Then
                Printer.NewPage
                Call PrintHeader(spcCnt)
            End If
            
            Call Print_Setting(strNo, lngNo, LineSpace, , , "C", False)
            Call Print_Setting(strWorkNo, lngWorkNo, LineSpace, , , "C", False)
            Call Print_Setting(strPtNm, lngPtNm, LineSpace, , , "C", False)
            Call Print_Setting(strPtID, lngPtid, LineSpace, , , "C", False)
            Call Print_Setting(strSa, lngSex, LineSpace, , , "C", False)
            Call Print_Setting(strHosilId, lngHos, LineSpace, , , "C", False)
            Call Print_Setting(strColdt, lngColdt, LineSpace, , , "C", False)
            If Len(strTest) < 26 Then
                Call Print_Setting(strTest, lngTest, LineSpace, , , "C", False)
                Call Print_Setting(strSpcnm, lngSpc, LineSpace, , , "C")
            Else
                Call Print_Setting(Mid(strTest, 1, 25), lngTest, LineSpace, , , "C", False)
                Call Print_Setting(strSpcnm, lngSpc, LineSpace, , , "C")
                Call Print_Setting(Mid(strTest, 26), lngTest, LineSpace, , , "C")
            End If

        Next
    End With
    Printer.EndDoc
    Me.MousePointer = 0
End Sub
