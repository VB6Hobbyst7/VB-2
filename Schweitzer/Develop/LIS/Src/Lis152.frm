VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm152WardAccession 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Accession - Ward Collection"
   ClientHeight    =   9450
   ClientLeft      =   -315
   ClientTop       =   705
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis152.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   11400
   WindowState     =   2  '최대화
   Begin VB.OptionButton optDiv 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "아침채혈대상"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   2445
      TabIndex        =   49
      Top             =   45
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.OptionButton optDiv 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "미채혈내역"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   5235
      TabIndex        =   5
      Top             =   45
      Width           =   1410
   End
   Begin VB.OptionButton optDiv 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "검체채취내역"
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   3840
      TabIndex        =   4
      Top             =   45
      Width           =   1410
   End
   Begin VB.CommandButton cmeExit 
      BackColor       =   &H00F4F0F2&
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
      Height          =   495
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CheckBox chkALL 
      BackColor       =   &H00800000&
      Caption         =   "전체"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1635
      TabIndex        =   2
      Top             =   60
      Width           =   690
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "병동정보"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblQuery 
      Height          =   300
      Left            =   6675
      TabIndex        =   1
      Top             =   45
      Width           =   7755
      _ExtentX        =   13679
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
      Caption         =   "검체 채취내역"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblWardList 
      Height          =   8040
      Left            =   75
      TabIndex        =   50
      Top             =   375
      Width           =   2340
      _Version        =   196608
      _ExtentX        =   4128
      _ExtentY        =   14182
      _StockProps     =   64
      AllowUserFormulas=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   3
      MaxRows         =   33
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "Lis152.frx":000C
      VisibleCols     =   3
      VisibleRows     =   10
   End
   Begin VB.PictureBox picQuery 
      BackColor       =   &H00DBE6E6&
      Height          =   8055
      Index           =   0
      Left            =   2445
      ScaleHeight     =   7995
      ScaleWidth      =   11955
      TabIndex        =   44
      Top             =   390
      Width           =   12015
      Begin VB.TextBox txtCopy 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   285
         Left            =   10890
         TabIndex        =   64
         Text            =   "1"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkPrint 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈리스트 출력"
         Height          =   390
         Left            =   9240
         TabIndex        =   63
         Top             =   105
         Value           =   1  '확인
         Width           =   2100
      End
      Begin VB.OptionButton optSc 
         BackColor       =   &H00D1D8D3&
         Caption         =   "미적용"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2730
         TabIndex        =   61
         Top             =   480
         Width           =   1170
      End
      Begin VB.OptionButton optSc 
         BackColor       =   &H00D1D8D3&
         Caption         =   "적용"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1605
         TabIndex        =   60
         Top             =   480
         Width           =   1170
      End
      Begin VB.ComboBox cboSaveTime 
         Height          =   345
         Left            =   5565
         TabIndex        =   52
         Text            =   "Combo1"
         Top             =   90
         Width           =   2490
      End
      Begin VB.CommandButton cmdCollection 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈처리"
         Height          =   510
         Left            =   10545
         Style           =   1  '그래픽
         TabIndex        =   46
         Top             =   585
         Width           =   1320
      End
      Begin VB.CommandButton cmdReqQuery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회"
         Height          =   510
         Left            =   9240
         Style           =   1  '그래픽
         TabIndex        =   45
         Top             =   585
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpReqdt 
         Height          =   315
         Left            =   1545
         TabIndex        =   47
         Top             =   90
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20774912
         CurrentDate     =   37935
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   45
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "희망채혈일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   4080
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "아침채혈스케줄"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   4080
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   435
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "채혈자수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   4080
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "업무분담구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColCnt 
         Height          =   315
         Left            =   5565
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   435
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
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
      Begin MedControls1.LisLabel lblBuss 
         Height          =   315
         Left            =   5565
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   780
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   45
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   435
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "스케줄적용여부"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   315
         Left            =   1545
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   435
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
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
      Begin FPSpread.vaSpread tblPtList 
         Height          =   6750
         Left            =   60
         TabIndex        =   62
         Top             =   1170
         Width           =   11820
         _Version        =   196608
         _ExtentX        =   20849
         _ExtentY        =   11906
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
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
         MaxCols         =   18
         MaxRows         =   28
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis152.frx":0728
         TextTip         =   4
         ScrollBarTrack  =   3
      End
      Begin MedControls1.LisLabel lblPage 
         Height          =   255
         Left            =   11370
         TabIndex        =   65
         Top             =   180
         Visible         =   0   'False
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
      Begin VB.Label lblCnt 
         BackColor       =   &H00DBE6E6&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   8085
         TabIndex        =   53
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox picQuery 
      BackColor       =   &H00DBE6E6&
      Height          =   8055
      Index           =   2
      Left            =   2445
      ScaleHeight     =   7995
      ScaleWidth      =   11955
      TabIndex        =   32
      Top             =   390
      Width           =   12015
      Begin VB.CommandButton cmdQuery1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회"
         Height          =   510
         Left            =   9330
         Style           =   1  '그래픽
         TabIndex        =   38
         Top             =   0
         Width           =   1320
      End
      Begin VB.CommandButton cmdCollectSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈처리"
         Height          =   510
         Left            =   10620
         Style           =   1  '그래픽
         TabIndex        =   37
         Top             =   0
         Width           =   1320
      End
      Begin FPSpread.vaSpread tblNoColList 
         Height          =   7485
         Left            =   -135
         TabIndex        =   33
         Tag             =   "10114"
         Top             =   525
         Width           =   12090
         _Version        =   196608
         _ExtentX        =   21325
         _ExtentY        =   13203
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
         MaxCols         =   25
         MaxRows         =   29
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis152.frx":101B
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   29
         TextTip         =   2
      End
      Begin MSComCtl2.DTPicker dtpFColDt 
         Height          =   315
         Left            =   1065
         TabIndex        =   34
         Top             =   105
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20774912
         CurrentDate     =   37935
      End
      Begin MSComCtl2.DTPicker dtpTColdt 
         Height          =   315
         Left            =   3840
         TabIndex        =   35
         Top             =   105
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20774912
         CurrentDate     =   37935
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   300
         Left            =   3555
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   105
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   529
         BackColor       =   14411494
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
         Caption         =   "~"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   39
         Top             =   105
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
         Caption         =   "등록일자"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picQuery 
      BackColor       =   &H00DBE6E6&
      Height          =   8055
      Index           =   1
      Left            =   2445
      ScaleHeight     =   7995
      ScaleWidth      =   11955
      TabIndex        =   6
      Top             =   390
      Width           =   12015
      Begin VB.CommandButton cmdExecute 
         BackColor       =   &H00F4F0F2&
         Caption         =   "일괄접수실행(&S)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10455
         Style           =   1  '그래픽
         TabIndex        =   43
         Top             =   15
         Width           =   1470
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회(&Q)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9135
         Style           =   1  '그래픽
         TabIndex        =   42
         Top             =   15
         Width           =   1320
      End
      Begin DRcontrol1.DrFrame fraColSave 
         Height          =   2880
         Left            =   2970
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1965
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5080
         Title           =   "= 미채혈 사유를 입력하세요(등록시 채혈취소작업이 진행됩니다.) ="
         BackColor       =   16776439
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton cmdPopup 
            Caption         =   "...."
            Height          =   330
            Left            =   2505
            TabIndex        =   11
            Top             =   1620
            Width           =   375
         End
         Begin VB.CheckBox chkCancel 
            BackColor       =   &H00FFFCF7&
            Caption         =   "희망채혈일시변경"
            Height          =   285
            Left            =   4200
            TabIndex        =   10
            Top             =   495
            Width           =   1845
         End
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00FFFCF7&
            Caption         =   "확인"
            Height          =   420
            Left            =   4110
            Style           =   1  '그래픽
            TabIndex        =   9
            Top             =   2355
            Width           =   960
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00FFFCF7&
            Caption         =   "닫기"
            Height          =   420
            Left            =   5100
            Style           =   1  '그래픽
            TabIndex        =   8
            Top             =   2355
            Width           =   960
         End
         Begin MedControls1.LisLabel LisLabel11 
            Height          =   360
            Left            =   165
            TabIndex        =   12
            Top             =   480
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
            Caption         =   "접수 번호"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblWorkarea 
            Height          =   285
            Left            =   1530
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   510
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblAccdt 
            Height          =   285
            Left            =   2235
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   510
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   503
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "20031102"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblAccSeq 
            Height          =   285
            Left            =   3510
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   510
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "1234"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblReasonCd 
            Height          =   315
            Left            =   1470
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1635
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblReasonNm 
            Height          =   315
            Left            =   2895
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1635
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   556
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   0
            Left            =   165
            TabIndex        =   18
            Top             =   858
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
         Begin MedControls1.LisLabel LisLabel3 
            Height          =   360
            Left            =   165
            TabIndex        =   19
            Top             =   1605
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
            Caption         =   "미채혈사유"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblPtid 
            Height          =   315
            Left            =   1470
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   885
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblPtnm 
            Height          =   315
            Left            =   2505
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   885
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel8 
            Height          =   360
            Left            =   165
            TabIndex        =   22
            Top             =   1236
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
            Height          =   315
            Left            =   1470
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1260
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel10 
            Height          =   360
            Left            =   165
            TabIndex        =   24
            Top             =   1980
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
            Caption         =   "변경일시"
            Appearance      =   0
         End
         Begin MSComCtl2.DTPicker dtpReqdate 
            Height          =   360
            Left            =   1470
            TabIndex        =   25
            Top             =   1995
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   635
            _Version        =   393216
            CustomFormat    =   "yyy-MM-dd HH:mm:ss"
            Format          =   20774915
            CurrentDate     =   36328
         End
         Begin MedControls1.LisLabel lblChangeReqdate 
            Height          =   315
            Left            =   3555
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   2010
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   556
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblSpcnm 
            Height          =   315
            Left            =   4470
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   885
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            BackColor       =   15857140
            ForeColor       =   0
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
            Caption         =   "01"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblSave 
            Height          =   360
            Left            =   165
            TabIndex        =   28
            Top             =   2400
            Visible         =   0   'False
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   635
            BackColor       =   15857140
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
            Alignment       =   1
            Caption         =   "미채혈등록처리되었습니다."
            Appearance      =   0
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2010
            TabIndex        =   30
            Top             =   555
            Width           =   195
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3300
            TabIndex        =   29
            Top             =   555
            Width           =   195
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00F1F5F4&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            Height          =   360
            Left            =   1470
            Shape           =   4  '둥근 사각형
            Top             =   480
            Width           =   2595
         End
      End
      Begin FPSpread.vaSpread tblColList 
         Height          =   8040
         Left            =   -30
         TabIndex        =   31
         Tag             =   "10114"
         Top             =   525
         Width           =   11985
         _Version        =   196608
         _ExtentX        =   21140
         _ExtentY        =   14182
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
         GrayAreaBackColor=   16777215
         MaxCols         =   25
         MaxRows         =   29
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis152.frx":1C32
         StartingColNumber=   2
         UserResize      =   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   29
      End
      Begin MSComCtl2.DTPicker dtpColdt 
         Height          =   330
         Left            =   1050
         TabIndex        =   40
         Top             =   120
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   582
         _Version        =   393216
         Format          =   20774912
         CurrentDate     =   37935
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   41
         Top             =   120
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
         Caption         =   "채혈일자"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frm152WardAccession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents objMyList    As clspopuplist
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private Const TAG_RSN& = 1
Private Const TAG_EMP& = 2
'Private WithEvents mnuPopup     As menu
'Private WithEvents mnuReg       As menu
Private Const MENU_REG& = 1
Private Const MENU_COL& = 2
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1

Private sWorkDt                 As String
Private sWorkTm                 As String
Private lngSelRow               As Long
Private blnCol                  As Boolean
Private blnNoCol                As Boolean

Private Enum TblCol
    tcChk = 1
    tcWARDID
    tcPTID
    tcPTNM
    tcLABNO
    tcTESTNM
    tcSPCNM
    tcCOLDATE
    tcColNm
    tcDonefg
    tcReg
    tcAccFg
    tcReqdt
    tcReqtm
    tcColdt
    tcColtm
    tcColID
End Enum

Private Enum TblCol1
    enDate = 1
    enWardID
    enPtid
    enPtNm
    enLabno
    enTestNm
    enSpcNm
    enReason
    enSaveNm
    enColDate
    enDone
    enColID
    enReasonNm
End Enum

Private Enum tblCol2
    enChk = 1
    enLocation
    enPtid
    enPtNm
    enTest
    
    enSpccd
    enDOB
    enBedInDt
    enDept
    enOrdDoct
    
    enMajDoct
    enWardID
    enHosil
    enRoom
    enRndFg
    
    enSEX
    enColID
    enColNm
End Enum
Private objSC   As clsDictionary

Private Sub FormInitialize()
    dtpFColDt.Value = GetSystemDate
    dtpTColdt.Value = GetSystemDate
    dtpColdt.Value = GetSystemDate
    dtpReqdt.Value = GetSystemDate
    dtpReqdate.Value = GetSystemDate
    
    Call medClearTable(tblPtList)
    Call medClearTable(tblColList)
    Call medClearTable(tblNoColList)
    
    lblColCnt.Caption = "": lblCnt.Caption = "": lblBuss.Caption = ""

    
    lblPtId.Caption = "":           lblPtNm.Caption = "":       lblReqdt.Caption = ""
    lblChangeReqdate.Caption = "":  lblReasonCd.Caption = "":   lblReasonNm.Caption = ""
    lblWorkarea.Caption = "":       lblAccdt.Caption = "":      lblAccSeq.Caption = "":     lblSpcnm.Caption = ""
    
    chkCancel.Value = 0
    optDiv(0).Value = True
    optSc(0).Value = True
    fraColSave.Visible = False
    If dtpReqdt.Enabled Then dtpReqdt.SetFocus
    
End Sub
Private Sub SaveClear()
    dtpReqdate.Value = GetSystemDate
    lblPtId.Caption = "": lblPtNm.Caption = "": lblReqdt.Caption = ""
    lblChangeReqdate.Caption = "": lblReasonCd.Caption = "": lblReasonNm.Caption = ""
    lblWorkarea.Caption = "": lblAccdt.Caption = "": lblAccSeq.Caption = ""
    chkCancel.Value = 0: lblSpcnm.Caption = "": lblSave.Visible = False
    dtpReqdate.Enabled = False
End Sub





Private Sub chkAll_Click()
    Dim ii  As Integer
    
    With tblWardList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1
            .Value = chkALL.Value
        Next
    End With
End Sub

Private Sub chkCancel_Click()
    lblChangeReqdate.Caption = ""
    If chkCancel.Value = 1 Then
        dtpReqdate.Enabled = True
        lblChangeReqdate.Caption = Format(dtpReqdate.Value, "YYYY-MM-DD   HH:MM:SS")
    Else
        dtpReqdate.Enabled = False
    End If
End Sub

Private Sub cmdClose_Click()
    fraColSave.Visible = False
End Sub

Private Sub cmdOk_Click()
    Dim sTmp As String
    Dim blnSave As Boolean
    
    If lblReasonCd.Caption = "" Then
        MsgBox "미채혈 사유가 등록되지 않았습니다.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    Me.MousePointer = 11
    If chkCancel.Value = 1 Then
        If lblChangeReqdate.Caption = "" Then
            MsgBox "채혈일시변경이 되지 않았습니다.", vbInformation + vbOKOnly, "Info"
            Me.MousePointer = 0
            Exit Sub
        End If
        
        sTmp = MsgBox("채혈일시를 변경할경우 채혈작업이 취소되며," & vbCRLF & " 처방상태로 되돌아갑니다." & vbCRLF & _
                    "변경하시겠습니까?", vbInformation + vbYesNo, "Info")
        If sTmp = vbYes Then
            blnSave = ChangeReqdtTm
        End If
    Else
        sTmp = MsgBox("미채혈등록을 할경우 접수번호는 유효합니다." & vbCRLF & _
                    "등록하시겠습니까?", vbInformation + vbYesNo, "Info")
        If sTmp = vbYes Then
            blnSave = NoCollect
        End If
    End If
    
    Me.MousePointer = 0
    If blnSave = False Then Exit Sub
    
    With tblColList
        .Row = lngSelRow: .Row2 = lngSelRow
        .Col = 1: .COL2 = .MaxRows
        .BlockMode = True
        .Action = ActionDeleteRow
        .BlockMode = False
    End With
    
    
End Sub
Private Function ChangeReqdtTm() As Boolean
    
    Dim Rs          As Recordset
    Dim ObjDic      As clsDictionary
    Dim sPtid       As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim SSQL        As String
    Dim sReqdt      As String
    Dim sReqtm      As String
    Dim ii          As Integer
    
    
    
    sPtid = lblPtId.Caption
    sWorkArea = Trim(lblWorkarea.Caption)
    sAccDt = Trim(lblAccdt.Caption)
    sAccSeq = Trim(lblAccSeq.Caption)
    sReqdt = Format(dtpReqdate.Value, "YYYYMMDD")
    sReqtm = Format(dtpReqdate.Value, "HHMMSS")
    
    Set Rs = New Recordset
    Set ObjDic = New clsDictionary
    
    ObjDic.Clear
    ObjDic.FieldInialize "seq", "ptid,orddt,ordno"
    SSQL = " SELECT b.ptid,b.orddt,b.ordno FROM " & T_LAB101 & " b," & T_LAB102 & " a " & _
           " WHERE " & _
                     DBW("a.workarea=", sWorkArea) & _
           " AND " & DBW("a.accdt=", sAccDt) & _
           " AND " & DBW("a.accseq=", sAccSeq) & _
           " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno"
         
    Rs.Open SSQL, DBConn
    If Not Rs.EOF Then
        Do Until Rs.EOF
            ii = ii + 1
            ObjDic.AddNew ii, Rs.Fields("ptid").Value & "" & COL_DIV & _
                             Rs.Fields("orddt").Value & "" & COL_DIV & _
                             Rs.Fields("ordno").Value & ""
            Rs.MoveNext
        Loop
        ObjDic.MoveFirst
    End If
    
    Set Rs = Nothing
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    If ObjDic.RecordCount > 0 Then
        Do Until ObjDic.EOF
            SSQL = " update " & T_LAB101 & _
                 " set donefg='0' ,reqdt='" & sReqdt & "',reqtm='" & sReqtm & "'" & _
                 " WHERE " & DBW("ptid=", ObjDic.Fields("ptid")) & _
                 " AND " & DBW("orddt=", ObjDic.Fields("orddt")) & _
                 " AND " & DBW("ordno=", ObjDic.Fields("ordno"))
            DBConn.Execute SSQL
            ObjDic.MoveNext
        Loop
    End If
    
    
    SSQL = " update " & T_LAB102 & _
                " set   stscd ='0', donefg='0', rcvdt='',   rcvtm='', " & _
                "       examdt='',  examtm='',  examdoct=0, workarea='', accdt='', accseq=0 " & _
                " WHERE " & DBW("workarea = ", sWorkArea) & " AND " & DBW("accdt=", sAccDt) & " AND " & DBW("accseq=", sAccSeq)

    DBConn.Execute SSQL
    SSQL = " update " & T_LAB201 & _
           " set    " & DBW("stscd ", enStsCd.StsCd_LIS_Cancel, 3) & _
                        DBW("reqinputcnt ", 0, 3) & _
                    " vfydt = '', vfytm = '', vfyid = 0, " & _
                    " footnotefg='0', rmkcd='' " & _
           " WHERE  " & DBW("workarea = ", sWorkArea) & _
           " AND    " & DBW("accdt    = ", sAccDt) & _
           " AND    " & DBW("accseq   = ", sAccSeq)
    DBConn.Execute SSQL
    
    SSQL = " insert into " & T_LAB304 & " (workarea,accdt,accseq,seq,vfyid,rsttxt) " & _
           " values (" & DBV("workarea", sWorkArea, 1) & _
                         DBV("accdt", sAccDt, 1) & _
                         DBV("accseq", sAccSeq, 1) & _
                         DBV("seq", "1", 1) & _
                         DBV("vfyid", ObjSysInfo.EmpId, 1) & _
                         DBV("rsttxt", lblReasonNm.Caption) & _
                    ")"

    DBConn.Execute SSQL
    DBConn.CommitTrans
    Set ObjDic = Nothing
    ChangeReqdtTm = True
    fraColSave.Visible = False
    Exit Function
    
SAVE_ERROR:
    DBConn.RollbackTrans
    Set ObjDic = Nothing
End Function
Private Function NoCollect() As Boolean
'미채혈목록을 저장
'lab206테이블에 저장
'희망채혈일시가 변경된건 제외한다.

    Dim SSQL    As String
    Dim sWardID As String
    Dim sColDt  As String
    Dim sColTm  As String
    Dim sCancelDt As String
    Dim sCancelTm   As String
    Dim sColID      As String
    
    
    With tblColList
        .Row = lngSelRow
        .Col = TblCol.tcWARDID: sWardID = Trim(.Value)
        .Col = TblCol.tcColdt:  sColDt = Trim(.Value)
        .Col = TblCol.tcColtm:  sColTm = Trim(.Value)
        .Col = TblCol.tcReqdt:  sCancelDt = Trim(.Value)
        .Col = TblCol.tcReqtm:  sCancelTm = Trim(.Value)
        .Col = TblCol.tcColID:  sColID = Trim(.Value)
    End With
    
    On Error GoTo SAVE_ERROR
    
    DBConn.BeginTrans
    
    SSQL = " insert into " & T_LAB206 & _
           " (workarea,accdt,accseq,ptid,wardid,coldt,coltm,colid,vfydt,vfytm,vfyid," & _
           " reqdt,reqtm,canceldt,canceltm,rmkcd,rsttxt)" & _
           " values (" & _
            DBV("workarea", lblWorkarea.Caption, 1) & DBV("accdt", lblAccdt.Caption, 1) & _
            DBV("accseq", lblAccSeq.Caption, 1) & DBV("ptid", lblPtId.Caption, 1) & _
            DBV("wardid", sWardID, 1) & DBV("coldt", sColDt, 1) & _
            DBV("coltm", sColTm, 1) & DBV("colid", sColID, 1) & _
            DBV("vfydt", Format(GetSystemDate, "YYYYMMdd"), 1) & _
            DBV("vfytm", Format(GetSystemDate, "HHMMSS"), 1) & _
            DBV("vfyid", ObjSysInfo.EmpId, 1) & _
            DBV("reqdt", Format(dtpReqdate.Value, "YYYYMMDD"), 1) & _
            DBV("reqtm", Format(dtpReqdate.Value, "HHMMSS"), 1) & _
            DBV("canceldt", sCancelDt, 1) & DBV("canceltm", sCancelTm, 1) & _
            DBV("rmkcd", lblReasonCd.Caption, 1) & DBV("rsttxt", lblReasonNm.Caption, 0) & _
            " ) "
    
    DBConn.Execute SSQL
    DBConn.CommitTrans

    fraColSave.Visible = False
    NoCollect = True
    Exit Function
    
SAVE_ERROR:
    DBConn.RollbackTrans

End Function


Private Sub cmdPopup_Click()
    Dim objSQL As clsLISSqlAccession
    
    Set objSQL = New clsLISSqlAccession
    
'    Set objMyList = New clspopuplist
    Set objMyList = New clsPopUpList
    
    With objMyList
        .FormCaption = "미채혈 사유찾기"
        .ColumnHeaderText = "코드;코드명"
        .Tag = TAG_RSN
        .LoadPopUp objSQL.SQLGetCancelReason
        
'        .Caption = "미채혈사유찾기"
'        .HeadName = "코드,코드명"
'        .Tag = "reason"
'        Call .ListPop(objSQL.SQLGetCancelReason, 10000, 10000)
        
    End With
    Set objSQL = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmeExit_Click()
    Set objSC = Nothing
    Unload Me
End Sub

Private Sub dtpReqdate_LostFocus()
    lblChangeReqdate.Caption = Format(dtpReqdate.Value, "YYYY-MM-DD   HH:MM:SS")
End Sub

Private Sub Form_Activate()
    Call FormInitialize
End Sub

Private Sub Form_Load()
'    Dim objData As clsBasisData
    Dim Rs As Recordset
    
'    Set objData = New clsBasisData
    Set Rs = New Recordset
    
    Rs.Open GetSQLWardList, DBConn
    
    With Rs 'objLisComCode.WardId
        .MoveFirst
        Do Until .EOF
            If tblWardList.DataRowCnt + 1 > tblWardList.MaxRows Then
                tblWardList.MaxRows = tblWardList.MaxRows + 1
            End If
            tblWardList.Row = tblWardList.DataRowCnt + 1
            tblWardList.Col = 1
            tblWardList.CellType = CellTypeCheckBox: tblWardList.TypeHAlign = TypeHAlignCenter
            tblWardList.Value = 0
            tblWardList.Col = 2: tblWardList.Value = .Fields("wardnm").Value & ""
            tblWardList.Col = 3: tblWardList.Value = .Fields("wardid").Value & ""
            .MoveNext
        Loop
    End With
    
    Set Rs = Nothing
'    Set objData = Nothing
End Sub
Private Sub cmdQuery_Click()
    blnCol = False
    If optDiv(1).Value Then Call CollectListQuery
End Sub

Private Sub CollectListQuery()
    Dim objPro  As clsProgress
    Dim Rs      As Recordset
    Dim sWorkDt As String
    Dim sWardID As String
    Dim sTmp    As String
    Dim sDupWard As String
    Dim sDupPtid As String
    Dim ii      As Integer
    Dim jj      As Integer
    
    With tblColList
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .CellType = CellTypeStaticText
        .Value = ""
        .BlockMode = False
    End With
    
    With tblWardList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1
            If .Value = 1 Then
                .Col = 3
                sWardID = sWardID & "'" & .Value & "',"
            End If
        Next
        If sWardID <> "" Then
            sWardID = Mid(sWardID, 1, Len(sWardID) - 1)
        Else
            MsgBox " 병동을 선택한후 조회하세요", vbInformation + vbOKOnly, "Info"
            Exit Sub
        End If
    End With
    
    Set Rs = New Recordset
    sWorkDt = Format(dtpColdt.Value, "yyyymmdd")
    Rs.Open QuerySQL(sWorkDt, sWardID), DBConn
    
    If Not Rs.EOF Then
        Set objPro = New clsProgress
        With objPro
'            .SetStsBar MainFrm.stsBar
            .Container = MainFrm.stsBar
            .Max = Rs.RecordCount
            .Message = "자료를 수집하고 있습니다."
        End With
        With tblColList
            .ReDraw = False
            Do Until Rs.EOF
                
                .Col = TblCol.tcTESTNM
                If Rs.Fields("labno").Value & "" = sTmp Then
                    .Value = .Value & "," & Rs.Fields("testnm").Value & ""
                
                
                Else
                    If .DataRowCnt + 1 > .MaxRows Then
                        .MaxRows = .MaxRows + 1
                    End If
                    .Row = .DataRowCnt + 1
                    .Value = Rs.Fields("testnm").Value & ""
                    .Col = TblCol.tcChk: .CellType = CellTypeCheckBox: .TypeCheckCenter = True
                End If
                
                .Col = TblCol.tcCOLDATE:    .Value = Format(Rs.Fields("coldt").Value & "", "####-##-##") & " " & _
                                                     Format(Mid(Rs.Fields("coltm").Value & "", 1, 4), "0#:##")
                .Col = TblCol.tcColNm:      .Value = GetEmpNm(Rs.Fields("colid").Value & "")
                .Col = TblCol.tcDonefg:     .Value = "√": .ForeColor = DCM_LightRed: .TypeHAlign = TypeHAlignCenter
                .Col = TblCol.tcLABNO:      .Value = Rs.Fields("labno").Value & ""
                .Col = TblCol.tcPTID:       .Value = Rs.Fields("ptid").Value & ""
                .Col = TblCol.tcPTNM:       .Value = Rs.Fields("ptnm").Value & ""
                .Col = TblCol.tcSPCNM:      .Value = Rs.Fields("spcnm").Value & ""

                .Col = TblCol.tcReqdt:      .Value = Rs.Fields("reqdt").Value & ""
                .Col = TblCol.tcReqtm:      .Value = Rs.Fields("reqtm").Value & ""
                .Col = TblCol.tcColdt:      .Value = Rs.Fields("coldt").Value & ""
                .Col = TblCol.tcColtm:      .Value = Rs.Fields("coltm").Value & ""
                .Col = TblCol.tcColID:      .Value = Rs.Fields("colid").Value & ""
                
                sTmp = Rs.Fields("labno").Value & ""
                .Col = TblCol.tcWARDID: .Value = Rs.Fields("wardid").Value & ""
                jj = jj + 1
                objPro.Value = jj
                Rs.MoveNext
            Loop
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = TblCol.tcWARDID
                If sDupWard = .Value Then
                    .ForeColor = .BackColor
                Else
                    .ForeColor = vbBlack
                End If
                sDupWard = .Value
                
                .Col = TblCol.tcPTID
                If sDupPtid = .Value Then
                    .ForeColor = .BackColor
                    .Col = TblCol.tcPTNM: .ForeColor = .BackColor
                Else
                    .ForeColor = vbBlack
                    .Col = TblCol.tcPTNM: .ForeColor = vbBlack
                End If
                .Col = TblCol.tcPTID
                sDupPtid = .Value
            Next
            .ReDraw = True
        End With
    Else
        MsgBox "채혈된 대상이 없거나 채혈된 검체에 대한 접수작업이 모두 이루어졌습니다.", vbInformation + vbOKOnly, "Info"
    End If
    Set Rs = Nothing
    Set objPro = Nothing
End Sub
Private Function QuerySQL(ByVal sWorkDt As String, ByVal WardId As String) As String
    Dim SSQL As String
    
    SSQL = "SELECT a.wardid,f.coldt,f.coltm,f.colid,g.reqdt,g.reqtm,a.workarea||'-'||a.accdt||'-'||a.accseq as labno ," & _
           "       b.ordcd,b.spccd,d.field5 as spcnm,b.ptid,c.abbrnm5 as testnm,e." & F_PTNM & " as ptnm" & _
           " FROM " & T_LAB204 & " a," & T_LAB102 & " b," & T_LAB001 & " c," & _
                      T_LAB032 & " d," & T_HIS001 & " e," & T_LAB201 & " f," & _
                      T_LAB101 & " g" & _
           " WHERE " & _
                      DBW("a.workdt=", sWorkDt)
    If WardId <> "" Then
        SSQL = SSQL & " AND a.wardid in (" & WardId & ")"
    End If
    
    SSQL = SSQL & " AND " & DBW("f.stscd=", "1")
    SSQL = SSQL & " AND " & DBW("a.mornfg=", "1")
    
    SSQL = SSQL & " AND NOT EXISTS (SELECT * FROM " & T_LAB206 & " z WHERE a.workarea=z.workarea AND a.accdt=z.accdt AND a.accseq=z.accseq)" & _
                  " AND  a.workarea=b.workarea " & _
                  " AND  a.accdt=b.accdt " & _
                  " AND  a.accseq=b.accseq" & _
                  " AND  b.ordcd=c.testcd " & _
                  " AND  a.workarea=f.workarea " & _
                  " AND  a.accdt=f.accdt " & _
                  " AND  a.accseq=f.accseq" & _
                  " AND  b.ptid=g.ptid " & _
                  " AND  b.orddt=g.orddt " & _
                  " AND b.ordno=g.ordno" & _
                  " AND  " & DBW("d.cdindex=", "C215") & _
                  " AND  b.spccd=d.cdval1 " & _
                  " AND  b.ptid=e." & F_PTID & _
                  " ORDER BY wardid, ptid,labno desc"
    QuerySQL = SSQL
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set objSC = Nothing
End Sub



'Private Sub mnuReg_Click()
'
'    If mnuReg.Caption = "미채혈등록" Then
'        Call SaveClear
'        fraColSave.Visible = True
'        With tblColList
'            .Row = lngSelRow
'            .Col = TblCol.tcPTID:   lblPtid.Caption = .Value
'            .Col = TblCol.tcPTNM:   lblPtnm.Caption = .Value
'            .Col = TblCol.tcSPCNM:  lblSpcnm.Caption = .Value
'            .Col = TblCol.tcLABNO:  lblWorkarea.Caption = medGetP(.Value, 1, "-")
'                                    lblAccdt.Caption = medGetP(.Value, 2, "-")
'                                    lblAccSeq.Caption = medGetP(.Value, 3, "-")
'            .Col = TblCol.tcReqdt:  lblReqdt.Caption = Format(.Value, "####-##-##")
'            .Col = TblCol.tcReqtm:  lblReqdt.Caption = lblReqdt.Caption & "   " & Format(.Value, "0#:##:##")
'
'            .Col = TblCol.tcReg:
'            If .Value <> "" Then
'                lblSave.Visible = True
'                cmdOk.Enabled = False
'            Else
'                cmdOk.Enabled = True
'            End If
'
'        End With
'    Else
'        Dim SSQL        As String
'        Dim strTmp      As String
'        Dim strColdt    As String
'        Dim strColTm    As String
'
'        strColdt = Format(dtpReqdt.Value, "YYYYMMDD")
'        strColTm = Replace(medGetP(cboSaveTime.Text, 1, " "), ":", "")
'
'        If lblCnt.Caption = "" Then
'            strTmp = MsgBox("스케줄 작성이 되지 않았습니다." & vbCRLF & _
'                            "채혈자를 변경하시겠습니까?", vbYesNo + vbInformation, "info")
'            If strTmp = vbNo Then Exit Sub
'            SSQL = "SELECT empid,empnm FROM " & T_COM006
'        Else
'            SSQL = "SELECT colid,empnm FROM " & t_lab901 & " WHERE " & DBW("coldt=", strColdt) & " AND " & DBW("coltm=", strColTm)
'        End If
'
'        Set objMyList = New clspopuplist
'        With objMyList
'            .Caption = "직원정보"
'            .HeadName = "사번,직원명"
'            .Tag = "empid"
'            Call .ListPop(SSQL, 8200, 7900)
'        End With
'        Set objMyList = Nothing
'    End If
'
'End Sub

'Private Sub objMyList_SendCode(ByVal SelString As String)
'    If objMyList.Tag = "reason" Then
'        lblReasonCd.Caption = "": lblReasonNm.Caption = ""
'        If SelString <> "" Then
'            lblReasonCd.Caption = medGetP(SelString, 1, ";")
'            lblReasonNm.Caption = medGetP(SelString, 2, ";")
'        End If
'    Else
'        If SelString = "" Then Exit Sub
'        With tblPtList
'            .Row = lngSelRow
'            .Col = tblCol2.enColID: .Value = medGetP(SelString, 1, ";")
'            .Col = tblCol2.enColNm: .Value = medGetP(SelString, 2, ";")
'        End With
'    End If
'End Sub

Private Sub objMyList_SelectedItem(ByVal pSelectedItem As String)
    Select Case objMyList.Tag
        Case TAG_RSN
            lblReasonCd.Caption = objMyList.SelectedItems(0)
            lblReasonNm.Caption = objMyList.SelectedItems(1)
        Case TAG_EMP
            With tblPtList
                .Row = lngSelRow
                .Col = tblCol2.enColID: .Value = objMyList.SelectedItems(0)
                .Col = tblCol2.enColNm: .Value = objMyList.SelectedItems(1)
            End With
    End Select
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_REG
            Call SaveClear
            fraColSave.Visible = True
            With tblColList
                .Row = lngSelRow
                .Col = TblCol.tcPTID:   lblPtId.Caption = .Value
                .Col = TblCol.tcPTNM:   lblPtNm.Caption = .Value
                .Col = TblCol.tcSPCNM:  lblSpcnm.Caption = .Value
                .Col = TblCol.tcLABNO:  lblWorkarea.Caption = medGetP(.Value, 1, "-")
                                        lblAccdt.Caption = medGetP(.Value, 2, "-")
                                        lblAccSeq.Caption = medGetP(.Value, 3, "-")
                .Col = TblCol.tcReqdt:  lblReqdt.Caption = Format(.Value, "####-##-##")
                .Col = TblCol.tcReqtm:  lblReqdt.Caption = lblReqdt.Caption & "   " & Format(.Value, "0#:##:##")
    
                .Col = TblCol.tcReg:
                If .Value <> "" Then
                    lblSave.Visible = True
                    cmdOk.Enabled = False
                Else
                    cmdOk.Enabled = True
                End If
    
            End With
        Case MENU_COL
            Dim SSQL        As String
            Dim strTmp      As String
            Dim strColdt    As String
            Dim strColTm    As String
    
            strColdt = Format(dtpReqdt.Value, "YYYYMMDD")
            strColTm = Replace(medGetP(cboSaveTime.Text, 1, " "), ":", "")
    
            If lblCnt.Caption = "" Then
                strTmp = MsgBox("스케줄 작성이 되지 않았습니다." & vbCRLF & _
                                "채혈자를 변경하시겠습니까?", vbYesNo + vbInformation, "info")
                If strTmp = vbNo Then Exit Sub
                SSQL = "SELECT empid,empnm FROM " & T_COM006
            Else
                SSQL = "SELECT colid,empnm FROM " & T_LAB901 & " WHERE " & DBW("coldt=", strColdt) & " AND " & DBW("coltm=", strColTm)
            End If
    
'            Set objMyList = New clspopuplist
            Set objMyList = New clsPopUpList
            With objMyList
                .Connection = DBConn
                .FormCaption = "직원정보"
                .ColumnHeaderText = "사번;직원명"
                .Tag = TAG_EMP
                .LoadPopUp SSQL
                
'                .Caption = "직원정보"
'                .HeadName = "사번,직원명"
'                .Tag = "empid"
'                Call .ListPop(SSQL, 8200, 7900)
            End With
            Set objMyList = Nothing
    End Select
End Sub

Private Sub optDiv_Click(Index As Integer)
    lblQuery.Caption = optDiv(Index).Caption
    picQuery(Index).ZOrder 0
    
    If Index = 0 Then
        tblPtList.MaxRows = 0
        lblCnt.Caption = "": lblColCnt.Caption = "": optSc(0).Value = True
        lblBuss.Caption = ""
        cboSaveTime.Clear
        If dtpReqdt.Enabled Then dtpReqdt.SetFocus
    End If
End Sub


Private Sub tblColList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ii As Integer
    Dim blnAccFg As Boolean
    Dim blnSavFG As Boolean
    
    If Col <> 1 Then Exit Sub
    If Row <> 0 Then Exit Sub
    With tblColList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblCol.tcAccFg
            If .Value = "" Then blnAccFg = True
            .Col = TblCol.tcReg
            If .Value = "" Then blnSavFG = True
            .Col = TblCol.tcChk
            If .CellType = CellTypeCheckBox And blnAccFg = True And blnSavFG = True Then
                .Col = TblCol.tcChk
                If blnCol Then
                    .Value = 0
                Else
                    .Value = 1
                End If
            End If
            blnAccFg = False: blnSavFG = False
        Next
        blnCol = IIf(blnCol = True, False, True)
        
    End With
End Sub

Private Sub tblColList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > tblColList.DataRowCnt Then Exit Sub
    With tblColList
        .Row = Row
        .Col = TblCol.tcAccFg
        If .Value <> "" Then
            MsgBox "접수처리된 검체입니다.", vbInformation + vbOKOnly, "Info"
            Exit Sub
        End If
        .Col = TblCol.tcReg
'        If .Value <> "" Then
'            MsgBox "미채혈등록된 검체입니다.", vbInformation + vbOKOnly, "Info"
'            Exit Sub
'        End If
    End With
    
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuReg = frmControls.mnuSub
'
'    mnuReg.Caption = "미채혈등록"
'    frmControls.mnuSub1.Visible = False
'    frmControls.mnuSub2.Visible = False
    lngSelRow = Row
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuReg = Nothing
    
    Set objPop = Nothing
    Set objPop = New clsPopupMenu
    
    With objPop
        .AddMenu MENU_REG, "미채혈 등록"
        
        .PopupMenus Me.hwnd
    End With
    
    Set objPop = Nothing
End Sub

Private Sub cmdExecute_Click()
    Dim ii As Integer
    
    Me.MousePointer = 11
    With tblColList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblCol.tcChk
            If .CellType = CellTypeCheckBox Then
                If .Value = 1 Then
                    .Col = TblCol.tcReg
                    If .Value = "" Then
                        Call DoAccession(ii)
                    End If
                End If
            End If
        Next
    End With
    Me.MousePointer = 0
End Sub
'% 접수Procedure를 수행한다.
Private Sub DoAccession(Optional ByVal ii As Integer = 0)

    Dim objAccess   As New clsLISAccession
    Dim blnSuccess  As Boolean
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String

    MouseRunning  '13
    With tblColList
        .Row = ii
        .Col = TblCol.tcLABNO:
                                sWorkArea = medGetP(.Value, 1, "-")
                                sAccDt = medGetP(.Value, 2, "-")
                                sAccSeq = medGetP(.Value, 3, "-")
        blnSuccess = objAccess.DoAccession(sWorkArea, sAccDt, CInt(sAccSeq), ObjSysInfo.EmpId)
        .Col = TblCol.tcAccFg
        If blnSuccess Then
            .Value = "접수": .ForeColor = DCM_LightBlue
        Else
            .Value = "오류": .ForeColor = DCM_LightRed
        End If
        .Col = TblCol.tcChk: .CellType = CellTypeStaticText: .Value = "√": .ForeColor = DCM_LightRed
                             .TypeHAlign = TypeHAlignCenter
    End With
    
    Set objAccess = Nothing
    MouseDefault
    
End Sub
Private Sub cmdQuery1_Click()
    Dim objPro   As clsProgress
    Dim Rs       As Recordset
    Dim sTmp     As String
    Dim sDupWard As String
    Dim sDupPtid As String
    Dim sDupDate As String
    Dim ii       As Integer
    Dim jj       As Integer
    
    blnNoCol = False
    With tblNoColList
        .ReDraw = False
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .CellType = CellTypeStaticText
        .Value = ""
        .BlockMode = False
        .ReDraw = True
    End With
    
    Set Rs = New Recordset
    Rs.Open QuerySQL1, DBConn
    If Not Rs.EOF Then
        Set objPro = New clsProgress
        With objPro
'            .SetStsBar MAINFRM.stsBar
            .Container = MainFrm.stsBar
            .Max = Rs.RecordCount
            .Message = "자료를 수집하고 있습니다."
        End With
        
        With tblNoColList
            .ReDraw = False
            Do Until Rs.EOF
                .Col = TblCol1.enTestNm
                If Rs.Fields("labno").Value & "" = sTmp Then
                    .Value = .Value & "," & Rs.Fields("testnm").Value & ""
                Else
                    If .DataRowCnt + 1 > .MaxRows Then
                        .MaxRows = .MaxRows + 1
                    End If
                    .Row = .DataRowCnt + 1
                    .Value = Rs.Fields("testnm").Value & ""
                    .Col = TblCol1.enDone:
                    If Rs.Fields("donefg").Value & "" <> "1" Then
                        .CellType = CellTypeCheckBox: .TypeCheckCenter = True
                    Else
                        .Value = "√": .TypeHAlign = TypeHAlignCenter: .ForeColor = DCM_LightRed
                    End If
                End If
                
                .Col = TblCol1.enWardID:    .Value = Rs.Fields("wardid").Value & ""
                .Col = TblCol1.enColDate:   .Value = Format(Rs.Fields("coldt").Value & "", "####-##-##") & " " & _
                                                     Format(Mid(Rs.Fields("coltm").Value & "", 1, 4), "0#:##")
                .Col = TblCol1.enColID:     .Value = Rs.Fields("colid").Value & ""
                .Col = TblCol1.enDate:      .Value = Format(Rs.Fields("vfydt").Value & "", "####-##-##")
                .Col = TblCol1.enLabno:     .Value = Rs.Fields("labno").Value & ""
                .Col = TblCol1.enPtid:      .Value = Rs.Fields("ptid").Value & ""
                .Col = TblCol1.enPtNm:      .Value = Rs.Fields("ptnm").Value & ""
                .Col = TblCol1.enReason:    .Value = Rs.Fields("rmkcd").Value & ""
                .Col = TblCol1.enReasonNm:  .Value = Rs.Fields("rsttxt").Value & ""
                .Col = TblCol1.enSaveNm:    .Value = GetEmpNm(Rs.Fields("vfyid").Value & "")
                .Col = TblCol1.enSpcNm:     .Value = Rs.Fields("spcnm").Value & ""
                '.Col = TblCol1.enTestNm:    .Value = RS.Fields("testnm").Value & ""
                sTmp = Rs.Fields("labno").Value & ""
                jj = jj + 1
                objPro.Value = jj
                Rs.MoveNext
            Loop

            
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = TblCol1.enDate
                If sDupDate = .Value Then
                    .ForeColor = .BackColor
                Else
                    .ForeColor = vbBlack
                End If
                sDupDate = .Value
                
                .Col = TblCol1.enWardID
                If sDupWard = .Value Then
                    .ForeColor = .BackColor
                Else
                    .ForeColor = vbBlack
                End If
                sDupWard = .Value
                .Col = TblCol1.enPtid
                If sDupPtid = .Value Then
                    .ForeColor = .BackColor
                    .Col = TblCol1.enPtNm: .ForeColor = .BackColor
                Else
                    .ForeColor = vbBlack
                    .Col = TblCol1.enPtNm: .ForeColor = vbBlack
                End If
                .Col = TblCol1.enPtid
                sDupPtid = .Value
            Next
            .ReDraw = True
        End With
    Else
        MsgBox "미채혈된 대상이 없습니다.", vbInformation + vbOKOnly, "Info"
    End If
    Set Rs = Nothing
    Set objPro = Nothing
End Sub
Private Sub cmdCollectSave_Click()
    Dim ii  As Integer
    Dim blnSave As Boolean
    
    With tblNoColList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblCol1.enDone
            If .CellType = CellTypeCheckBox Then
                If .Value = 1 Then
                    blnSave = CollectSave(ii)
                    If blnSave = True Then
                        .Col = TblCol1.enDone
                        .CellType = CellTypeStaticText: .Value = "√":
                        .TypeHAlign = TypeHAlignCenter
                        .ForeColor = DCM_LightRed
                    End If
                End If
            End If
        Next
    End With
End Sub
Private Function CollectSave(ByVal ii As Integer) As Boolean
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim SSQL        As String
    
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    With tblNoColList
        .Row = ii
        .Col = TblCol1.enLabno: sWorkArea = medGetP(.Value, 1, "-")
                                sAccDt = medGetP(.Value, 2, "-")
                                sAccSeq = medGetP(.Value, 3, "-")
    End With
    
    SSQL = " update " & T_LAB206 & " set " & _
                     DBW("donefg", "1", 3) & _
                     DBW("vfydt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
                     DBW("vfytm", Format(GetSystemDate, "HHMMSS"), 2) & _
           " WHERE " & _
                     DBW("workarea=", sWorkArea) & _
           " AND " & DBW("accdt=", sAccDt) & _
           " AND " & DBW("accseq=", sAccSeq)
        
    DBConn.Execute SSQL
    DBConn.CommitTrans
    CollectSave = True
    Exit Function
    
SAVE_ERROR:
    DBConn.RollbackTrans
End Function
Private Function QuerySQL1() As String
    Dim SSQL    As String
    Dim sFdate  As String
    Dim sTdate  As String
    
    sFdate = Format(dtpFColDt.Value, "YYYYMMDD")
    sTdate = Format(dtpTColdt.Value, "YYYYMMDD")
    
    SSQL = "SELECT a.workarea||'-'||a.accdt||'-'||a.accseq as labno,a.ptid,a.wardid,a.donefg," & _
           " a.coldt,a.coltm,a.colid,a.vfydt,a.vfytm,a.vfyid,a.rmkcd,a.rsttxt," & _
           " b.ordcd as testcd,c.abbrnm10 as testnm,d.field5 as spcnm,e." & F_PTNM & " as ptnm" & _
           " FROM " & T_LAB206 & " a," & T_LAB102 & " b," & _
                      T_LAB001 & " c," & T_LAB032 & " d," & _
                      T_HIS001 & " e" & _
           " WHERE " & _
                    DBW("a.vfydt>=", sFdate) & _
           " AND   " & DBW("a.vfydt<=", sTdate) & _
           " AND  a.workarea=b.workarea" & _
           " AND   a.accdt=b.accdt" & _
           " AND   a.accseq=b.accseq" & _
           " AND   b.ordcd=c.testcd" & _
           " AND   d.cdindex='C215'" & _
           " AND   d.cdval1=b.spccd" & _
           " AND   a.ptid=e." & F_PTID & _
           " ORDER BY vfydt,ptid,labno"
    QuerySQL1 = SSQL
End Function

Private Sub tblNoColList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ii As Integer
    Dim blnAccFg As Boolean
    Dim blnSavFG As Boolean
    
    If Col <> TblCol1.enDone Then Exit Sub
    If Row <> 0 Then Exit Sub
    With tblNoColList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblCol1.enDone
            If .CellType = CellTypeCheckBox Then
                .Col = TblCol1.enDone
                If blnNoCol Then
                    .Value = 1
                Else
                    .Value = 0
                End If
            End If
        Next
        blnNoCol = IIf(blnNoCol = True, False, True)
    End With
End Sub

Private Sub tblNoColList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim sPtInfo     As String
    Dim sLabNo      As String
    Dim sColDate    As String
    Dim sReason     As String
    Dim sDone       As String
    Dim sColNm      As String
    Dim sToolTip      As String
    
    With tblNoColList
        If Row < 1 Then Exit Sub
        If Row > .DataRowCnt Then Exit Sub
        .Row = Row
        Call .SetTextTipAppearance("굴림체", 9, False, False, &HFFFFC0, vbBlack)
        
        .Col = TblCol1.enPtid: sPtInfo = .Value
        .Col = TblCol1.enPtNm: sPtInfo = sPtInfo & " [ " & .Value & " ]"
        .Col = TblCol1.enLabno: sLabNo = .Value
        .Col = TblCol1.enColDate: sColDate = .Value
        .Col = TblCol1.enDone
        If .CellType = CellTypeStaticText Then
            sDone = "YES"
        Else
            sDone = "NO"
        End If
        .Col = TblCol1.enReason:   sReason = .Value
        .Col = TblCol1.enReasonNm: sReason = sReason & " " & .Value
        .Col = TblCol1.enColID: sColNm = GetEmpNm(.Value)
        
        sToolTip = vbNewLine & "             ♣ 미채혈리스트 세부내역 ♣  " & vbNewLine
        sToolTip = sToolTip & vbNewLine & "  환 자  정 보 : " & sPtInfo
        sToolTip = sToolTip & vbNewLine & "  접 수  번 호 : " & sLabNo
        sToolTip = sToolTip & vbNewLine & "  채 혈  일 시 : " & sColDate
        sToolTip = sToolTip & vbNewLine & "  채   혈   자 : " & sColNm
        sToolTip = sToolTip & vbNewLine & "  검체채취여부 : " & sDone
        sToolTip = sToolTip & vbNewLine & "  미 채혈 사유 : " & sReason

        
        TipWidth = 5000
        MultiLine = 1
        TipText = sToolTip & vbNewLine
        ShowTip = True
    End With
End Sub

'아침채혈 관련 부분~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub cmdReqQuery_Click()
    Dim objMySql    As clsLISSqlCollection
    Dim Rs          As Recordset
    Dim strDate     As String
    Dim strTime     As String
    Dim strWardID   As String
    Dim SSQL        As String
    Dim tmpPtid     As String
    Dim tmpSpcCd    As String
    Dim tmpTest     As String
    Dim intPtCount  As Integer
    Dim ii          As Integer
    
    With tblWardList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1
            If .Value = 1 Then
                .Col = 3
                strWardID = strWardID & "'" & .Value & "',"
            End If
        Next
        If strWardID <> "" Then
            strWardID = Mid(strWardID, 1, Len(strWardID) - 1)
        Else
            MsgBox " 병동을 선택한후 조회하세요", vbInformation + vbOKOnly, "Info"
            Exit Sub
        End If
    End With
    Me.MousePointer = 11
    
    strDate = Format(dtpReqdt.Value, CS_DateDbFormat)
    strTime = Format(dtpReqdt.Value, CS_TimeDbFormat)
    
    
    Set objMySql = New clsLISSqlCollection
    SSQL = SqlOrderForMornCol(strDate, strTime, strWardID)
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    With tblPtList
        .MaxRows = 0
        .ReDraw = False
        If Not Rs.EOF Then
            Do Until Rs.EOF
                If tmpPtid <> Trim(Rs.Fields("PtId").Value & "") Then
                    intPtCount = intPtCount + 1
                    
                    If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .Col = tblCol2.enLocation: .Value = Rs.Fields("wardid").Value & ""
                                               If Rs.Fields("hosilid").Value & "" <> "" Then .Value = .Value & "-" & Rs.Fields("hosilid").Value & ""
                    .Col = tblCol2.enWardID:    .Value = Rs.Fields("wardid").Value & ""
                    .Col = tblCol2.enHosil:     .Value = Rs.Fields("hosilid").Value & ""
                    .Col = tblCol2.enRoom:      .Value = Rs.Fields("roomid").Value & ""
                    .Col = tblCol2.enPtid:      .Value = Rs.Fields("ptid").Value & ""
                    .Col = tblCol2.enPtNm:      .Value = Rs.Fields("ptnm").Value & ""
                    .Col = tblCol2.enDOB:       .Value = Rs.Fields("dob").Value & ""
                    .Col = tblCol2.enBedInDt:   .Value = Rs.Fields("bedindt").Value & ""
                    .Col = tblCol2.enSEX:       .Value = Trim("" & Rs.Fields("Sex").Value)
                                                If IsNumeric(.Text) Then .Value = Choose((Val(.Text) Mod 2) + 1, "F", "M")
                    .Col = tblCol2.enDept:      .Value = "" & Rs.Fields("DeptCd").Value                      '진료과
                    .Col = tblCol2.enOrdDoct:   .Value = "" & Rs.Fields("OrdDoct").Value                     '처방의
                    .Col = tblCol2.enMajDoct:   .Value = "" & Rs.Fields("MajDoct").Value                    '주치의
                    .Col = tblCol2.enColID:     .Value = ObjSysInfo.EmpId
                    .Col = tblCol2.enColNm:     .Value = ObjSysInfo.EmpNm
                    .Col = tblCol2.enRndFg:     .Value = "1"
                    tmpPtid = "" & Rs.Fields("PtId").Value
                End If
                .Col = tblCol2.enSpccd
                '검체
                tmpSpcCd = Rs.Fields("spccd").Value & ""
                tmpTest = Rs.Fields("ordcd").Value & ""
                
                Dim strSpcAbbr As String
                Dim strLabRng As String
                
                Call GetSpcInfo(tmpSpcCd, strSpcAbbr, strLabRng)
                
                If strSpcAbbr <> "" Then
                    tmpSpcCd = strSpcAbbr
                Else
                    tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, lis_orddiv)
                End If
                
'                If objLisComCode.LisSpc.Exists(tmpSpcCd) Then
'                    objLisComCode.LisSpc.KeyChange (tmpSpcCd)
'                    tmpSpcCd = objLisComCode.LisSpc.Fields("spcbarnm")
'                Else
'                    tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, lis_orddiv)
'                End If
                
                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpSpcCd & ", "
                End If
                .ForeColor = DCM_LightBlue
                
                .Col = tblCol2.enTest:                          '검사명
'                If objLisComCode.LisItem.Exists(tmpTest) Then
'                    objLisComCode.LisItem.KeyChange (tmpTest)
'                    tmpTest = objLisComCode.LisItem.Fields("abbrnm5")
'                Else
'                    tmpTest = tmpTest
'                End If
                
                Dim strTmp As String
                strTmp = GetAbbrNm(tmpTest)
                
                If strTmp = "" Then
                    tmpTest = tmpTest
                Else
                    tmpTest = strTmp
                End If
                
                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpTest & ", "
                End If
                .ForeColor = DCM_LightRed
                
                Rs.MoveNext
            Loop
            Call MornDSPCollector
        Else
            MsgBox "아침채혈 대상이 없습니다.", vbInformation + vbOKOnly, "Info"
        End If
        .ReDraw = True
    End With
    
   
    Set Rs = Nothing
    Set objMySql = Nothing
    Me.MousePointer = 0
End Sub

Private Function GetAbbrNm(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select abbrnm5 from " & T_LAB001 & _
            " where " & DBW("testcd=", vTestCd)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    GetAbbrNm = Rs.Fields("abbrnm5").Value & ""
    
    Set Rs = Nothing
End Function

Private Sub MornDSPCollector()
    Dim Rs          As Recordset
    Dim ObjDic      As clsDictionary
    Dim aryTmp()    As String
    Dim strTmp      As String
    Dim sColDt      As String
    Dim sColTm      As String
    Dim SSQL        As String
    
    
    Dim ii          As Integer
    Dim jj          As Integer
    
    On Error GoTo Errors
    
    If cboSaveTime.ListCount < 0 Then Exit Sub
    
    Set ObjDic = New clsDictionary
    
    ObjDic.Clear
    ObjDic.FieldInialize "wardid", "empid,empnm"
    
    sColDt = Format(dtpReqdt.Value, "YYYYMMDD")
    sColTm = Replace(medGetP(cboSaveTime.Text, 1, " "), ":", "")
    
    SSQL = " SELECT colid,bussdiv,wardid,empnm,cnt FROM " & T_LAB901 & _
           " WHERE  " & DBW("coldt=", sColDt) & _
           " AND " & DBW("coltm=", sColTm)
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        Do Until Rs.EOF
            If Rs.Fields("bussdiv").Value & "" = "2" And Rs.Fields("wardid").Value & "" <> "" Then
                aryTmp = Split(Rs.Fields("wardid").Value & "", ",")
                For ii = LBound(aryTmp) To UBound(aryTmp)
                    If Not ObjDic.Exists(aryTmp(ii)) Then
                        ObjDic.AddNew aryTmp(ii), Rs.Fields("colid").Value & "" & COL_DIV & _
                                                 Rs.Fields("empnm").Value & ""
                    End If
                Next
            Else
                strTmp = strTmp & Rs.Fields("colid").Value & "" & vbTab & Rs.Fields("empnm").Value & "" & COL_DIV
            End If
            Rs.MoveNext
        Loop
    End If
    
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        aryTmp() = Split(strTmp, COL_DIV)
    End If
    
    '병동별 분담
    If medGetP(lblBuss.Tag, 1, vbTab) = "2" Then
        If optSc(1).Value Then
            With tblPtList
                For ii = 1 To .DataRowCnt
                    .Row = ii
                    .Col = tblCol2.enColID: .Value = ObjSysInfo.EmpId
                    .Col = tblCol2.enColNm: .Value = ObjSysInfo.EmpNm
                Next
            End With
        Else
            If ObjDic.RecordCount < 1 Then Exit Sub
            With tblPtList
                For ii = 1 To .DataRowCnt
                    .Row = ii
                    .Col = tblCol2.enWardID
                    If ObjDic.Exists(.Value) Then
                        ObjDic.KeyChange .Value
                        .Col = tblCol2.enColID: .Value = ObjDic.Fields("empid")
                        .Col = tblCol2.enColNm: .Value = ObjDic.Fields("empnm")
                    End If
                Next
            End With
        End If
    Else
        Dim lngDiv  As Long
        
        If strTmp = "" Then Exit Sub
        
        lngDiv = UBound(aryTmp) + 1
        
        '채혈자수 분담.
        '채혈환자가 채혈자수보다 적을경우 한명이 모두 처리한다.
        If optSc(1).Value Then
            With tblPtList
                For ii = 1 To .DataRowCnt
                    .Row = ii
                    .Col = tblCol2.enColID: .Value = ObjSysInfo.EmpId
                    .Col = tblCol2.enColNm: .Value = ObjSysInfo.EmpNm
                Next
            End With
        Else
            With tblPtList
                If UBound(aryTmp) + 1 > .DataRowCnt Then
                    For ii = 1 To .DataRowCnt
                        .Row = ii
                        .Col = tblCol2.enColID: .Value = medGetP(aryTmp(jj), 1, vbTab)
                        .Col = tblCol2.enColNm: .Value = medGetP(aryTmp(jj), 2, vbTab)
                        If ii Mod UBound(aryTmp) + 1 = 0 Then
                            jj = jj + 1
                        End If
                    Next
                Else
                    lngDiv = CLng(.DataRowCnt / (UBound(aryTmp) + 1))
                    
                    For ii = 1 To .DataRowCnt
                        .Row = ii
                        .Col = tblCol2.enColID: .Value = medGetP(aryTmp(jj), 1, vbTab)
                        .Col = tblCol2.enColNm: .Value = medGetP(aryTmp(jj), 2, vbTab)
                        If ii Mod lngDiv = 0 Then
                            jj = jj + 1
                        End If
                        
                    Next
                End If
            End With
        End If
    End If
    

Errors:

    Set Rs = Nothing
    Set ObjDic = Nothing
End Sub

Private Sub optSc_Click(Index As Integer)
    Call MornDSPCollector
End Sub

Private Function SqlReadOrderForMornCol(ByVal PtId As String, ByVal ReqDt As String, ByVal ReqTm As String) As String
 
    Dim tmpStr As String
    Dim strSQL(3) As String

    tmpStr = ""

     '임상병리 아침채혈
    SqlReadOrderForMornCol = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, " & _
                         FUNC_CONVERT("int", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & _
                           T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND    a.donefg = '0' " & _
                " AND  " & DBW("a.reqdt < ", ReqDt) & _
                " AND  " & DBW("a.orddiv = ", lis_orddiv) & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND    b.donefg = '0' " & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd  = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                " AND  " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) " & _
                " AND    f.rndfg = '1' AND " & DBW("a.bussdiv=", enBussDiv.BussDiv_InPatient)
    SqlReadOrderForMornCol = SqlReadOrderForMornCol & " union all " & _
                "SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, e." & F_DOCTNM & " as DoctNm, a.deptcd,  f.statflags, " & _
                         FUNC_CONVERT("int", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_LAB101 & " a, " & T_LAB102 & " b, " & T_LAB001 & " c, " & _
                           T_LAB032 & " d, " & T_HIS005 & " e, " & T_LAB004 & " f " & _
                " WHERE " & DBW("a.ptid = ", PtId) & _
                " AND    a.donefg = '0' " & _
                " AND  " & DBW("a.reqdt = ", ReqDt) & _
                " AND  " & DBW("a.reqtm <= ", ReqTm) & _
                " AND  " & DBW("a.orddiv = ", lis_orddiv) & _
                " AND    b.ptid  = a.ptid  AND    b.orddt = a.orddt  AND    b.ordno = a.ordno " & _
                " AND    b.donefg = '0' " & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd  = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                " AND  " & DBJ("e." & F_DOCTID & " =* a.orddoct") & _
                " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & _
                " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) " & _
                " AND    f.rndfg = '1' AND " & DBW("a.bussdiv=", enBussDiv.BussDiv_InPatient)
                
                

    SqlReadOrderForMornCol = SqlReadOrderForMornCol & " ORDER BY ColTm, orddt, ordno, ordcd "                         '<< D/C 처방 제외 >>
      
End Function

Private Function SqlOrderForMornCol(ByVal ReqDt As String, ByVal ReqTm As String, ByVal WardId As String) As String

' 임상병리 아침채혈
' 기준 : 작업시점을 기준으로 발생되어 있는 모든처방
'        임상병리일반검사 중 검체가 혈액인 처방 (임상병리 마스터에서 정의 : 검사-검체)
'        응급처방이 하나라도 포함되어 있으면 표시 --> 간호사채혈
'        비응급처방만 임상병리사가 채혈함

    SqlOrderForMornCol = " SELECT  b.wardid as WardId, a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                                   F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                                   F_DOB2("a") & " as Dob, b.bedindt, b.deptcd, " & _
                         "        b.orddoct, b.majdoct, b.roomid,c.ordcd,'' as bedid, b.hosilid, b.orddiv  " & _
                         " FROM   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c, " & T_LAB004 & " d " & _
                         " WHERE  b.wardid in (" & WardId & ")" & _
                         " AND    " & DBW("b.donefg", "0", 2) & _
                         " AND    " & DBW("b.reqdt < ", ReqDt) & _
                         " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
                         " AND    " & DBW("b.orddiv", lis_orddiv, 2) & _
                         " AND    a." & F_PTID & " = b.ptid " & _
                         " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & _
                         " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                         " AND    d.testcd = c.ordcd " & _
                         " AND    d.spccd  = c.spccd " & _
                         " AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                                             " WHERE testcd = d.testcd  AND  spccd = d.spccd)  " & _
                         " AND    " & DBW("c.statfg<>", "1") & _
                         " AND    d.rndfg = '1' AND " & DBW("b.bussdiv=", enBussDiv.BussDiv_InPatient)

    SqlOrderForMornCol = SqlOrderForMornCol & " union all " & _
                         " SELECT  b.wardid as WardId, a." & F_PTID & " as PtId, a." & F_PTNM & " as PtNm, " & _
                                   F_SEX2("a") & " as Sex, c.statfg, c.spccd, " & _
                                   F_DOB2("a") & " as Dob, b.bedindt, b.deptcd, " & _
                         "        b.orddoct, b.majdoct, b.roomid,c.ordcd, '' as bedid, b.hosilid, b.orddiv  " & _
                         " FROM   " & T_HIS001 & " a, " & T_LAB101 & " b, " & T_LAB102 & " c, " & T_LAB004 & " d " & _
                         " WHERE   b.wardid in (" & WardId & ")" & _
                         " AND    " & DBW("b.donefg", "0", 2) & _
                         " AND    " & DBW("b.reqdt = ", ReqDt) & _
                         " AND    " & DBW("b.reqtm <= ", ReqTm) & _
                         " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
                         " AND    " & DBW("b.orddiv", lis_orddiv, 2) & _
                         " AND    a." & F_PTID & " = b.ptid " & _
                         " AND    c.ptid = b.ptid    AND   c.orddt = b.orddt  AND   c.ordno = b.ordno  " & _
                         " AND    c.stscd = '0'  AND  ( c.dcfg = '' or c.dcfg is null ) " & _
                         " AND    d.testcd = c.ordcd " & _
                         " AND    d.spccd  = c.spccd " & _
                         " AND    d.applydt = (SELECT max(applydt) FROM " & T_LAB004 & _
                                             " WHERE testcd = d.testcd  AND  spccd = d.spccd)  " & _
                         " AND    " & DBW("c.statfg<>", "1") & _
                         " AND    d.rndfg = '1' AND " & DBW("b.bussdiv=", enBussDiv.BussDiv_InPatient) & _
                         " Order  By WardId,hosilid, PtId, statfg, spccd "

End Function

Private Sub dtpReqdt_Change()
    tblPtList.MaxRows = 0
    Call GetSchedule
End Sub

Private Sub dtpReqdt_LostFocus()
    Call GetSchedule
End Sub

Private Sub GetSchedule()
    Dim Rs      As Recordset
    Dim SSQL    As String
    
    On Error GoTo Errors
    
    cboSaveTime.Clear
    lblColCnt.Caption = "": lblCnt.Caption = "": lblBuss.Caption = ""
    optSc(1).Value = True
    
    SSQL = " SELECT distinct a.coltm ,b.field1 FROM " & T_LAB032 & " b," & T_LAB901 & " a " & _
           " WHERE " & DBW("a.coldt=", Format(dtpReqdt.Value, "YYYYMMDD")) & _
           " AND " & DBW("b.cdindex=", LC3_RoundTime) & _
           " AND a.coltm=b.cdval1 " & _
           " ORDER BY coltm"
    
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        Do Until Rs.EOF
            cboSaveTime.AddItem Format(Rs.Fields("coltm").Value & "", "0#:##") & " [ " & Rs.Fields("field1").Value & "" & "]        "
            Rs.MoveNext
        Loop
        lblCnt.Caption = Rs.RecordCount:
        optSc(0).Value = True
        cboSaveTime.ListIndex = 0
    End If
Errors:
    Set Rs = Nothing
End Sub


Private Sub cboSaveTime_Click()
    Dim Rs          As Recordset
    Dim aryTmp()    As String
    Dim sColDt      As String
    Dim sColTm      As String
    Dim SSQL        As String
    Dim blnTF       As Boolean
    Dim ii          As Integer
    
    If cboSaveTime.ListCount < 0 Then Exit Sub
    
    Set objSC = New clsDictionary
    objSC.Clear
    objSC.FieldInialize "wardid", "colid"
    
    sColDt = Format(dtpReqdt.Value, "YYYYMMDD")
    sColTm = Replace(medGetP(cboSaveTime.Text, 1, " "), ":", "")
    
    SSQL = " SELECT colid,bussdiv,wardid,empnm,cnt FROM " & T_LAB901 & " " & _
           " WHERE  " & DBW("coldt=", sColDt) & _
           " AND " & DBW("coltm=", sColTm)
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        lblColCnt.Caption = Rs.Fields("empnm").Value & "" & " 외" & Rs.RecordCount & " 명"
        If Rs.Fields("bussdiv").Value & "" = "1" Then
            lblBuss.Caption = "채혈자수 비례"
            lblBuss.Tag = "1" & vbTab & Rs.Fields("cnt").Value & ""
        Else
            lblBuss.Caption = "병동별 담당"
            lblBuss.Tag = "2" & vbTab & Rs.Fields("cnt").Value & ""
        End If
        
        Rs.MoveFirst
        Do Until Rs.EOF
            If Rs.Fields("bussdiv").Value & "" = "2" And Rs.Fields("wardid").Value & "" <> "" Then
                aryTmp = Split(Rs.Fields("wardid").Value & "", ",")
                For ii = LBound(aryTmp) To UBound(aryTmp)
                    If Not objSC.Exists(aryTmp(ii)) Then
                        objSC.AddNew aryTmp(ii), Rs.Fields("colid").Value & "" & vbTab & _
                                                 Rs.Fields("empnm").Value & ""
                    End If
                Next
            End If
            Rs.MoveNext
        Loop
    End If
    
    Set Rs = Nothing
    Call MornDSPCollector
    
End Sub

Private Sub tblPtList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > tblPtList.DataRowCnt Then Exit Sub
    
    tblPtList.Row = Row: tblPtList.Col = tblCol2.enPtid
    If tblPtList.Value = "" Then Exit Sub
    
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuReg = frmControls.mnuSub
'
'    mnuReg.Caption = "채혈자변경"
'    frmControls.mnuSub1.Visible = False
'    frmControls.mnuSub2.Visible = False
    lngSelRow = Row
'    PopupMenu mnuPopup

'    Set mnuPopup = Nothing
'    Set mnuReg = Nothing
    
    Set objPop = Nothing
    Set objPop = New clsPopupMenu
    
    With objPop
        .AddMenu MENU_COL, "채혈자변경"
        
        .PopupMenus Me.hwnd
    End With
    
    Set objPop = Nothing
End Sub

Private Sub cmdCollection_Click()
    Dim strWardID   As String
    Dim aryTmp()    As String
    Dim tmpWardID   As String
    Dim strTmp      As String
    Dim i           As Integer
    Dim j           As Integer

    
    
    
    Me.MousePointer = 11
    sWorkDt = Format(GetSystemDate, CS_DateDbFormat)
    sWorkTm = Format(GetSystemDate, CS_TimeDbFormat)
    
'    Call PrintCollectList("21W", "개발자")
'    Exit Sub
    
    With tblPtList
        For i = 1 To .DataRowCnt
            .Row = i
            '* 제외버튼 Check
            .Col = tblCol2.enChk: If .Value = 1 Then GoTo Skip
            .Col = tblCol2.enWardID: strWardID = .Value
            '* 채혈수행
            Call DoCollectionForLIS(strWardID, i)
Skip:
        Next
        
    End With
    
    strWardID = ""
    If chkPrint.Value = 1 Then
        Call PrintIntionlize
        With tblPtList
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = tblCol2.enChk
                If .Value <> 1 Then
                    .Col = tblCol2.enWardID
                    If .ForeColor <> vbRed Then
                        If .Value <> tmpWardID Then
                            strTmp = .Value
                            .Col = tblCol2.enColNm
                            strWardID = strWardID & strTmp & vbTab & .Value & ","
                            Debug.Print "strwardid =======> " & strWardID
'                            Debug.Print "tmpwardid =======> " & tmpWardID
'                            Debug.Print "strtmp    =======> " & strTmp
                            tmpWardID = strTmp
                        End If
                    End If
                End If
            Next
        End With
        If strWardID <> "" Then
            strWardID = Mid(strWardID, 1, Len(strWardID) - 1)
            aryTmp() = Split(strWardID, ",")
            For i = LBound(aryTmp) To UBound(aryTmp)
                '출력 함수
                For j = 1 To Val(txtCopy.Text)
                    Call PrintCollectList(medGetP(aryTmp(i), 1, vbTab), medGetP(aryTmp(i), 2, vbTab))
                    If i <> UBound(aryTmp) Then Printer.NewPage
                    
                Next
            Next
            Printer.EndDoc
        End If
        
    End If
    tblPtList.MaxRows = 0
    MsgBox "정상적으로 처리되었습니다.", vbInformation + vbOKOnly, "Info"
    Me.MousePointer = 0
    
End Sub


'& 채혈 클래스 MyCollect 를 이용하여 해당 환자들의 처방을 채혈수행한다.
Private Sub DoCollectionForLIS(ByVal sWardID As String, ByVal Row As Long)
    Dim objLISCollect   As clsLISCollectioin
    Dim tmpRs           As Recordset
    Dim tmpData()       As String
    Dim tmpDate         As String
    Dim tmpTime         As String
    Dim tmpStatFg       As String
    Dim tmpTestFg       As String
    Dim SqlStmt         As String
    Dim tmpDeptCd       As String
    Dim tmpOrdDoct      As String
    Dim tmpMajDoct      As String
    Dim blnSuccess      As Boolean
    
    Dim i As Integer
    Dim j As Integer
    
    '* 채혈 Class Initialize
    Set objLISCollect = New clsLISCollectioin
    With objLISCollect
'        Call .InitRtn
        Call .SetWardCol(sWorkDt, sWorkTm, sWardID)                             '병동채혈내역 Seq구하기
        .MornFg = "1"                                                           '아침채혈여부
        .ColTm = sWorkTm                                                        '채혈시간
    End With
    
    tmpDate = Format(dtpReqdt.Value, CS_DateDbFormat)
    tmpTime = "235959"
    
    
    ReDim tmpData(0 To 16)
    
    With tblPtList
        .Row = Row
                                    tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)              '
        .Col = tblCol2.enPtid:      tmpData(1) = .Value                                     '환자ID
        .Col = tblCol2.enPtNm:      tmpData(2) = .Value                                                 '환자명
        .Col = tblCol2.enSEX:       tmpData(3) = .Value                                                 '환자성별
        .Col = tblCol2.enDOB:
                                    If IsDate(Format(.Value, CS_DateMask)) Then
                                        tmpData(4) = DateDiff("y", Format(.Value, CS_DateMask), Now)    '환자일령
                                    Else
                                        tmpData(4) = Mid(.Value, 1, 4) & "-01-01"
                                        If IsDate(tmpData(4)) Then
                                            tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
                                        Else
                                            tmpData(4) = 0
                                        End If
                                    End If
        .Col = tblCol2.enBedInDt:   tmpData(5) = .Value                                                 '입원일
                                    tmpData(6) = sWorkDt                                                '입력일
                                    tmpData(7) = sWorkTm                                                '입력시간
        .Col = tblCol2.enColID:     tmpData(8) = .Value                                                 '입력자
                                    tmpData(9) = ""                                                     '원접수번호
                                    tmpData(10) = sWorkDt                                               '채혈일
        .Col = tblCol2.enColID:     tmpData(11) = .Value                                                '채혈자
        .Col = tblCol2.enWardID:    tmpData(12) = .Value                                                '병동ID
        .Col = tblCol2.enHosil:     tmpData(13) = .Value                                                '병실ID
        .Col = tblCol2.enRoom:      tmpData(14) = .Value                                                '호실ID
                                    tmpData(15) = ""                                                    '침상ID
                                    tmpData(16) = ObjSysInfo.BuildingCd                                 '채혈이 수행되는 건물코드
        Call objLISCollect.SetColData(tmpData)                                                          '채혈준비작업
        .Col = tblCol2.enDept:      tmpDeptCd = .Value                                                  '진료과
        .Col = tblCol2.enOrdDoct:   tmpOrdDoct = .Value                                                 '처방의
        .Col = tblCol2.enMajDoct:   tmpMajDoct = .Value                                                 '주치의
    End With
    SqlStmt = SqlReadOrderForMornCol(objLISCollect.PtId, tmpDate, tmpTime)                               ' 처방내역 검색
    
    blnSuccess = False
    On Error GoTo Err_Trap
    
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn

    ReDim tmpData(0 To 20)
    With tmpRs
        For i = 1 To .RecordCount
            tmpStatFg = medGetP("" & .Fields("StatFlags").Value, 1, ";")    '건물별 응급가능 여부
            tmpTestFg = medGetP("" & .Fields("StatFlags").Value, 2, ";")    '건물별 검사가능 여부
        '***건물정보 사용
            If P_ApplyBuildingInfo Then
                If Trim(.Fields("StatFg").Value) = "1" Then
                   If Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1) = "1" Then   '응급검사 가능
                      If ObjSysInfo.BuildingCd = CentralLab Or _
                         ObjSysInfo.BuildingCd = AneLab Then                '중앙/안이센터에서 응급검사가 발생하면..
                         tmpData(0) = EmergencyLab                          '응급센터로...
                      Else
                         tmpData(0) = ObjSysInfo.BuildingCd                 '해당건물에서 응급검사 가능함
                      End If
                      tmpData(4) = "1"                                      'StatFg
                      GoTo DataSet
                   Else
                   '*******************************************************************************************************
                   '** 여성/심장센터 : 응급검사가 가능하지 않을경우 응급실에서 검사가 가능하면 응급실로, 아니면 중앙으로...
                   '*******************************************************************************************************
                      If ObjSysInfo.BuildingCd = WomLab Or _
                         ObjSysInfo.BuildingCd = HrtLab Then                '여성/심장센터에서 응급검사가 발생하면..
                           If Mid(tmpStatFg, EmergencyNo, 1) = "1" Then     '응급실에서 응급검사 가능
                             tmpData(0) = EmergencyLab                      '응급센터로...
                             tmpData(4) = "1"                               'StatFg
                             GoTo DataSet
                           End If
                      End If
                   '*******************************************************************************************************
                   End If
                End If
                tmpData(4) = "0"                                            'StatFg
                If Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                   tmpData(0) = ObjSysInfo.BuildingCd                       '일반검사가능
                Else
                   tmpData(0) = CentralLab                                  '일반검사 불가능 --> 중앙검사실로...
                End If
    
        '***건물정보 사용하지 않음
            Else
                tmpData(0) = ObjSysInfo.BuildingCd
                tmpData(4) = Trim(.Fields("StatFg").Value)
            End If
                
DataSet:
            tmpData(1) = Trim("" & .Fields("WorkArea").Value)                               'WorkArea
            tmpData(2) = Trim("" & .Fields("SpcCd").Value)                                  'SpcCd
            tmpData(3) = Trim("" & .Fields("StoreCd").Value)                                'StoreCd
            tmpData(5) = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)          '희망채취일시
            tmpData(6) = Trim("" & .Fields("TestDiv").Value)                                'TestDiv
            tmpData(7) = Trim("" & .Fields("MultiFg").Value)                                'MultiFg
            tmpData(8) = Trim("" & .Fields("SpcGrp").Value)                                 'SpcGrp
            tmpData(9) = Trim("" & .Fields("OrdDt").Value)                                  'OrdDt
            tmpData(10) = Trim("" & .Fields("OrdNo").Value)                                 'OrdNo
            tmpData(11) = Trim("" & .Fields("OrdSeq").Value)                                'OrdSeq
            tmpData(12) = Trim("" & .Fields("OrdCd").Value)                                 'OrdCd
            tmpData(13) = tmpDeptCd
            tmpData(14) = tmpOrdDoct
            tmpData(15) = tmpMajDoct
            tmpData(16) = Trim("" & .Fields("AbbrNm5").Value)                               '처방 약어명
            tmpData(17) = Trim("" & .Fields("LabelCnt").Value)                              '라벨출력장수
'            Call objLisComCode.LisItem.KeyChange(tmpData(12))
            tmpData(18) = GetLabDiv(tmpData(12)) 'objLisComCode.LisItem.Fields("labdiv")                            'LabDiv
'            Call objLisComCode.LisSpc.KeyChange(tmpData(2))
            Call GetSpcInfo(tmpData(2), tmpData(19), tmpData(20))
'            tmpData(19) = objLisComCode.LisSpc.Fields("spcbarnm")                           '검체약어명
'            tmpData(20) = objLisComCode.LisSpc.Fields("labrange")                           '미생물접수번호범위
            Call objLISCollect.SetAddLabCollect(tmpData)
            .MoveNext
        Next
    End With

    ' 채혈 수행
    If tmpRs.RecordCount > 0 Then
        blnSuccess = objLISCollect.DoCollection
    Else
        GoTo Skip
    End If
Err_Trap:
    If Not blnSuccess Then
        tblPtList.Row = Row
        tblPtList.Col = -1
        tblPtList.ForeColor = vbRed       '빨간색
    End If
Skip:
    Set tmpRs = Nothing
    Set objLISCollect = Nothing
End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b "
    strSQL = strSQL & " where " & DBW("b.cdindex=", lc3_workarea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    GetLabDiv = Rs.Fields("field2").Value & ""
    
    Set Rs = Nothing
End Function

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    vSpcAbbr = Rs.Fields("spcbarnm").Value & ""
    vLabRng = Rs.Fields("labrange").Value & ""
    
    Set Rs = Nothing
End Sub

Private Sub PrintCollectList(ByVal sWardID As String, ByVal sEmpNm As String)
    Dim Rs      As Recordset
    Dim SSQL    As String
    
    Dim sNo     As String
    Dim sAccNo  As String
    Dim sPtNm   As String
    Dim sPtid   As String
    Dim sSEX    As String
    Dim sAge    As String
    Dim sHosil  As String
    Dim sSpcnm  As String
    Dim sTestNm As String
    Dim strTmp  As String
    Dim blnFirst As Boolean
    
    
    
'    sWorkDt = "20040302"
'    sWorkTm = "103703"
    
    SSQL = " SELECT b.ptid,a.workarea,a.accdt,b.accseq,c.ordcd,d.abbrnm5,e.field3,b.wardid,b.hosilid," & _
           " f." & F_PTNM & " as ptnm,f." & F_SEX & " as sex, " & F_DOB2("f") & " as dob" & _
           " FROM " & T_HIS001 & " f," & T_LAB032 & " e," & T_LAB001 & " d," & T_LAB102 & " c," & T_LAB201 & " b," & T_LAB204 & " a" & _
           " WHERE " & DBW("a.workdt=", sWorkDt) & _
           " AND " & DBW("a.wardid=", sWardID) & _
           " AND " & DBW("a.workTm=", sWorkTm) & _
           " AND " & DBW("a.orddiv=", lis_orddiv) & _
           " AND " & DBW("a.mornfg=", "1") & _
           " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
           " AND a.workarea=c.workarea AND a.accdt=c.accdt AND a.accseq=c.accseq" & _
           " AND c.ordcd=d.testcd" & _
           " AND " & DBW("e.cdindex=", LC3_Specimen) & _
           " AND e.cdval1=c.spccd" & _
           " AND f." & F_PTID & "=b.ptid" & _
           " ORDER BY ptid,workarea,accdt,accseq"
    Set Rs = New Recordset
    Rs.Open SSQL, DBConn
    
    If Not Rs.EOF Then
        Call PrinterHeader(sWardID, sEmpNm)
        sNo = "1"
        
        Do Until Rs.EOF
            sAccNo = Rs.Fields("workarea").Value & "" & "-" & Rs.Fields("accdt").Value & "" & "-" & Rs.Fields("accseq").Value & ""

            If strTmp = sAccNo Then
                sTestNm = sTestNm & Rs.Fields("abbrnm5").Value & "" & ","
            Else

                
                If blnFirst = True Then
                    sTestNm = Mid(sTestNm, 1, Len(sTestNm) - 1)
                    Call PrintBody(sNo, strTmp, sPtNm, sPtid, sSEX & "/" & sAge, sHosil, sSpcnm, sTestNm, sWardID, sEmpNm)
                    sNo = Val(sNo) + 1
                    sTestNm = ""
                Else
                    
                    blnFirst = True
                End If
                
                sPtNm = Rs.Fields("ptnm").Value & ""
                sPtid = Rs.Fields("ptid").Value & ""
                sHosil = Rs.Fields("hosilid").Value & ""
                sSpcnm = Rs.Fields("field3").Value & ""
                sSEX = Rs.Fields("sex").Value & ""
                sTestNm = sTestNm & Rs.Fields("abbrnm5").Value & "" & ","
                
                If IsNumeric(sSEX) Then sSEX = Choose((Val(sSEX) Mod 2) + 1, "F", "M")
                
                If IsDate(Format(Rs.Fields("dob").Value & "", CS_DateMask)) Then
                    sAge = DateDiff("y", Format(Rs.Fields("dob").Value & "", CS_DateMask), GetSystemDate)
                Else
                    sAge = Mid(sAge, 1, 4) & "-01-01"
                    If IsDate(sAge) Then
                        sAge = DateDiff("y", sAge, GetSystemDate)
                    Else
                        sAge = "0"
                    End If
                End If
                sAge = CLng((Val(sAge) / 365) + 1)
                strTmp = sAccNo
            End If
        
            Rs.MoveNext
        Loop
        sTestNm = Mid(sTestNm, 1, Len(sTestNm) - 1)
        Call PrintBody(sNo, strTmp, sPtNm, sPtid, sSEX & "/" & sAge, sHosil, sSpcnm, sTestNm, sWardID, sEmpNm)
    End If
    
'    Printer.EndDoc
    
    Set Rs = Nothing
End Sub

Private Sub PrintIntionlize()
    Printer.Font = "굴림체"
    Printer.FontSize = 9
    Printer.Orientation = vbPRORPortrait
    Printer.ScaleMode = vbMillimeters
    Printer.DrawWidth = 6
End Sub

Private Sub PrinterHeader(ByVal sWardID As String, ByVal sEmpNm As String)
    Dim strBase As String

    lngCurYPos = 0
    Printer.FontSize = 20: Printer.FontBold = True: Printer.FontUnderline = True
    Call Print_Setting("병동채혈리스트", 0, 10, Printer.ScaleWidth, "C", "C")
    Printer.FontSize = 9: Printer.FontBold = False: Printer.FontUnderline = False
    
    lngCurYPos = lngCurYPos + 10
    strBase = "채혈장소: " & sWardID & _
              "  작업일시 : " & Format(sWorkDt, "####-##-##") & "  " & _
                                Format(Mid(sWorkTm, 1, 4), "0#:##") & _
              "     채혈자 : " & Trim(sEmpNm)
    
    Call Print_Setting(strBase, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.ScaleWidth, lngCurYPos)
    
    Call Print_Setting("No", PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("Work No", 15, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("환자명", 45, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("환자ID", 65, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("S/A", 85, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("호실", 100, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("검체", 115, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting("검사종목", 135, LineSpace, Printer.ScaleWidth, "L", "C")
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.ScaleWidth, lngCurYPos)
    
    
End Sub
Private Sub PrintBody(ByVal No As String, ByVal AccNo As String, ByVal sPtNm As String, ByVal sPtid As String, _
                      ByVal sSexAge As String, ByVal sHosil As String, ByVal sSpcnm As String, ByVal sTestNm As String, _
                      ByVal sWardID As String, ByVal sEmpNm As String)

    If lngCurYPos >= Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call PrinterHeader(sWardID, sEmpNm)
    End If
    
    Call Print_Setting(No, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting(AccNo, 15, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting(sPtNm, 45, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting(sPtid, 65, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting(sSexAge, 85, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting(sHosil, 100, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting(sSpcnm, 115, LineSpace, Printer.ScaleWidth, "L", "C", False)
    Call Print_Setting(sTestNm, 135, LineSpace, Printer.ScaleWidth, "L", "C")


End Sub
