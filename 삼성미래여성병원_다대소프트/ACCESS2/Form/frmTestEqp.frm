VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmTestEqp 
   Caption         =   " 장비 VS 검사코드 설정"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16545
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   16545
   WindowState     =   2  '최대화
   Begin VB.Frame Frame7 
      BackColor       =   &H00F8E4D8&
      Height          =   2220
      Left            =   6150
      TabIndex        =   64
      Top             =   6630
      Width           =   9195
      Begin TabDlg.SSTab SSTab1 
         Height          =   2025
         Left            =   30
         TabIndex        =   65
         Top             =   120
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   3572
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         BackColor       =   16311512
         TabCaption(0)   =   "Cuttoff 설정"
         TabPicture(0)   =   "frmTestEqp.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame5"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "특정결과변환"
         TabPicture(1)   =   "frmTestEqp.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame5 
            BackColor       =   &H00F8E4D8&
            Height          =   2295
            Left            =   -74970
            TabIndex        =   70
            Top             =   270
            Width           =   9075
            Begin BHButton.BHImageButton cmdAdd_Cuttoff 
               Height          =   435
               Left            =   6120
               TabIndex        =   71
               Top             =   180
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   767
               Enabled         =   0   'False
               Caption         =   "저장"
               CaptionChecked  =   "BHImageButton1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ImgOutLineSize  =   3
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   2100
               Left            =   60
               TabIndex        =   72
               Top             =   150
               Width           =   6000
               _Version        =   524288
               _ExtentX        =   10583
               _ExtentY        =   3704
               _StockProps     =   64
               BackColorStyle  =   1
               ColsFrozen      =   1
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
               GrayAreaBackColor=   16777215
               MaxCols         =   2
               MaxRows         =   30
               ScrollBars      =   2
               ShadowColor     =   14737632
               SpreadDesigner  =   "frmTestEqp.frx":0038
               UserResize      =   1
               TextTip         =   2
            End
            Begin BHButton.BHImageButton cmdDel_Cuffoff 
               Height          =   435
               Left            =   6120
               TabIndex        =   73
               Top             =   660
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   767
               Enabled         =   0   'False
               Caption         =   "삭제"
               CaptionChecked  =   "BHImageButton1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ImgOutLineSize  =   3
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00F8E4D8&
            Height          =   1710
            Left            =   30
            TabIndex        =   66
            Top             =   270
            Width           =   9075
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   1440
               Left            =   60
               TabIndex        =   67
               Top             =   150
               Width           =   6000
               _Version        =   524288
               _ExtentX        =   10583
               _ExtentY        =   2540
               _StockProps     =   64
               BackColorStyle  =   1
               ColsFrozen      =   1
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
               GrayAreaBackColor=   16777215
               MaxCols         =   2
               MaxRows         =   30
               ScrollBars      =   2
               ShadowColor     =   14737632
               SpreadDesigner  =   "frmTestEqp.frx":05B4
               UserResize      =   1
               TextTip         =   2
            End
            Begin BHButton.BHImageButton cmdAdd_Result 
               Height          =   435
               Left            =   6120
               TabIndex        =   68
               Top             =   180
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   767
               Enabled         =   0   'False
               Caption         =   "저장"
               CaptionChecked  =   "BHImageButton1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ImgOutLineSize  =   3
            End
            Begin BHButton.BHImageButton cmdDel_Result 
               Height          =   435
               Left            =   6120
               TabIndex        =   69
               Top             =   660
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   767
               Enabled         =   0   'False
               Caption         =   "삭제"
               CaptionChecked  =   "BHImageButton1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ImgOutLineSize  =   3
            End
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F8E4D8&
      Caption         =   "참고치정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1725
      Left            =   6150
      TabIndex        =   20
      Top             =   4860
      Width           =   9195
      Begin VB.CheckBox chkDelta 
         BackColor       =   &H00F8E4D8&
         Caption         =   "Delta Check"
         Height          =   315
         Left            =   6570
         TabIndex        =   63
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtDelta 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   7980
         MaxLength       =   10
         TabIndex        =   62
         Text            =   "1234567890"
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtPRefH 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5220
         MaxLength       =   10
         TabIndex        =   61
         Text            =   "1234567890"
         Top             =   1320
         Width           =   1020
      End
      Begin VB.TextBox txtFRefH 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5220
         MaxLength       =   10
         TabIndex        =   60
         Text            =   "1234567890"
         Top             =   960
         Width           =   1020
      End
      Begin VB.TextBox txtMRefH 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5220
         MaxLength       =   10
         TabIndex        =   59
         Text            =   "1234567890"
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtPRefL 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   58
         Text            =   "1234567890"
         Top             =   1320
         Width           =   1020
      End
      Begin VB.TextBox txtFRefL 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   57
         Text            =   "1234567890"
         Top             =   960
         Width           =   1020
      End
      Begin VB.TextBox txtMRefL 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   56
         Text            =   "1234567890"
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtRefH 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5220
         MaxLength       =   10
         TabIndex        =   47
         Text            =   "1234567890"
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox txtRefL 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   46
         Text            =   "1234567890"
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Panic High value : "
         Height          =   180
         Left            =   3420
         TabIndex        =   55
         Top             =   1350
         Width           =   1605
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Panic Low value : "
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   1350
         Width           =   1590
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Female High value : "
         Height          =   180
         Left            =   3420
         TabIndex        =   53
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Female Low value : "
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   1020
         Width           =   1740
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Male High value : "
         Height          =   180
         Left            =   3420
         TabIndex        =   51
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Male Low value : "
         Height          =   180
         Left            =   240
         TabIndex        =   50
         Top             =   660
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Both High value : "
         Height          =   180
         Left            =   3420
         TabIndex        =   49
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "Both Low value : "
         Height          =   180
         Left            =   240
         TabIndex        =   48
         Top             =   300
         Width           =   1485
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F8E4D8&
      Caption         =   "결과속성"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2745
      Left            =   6150
      TabIndex        =   19
      Top             =   2070
      Width           =   9195
      Begin VB.Frame frmStrSet 
         BackColor       =   &H00F8E4D8&
         Caption         =   "[문자판정]"
         Height          =   2025
         Left            =   90
         TabIndex        =   27
         Top             =   660
         Width           =   9015
         Begin VB.TextBox txtRstMidStr 
            Appearance      =   0  '평면
            Height          =   285
            Left            =   1170
            TabIndex        =   75
            Top             =   1080
            Width           =   2385
         End
         Begin VB.ComboBox cboResult 
            Height          =   300
            Left            =   1170
            TabIndex        =   43
            Text            =   "Combo2"
            Top             =   1470
            Width           =   5205
         End
         Begin VB.TextBox txtRefLow 
            Appearance      =   0  '평면
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   270
            Width           =   735
         End
         Begin VB.TextBox txtRefLowStr 
            Appearance      =   0  '평면
            Height          =   285
            Left            =   3870
            TabIndex        =   30
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox txtRefHigh 
            Appearance      =   0  '평면
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox txtRefHighStr 
            Appearance      =   0  '평면
            Height          =   285
            Left            =   3870
            TabIndex        =   28
            Top             =   660
            Width           =   2385
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Left            =   1440
            TabIndex        =   31
            Top             =   630
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   661
            _StockProps     =   15
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption optRefHigh 
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Top             =   30
               Width           =   675
               _Version        =   65536
               _ExtentX        =   1191
               _ExtentY        =   556
               _StockProps     =   78
               Caption         =   "초과"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9.01
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption optRefHigh 
               Height          =   315
               Index           =   1
               Left            =   870
               TabIndex        =   33
               Top             =   30
               Width           =   675
               _Version        =   65536
               _ExtentX        =   1191
               _ExtentY        =   556
               _StockProps     =   78
               Caption         =   "이상"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9.01
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   -1  'True
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Left            =   1440
            TabIndex        =   35
            Top             =   240
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   661
            _StockProps     =   15
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption optRefLow 
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   30
               Width           =   675
               _Version        =   65536
               _ExtentX        =   1191
               _ExtentY        =   556
               _StockProps     =   78
               Caption         =   "미만"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9.01
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   -1  'True
            End
            Begin Threed.SSOption optRefLow 
               Height          =   315
               Index           =   1
               Left            =   870
               TabIndex        =   37
               Top             =   30
               Width           =   675
               _Version        =   65536
               _ExtentX        =   1191
               _ExtentY        =   556
               _StockProps     =   78
               Caption         =   "이하"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9.01
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Label Label23 
            BackColor       =   &H00F8E4D8&
            Caption         =   "중  간  값 :"
            Height          =   255
            Left            =   180
            TabIndex        =   74
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00F8E4D8&
            Caption         =   "결과 표기 :"
            Height          =   180
            Left            =   180
            TabIndex        =   42
            Top             =   1560
            Width           =   900
         End
         Begin VB.Label Label21 
            BackColor       =   &H00F8E4D8&
            Caption         =   "보다"
            Height          =   225
            Left            =   930
            TabIndex        =   41
            Top             =   330
            Width           =   405
         End
         Begin VB.Label Label20 
            BackColor       =   &H00F8E4D8&
            Caption         =   "이면"
            Height          =   255
            Left            =   3330
            TabIndex        =   40
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label19 
            BackColor       =   &H00F8E4D8&
            Caption         =   "보다"
            Height          =   225
            Left            =   930
            TabIndex        =   39
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label18 
            BackColor       =   &H00F8E4D8&
            Caption         =   "이면"
            Height          =   255
            Left            =   3330
            TabIndex        =   38
            Top             =   720
            Width           =   555
         End
      End
      Begin VB.TextBox txtUnit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   26
         Top             =   240
         Width           =   1365
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1170
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txtResultLen 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   24
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "결과 단위 :"
         Height          =   180
         Left            =   5850
         TabIndex        =   23
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "소  수  점 :"
         Height          =   180
         Left            =   3090
         TabIndex        =   22
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "결과 유형 :"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F8E4D8&
      Caption         =   "검사정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   6150
      TabIndex        =   8
      Top             =   570
      Width           =   9195
      Begin VB.CheckBox chkuse 
         BackColor       =   &H00F8E4D8&
         Caption         =   "사용안함"
         Height          =   240
         Left            =   6975
         TabIndex        =   76
         Top             =   1035
         Width           =   1455
      End
      Begin VB.TextBox txtOutSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5430
         MaxLength       =   3
         TabIndex        =   44
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox txtTestCd 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   18
         Text            =   "1234567890"
         Top             =   990
         Width           =   2520
      End
      Begin VB.TextBox txtTestNm 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1170
         MaxLength       =   40
         TabIndex        =   16
         Top             =   630
         Width           =   2520
      End
      Begin VB.TextBox txtTestCdEqp 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5430
         MaxLength       =   20
         TabIndex        =   15
         Text            =   "1234567890"
         Top             =   630
         Width           =   2505
      End
      Begin VB.TextBox txtEQPCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1170
         MaxLength       =   100
         TabIndex        =   12
         Text            =   "1234567890"
         Top             =   270
         Width           =   2520
      End
      Begin VB.TextBox txtEQPNM 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   11
         Text            =   "1234567890"
         Top             =   270
         Width           =   2520
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "정렬 순서 :"
         Height          =   180
         Left            =   4410
         TabIndex        =   45
         Top             =   1035
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "검사 코드 :"
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "검사 채널 :"
         Height          =   180
         Left            =   4410
         TabIndex        =   14
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "검  사  명 :"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "장  비  명 :"
         Height          =   180
         Left            =   4410
         TabIndex        =   10
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "장비 코드 :"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F8E4D8&
      Caption         =   "장비별검사항목"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8265
      Left            =   0
      TabIndex        =   5
      Top             =   570
      Width           =   6135
      Begin VB.OptionButton optOutSeq 
         BackColor       =   &H00F8E4D8&
         Caption         =   "정렬 순번"
         Height          =   180
         Left            =   2295
         TabIndex        =   79
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton optSortNm 
         BackColor       =   &H00F8E4D8&
         Caption         =   "검사명"
         Height          =   180
         Left            =   1170
         TabIndex        =   78
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
      Begin FPSpreadADO.fpSpread spdEqInfo 
         Height          =   7575
         Left            =   60
         TabIndex        =   6
         Top             =   630
         Width           =   6000
         _Version        =   524288
         _ExtentX        =   10583
         _ExtentY        =   13361
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   1
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
         GrayAreaBackColor=   16777215
         MaxCols         =   5
         MaxRows         =   30
         ScrollBars      =   2
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmTestEqp.frx":0B29
         UserResize      =   1
         TextTip         =   2
      End
      Begin BHButton.BHImageButton cmdSort 
         Height          =   420
         Left            =   4680
         TabIndex        =   80
         Top             =   180
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "재정렬 "
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00F8E4D8&
         Caption         =   "정렬 순서 :"
         Height          =   180
         Left            =   90
         TabIndex        =   77
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Frame fraCmdBar 
      BackColor       =   &H00F8E4D8&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   15
      TabIndex        =   0
      Top             =   9060
      Width           =   15360
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Delete"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   1575
         TabIndex        =   2
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Save"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   3015
         TabIndex        =   3
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   4455
         TabIndex        =   4
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":119D
            Key             =   "TST_E"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":1737
            Key             =   "TST_M"
         EndProperty
      EndProperty
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   16545
      _ExtentX        =   29184
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmTestEqp.frx":1CD1
      Caption         =   " Instruments Test Item Link ."
      SubCaption      =   "검사실 검사항목과 장비 검사항목을 연결 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   360
         Left            =   11520
         Top             =   75
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmTestEqp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const OBJTAG_EQP    As String = "EQP"
Private Const OBJTAG_TST    As String = "TST"
Private Const AUTO_VEFY     As String = "YES"
Private Const AUTO_VEFN     As String = "NO"

Private Const TLB_TEMP      As String = "TEMPTEABLE"
Private Const TLB_RESULT    As String = "INTERFACE003"

Private mAdoRs              As ADODB.Recordset
Private WithEvents PopUp_List As Listview
Attribute PopUp_List.VB_VarHelpID = -1

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
        Case 0: Call cmdPrint
        Case 1: Call cmdSave
        Case 2: Call cmdClear
        Case 3: Call cmdClose
        Case Else
    End Select
    
End Sub

Private Sub cmdPrint()
    Dim strSql As String
    
    Select Case MsgBox("선택 된 장비채널을 삭제 하겠습니까? ", vbYesNo Or vbExclamation Or vbDefaultButton1, "채널 삭제")
        Case vbYes
            strSql = ""
            strSql = strSql & vbLf & " DELETE FROM INTERFACE002"
            strSql = strSql & vbLf & "  WHERE EQP_CD = '" & Trim(txtEQPCD.Text) & "' AND TESTCD_EQP = '" & Trim(txtTestCdEqp.Text) & "' "
        
            AdoCn_Jet.Execute strSql
            
            Call cmdClear
            If optOutSeq.Value = True Then
                Call f_subSet_EqpData(INS_CODE, 1)
            Else
                Call f_subSet_EqpData(INS_CODE, 2)
            End If
        Case vbNo
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose()
    Unload Me
End Sub

Private Sub cmdClear()
    
    Call f_subClear_Form
       
    txtEQPCD.Text = INS_CODE
    txtEQPNM.Text = INS_NAME
    txtResultLen.Text = "9"
    
    chkuse.Value = 0
End Sub


Private Sub cmdSave()

    On Error GoTo frmTestEqp_Add_Error
    
    Dim tmpRS   As ADODB.Recordset
    Dim sqlDoc  As String
    Dim sqlRet  As Integer
    Dim itemX   As ListItem
    Dim optLow_Int As Integer
    Dim optHigh_Int As Integer
    
    If optRefLow(0).Value = True Then
        optLow_Int = 1
    End If
    
    If optRefLow(1).Value = True Then
        optLow_Int = 2
    End If
    
    If optRefHigh(0).Value = True Then
        optHigh_Int = 1
    End If
    
    If optRefHigh(1).Value = True Then
        optHigh_Int = 2
    End If
    
    sqlDoc = ""
    sqlDoc = sqlDoc & vbLf & " SELECT * FROM INTERFACE002"
    sqlDoc = sqlDoc & vbLf & "  WHERE EQP_CD = '" & Trim(txtEQPCD.Text) & "' AND TESTCD_EQP = '" & Trim(txtTestCdEqp.Text) & "' "
    
    Set tmpRS = New ADODB.Recordset
    
    tmpRS.CursorLocation = adUseClient
    tmpRS.Open sqlDoc, AdoCn_Jet
    
    With tmpRS
        If .RecordCount > 0 Then
            sqlDoc = ""
            sqlDoc = sqlDoc & vbLf & "Update INTERFACE002"
            sqlDoc = sqlDoc & vbLf & "   set TESTNM_EQP = '" & Trim(txtTestCdEqp.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "          OUT_SEQ = '" & Trim(txtOutSeq.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "           TESTCD = '" & Trim(txtTestCd.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "           TESTNM = '" & Trim(txtTestNm.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "       AUTOVERIFY = '',"
            sqlDoc = sqlDoc & vbLf & "           REMARK = '" & Trim(chkuse.Value) & "',"
            sqlDoc = sqlDoc & vbLf & "            DELTA = '" & Trim(txtDelta.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "         DELTAGBN = '" & Trim(chkDelta.Value) & "',"
            sqlDoc = sqlDoc & vbLf & "      RESULT_TYPE = '" & Trim(cboType.ListIndex) & "',"
            sqlDoc = sqlDoc & vbLf & "       RESULT_LOW = '" & Trim(txtRefLow.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "   RESULT_LOW_INT = '" & optLow_Int & "', "
            sqlDoc = sqlDoc & vbLf & "   RESULT_LOW_CHR = '" & Trim(txtRefLowStr.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "      RESULT_HIGH = '" & Trim(txtRefHigh.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "  RESULT_HIGH_INT = '" & optHigh_Int & "', "
            sqlDoc = sqlDoc & vbLf & "  RESULT_HIGH_CHR = '" & Trim(txtRefHighStr.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "  RESULT_MID_CHR  = '" & Trim(txtRstMidStr.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "            REFL  = '" & Trim(txtRefL.Text) & "', REFH       = '" & Trim(txtRefH.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "           MREFL  = '" & Trim(txtMRefL.Text) & "', MREFH     = '" & Trim(txtMRefH.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "           FREFL  = '" & Trim(txtFRefL.Text) & "', FREFH     = '" & Trim(txtFRefH.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "          PANICL  = '" & Trim(txtPRefL.Text) & "', PANICH    = '" & Trim(txtPRefH.Text) & "', "
            sqlDoc = sqlDoc & vbLf & "          EQP_NM  = '" & Trim(txtEQPNM.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "            UNIT  = '" & Trim(txtUnit.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "    ResultLength  = '" & Trim(txtResultLen.Text) & "',"
            sqlDoc = sqlDoc & vbLf & "      RESULT_DSP  = '" & Trim(cboResult.ListIndex) & "'"
            sqlDoc = sqlDoc & vbLf & " where EQP_CD     = '" & Trim(txtEQPCD.Text) & "'"
            sqlDoc = sqlDoc & vbLf & "   and TESTCD_EQP = '" & Trim(txtTestCdEqp.Text) & "' "

            AdoCn_Jet.Execute sqlDoc, sqlRet
        Else
            sqlDoc = ""
            sqlDoc = sqlDoc & vbLf & " Insert into INTERFACE002( "
            sqlDoc = sqlDoc & vbLf & "             EQP_CD, TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK, DELTA, DELTAGBN, "
            sqlDoc = sqlDoc & vbLf & "             RESULT_LOW, RESULT_LOW_INT, RESULT_LOW_CHR, RESULT_HIGH, RESULT_HIGH_INT, RESULT_HIGH_CHR,"
            sqlDoc = sqlDoc & vbLf & "             REFL , REFH, MREFL, MREFH, FREFL, FREFH, PANICL, PANICH, EQP_NM, UNIT, ResultLength, RESULT_DSP, RESULT_MID_CHR) Values ("
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtEQPCD.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtTestCdEqp.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtTestCdEqp.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtOutSeq.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtTestCd.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtTestNm.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(chkuse.Value) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtDelta.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(chkDelta.Value) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtRefLow.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(optLow_Int) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtRefLowStr.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtRefHigh.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(optHigh_Int) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtRefHighStr.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtRefL.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtRefH.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtMRefL.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtMRefH.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtFRefL.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtFRefH.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtPRefL.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtPRefH.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtEQPNM.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtUnit.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtResultLen.Text) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(cboResult.ListIndex) & "' ,"
            sqlDoc = sqlDoc & vbLf & " '" & Trim(txtRstMidStr.Text) & "') "
            
            AdoCn_Jet.Execute sqlDoc, sqlRet
        End If
    End With
        
    Set tmpRS = Nothing
    Call cmdClear
    If optOutSeq.Value = True Then
        Call f_subSet_EqpData(INS_CODE, 1)
    Else
        Call f_subSet_EqpData(INS_CODE, 2)
    End If
    
    Exit Sub
frmTestEqp_Add_Error:

    Call ErrMsgProc("frmTestEqp - Private Sub cmdSave()")

End Sub

Private Sub cmdSort_Click()
    If optOutSeq.Value = True Then
        Call f_subSet_EqpData(INS_CODE, 1)
    Else
        Call f_subSet_EqpData(INS_CODE, 2)
    End If
End Sub

Private Sub Form_Load()
    
    CaptionBar1.Caption = INS_NAME & " Instruments Test Item Link ."
    Call cmdClear
    
    cboType.Text = ""
    cboType.AddItem "숫자"
    cboType.AddItem "문자"
    
    cboType.ListIndex = 0
    
    cboResult.Text = ""
    cboResult.AddItem "문자"
    cboResult.AddItem "문자 숫자"
    cboResult.AddItem "문자(숫자)"
    cboResult.AddItem "숫자"
    cboResult.AddItem "숫자 문자"
    cboResult.AddItem "숫자(문자)"
    
    cboResult.ListIndex = 3
    
    If optOutSeq.Value = True Then
        Call f_subSet_EqpData(INS_CODE, 1)
    Else
        Call f_subSet_EqpData(INS_CODE, 2)
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If PopUp_List Is Nothing Then Set PopUp_List = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub

Private Sub f_subSet_EqpData(ByVal strEqp_Cd As String, ByVal intSort As String)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim intCnt  As Integer
       
    With spdEqInfo
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
    End With
    
    sqlDoc = "select *" & _
             "  from INTERFACE002" & _
             " where EQP_CD = '" & INS_CODE & "'"
                 
    If intSort = "1" Then
        sqlDoc = sqlDoc & " order by OUT_SEQ, TESTCD_EQP, TESTCD DESC"
    Else
        sqlDoc = sqlDoc & " order by TESTNM, TESTCD_EQP, TESTCD"
    End If
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    
    If adoRS.RecordCount > 0 Then
        spdEqInfo.maxrows = adoRS.RecordCount
        adoRS.MoveFirst
        With spdEqInfo
            For intCnt = 1 To adoRS.RecordCount
                .SetText 1, intCnt, adoRS.Fields("EQP_CD") & ""
                .SetText 2, intCnt, adoRS.Fields("EQP_NM") & ""
                .SetText 3, intCnt, adoRS.Fields("TESTCD_EQP") & ""
                .SetText 4, intCnt, adoRS.Fields("TESTNM") & ""
                .SetText 5, intCnt, adoRS.Fields("REMARK") & ""
                adoRS.MoveNext
            Next
        End With
    End If

    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subClear_Form()
        
    txtRefL = "":   txtRefH = ""
    txtMRefL = "":   txtMRefH = ""
    txtFRefL = "":   txtFRefH = ""
    txtPRefL = "":   txtPRefH = ""
    txtDelta = ""
    txtOutSeq = ""
    chkDelta.Value = 0
    txtEQPNM = ""
    txtTestCd = ""
    txtTestNm = ""
    txtTestCdEqp = ""
    txtResultLen = ""
    txtUnit = ""
    txtDelta = ""
    txtRefLowStr = ""
    txtRefHighStr = ""
    optRefHigh(1).Value = True
    optRefLow(0).Value = True
    txtRefLow = ""
    txtRefHigh = ""
    txtRstMidStr = ""
    
End Sub

Private Sub spdEqInfo_Click(ByVal Col As Long, ByVal Row As Long)
    Dim tmpRS As ADODB.Recordset
    Dim strSql As String
    Dim varTmp
    Dim strEqpCd As String
    Dim strTestEqpCd As String
    
    With spdEqInfo
        .GetText 1, .ActiveRow, varTmp: strEqpCd = Trim(varTmp)
        .GetText 3, .ActiveRow, varTmp: strTestEqpCd = Trim(varTmp)
    End With
    
    strSql = ""
    strSql = strSql & vbLf & " SELECT * FROM INTERFACE002"
    strSql = strSql & vbLf & "  WHERE EQP_CD = '" & strEqpCd & "' AND TESTCD_EQP = '" & strTestEqpCd & "' "
    
    Set tmpRS = New ADODB.Recordset
    
    tmpRS.CursorLocation = adUseClient
    tmpRS.Open strSql, AdoCn_Jet
    
    With tmpRS
        If tmpRS.RecordCount > 0 Then
            txtTestCd.Text = "" & .Fields("TESTCD")
            txtTestNm.Text = "" & .Fields("TESTNM")
            txtTestCdEqp.Text = "" & .Fields("TESTCD_EQP")
            txtOutSeq.Text = "" & .Fields("OUT_SEQ")
            cboType.ListIndex = Val("" & .Fields("RESULT_TYPE"))
            txtResultLen.Text = "" & .Fields("ResultLength")
            txtUnit.Text = "" & .Fields("UNIT")
            txtRefLow.Text = "" & .Fields("RESULT_LOW")
            
            Select Case .Fields("RESULT_LOW_INT")
                Case "1": optRefLow(0).Value = True
                Case "2": optRefLow(1).Value = True
            End Select
            
            txtRefLowStr.Text = "" & .Fields("RESULT_LOW_CHR")
            
            txtRefHigh.Text = "" & .Fields("RESULT_HIGH")
            
            Select Case .Fields("RESULT_HIGH_INT")
                Case "1": optRefHigh(0).Value = True
                Case "2": optRefHigh(1).Value = True
            End Select
            
            txtRefHighStr.Text = "" & .Fields("RESULT_HIGH_CHR")
            cboResult.ListIndex = Val("" & .Fields("RESULT_DSP"))
            txtRefL.Text = "" & .Fields("REFL")
            txtRefH.Text = "" & .Fields("REFH")
            txtMRefL.Text = "" & .Fields("MREFL")
            txtMRefH.Text = "" & .Fields("MREFH")
            txtFRefL.Text = "" & .Fields("FREFL")
            txtFRefH.Text = "" & .Fields("FREFH")
            txtPRefL.Text = "" & .Fields("PANICL")
            txtPRefH.Text = "" & .Fields("PANICH")
            txtRstMidStr.Text = "" & .Fields("RESULT_MID_CHR")
            
            If "" & .Fields("REMARK") = "" Then
                chkuse.Value = "0"
            Else
                chkuse.Value = "" & .Fields("REMARK")
            End If
            
            
            If "" & .Fields("DELTAGBN") = 1 Then
                chkDelta.Value = 1
            Else
                chkDelta.Value = 0
            End If
            txtDelta.Text = "" & .Fields("DELTA")
            .MoveNext
        End If
    End With
    
    Set tmpRS = Nothing
End Sub

Private Sub txtDelta_GotFocus()

    With txtDelta
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtDelta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub txtRefH_GotFocus()

    With txtRefH
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtRefH_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub


Private Sub txtRefL_GotFocus()

    With txtRefL
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtRefL_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub



