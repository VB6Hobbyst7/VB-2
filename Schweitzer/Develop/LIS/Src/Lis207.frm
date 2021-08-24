VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm207WBCDiffCnt 
   BackColor       =   &H00DBE6E6&
   Caption         =   "WBC Differential Count"
   ClientHeight    =   9195
   ClientLeft      =   285
   ClientTop       =   1980
   ClientWidth     =   14490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14490
   Tag             =   "20700"
   WindowState     =   2  '최대화
   Begin MSComctlLib.ListView lvwPatient 
      Height          =   555
      Left            =   75
      TabIndex        =   8
      Tag             =   "20113"
      Top             =   915
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   979
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15857140
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   4455
      Left            =   75
      TabIndex        =   2
      Top             =   1365
      Width           =   7245
      Begin VB.PictureBox picRst 
         Height          =   4230
         Left            =   60
         ScaleHeight     =   4170
         ScaleWidth      =   7065
         TabIndex        =   4
         Top             =   180
         Width           =   7125
         Begin FPSpread.vaSpread ssRst 
            CausesValidation=   0   'False
            Height          =   4155
            Left            =   -15
            TabIndex        =   61
            Tag             =   "20001"
            Top             =   15
            Width           =   7080
            _Version        =   196608
            _ExtentX        =   12488
            _ExtentY        =   7329
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
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
            GrayAreaBackColor=   15857140
            MaxCols         =   15
            MaxRows         =   10
            OperationMode   =   2
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "Lis207.frx":0000
            VisibleCols     =   10
            VisibleRows     =   10
            TextTip         =   2
         End
      End
   End
   Begin VB.Frame fraKey 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Index           =   0
      Left            =   75
      TabIndex        =   6
      Top             =   75
      Width           =   3015
      Begin VB.Frame fraAccNo 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   364
         Left            =   60
         TabIndex        =   7
         Top             =   270
         Width           =   2805
         Begin MSMask.MaskEdBox mskAccNo 
            Height          =   330
            Left            =   1095
            TabIndex        =   0
            Top             =   0
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "&&-######-#####"
            PromptChar      =   "_"
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   0
            Left            =   30
            TabIndex        =   128
            Top             =   0
            Width           =   1050
            _ExtentX        =   1852
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
            Caption         =   "접수번호"
            Appearance      =   0
         End
      End
   End
   Begin VB.Frame fraKeyMap 
      BackColor       =   &H00DBE6E6&
      Caption         =   "KeyMap"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   75
      TabIndex        =   1
      Top             =   5835
      Width           =   7245
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   0
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   165
         Width           =   1035
      End
      Begin VB.PictureBox picKeyBoard 
         BackColor       =   &H00DBE6E6&
         Height          =   2385
         Left            =   7350
         ScaleHeight     =   2325
         ScaleWidth      =   6705
         TabIndex        =   66
         Top             =   165
         Width           =   6765
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "SPACE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   46
            Left            =   1395
            TabIndex        =   113
            Tag             =   "32"
            Top             =   1815
            Width           =   3990
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   45
            Left            =   5220
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   112
            Tag             =   "47"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   44
            Left            =   4770
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   111
            Tag             =   "46"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   ","
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   43
            Left            =   4320
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   110
            Tag             =   "44"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   42
            Left            =   3870
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   109
            Tag             =   "77"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   41
            Left            =   3420
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   108
            Tag             =   "78"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   40
            Left            =   2970
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   107
            Tag             =   "66"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   39
            Left            =   2520
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   106
            Tag             =   "86"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   38
            Left            =   2070
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   105
            Tag             =   "67"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   37
            Left            =   1620
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   104
            Tag             =   "88"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   36
            Left            =   1170
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   103
            Tag             =   "90"
            Top             =   1440
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "'"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   35
            Left            =   5445
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   102
            Tag             =   "39"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   ";"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   34
            Left            =   4995
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   101
            Tag             =   "59"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   33
            Left            =   4545
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   100
            Tag             =   "76"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "K"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   32
            Left            =   4095
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   99
            Tag             =   "75"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "J"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   31
            Left            =   3645
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   98
            Tag             =   "74"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "H"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   30
            Left            =   3195
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   97
            Tag             =   "72"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "G"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   29
            Left            =   2745
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   96
            Tag             =   "71"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   28
            Left            =   2295
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   95
            Tag             =   "70"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   27
            Left            =   1845
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   94
            Tag             =   "68"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   26
            Left            =   1395
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   93
            Tag             =   "83"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   25
            Left            =   945
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   92
            Tag             =   "65"
            Top             =   1065
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   24
            Left            =   5670
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   91
            Tag             =   "93"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&["
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   23
            Left            =   5220
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   90
            Tag             =   "91"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   22
            Left            =   4770
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   89
            Tag             =   "80"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   21
            Left            =   4320
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   88
            Tag             =   "79"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   20
            Left            =   3870
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   87
            Tag             =   "73"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   19
            Left            =   3420
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   86
            Tag             =   "85"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   18
            Left            =   2970
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   85
            Tag             =   "89"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "T"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   2520
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   84
            Tag             =   "84"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   2070
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   83
            Tag             =   "82"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "E"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   1620
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   82
            Tag             =   "69"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "W"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   1170
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   81
            Tag             =   "87"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Q"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   720
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   80
            Tag             =   "81"
            Top             =   690
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "\"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   5895
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   79
            Tag             =   "92"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   5445
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   78
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   4995
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   77
            Tag             =   "45"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   4545
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   76
            Tag             =   "48"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   4095
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   75
            Tag             =   "57"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   3645
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   74
            Tag             =   "56"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   3195
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   73
            Tag             =   "55"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   2745
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   72
            Tag             =   "54"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2295
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   71
            Tag             =   "53"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1845
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   70
            Tag             =   "52"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1395
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   69
            Tag             =   "51"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   945
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   68
            Tag             =   "50"
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   495
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   67
            Tag             =   "49"
            Top             =   315
            Width           =   390
         End
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   23
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   60
         Top             =   2265
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   22
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   58
         Top             =   1950
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   21
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   56
         Top             =   1665
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   20
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   54
         Top             =   1350
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   19
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   52
         Top             =   1065
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   18
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   750
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   17
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   48
         Top             =   465
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   16
         Left            =   6045
         Style           =   1  '그래픽
         TabIndex        =   46
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   15
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   44
         Top             =   2280
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   14
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   42
         Top             =   1965
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   13
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   40
         Top             =   1680
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   12
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   38
         Top             =   1365
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   11
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   36
         Top             =   1080
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   10
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   34
         Top             =   765
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   9
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   480
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   8
         Left            =   3765
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   165
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   7
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   2280
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   6
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   1965
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   5
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   1680
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   4
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   1365
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   3
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   1080
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   315
         Index           =   2
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   765
         Width           =   1035
      End
      Begin VB.CommandButton cmdTestKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "F"
         Height          =   285
         Index           =   1
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   480
         Width           =   1035
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   780
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1380
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   1980
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   8
         Left            =   2520
         TabIndex        =   29
         Top             =   180
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   9
         Left            =   2520
         TabIndex        =   31
         Top             =   480
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   10
         Left            =   2520
         TabIndex        =   33
         Top             =   780
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   11
         Left            =   2520
         TabIndex        =   35
         Top             =   1080
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   12
         Left            =   2520
         TabIndex        =   37
         Top             =   1380
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   13
         Left            =   2520
         TabIndex        =   39
         Top             =   1680
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   14
         Left            =   2520
         TabIndex        =   41
         Top             =   1980
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   15
         Left            =   2520
         TabIndex        =   43
         Top             =   2280
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   16
         Left            =   4815
         TabIndex        =   45
         Top             =   165
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   17
         Left            =   4815
         TabIndex        =   47
         Top             =   465
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   18
         Left            =   4815
         TabIndex        =   49
         Top             =   765
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   19
         Left            =   4815
         TabIndex        =   51
         Top             =   1065
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   20
         Left            =   4815
         TabIndex        =   53
         Top             =   1365
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   21
         Left            =   4815
         TabIndex        =   55
         Top             =   1665
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   22
         Left            =   4815
         TabIndex        =   57
         Top             =   1965
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   23
         Left            =   4815
         TabIndex        =   59
         Top             =   2265
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
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
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel lblDisease 
      Height          =   270
      Left            =   7575
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   350
      Width           =   6850
      _ExtentX        =   12091
      _ExtentY        =   476
      BackColor       =   16777215
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
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.CommandButton cmdRmk 
      BackColor       =   &H008080FF&
      Caption         =   "처방비고"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11835
      Style           =   1  '그래픽
      TabIndex        =   137
      Top             =   45
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdMicro 
      BackColor       =   &H00DBE6E6&
      Caption         =   "미생물"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13575
      Style           =   1  '그래픽
      TabIndex        =   138
      Top             =   45
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H00DBE6E6&
      Caption         =   "특  수"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12735
      Style           =   1  '그래픽
      TabIndex        =   139
      Top             =   45
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Comment by Accession No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7320
      TabIndex        =   114
      Tag             =   "20003"
      Top             =   5805
      Width           =   7125
      Begin VB.CommandButton cmdCommentTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6660
         Picture         =   "Lis207.frx":07B1
         Style           =   1  '그래픽
         TabIndex        =   116
         Top             =   1680
         Width           =   315
      End
      Begin VB.CommandButton cmdRemarkTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6660
         Picture         =   "Lis207.frx":0CE3
         Style           =   1  '그래픽
         TabIndex        =   115
         Top             =   2205
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   1725
         Left            =   90
         TabIndex        =   117
         Top             =   270
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   3043
         _Version        =   393217
         BackColor       =   15857140
         ScrollBars      =   2
         TextRTF         =   $"Lis207.frx":1215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   360
         Left            =   90
         TabIndex        =   118
         Top             =   2220
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   635
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"Lis207.frx":1444
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCapRemark 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   119
         Top             =   1980
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdSetup 
      BackColor       =   &H00FCEFE9&
      Height          =   585
      Left            =   3660
      Picture         =   "Lis207.frx":1678
      Style           =   1  '그래픽
      TabIndex        =   65
      Tag             =   "20714"
      Top             =   195
      Width           =   705
   End
   Begin VB.ComboBox cboRelTest 
      BackColor       =   &H00FFF9F7&
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
      Left            =   7575
      Style           =   2  '드롭다운 목록
      TabIndex        =   64
      Top             =   615
      Width           =   6885
   End
   Begin VB.CommandButton cmdPopupList 
      BackColor       =   &H00DEDBDD&
      Caption         =   "찾기"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3075
      MousePointer    =   14  '화살표와 물음표
      Picture         =   "Lis207.frx":2242
      Style           =   1  '그래픽
      TabIndex        =   62
      Top             =   195
      Width           =   570
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
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
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
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
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "확인(&S)"
      CausesValidation=   0   'False
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
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   4455
      Left            =   7320
      TabIndex        =   3
      Top             =   1365
      Width           =   7140
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3750
         TabIndex        =   126
         Top             =   4065
         Width           =   1110
      End
      Begin VB.TextBox txtMaxCount 
         Height          =   360
         Left            =   6030
         TabIndex        =   125
         Top             =   4065
         Width           =   1050
      End
      Begin VB.PictureBox picdiff 
         Height          =   3870
         Left            =   45
         ScaleHeight     =   3810
         ScaleWidth      =   7005
         TabIndex        =   123
         Top             =   195
         Width           =   7065
         Begin MSComctlLib.ProgressBar prgRst 
            Height          =   240
            Left            =   0
            TabIndex        =   127
            ToolTipText     =   "자료를 가져오고 있읍니다."
            Top             =   3570
            Visible         =   0   'False
            Width           =   6810
            _ExtentX        =   12012
            _ExtentY        =   423
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin FPSpread.vaSpread tblDiff 
            CausesValidation=   0   'False
            Height          =   3795
            Left            =   0
            TabIndex        =   124
            Tag             =   "20001"
            Top             =   15
            Width           =   7005
            _Version        =   196608
            _ExtentX        =   12356
            _ExtentY        =   6694
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
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
            GrayAreaBackColor=   15857140
            MaxCols         =   19
            MaxRows         =   10
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "Lis207.frx":27CC
            VisibleCols     =   10
            VisibleRows     =   10
            TextTip         =   1
         End
      End
      Begin VB.TextBox txtCorrectWBC 
         BackColor       =   &H00F1F5F4&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1575
         TabIndex        =   5
         Top             =   4065
         Width           =   915
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   4
         Left            =   60
         TabIndex        =   132
         Top             =   4065
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
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
         Caption         =   "Corrected WBC"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   5
         Left            =   2550
         TabIndex        =   133
         Top             =   4065
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
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
         Caption         =   "Input count"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   6
         Left            =   4890
         TabIndex        =   134
         Top             =   4065
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
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
         Caption         =   "Max count"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraText 
      BackColor       =   &H00DBE6E6&
      Caption         =   " Text Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   7320
      TabIndex        =   120
      Tag             =   "20002"
      Top             =   7395
      Visible         =   0   'False
      Width           =   7140
      Begin VB.CommandButton cmdTextTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6645
         Picture         =   "Lis207.frx":2FBC
         Style           =   1  '그래픽
         TabIndex        =   121
         Top             =   300
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   675
         Left            =   60
         TabIndex        =   122
         Top             =   300
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   1191
         _Version        =   393217
         BackColor       =   15663102
         Enabled         =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis207.frx":34EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   1
      Left            =   6225
      TabIndex        =   129
      Top             =   45
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
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
      Caption         =   "연 락 처"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   2
      Left            =   6225
      TabIndex        =   130
      Top             =   330
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
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
      Caption         =   "상 병 명"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   3
      Left            =   6225
      TabIndex        =   131
      Top             =   615
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
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
      Caption         =   "관련검사 결과"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblTelno 
      Height          =   270
      Left            =   7575
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   45
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   476
      BackColor       =   16777215
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
      Caption         =   ""
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2415
      Left            =   75
      TabIndex        =   63
      Top             =   660
      Visible         =   0   'False
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "환자ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "환자명"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "접수번호"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "접수일시"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame fraMesg 
      BackColor       =   &H00DBE6E6&
      Height          =   2655
      Left            =   10350
      TabIndex        =   140
      Top             =   525
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtMesg 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   142
         Top             =   390
         Width           =   4050
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "확인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2940
         Style           =   1  '그래픽
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1095
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   2
         Left            =   15
         TabIndex        =   143
         Top             =   90
         Width           =   4050
         _ExtentX        =   7144
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
         Caption         =   "처방 비고사항 조회"
         Appearance      =   0
         LeftGab         =   200
      End
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDF0F5&
      BackStyle       =   0  '투명
      Caption         =   "오류가 발생했습니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00313D46&
      Height          =   180
      Left            =   255
      TabIndex        =   12
      Top             =   8685
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   8610
      Width           =   9870
   End
End
Attribute VB_Name = "frm207WBCDiffCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명 : frm207WBCDiffCnt.frm
'   작성자 : 김정규
'   내  용 : WBC Diff Count폼
'   작성일 :
'   버  전 :
'-----------------------------------------------------------------------------'

Option Explicit

Public WithEvents clsTemplete   As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private objPtInfo       As clsPatientInfo
Private ClearFg         As Boolean
Private strCombo        As String
Private blnFirst        As Boolean
Private gintTemplete    As Integer
Private strPreTestcd    As String
Private strWBCResult    As String

Private objKey          As clsDictionary
Private objWBC          As clsDictionary    'WBC검사코드 Dictionary
Private objNRBC         As clsDictionary    'NRBC검사코드 Dictionary

Private mCurWBCCode     As String           '현재 접수번호의 WBC 검사코드
Private mCurNRBCCode    As String           '현재 접수번호의 NRBC 검사코드

Private Sub Form_Load()
    Dim ii As Integer
    
    '## WBC, NRBC 검사코드 로딩
    Call GetWBCInfo
    DoEvents
    
    blnFirst = False
    For ii = cmdKey.LBound To cmdKey.UBound
        cmdKey(ii).FontBold = False
    Next
    DoEvents
    
    For ii = lblTestNm.LBound To lblTestNm.UBound
        lblTestNm(ii).Caption = ""
        cmdTestKey(ii).Caption = ""
    Next
    Call ClearData
    DoEvents
    
    Set objKey = New clsDictionary
    
    '## 검사코드, 검사명,Key,ASCII,,키순번
    objKey.Clear
    objKey.FieldInialize "testcd", "testnm,chr,asc,cnt,seq"
    
    Call DiffKeySetting
    DoEvents
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
    
    If blnFirst = False Then
        Call LoadLvwHead
        blnFirst = True
        ClearData
    End If
   
   If mskAccNo.Enabled Then mskAccNo.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim lngASC      As Integer
    
    If ActiveControl.Name = mskAccNo.Name Then Exit Sub
    If ActiveControl.Name = txtMaxCount.Name Then Exit Sub
    
On Error Resume Next
    If KeyAscii > 96 And KeyAscii < 123 Then
        lngASC = KeyAscii - 32
    Else
        lngASC = KeyAscii
    End If
    
    '스페이스 입력시 바로 이전값 롤백
    If lngASC = 32 Then
        objKey.KeyChange strPreTestcd
        If Val(objKey.Fields("cnt")) > 0 Then
            objKey.Fields("cnt") = Val(objKey.Fields("cnt")) - 1
            If Not objNRBC.Exists(strPreTestcd) Then
                txtInput.Text = Val(txtInput.Text) - 1
                txtInput.Text = IIf(txtInput.Text = "0", "", txtInput.Text)
            End If
            Call DiffResult(strPreTestcd, objKey.Fields("cnt"))
        End If
        Exit Sub
    End If
    strPreTestcd = ""
    
    '입력된 값에 해당되는 검사코드 찾기
    objKey.MoveFirst
    Do Until objKey.EOF
        If Val(objKey.Fields("asc")) = lngASC Then
            strPreTestcd = objKey.Fields("testcd")
            Exit Do
        End If
        objKey.MoveNext
    Loop
    
    If strPreTestcd = "" Then Exit Sub
    
    'MaxCount 체크
    '카운트 체크(nRBC는 카운트하지 않는다)
    If Not objNRBC.Exists(strPreTestcd) Then
        If Val(txtInput.Text) + 1 > Val(txtMaxCount.Text) Then
            MsgBox "MaxCout값이 초과되었습니다.", vbInformation + vbOKOnly, "Info"
            Exit Sub
        End If
        
        txtInput.Text = Val(txtInput.Text) + 1
        txtInput.Text = IIf(txtInput.Text = "0", "", txtInput.Text)
    End If
    
    '개수 업데이트 및..화면 display
    objKey.KeyChange strPreTestcd
    objKey.Fields("cnt") = Val(objKey.Fields("cnt")) + 1
    Call DiffResult(strPreTestcd, objKey.Fields("cnt"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsTemplete = Nothing
    Set objCodeList = Nothing
    Set objWBC = Nothing
    Set objNRBC = Nothing
    Set objKey = Nothing
    Set objPtInfo = Nothing
    Call ICSPatientMark
End Sub

Private Sub cmdPopupList_Click()
    Dim RS          As Recordset
    Dim iTmx        As ListItem
    Dim strTestCds  As String       'Diff 검사코드들
    Dim SSQL        As String

    lvw.ListItems.Clear
    If objKey.BOF Then
        MsgBox "WBC Diff 검사코드 설정이 되지 않았습니다.", vbInformation + vbOKOnly, "정보"
        Exit Sub
    End If
    
    '## 5.1.0: 이상대(2005-01-03)
    '   - Diff검사코드가 꼭 상세대표나, 그룹코드가 아닐수도 있기때문에 Diff에 포함된 코드로
    '     IN 을 거는 방식으로 쿼리수정
    '   - Workarea는 Const에 등록된 변수를 사용하는 방식으로 변경
    With objKey
        .MoveFirst
        Do Until .EOF
            strTestCds = strTestCds & "'" & .Fields("testcd") & "',"
            .MoveNext
        Loop
        strTestCds = Mid$(strTestCds, 1, Len(strTestCds) - 1)
    End With
    
    SSQL = " SELECT distinct a.workarea,a.accdt,a.accseq ,a.stscd,a.ptid,c." & F_PTNM & " as ptnm,a.rcvdt,a.rcvtm " & _
           " FROM " & T_HIS001 & " c," & T_LAB201 & " a," & T_LAB302 & " b" & _
           " WHERE " & DBW("a.rcvdt>=", Format(DateAdd("d", -7, GetSystemDate), "YYYYMMDD")) & _
           " AND  " & DBW("a.workarea=", CBC_WorkArea) & _
           " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_Accession) & _
           " AND " & DBW("a.stscd<", enStsCd.StsCd_LIS_FinRst) & _
           " AND a.workarea=b.workarea" & _
           " AND a.accdt=b.accdt" & _
           " AND a.accseq=b.accseq" & _
           " AND b.testcd IN (" & strTestCds & ")" & _
           " AND a.ptid=c." & F_PTID & _
           " AND (b.vfydt =' ' or b.vfydt is null)"
    
On Error GoTo Errors
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Do Until RS.EOF
            Set iTmx = lvw.ListItems.Add(, , RS.Fields("ptid").Value & "")
            iTmx.SubItems(1) = RS.Fields("ptnm").Value & ""
            iTmx.SubItems(2) = RS.Fields("workarea").Value & "" & "-" & _
                               RS.Fields("accdt").Value & "" & "-" & _
                               RS.Fields("accseq").Value & ""
            iTmx.SubItems(3) = Format(RS.Fields("rcvdt").Value & "", "####-0#-0#") & " " & _
                               Format(RS.Fields("rcvtm").Value & "", "0#:0#:0#")
            RS.MoveNext
        Loop
        lvw.Visible = True
        lvw.ZOrder 0
    Else
        MsgBox "검사대상이 없습니다.", vbInformation, "정보"
    End If
    RS.Close
    Set RS = Nothing
    Exit Sub
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub

Private Sub cmdSetup_Click()
    If fraKeyMap.Width = 14310 Then
        fraKeyMap.Width = 7245
    Else
        fraKeyMap.Width = 14310
        fraKeyMap.ZOrder 0
    End If
End Sub

Private Sub cmdSave_Click()
    Dim strMsg          As String
    Dim blnDBSuccess    As Boolean
    
    If txtMaxCount.Text <> txtInput.Text Then
        strMsg = MsgBox("Max Count 와 Input Count가 일치하지 않습니다." & vbCRLF & " 진행하시겠습니까?", vbInformation + vbYesNo, "Info")
        If strMsg = vbNo Then Exit Sub
    End If

    If strWBCResult <> "" Then
        If Val(strWBCResult) <> Val(txtCorrectWBC.Text) Then
            If Not SaveCorrectWBC Then
                MsgBox "Correct WBC Count결과저장시 오류발생입니다.", vbInformation + vbOKOnly, "Info"
                Exit Sub
            End If
        End If
    End If
    
    Me.MousePointer = 11
    With objPtInfo
        .FootNote = rtfComment.Text
        blnDBSuccess = .DataEntry
    End With
    Me.MousePointer = 0
    If blnDBSuccess = False Then
        MsgBox objPtInfo.ErrNo & _
                " - " & objPtInfo.ErrText, vbCritical + vbOKOnly, "Info"
        Exit Sub
    Else
        Call ClearData
    End If
    
    ssRst.MaxRows = 0
    tblDiff.MaxRows = 0
    lvwPatient.ListItems.Clear
    rtfText.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""
    If mskAccNo.Enabled Then mskAccNo.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call ClearData
    If mskAccNo.Enabled Then mskAccNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdTestKey_Click(Index As Integer)
    Dim strTmp  As String
    Dim lngASC  As Integer
    Dim strChr  As String
    Dim strPChr As String
    Dim strMsg  As String
    Dim strIChr As String
    Dim jj      As Integer
    
    strTmp = InputBox(lblTestNm(Index).Caption & " 의 KeyMap을 선택하세요.", "검사항목별 KeyMap 설정", cmdTestKey(Index).Caption)
    
    
    If Len(strTmp) > 1 Then
        MsgBox "KeyMap은 한자리만 가능합니다.", vbInformation, "Info"
        Exit Sub
    End If
    
    If strTmp = "" Then
        strMsg = MsgBox("Mapping Key를 삭제하시겠습까?", vbInformation + vbYesNo, "Info")
        If strMsg = vbYes Then
            For jj = cmdKey.LBound To cmdKey.UBound
                If cmdKey(jj).Caption = cmdTestKey(Index).Caption Then
                    If DiffSaveSQL(lblTestNm(Index).Tag, lblTestNm(Index).Caption, strChr, CStr(lngASC)) Then
                        cmdKey(jj).FontBold = False
                        cmdKey(jj).FontSize = 9
                        objKey.KeyChange lblTestNm(Index).Tag
                        objKey.Fields("chr") = ""
                        objKey.Fields("asc") = ""
                        cmdTestKey(Index).Caption = "KEY"
                        cmdTestKey(Index).Tag = ""
                    End If
                End If
            Next
            Exit Sub
        Else
            Exit Sub
        End If
    Else
        If Asc(strTmp) > 96 And Asc(strTmp) < 123 Then
            strIChr = Chr(Asc(strTmp) - 32)
        Else
            strIChr = Chr(Asc(strTmp))
        End If
        If strIChr <> cmdTestKey(Index).Caption Then
            objKey.MoveFirst
            Do Until objKey.EOF
                If objKey.Fields("chr") = UCase(strTmp) Then
                    MsgBox objKey.Fields("testnm") & " 에서 사용되는 Key입니다.", vbInformation, "Info"
                    Exit Sub
                End If
                objKey.MoveNext
            Loop
        End If
    End If
    
    lngASC = Asc(strTmp)
    If lngASC > 96 And lngASC < 123 Then lngASC = lngASC - 32
    strChr = Chr(lngASC)
    
    If DiffSaveSQL(lblTestNm(Index).Tag, lblTestNm(Index).Caption, strChr, CStr(lngASC)) Then
        strPChr = cmdTestKey(Index).Caption
        objKey.KeyChange lblTestNm(Index).Tag
        objKey.Fields("chr") = strChr
        objKey.Fields("asc") = lngASC
        cmdTestKey(Index).Caption = strChr
        cmdTestKey(Index).Tag = lngASC
        For jj = cmdKey.LBound To cmdKey.UBound
            If cmdKey(jj).Caption = cmdTestKey(Index).Caption Then
                cmdKey(jj).FontBold = True
                cmdKey(jj).FontSize = 11
            ElseIf cmdKey(jj).Caption = strPChr Then
                cmdKey(jj).FontBold = False
                cmdKey(jj).FontSize = 9
            End If
        Next
    End If
    
End Sub

Private Sub cmdCommentTemplete_Click()
    If ssRst.MaxRows < 1 Then Exit Sub
    Call CallTemplete(3, 0)
End Sub

Private Sub cmdTextTemplete_Click()
    If rtfText.Enabled = False Then Exit Sub
    Call CallTemplete(2, 0)
End Sub

Private Sub cmdRmk_Click()
    Dim objSQL   As clsLISResultReview
    Dim RS       As Recordset
    Dim aryTmp() As String
    Dim strTmp   As String
    Dim SSQL     As String
    Dim ii       As Integer
    
    txtMesg.Text = ""
    Set objSQL = New clsLISResultReview
    SSQL = objSQL.GetOrderRemark(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        strTmp = "처방일자 : " & Format(RS.Fields("orddt").Value & "", "####-##-##") & vbCRLF
        strTmp = strTmp & "처방번호 : " & RS.Fields("ordno").Value & "" & vbCRLF
        strTmp = strTmp & "처방비고  " & vbCRLF
        txtMesg.Text = strTmp
        strTmp = ""
        aryTmp = Split(RS.Fields("mesg").Value & "", vbCRLF)
        For ii = LBound(aryTmp) To UBound(aryTmp)
            strTmp = strTmp & " " & aryTmp(ii) & vbCRLF
        Next
        txtMesg.Text = txtMesg.Text & strTmp
        fraMesg.Visible = True
    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub

Private Sub cmdOk_Click()
    fraMesg.Visible = False
    ssRst.SetFocus
End Sub

Private Sub mskAccNo_KeyPress(KeyAscii As Integer)
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyReturn Then lvwPatient.SetFocus
End Sub

Private Sub mskAccNo_LostFocus()
    If Trim(mskAccNo.ClipText) = "" Then
        lblErr.Caption = ""
        Exit Sub
    End If
    Call Data_Load
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과값 체크, 화면표시
'-----------------------------------------------------------------------------'
Private Sub DiffResult(ByVal sTestcd As String, ByVal lngCnt As Integer)
    Dim strTmp      As String
    Dim lngWBC      As Integer
    Dim lngRBC      As Integer
    Dim lngCount    As Integer
    
    Dim ii          As Long
    Dim jj          As Long
    
    With tblDiff
        .Col = 2: .COL2 = 2
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .BackColor = vbWhite
        .BlockMode = False
        
        '## 5.1.0: 이상대(2005-01-03)
        '   - Diff결과값은 모두 %이기 때문에 정수로 반올림함!
        For ii = 1 To objPtInfo.Result.Count
            .Row = ii
            
            'input값이 일치할때..
            If objPtInfo.Result.Item(ii).TestCd = sTestcd Then
                .Col = 19: .Value = lngCnt
                .Col = 2
                'nRBC가 아닐때(%계산, rstcd값 삽입)
                If Not objNRBC.Exists(sTestcd) Then
                    If Val(txtInput.Text) = 0 Then
                        .Value = 0
                    Else
                        .Value = Round((Val(lngCnt) / Val(txtInput.Text)) * 100)
                    End If
                    .Value = IIf(.Value = "0", "", .Value)
                    .ForeColor = DCM_Red
                    .BackColor = DCM_Yellow
                Else
                    .ForeColor = DCM_Red
                    .BackColor = DCM_Yellow
                    .Value = lngCnt
                End If
            Else
                'Input값이 검사코드와 다를경우 % 계산
                .Col = 19: lngCount = Val(.Value)
                .Col = 2
                If Not objNRBC.Exists(objPtInfo.Result.Item(ii).TestCd) Then
                    If Val(txtInput.Text) = 0 Then
                        .Value = "0"
                    Else
                        .Value = Round((Val(lngCount) / Val(txtInput.Text)) * 100)
                    End If
                Else
                    .Value = IIf(lngCount = "0", "", lngCount)
                End If
                .Value = IIf(.Value = "0", "", .Value)
            End If
            
            .Col = 2:
            .Value = IIf(.Value = "0", "", .Value)
            objPtInfo.Result.Item(ii).RstCd = .Value
            
            'High/Low,Delta/Panic Check를 위해서 Col=18에 결과치를 넣어준다.
            'High/Low,Delta/Panic Check
            .Col = 18: .Value = objPtInfo.Result.Item(ii).RstCd
            Call ResultCheck(ii)
        Next
        
        '## CorrectedWBC 계산
        '## NRBC값이 10미만이면 CorrectedWBC를 계산하지 않음
        objKey.KeyChange mCurNRBCCode
        If Val(objKey.Fields("cnt")) > 11 Then
            lngRBC = Val(txtInput.Text) + Val(objKey.Fields("cnt"))
            lngWBC = (Val(strWBCResult) * Val(txtInput.Text)) / lngRBC
            txtCorrectWBC.Text = IIf(lngWBC = 0, "", lngWBC)
        End If
        
        'FootNote처리
        strTmp = "▶ WBC Differential Count Information" & vbCRLF & _
                 "1.Max Count : " & txtMaxCount.Text & vbCRLF & _
                 "2.Input Count : " & txtInput.Text & vbCRLF & _
                 "3.NRBC Count : " & IIf(objKey.Fields("cnt") = "", "0", objKey.Fields("cnt")) & "/" & txtInput.Text & vbCRLF & _
                 "4.Before WBC Count : " & strWBCResult
        
        If objPtInfo.FootNote <> "" Then
            rtfComment.Text = objPtInfo.FootNote & vbCr & strTmp
        Else
            rtfComment.Text = strTmp
        End If
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과값 H/L, D/P Check
'-----------------------------------------------------------------------------'
Private Sub ResultCheck(ByVal lngRow As Long)
    Dim strRstType  As String
    Dim strErr      As String
    
On Error GoTo ErrLevaeCell:
    Call objPtInfo.ResultCheck(lngRow)
    strRstType = objPtInfo.Result.Item(lngRow).RstType
    If strRstType = "N" Then
        strErr = objPtInfo.Result.Item(lngRow).AvalVal
        If objPtInfo.IsAvalVal(lngRow) = False Then
            If strErr <> "0" Then
                strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(lngRow).AvalVal & "자리)"
            Else
                strErr = "유효숫자 입력 오류. (정수형만 입력)"
            End If
            GoTo ErrLevaeCell
        Else
            lblErr.Caption = ""
            objPtInfo.NumValCheck (lngRow)
        End If
    ElseIf strRstType = "A" Then
        If objPtInfo.IsAlphaCd(lngRow) = False Then
            strErr = "결과 입력 오류!"
            GoTo ErrLevaeCell
        Else
            lblErr.Caption = ""
        End If
    ElseIf strRstType = "R" Then
        If objPtInfo.IsRateCd(lngRow) = False Then
            strErr = "비율결과 입력 오류!"
            GoTo ErrLevaeCell
        Else
            lblErr.Caption = ""
        End If
    ElseIf strRstType = "F" Then
        If objPtInfo.IsFreeResult(lngRow) = False Then
            strErr = "FREE결과 입력 오류! (10자리이내)"
            GoTo ErrLevaeCell
        Else
            objPtInfo.NumValCheck (lngRow)
            lblErr.Caption = ""
        End If
    End If
    Exit Sub

ErrLevaeCell:
    lblErr.Caption = strErr
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    DoEvents
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Diff Count 검사코드(상세자코드)와 Mapping된 검사명을 표시
'-----------------------------------------------------------------------------'
Private Sub DiffKeySetting()
    Dim RS      As Recordset
    Dim strTmp  As String
    Dim ii      As Integer
    Dim jj      As Integer
    
    For ii = lblTestNm.LBound To lblTestNm.UBound
        lblTestNm(ii).Visible = False
        cmdTestKey(ii).Visible = False
    Next
    
On Error GoTo Errors
    Set RS = New Recordset
    RS.Open GetDiffCodeSQL, DBConn
    If Not RS.EOF Then
        ii = 0
        Do Until RS.EOF
            lblTestNm(ii).Caption = RS.Fields("testnm").Value & ""
            lblTestNm(ii).Tag = RS.Fields("testcd").Value & ""
            
            If Not objKey.Exists(RS.Fields("testcd").Value & "") Then
                objKey.AddNew RS.Fields("testcd").Value & "", RS.Fields("testnm").Value & _
                    "" & COL_DIV & "" & COL_DIV & "" & COL_DIV & "" & COL_DIV & ii
            End If
            
            ii = ii + 1
            RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS = Nothing
    
    objKey.MoveFirst
    Do Until objKey.EOF
        strTmp = GetDiffMapping(objKey.Fields("testcd"))
        cmdTestKey(objKey.Fields("seq")).Visible = True
        lblTestNm(objKey.Fields("seq")).Visible = True
        
        If strTmp <> "" Then
            objKey.Fields("chr") = medGetP(strTmp, 1, COL_DIV)
            objKey.Fields("asc") = medGetP(strTmp, 2, COL_DIV)
            cmdTestKey(objKey.Fields("seq")).Caption = medGetP(strTmp, 1, COL_DIV)
            cmdTestKey(objKey.Fields("seq")).Tag = medGetP(strTmp, 2, COL_DIV)
            
            If cmdTestKey(objKey.Fields("seq")).Caption = "" Then
                cmdTestKey(objKey.Fields("seq")).Caption = "KEY"
                cmdTestKey(objKey.Fields("seq")).Tag = ""
            End If
            
            '## 자판배열에 일치하는게 있을면 표시!
            For jj = cmdKey.LBound To cmdKey.UBound
                If cmdKey(jj).Caption = cmdTestKey(objKey.Fields("seq")).Caption Then
                    cmdKey(jj).FontBold = True
                    cmdKey(jj).FontSize = 11
                    Exit For
                End If
            Next
        Else
            cmdTestKey(objKey.Fields("seq")).Caption = "KEY"
        End If
        
        objKey.MoveNext
    Loop
    Exit Sub
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당검사코드에 Mapping된 키값을 조회
'-----------------------------------------------------------------------------'
Private Function GetDiffMapping(ByVal sTestcd As String) As String
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = " SELECT field2, field3 FROM " & T_LAB032 & _
           " WHERE " & DBW("cdindex=", LC3_DiffKeyMap) & _
           " AND " & DBW("cdval1=", sTestcd) & _
           " AND " & DBW("field4=", ObjSysInfo.EmpId)
           
On Error GoTo Errors
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        GetDiffMapping = RS.Fields("field2").Value & "" & COL_DIV & _
                         RS.Fields("field3").Value & ""
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : WBC Diff Count에 포함된 상세코드를 조회
'-----------------------------------------------------------------------------'
Private Function GetDiffCodeSQL() As String
    Dim SSQL As String
    
    SSQL = " SELECT  b.testcd,b.abbrnm10 as testnm " & _
           " FROM   " & T_LAB032 & " a," & T_LAB001 & " b " & _
           " WHERE  " & DBW("a.cdindex=", LC3_WBCDiffCode) & _
           "    AND a.cdval1=b.testcd" & _
           "    AND b.applydt=(SELECT max(applydt) FROM " & T_LAB001 & _
           "                   WHERE testcd=a.cdval1 AND (panelfg ='' or panelfg is null))"
           
    GetDiffCodeSQL = SSQL
End Function

Private Function DiffSaveSQL(ByVal sTestcd As String, ByVal sTestNm As String, _
                             ByVal sChr As String, ByVal sASC As String) As Boolean
    Dim SSQL As String
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    SSQL = "delete FROM " & T_LAB032 & _
           " WHERE " & _
                      DBW("cdindex=", LC3_DiffKeyMap) & _
           " AND " & DBW("cdval1=", sTestcd)
    DBConn.Execute SSQL
    
    SSQL = " insert into " & T_LAB032 & _
           " (cdindex,cdval1,field1,field2,field3,field4) " & _
           " values( " & _
                DBV("cdindex", LC3_DiffKeyMap, 1) & _
                DBV("cdval1", sTestcd, 1) & _
                DBV("field1", sTestNm, 1) & _
                DBV("field2", sChr, 1) & _
                DBV("field3", sASC, 1) & _
                DBV("field4", ObjSysInfo.EmpId) & _
           " )"
    DBConn.Execute SSQL
    DBConn.CommitTrans
    DiffSaveSQL = True
    Exit Function
    
SAVE_ERROR:
    DBConn.RollbackTrans
End Function

Private Sub ClearData()
    mskAccNo.Mask = "&&-######-#####"
    mskAccNo.Text = "__-______-_____"
    
    mskAccNo.BackColor = vbWhite
    
    ClearFg = True
    txtInput.Text = ""
    lblErr.Caption = "":    lblDisease.Caption = "": lblTelno.Caption = ""
    ssRst.MaxRows = 0:      ssRst.Enabled = False
    tblDiff.MaxRows = 0:    tblDiff.Enabled = False
    cmdSave.Enabled = False
    Call CmdTemplete(False)
   '
    cboRelTest.Clear
    lvwPatient.ListItems.Clear
    lvwPatient.BackColor = DCM_LightGray
    rtfComment.BackColor = DCM_LightGray
    rtfText.BackColor = DCM_LightGray
    fraComment.Enabled = False
    lblCapRemark.Enabled = False
    rtfComment.Text = ""
    rtfText.Text = ""
    rtfRemark.Text = ""
    lvw.Visible = False
    txtMaxCount.Text = 100
    cmdRmk.Visible = False
    fraMesg.Visible = False
End Sub

Private Sub EditData()
    ssRst.Enabled = True
    mskAccNo.BackColor = DCM_LightGray
    cmdSave.Enabled = True
    lvwPatient.BackColor = vbWhite
    fraComment.Enabled = True
    lblCapRemark.Enabled = True
End Sub

Private Sub CmdTemplete(ByVal blnVisible As Boolean)
    cmdTextTemplete.Enabled = blnVisible
    cmdRemarkTemplete.Enabled = blnVisible
    cmdCommentTemplete.Enabled = blnVisible
End Sub

Public Sub Data_Load()
    Dim strBk As String
    
    '검사항목별 키값초기화
    objKey.MoveFirst
    Do Until objKey.EOF
        objKey.Fields("cnt") = ""
        objKey.MoveNext
    Loop
    txtCorrectWBC.Text = ""
    txtInput.Text = ""
    
    strBk = mskAccNo.Text
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
    
    Call CBCRoutinDisplay
    Call PtResultLoad(Trim(mskAccNo.FormattedText))
    
    If fraMesg.Visible Then fraMesg.Visible = False
    If cmdRmk.Visible Then cmdRmk.Visible = False
    
    If objPtInfo.TestCount > 0 Then
        ClearFg = False
        Call EditData
        lblErr.Caption = ""
        lvwPatient.SetFocus
        SendKeys "{TAB}"
        
        Dim MyResult As New clsLISResultReview
        Dim RS       As Recordset
        Dim SSQL     As String
        Dim ii       As Integer
        
        Dim strCombo
        
        Call MyResult.GetRelTest(cboRelTest, mskAccNo.FormattedText)
        
        '------------------------추가 사항----------------------------
        strCombo = ""
        For ii = 0 To cboRelTest.ListCount - 1
            strCombo = strCombo & cboRelTest.List(ii) & COL_DIV
        Next
        If strCombo <> "" Then strCombo = Mid(strCombo, 1, Len(strCombo) - 1)
        Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(1).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
        
         '처방리마크 조회(있는지 여부만 조회)
        SSQL = MyResult.GetOrderRemark(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq)
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        If Not RS.EOF Then cmdRmk.Visible = True
        
        cboRelTest.ListIndex = 0
        Set RS = Nothing
        Set MyResult = Nothing
        fraAccNo.Enabled = False
        '------------------------추가 사항----------------------------
        
        ssRst.Row = 1
        ssRst.Col = objPtInfo.SSCol("RESULT")
        ssRst.Action = ActionActiveCell
    Else
        mskAccNo.Text = strBk
    '    ssRst.Visible = True
        MsgBox "해당 접수번호는 입력할 검사가 없습니다.", vbCritical + vbOKOnly, "결과등록 Message"
        Call ClearData
        If mskAccNo.Enabled Then mskAccNo.SetFocus
    End If
End Sub

Private Function CBCRoutionSQL() As String
    Dim SSQL As String
    Dim sWorkArea As String
    Dim sAccDt As String
    Dim sAccSeq As String
    
    
    sWorkArea = medGetP(mskAccNo.FormattedText, 1, "-")
    
    If Mid(medGetP(mskAccNo.FormattedText, 2, "-"), 1, 1) = "9" Then
        sAccDt = "19" & Trim(medGetP(mskAccNo.FormattedText, 2, "-"))
    Else
        sAccDt = "20" & Trim(medGetP(mskAccNo.FormattedText, 2, "-"))
    End If
    sAccSeq = Trim(medGetP(mskAccNo.FormattedText, 3, "-"))
    
    
    SSQL = "SELECT a.workarea,a.accdt,a.accseq,a.testcd, a.rstval,a.rstcd," & _
             "       a.rstunit,a.hldiv,a.dpdiv,a.spccd,a.statfg,a.lastrst,a.lastvfydt,a.lastvfytm," & _
             "       a.lastvfyid,a.vfydt,a.vfytm,a.vfyid,a.mfyfg,a.grpfg,a.txtfg,a.rsttype,a.rstdiv," & _
             "       a.ptid,a.orddt,a.ordno,a.ordseq,a.detailfg,a.autofg,a.eqpcd,g.coldt,g.spcyy,g.spcno," & _
             "       h.testnm,h.txttype,h.rptseq," & _
             "       x.avalval,x.panicfg,x.panicfrval,x.panictoval,x.deltafg,x.deltaval,x.deltaval2," & _
             "       y.eqpnm" & _
             " FROM  " & T_LAB001 & " h, " & T_LAB004 & " x, " & T_LAB006 & " y, " & _
                         T_LAB201 & " g, " & T_LAB302 & " a " & _
             " WHERE " & DBW("a.workarea", sWorkArea, 2) & _
             " AND   " & DBW("a.accdt", sAccDt, 2) & _
             " AND   " & DBW("a.accseq", sAccSeq, 2) & _
             " AND   a.workarea = g.workarea" & _
             " AND   a.accdt    = g.accdt" & _
             " AND   a.accseq   = g.accseq" & _
             " AND   " & DBJ("y.eqpcd =* a.eqpcd")

        SSQL = SSQL & _
                 " AND   NOT EXISTS(SELECT f.field1 FROM " & T_LAB032 & " f " & _
                                    " WHERE " & DBW("f.cdindex", LC3_WBCDiffCode, 2) & _
                                    " AND   f.cdval1 = a.testcd )" & _
                 " AND   h.testcd = a.testcd" & _
                 " AND   h.applydt = (SELECT MAX(i.applydt) FROM " & T_LAB001 & " i" & _
                                    " WHERE i.testcd = a.testcd)" & _
                 " AND   x.testcd = a.testcd" & _
                 " AND   x.spccd = a.spccd" & _
                 " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
                                    " WHERE  testcd = x.testcd" & _
                                    " AND    spccd  = x.spccd)" & _
                 " ORDER BY h.rptseq"
    CBCRoutionSQL = SSQL
End Function

Private Sub CBCRoutinDisplay()
    Dim RS   As Recordset
    Dim SSQL As String
    
    mCurWBCCode = ""
    
    Set RS = New Recordset
    ssRst.MaxRows = 0
    ssRst.Visible = False
    RS.Open CBCRoutionSQL, DBConn
    
    If Not RS.EOF Then
        With ssRst
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                
                '## 현재 접수번호의 WBC 검사코드, 결과값을 저장
                If objWBC.Exists(RS.Fields("testcd").Value & "") Then
                    mCurWBCCode = RS.Fields("testcd").Value & ""
                    strWBCResult = RS.Fields("rstcd").Value & ""
                End If
                
                If RS.Fields("rstdiv").Value & "" <> "*" Then
                    If RS.Fields("detailfg").Value & "" <> "" Then
                        .Col = 1: .Value = "     " & RS.Fields("testnm").Value & ""
                    Else
                        .Col = 1: .Value = RS.Fields("testnm").Value & ""
                    End If
                    .BackColor = vbWhite
                    .ForeColor = DCM_LightBlue
                    
                Else
                    .Col = 1: .Value = RS.Fields("testnm").Value & ""
                    .BackColor = vbGrayText
                    .ForeColor = vbWhite
                    
                End If
                
                .Col = 2: .Value = objPtInfo.GetRstCd(RS.Fields("testcd").Value & "", _
                                                      RS.Fields("rstcd").Value & "")
                
                .Col = 3: .Value = RS.Fields("lastrst").Value & "": .ForeColor = DCM_LightRed: .FontBold = True
                .Col = 4:
                        Select Case RS.Fields("hldiv").Value & ""
                            Case "H": .Value = "High▲": .ForeColor = vbRed
                            Case "L": .Value = "▼Low": .ForeColor = vbBlue
                        End Select
                .Col = 5: .Value = RS.Fields("dpdiv").Value & ""
                .Col = 6: .Value = RS.Fields("rstunit").Value & ""
                .Col = 7: .Value = "최근보고일시 : " & Format(RS.Fields("vfydt").Value & "", "####-##-##") & " " & Format(RS.Fields("vfytm").Value & "", "0#:##:##")
                
                
                RS.MoveNext
            Loop
        End With
    End If
    
    If strWBCResult = "" Then
        MsgBox "WBC Count의 결과가 없습니다.", vbInformation + vbOKOnly, "Info"
    End If
    ssRst.Visible = True
    Set RS = Nothing
End Sub

Private Sub PtResultLoad(ByVal strAccNo As String)
    Dim i As Long
    
    mCurNRBCCode = ""
    lvwPatient.ListItems.Clear
    tblDiff.Visible = False
    DoEvents
    
    MouseRunning
    
    Set objPtInfo.prgBar = prgRst
    objPtInfo.PrgBarInit
    
    With objPtInfo
        .PtType = RESULT_BY_DIFFCOUNT           '/* 결과등록 유형, 반드시 셋팅 해야 됨./
        .AccNo = strAccNo                       '/* 접수번호, 반드시 셋팅 해야 됨./
        .LoadTable , ObjMyUser.EmpId
        
        If .TestCount > 0 Then
            Call CmdTemplete(True)
            If lvwPatient.Enabled = False Then
                lvwPatient.Enabled = True
            End If
            medDataLoadLvw lvwPatient, vbNewLine, vbTab, .GetStringPtInfo
            
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.PtId = lvwPatient.ListItems(1).Text
            lblDisease.Caption = objDisease.Disease
            lblDisease.ToolTipText = objDisease.Disease
            Set objDisease = Nothing
                
            rtfRemark.Text = .RmkNm
            rtfComment.Text = .FootNote
            If objPtInfo.Result.Item(1).TxtType <> "0" Then
                rtfText.Text = objPtInfo.Result.Item(1).TextRst
                rtfText.Enabled = True
                rtfText.BackColor = &HEEFFFE    'vbWhite
                cmdTextTemplete.Enabled = True
            Else
                rtfText.Enabled = False
                rtfText.BackColor = DCM_LightGray
                cmdTextTemplete.Enabled = False
            End If
            .GetResultSpread tblDiff, RESULT_BY_ACCESSION
            '========================================================================================
            '감염관리
            Call ICSPatientMark(lvwPatient.ListItems(1).Text, enICSNum.LIS_ALL)
            '병동/진료과 연락처(환자ID,CONTROL)
            Call GetPtTelInfo(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq, lblTelno)
            '========================================================================================
            
            '## 현재 접수번호의 NRBC검사코드를 조회
            For i = 1 To .TestCount
                If objNRBC.Exists(.Result.Item(i).TestCd) Then
                    mCurNRBCCode = .Result.Item(i).TestCd
                    Exit For
                End If
            Next i
        Else
            Call ICSPatientMark
        End If
    End With
    
    Dim ii As Integer
    
    With ssRst
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 5: .ForeColor = DCM_LightRed: .FontBold = True
        Next
    End With
    
    Call MouseDefault
    
    objPtInfo.PrgBarClear
    DoEvents
End Sub

Private Sub LoadLvwHead()
    Dim colHead As ColumnHeader
   
'    medInitLvwHead lvwPatient, "환자ID,환자성명,성/나이,생년월일,병상,주치의,검체", _
'                               "-100,300,-400,0,100,100,0"
    medInitLvwHead lvwPatient, "환자ID,환자성명,성/나이,생년월일,병상,주치의,검체,비고(외부QC)", _
                               "-100,300,-400,0,100,100,100,0"
End Sub

Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    Dim strTitle As String
   
    Set clsTemplete = New frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
    With clsTemplete
        .Show
        If pintMode = 0 Then
            .lblName.Caption = "Edit " & strTitle
        Else
            .lblName.Caption = "Modify " & strTitle
        End If
        .Caption = strTitle & " " & "Templete Editor"
        .lblInfo.Caption = pintMode & "$" & pintPrg
        Select Case pintPrg
            Case 1:
                .lblCode.Caption = objPtInfo.RmkCd
                .rtfText = rtfRemark.Text
            Case 2:
                .rtfText = rtfText.Text
            Case 3:
                .rtfText = rtfComment.Text
        End Select
    End With
    gintTemplete = pintPrg
    
End Sub

Private Sub clsTemplete_CopyTemplete()
    If ssRst.MaxRows < 1 Then Exit Sub
    With objPtInfo
        Select Case gintTemplete
            Case 1:
                If clsTemplete.rtfText.Text <> "" Then
                    rtfRemark.Text = clsTemplete.rtfText.Text
                    .RmkCd = frm230TempSearch.lblCode.Caption
                    .RmkNm = rtfRemark.Text
                Else
                    rtfRemark.Text = ""
                    .RmkCd = ""
                    .RmkNm = ""
                End If
            Case 2:
                rtfText.Text = clsTemplete.rtfText.Text
                .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
                rtfText.SetFocus
            Case 3:
                rtfComment.Text = clsTemplete.rtfText.Text
                .FootNote = rtfComment.Text
                rtfComment.SetFocus
        End Select
    End With
    Set clsTemplete = Nothing
End Sub

Private Sub cmdRemarkTemplete_Click()
    Dim SqlStmt As String

    Set objCodeList = Nothing
    Set objCodeList = New clsPopUpList

    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Remark"
        .ColumnHeaderText = "Code;Remark"
        .LoadPopUp SqlStmt
    End With
End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    objPtInfo.RmkCd = objCodeList.SelectedItems(0)
    objPtInfo.RmkNm = objCodeList.SelectedItems(1)
    rtfRemark.Text = objPtInfo.RmkNm
End Sub

Private Sub tblDiff_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row < 1 Then Exit Sub
    objPtInfo.SpToolTip Row, Col, MultiLine, ShowTip, TipText, TipWidth
    tblDiff.TextTip = TextTipFloatingFocusOnly
End Sub

Private Function SaveCorrectWBC() As Boolean
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim SSQL        As String
    
    sWorkArea = medGetP(mskAccNo.FormattedText, 1, "-")
    
    If Mid(medGetP(mskAccNo.FormattedText, 2, "-"), 1, 1) = "9" Then
        sAccDt = "19" & Trim(medGetP(mskAccNo.FormattedText, 2, "-"))
    Else
        sAccDt = "20" & Trim(medGetP(mskAccNo.FormattedText, 2, "-"))
    End If
    sAccSeq = Trim(medGetP(mskAccNo.FormattedText, 3, "-"))
    
On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    SSQL = "UPDATE " & T_LAB302 & " SET  " & _
         DBW("rstcd", Val(txtCorrectWBC.Text), 3) & DBW("rstval", Val(txtCorrectWBC.Text), 3) & _
         DBW("vfydt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
         DBW("vfytm", Format(GetSystemDate, "HHMMSS"), 3) & _
         DBW("vfyid", ObjSysInfo.EmpId, 2) & _
         "WHERE " & _
         DBW("workarea=", CBC_WorkArea) & _
         " AND " & DBW("accdt=", sAccDt) & _
         " AND " & DBW("accseq=", sAccSeq) & _
         " AND " & DBW("testcd=", mCurWBCCode)
         
    DBConn.Execute SSQL
    DBConn.CommitTrans
    SaveCorrectWBC = True
    Exit Function
    
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, "오류"
End Function

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sSEQ As String
    Dim sWorkArea As String
    Dim sAccDt As String
    Dim sAccSeq As String
    

    sWorkArea = medGetP(Item.SubItems(2), 1, "-")
    sAccDt = medGetP(Item.SubItems(2), 2, "-")
    sAccSeq = medGetP(Item.SubItems(2), 3, "-")
    Call LabNoResult(sWorkArea, sAccDt, sAccSeq)
    
    lvw.Visible = False
End Sub

Private Sub LabNoResult(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String)
    Call ClearData

    mskAccNo.Mask = String(Len(WorkArea), "&") & "-"
    mskAccNo.Mask = mskAccNo.Mask & String(Len(Mid(AccDt, 3)), "#") & "-"
    mskAccNo.Mask = mskAccNo.Mask & String(Len(AccSeq), "#")
    
    mskAccNo.Text = WorkArea & "-" & Mid(AccDt, 3) & "-" & AccSeq
    Call Data_Load
End Sub

Private Sub txtMaxCount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cboRelTest.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : WBC, NRBC검사코드 Dictionary를 초기화
'-----------------------------------------------------------------------------'
Private Sub GetWBCInfo()
    Dim RS  As Recordset
    Dim SQL As String
    
    '## WBC
    Set objWBC = New clsDictionary
    objWBC.Clear
    objWBC.FieldInialize "testcd", "testnm"
        
    SQL = " SELECT a.cdval1, b.testnm FROM " & T_LAB032 & " a, " & T_LAB001 & " b" & _
          " WHERE " & DBW("a.cdindex=", LC3_WBCCode) & _
          " AND a.cdval1=b.testcd AND b.applydt=(SELECT MAX(applydt) FROM " & T_LAB001 & _
          " WHERE testcd=a.cdval1)"

On Error GoTo Errors
    Set RS = New Recordset
    RS.Open SQL, DBConn
    If Not (RS.BOF Or RS.EOF) Then
        Do Until RS.EOF
            If Not objWBC.Exists(RS.Fields("cdval1").Value & "") Then
                objWBC.AddNew RS.Fields("cdval1").Value & "", RS.Fields("testnm").Value & ""
            End If
            
            RS.MoveNext
        Loop
    End If
    RS.Close

    '## NRBC
    Set objNRBC = New clsDictionary
    objNRBC.Clear
    objNRBC.FieldInialize "testcd", "testnm"
    
    SQL = " SELECT a.cdval1, b.testnm FROM " & T_LAB032 & " a, " & T_LAB001 & " b" & _
          " WHERE " & DBW("a.cdindex=", LC3_NRBCCode) & _
          " AND a.cdval1=b.testcd AND b.applydt=(SELECT MAX(applydt) FROM " & T_LAB001 & _
          " WHERE testcd=a.cdval1)"

    RS.Open SQL, DBConn
    If Not (RS.BOF Or RS.EOF) Then
        Do Until RS.EOF
            If Not objNRBC.Exists(RS.Fields("cdval1").Value & "") Then
                objNRBC.AddNew RS.Fields("cdval1").Value & "", RS.Fields("testnm").Value & ""
            End If
            
            RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS = Nothing
    Exit Sub
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub
