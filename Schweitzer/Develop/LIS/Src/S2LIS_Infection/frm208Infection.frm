VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm208Infection 
   BackColor       =   &H00DBE6E6&
   Caption         =   "감염관리 결과등록"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14640
   Icon            =   "frm208Infection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14640
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00F4F0F2&
      Caption         =   "삭제"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   555
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   76
      Tag             =   "124"
      Top             =   8625
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Middle Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   9870
      TabIndex        =   59
      Tag             =   "20003"
      Top             =   4065
      Width           =   4710
      Begin VB.CommandButton cmdMQuery 
         BackColor       =   &H00FFF7FC&
         Caption         =   "&Query"
         Height          =   480
         Left            =   3585
         MaskColor       =   &H00808080&
         Style           =   1  '그래픽
         TabIndex        =   60
         Top             =   195
         Width           =   1035
      End
      Begin FPSpread.vaSpread tblMidReview 
         Height          =   3540
         Left            =   75
         TabIndex        =   61
         Top             =   720
         Width           =   4545
         _Version        =   196608
         _ExtentX        =   8017
         _ExtentY        =   6244
         _StockProps     =   64
         BackColorStyle  =   1
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
         GrayAreaBackColor=   14411494
         MaxCols         =   14
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   -2147483633
         ShadowText      =   0
         SpreadDesigner  =   "frm208Infection.frx":058A
         TextTip         =   4
      End
      Begin MSComCtl2.DTPicker dtpMFDate 
         Height          =   285
         Left            =   615
         TabIndex        =   62
         Top             =   315
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   21430275
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpMTDate 
         Height          =   285
         Left            =   2220
         TabIndex        =   63
         Top             =   315
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   21430275
         CurrentDate     =   36328
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1935
         TabIndex        =   65
         Tag             =   "40110"
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   64
         Tag             =   "40105"
         Top             =   345
         Width           =   495
      End
   End
   Begin VB.Frame fraTemp 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   60
      TabIndex        =   55
      Tag             =   "20003"
      Top             =   3945
      Visible         =   0   'False
      Width           =   6885
      Begin VB.CommandButton cmdTClear 
         Appearance      =   0  '평면
         BackColor       =   &H00F4F0F2&
         Caption         =   "Clear"
         Height          =   510
         Left            =   270
         Style           =   1  '그래픽
         TabIndex        =   72
         Top             =   4020
         Width           =   1320
      End
      Begin VB.TextBox txtCdindex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   71
         Top             =   4125
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtTemp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   69
         Top             =   2580
         Width           =   1365
      End
      Begin VB.CommandButton cmAddTmp 
         Appearance      =   0  '평면
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장"
         Height          =   510
         Left            =   3915
         Style           =   1  '그래픽
         TabIndex        =   67
         Top             =   4020
         Width           =   1320
      End
      Begin VB.CommandButton cmDelSelTmp 
         Appearance      =   0  '평면
         BackColor       =   &H00F4F0F2&
         Caption         =   "제거"
         Height          =   510
         Left            =   5250
         Style           =   1  '그래픽
         TabIndex        =   66
         Top             =   4020
         Width           =   1320
      End
      Begin RichTextLib.RichTextBox rtfTemp 
         Height          =   990
         Left            =   270
         TabIndex        =   56
         Top             =   2970
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   1746
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frm208Infection.frx":492B
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
      Begin MSComctlLib.ListView lvwTemplate 
         Height          =   1905
         Left            =   270
         TabIndex        =   57
         Top             =   540
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16775406
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "코드"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "내용"
            Object.Width           =   12435
         EndProperty
      End
      Begin MedControls1.LisLabel lblTitle 
         Height          =   300
         Left            =   270
         TabIndex        =   58
         Top             =   210
         Width           =   6315
         _ExtentX        =   11139
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
         Caption         =   "Foot Ward by Report"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Left            =   270
         TabIndex        =   68
         Top             =   2565
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "코드"
         Appearance      =   0
      End
      Begin VB.Label Label7 
         BackColor       =   &H00ACCDD0&
         Caption         =   "[코드 등록방법 : CF) 01, 02, 03, ...]"
         ForeColor       =   &H00313D46&
         Height          =   180
         Left            =   3120
         TabIndex        =   70
         Top             =   2655
         Width           =   3420
      End
   End
   Begin VB.ListBox lstDept 
      Appearance      =   0  '평면
      BackColor       =   &H00FFF7FC&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   4275
      TabIndex        =   8
      Top             =   555
      Visible         =   0   'False
      Width           =   4350
   End
   Begin VB.ListBox lstResult 
      Appearance      =   0  '평면
      BackColor       =   &H00F7F3F8&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   10410
      Sorted          =   -1  'True
      TabIndex        =   52
      Top             =   480
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.ListBox lstSpc 
      Appearance      =   0  '평면
      BackColor       =   &H00FFF7FC&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   7755
      TabIndex        =   51
      Top             =   480
      Visible         =   0   'False
      Width           =   4350
   End
   Begin VB.CommandButton cmdMfy 
      BackColor       =   &H00F4F0F2&
      Caption         =   "결과수정"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   10695
      Style           =   1  '그래픽
      TabIndex        =   50
      Tag             =   "124"
      Top             =   8610
      Width           =   1215
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   9375
      Style           =   1  '그래픽
      TabIndex        =   46
      Tag             =   "124"
      Top             =   8625
      Width           =   1215
   End
   Begin VB.CommandButton cmdMResult 
      BackColor       =   &H00F4F0F2&
      Caption         =   "중간결과"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   6795
      Style           =   1  '그래픽
      TabIndex        =   45
      Tag             =   "124"
      Top             =   8625
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Result Review"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   60
      TabIndex        =   32
      Tag             =   "20003"
      Top             =   4065
      Width           =   9795
      Begin VB.OptionButton optQueryKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "보고일"
         Height          =   255
         Index           =   1
         Left            =   1335
         TabIndex        =   36
         Tag             =   "15305"
         Top             =   345
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optQueryKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "의뢰일"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   35
         Tag             =   "15304"
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFF7FC&
         Caption         =   "&Query"
         Height          =   480
         Left            =   8460
         MaskColor       =   &H00808080&
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   195
         Width           =   1185
      End
      Begin FPSpread.vaSpread tblReview 
         Height          =   3540
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   9525
         _Version        =   196608
         _ExtentX        =   16801
         _ExtentY        =   6244
         _StockProps     =   64
         BackColorStyle  =   1
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
         GrayAreaBackColor=   14411494
         MaxCols         =   23
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   -2147483633
         ShadowText      =   0
         SpreadDesigner  =   "frm208Infection.frx":4B5D
         TextTip         =   4
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   285
         Left            =   3390
         TabIndex        =   37
         Top             =   315
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   21430275
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   285
         Left            =   5265
         TabIndex        =   38
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   21430275
         CurrentDate     =   36328
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2865
         TabIndex        =   40
         Tag             =   "40105"
         Top             =   345
         Width           =   495
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4980
         TabIndex        =   39
         Tag             =   "40110"
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame fraReport 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Foot Ward by Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   60
      TabIndex        =   29
      Tag             =   "20003"
      Top             =   2580
      Width           =   6885
      Begin VB.CommandButton cmdFDel 
         BackColor       =   &H00DB95FD&
         Caption         =   "X"
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
         Left            =   6480
         Style           =   1  '그래픽
         TabIndex        =   49
         ToolTipText     =   "Delete"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdFootWard 
         BackColor       =   &H00DEDBDD&
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
         Left            =   6480
         Picture         =   "frm208Infection.frx":94FF
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   945
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfFootWard 
         Height          =   990
         Left            =   90
         TabIndex        =   31
         Top             =   270
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   1746
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frm208Infection.frx":9A31
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
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Comment by Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   45
      TabIndex        =   25
      Tag             =   "20003"
      Top             =   1170
      Width           =   6885
      Begin VB.CommandButton cmdCDel 
         BackColor       =   &H00DB95FD&
         Caption         =   "X"
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
         Left            =   6480
         Style           =   1  '그래픽
         TabIndex        =   48
         ToolTipText     =   "Delete"
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton cmdComment 
         BackColor       =   &H00DEDBDD&
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
         Left            =   6480
         Picture         =   "frm208Infection.frx":9AE1
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   945
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   990
         Left            =   90
         TabIndex        =   27
         Top             =   270
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   1746
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frm208Infection.frx":A013
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
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   8055
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "124"
      Top             =   8625
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   13260
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8595
      Width           =   1215
   End
   Begin VB.CommandButton cmdFResult 
      BackColor       =   &H00F4F0F2&
      Caption         =   "최종결과"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   11970
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8610
      Width           =   1215
   End
   Begin VB.Frame fraWS 
      BackColor       =   &H00DBE6E6&
      Height          =   1170
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   14550
      Begin VB.TextBox txtSpcCd 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4230
         TabIndex        =   10
         Top             =   675
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton cmdSpc 
         BackColor       =   &H00DEDBDD&
         Caption         =   "▼"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   675
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtDeptCd 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4230
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton cmdDept 
         BackColor       =   &H00DEDBDD&
         Caption         =   "▼"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5640
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   210
         Width           =   285
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   5955
         TabIndex        =   7
         Top             =   225
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         BackColor       =   15265000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   5955
         TabIndex        =   11
         Top             =   675
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         BackColor       =   15265000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dptWorkDt 
         Height          =   315
         Left            =   1155
         TabIndex        =   12
         Top             =   240
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy'-'MM'-'dd"
         Format          =   21430275
         CurrentDate     =   36287
      End
      Begin MSComCtl2.DTPicker dptColDt 
         Height          =   315
         Left            =   1155
         TabIndex        =   14
         Top             =   690
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyy-MM-dd HH:mm"
         Format          =   21430275
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Left            =   11475
         TabIndex        =   17
         Top             =   210
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "최종보고자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblFvfyID 
         Height          =   360
         Left            =   12735
         TabIndex        =   18
         Top             =   210
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   360
         Left            =   11475
         TabIndex        =   19
         Top             =   675
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "최종보고일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblFvfyDt 
         Height          =   360
         Left            =   12735
         TabIndex        =   20
         Top             =   675
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Left            =   8445
         TabIndex        =   21
         Top             =   210
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "중간보고자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblMvfyID 
         Height          =   360
         Left            =   9705
         TabIndex        =   22
         Top             =   210
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   360
         Left            =   8445
         TabIndex        =   23
         Top             =   675
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "중간보고일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblMvfyDt 
         Height          =   360
         Left            =   9705
         TabIndex        =   24
         Top             =   675
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblFvfy 
         Height          =   360
         Left            =   12735
         TabIndex        =   73
         Top             =   405
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "의뢰검체 :"
         Height          =   180
         Left            =   3345
         TabIndex        =   75
         Top             =   750
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "채취일시 :"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   765
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "의뢰부서 :"
         Height          =   180
         Left            =   3345
         TabIndex        =   15
         Top             =   300
         Width           =   840
      End
      Begin VB.Label lblBuildDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "의뢰일자 :"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6960
      TabIndex        =   28
      Tag             =   "20003"
      Top             =   1170
      Width           =   7635
      Begin VB.ComboBox cboResult 
         BackColor       =   &H00F1F5F4&
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frm208Infection.frx":A245
         Left            =   5430
         List            =   "frm208Infection.frx":A24F
         TabIndex        =   47
         Text            =   "적합"
         Top             =   255
         Width           =   2055
      End
      Begin VB.TextBox txtWaterArea 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   42
         Top             =   255
         Width           =   2340
      End
      Begin VB.CheckBox chkWater 
         BackColor       =   &H00DBE6E6&
         Caption         =   "정수기"
         ForeColor       =   &H00737A58&
         Height          =   315
         Left            =   105
         TabIndex        =   41
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "판정결과"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   4605
         TabIndex        =   44
         Tag             =   "107"
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lblWaterArea 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "정수기위치"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1155
         TabIndex        =   43
         Tag             =   "107"
         Top             =   330
         Width           =   900
      End
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   1905
      Left            =   6990
      TabIndex        =   53
      Top             =   2070
      Width           =   7590
      _Version        =   196608
      _ExtentX        =   13388
      _ExtentY        =   3360
      _StockProps     =   64
      AutoCalc        =   0   'False
      BackColorStyle  =   1
      DAutoSizeCols   =   0
      EditEnterAction =   5
      EditModePermanent=   -1  'True
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
      GrayAreaBackColor=   14411494
      MaxCols         =   4
      MaxRows         =   499
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   -2147483633
      ShadowText      =   0
      SpreadDesigner  =   "frm208Infection.frx":A261
   End
   Begin FPSpread.vaSpread tblWater 
      Height          =   1905
      Left            =   6975
      TabIndex        =   54
      Top             =   2070
      Visible         =   0   'False
      Width           =   7590
      _Version        =   196608
      _ExtentX        =   13388
      _ExtentY        =   3360
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      EditEnterAction =   5
      EditModePermanent=   -1  'True
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
      GrayAreaBackColor=   14411494
      MaxCols         =   3
      MaxRows         =   499
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   -2147483633
      ShadowText      =   0
      SpreadDesigner  =   "frm208Infection.frx":E1E2
   End
   Begin VB.PictureBox picESign 
      Height          =   500
      Left            =   2940
      ScaleHeight     =   435
      ScaleWidth      =   1140
      TabIndex        =   74
      Top             =   8715
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDF0F5&
      BackStyle       =   0  '투명
      ForeColor       =   &H00313D46&
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   8700
      Width           =   60
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00C0C0C0&
      Height          =   390
      Left            =   15
      Shape           =   4  '둥근 사각형
      Top             =   8580
      Width           =   4410
   End
End
Attribute VB_Name = "frm208Infection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsSQL      As New clsInfectiontSQL
Private strImgPath  As String

'-- Header 결과 Type
Private Type InDataHeader
    WorkDt      As String   '의뢰일자
    WorkDept    As String   '의뢰부서
    ColDt       As String   '채취일자
    ColTm       As String   '채취시간
    TestMeth    As String   '검사방법
    WaterFg     As String   '정수기검사 Flag
    WaterArea   As String   '정수기위치
    RstVal      As String   '정수기 판정결과
    StsCd       As String   '진행상태
    MvfyID      As String   '중간결과등록자ID
    MvfyDt      As String   '중간결과확인일자
    MvfyTm      As String   '중간결과확인시간
    FvfyID      As String   '최종결과등록자ID
    FvfyDt      As String   '최종결과확인일자
    FvfyTm      As String   '중간결과확인시간
    MfyID       As String   '결과수정자ID
    MfyDt       As String   '결과수정일자
    MfyTm       As String   '결과수정시간
    RptFg       As String   '레포트 출력여부
    RptTxt      As String   '레포트 FootWard
    RptDt       As String   '출력일자
    RptTm       As String   '출력시간
    RptID       As String   '출력자ID
End Type

'-- Body 결과 Type
Private Type InDataBody
    WorkDt      As String   '의뢰일자
    WorkDept    As String   '의뢰부서
    ColDt       As String   '채취일자
    ColTm       As String   '채취시간
    SpcCd       As String   '의뢰검체
    RstCd       As String   '균종결과
    RstCount    As String   'Colony/물1ml
    RstTxt      As String   'Text결과(Nogrowth등)
    RstFg       As String   'Text결과여부
End Type

Private Const iCm = 10
Private Const iLineHeight = 10


Private HData   As InDataHeader
Private bData   As InDataBody

Private iPageWidth  As Integer
Private iPageHeight As Integer
Private PageNumber  As Integer
Private iCurY       As Integer


Private Sub chkWater_Click()
    With chkWater
        If .Value = 1 Then
            txtWaterArea.Enabled = True
            cboResult.Enabled = True
            txtSpcCd.Text = "정수기"
'            txtDeptCd.Text = "정수기"
'            lblDeptNm.Caption = "정수기"
            cmdDept.Enabled = False
            
            Call Spread_Use(.Value)
            'Call Spread_Init(.Value)
        Else
            txtWaterArea.Text = ""
            cboResult.ListIndex = 0
            txtWaterArea.Enabled = False
            cboResult.Enabled = False
            txtSpcCd.Text = ""
            cmdDept.Enabled = True
            
            Call Spread_Use(.Value)
            'Call Spread_Init(.Value)
        End If
    End With
End Sub

Private Sub Spread_Use(ByVal pFlag As String)
    Select Case pFlag
        Case "1"
            tblWater.Visible = True
            tblResult.Visible = False
        
        Case "0"
            tblWater.Visible = False
            tblResult.Visible = True
            
    End Select
End Sub

Private Sub Spread_Init(ByVal pFlag As String)
    Dim iCol As Integer
    
    With tblResult
        Select Case pFlag
            Case "1"
                .MaxRows = 0: .MaxCols = 4
                
                .Col = 1: .ColWidth(iCol) = 24.38
                .Col = 2: .ColWidth(iCol) = 32.38
                .Col = 3: .ColWidth(iCol) = 9
                .Col = 4: .ColWidth(iCol) = 9
                
                ' - Title 설정
                .Col = 1: .Col2 = 4
                '.CellType = CellTypeEdit
                .Row = -1: .Row2 = -1
                .Clip = "Specimen" & Chr$(9) & "(No)Growth" & Chr$(9) & "검체코드" & Chr$(9) & "균코드"
                        
            Case "0"
                .MaxRows = 0: .MaxCols = 2
                
                .Col = 1: .ColWidth(iCol) = 56.75
                .Col = 2: .ColWidth(iCol) = 9
                
                ' - Title 설정
                .Col = 1: .Col2 = 2
                '.CellType = CellTypeEdit
                .Row = -1: .Row2 = -1
                .Clip = "균명" & Chr$(9) & "균코드"
                
        End Select
    End With
End Sub

Private Sub cmAddTmp_Click()
    Dim strSql      As String
    Dim strCdindex  As String
    Dim strCode     As String
    Dim strTemp     As String
    
    strCdindex = txtCdindex.Text
    strCode = Trim(txtTemp.Text)
    strTemp = rtfTemp.Text
    
    If strCdindex = "" Then
        MsgBox "오류가 발생하였습니다. 전산담당자에게 연락 주시기 바랍니다!"
        Exit Sub
    End If
    
    If strCode = "" Then
        Exit Sub
    End If
    
    On Error GoTo ErrMsg
    
    DBConn.BeginTrans
    
    If clsSQL.Template_Insert_Update_Flag(strCdindex, strCode) = False Then
        '-- INSERT
        strSql = clsSQL.Template_Insert(strCdindex, strCode, strTemp)
        
        Call DBConn.Execute(strSql)
    Else
        '-- UPDATE
        strSql = clsSQL.Template_Update(strCdindex, strCode, strTemp)
        
        Call DBConn.Execute(strSql)
    End If
    
    DBConn.CommitTrans
    
    '-- Display
    Call Template_List(strCdindex)
    
    Exit Sub
    
ErrMsg:
    MsgBox Err.Description
    DBConn.RollbackTrans
End Sub

Private Sub cmdCDel_Click()
    rtfComment.Text = ""
End Sub

Private Sub cmdClear_Click()
    Call ClearData
End Sub

Private Sub cmdComment_Click()
    If fraTemp.Visible = False Then
        Call TempClear
        
        lblTitle.Caption = fraComment.Caption
        txtCdindex.Text = LC4_Infection
        
        '-- Template Display
        Call Template_List(LC4_Infection)
        
        '-- Disp Position
        With fraTemp
            .Left = fraComment.Left
            .Top = fraComment.Top + fraComment.Height
        End With
        
        fraTemp.Visible = True
    Else
        fraTemp.Visible = False
    End If
End Sub

Private Sub TempClear()
    lvwTemplate.ListItems.Clear
    txtTemp.Text = ""
    rtfTemp.Text = ""
End Sub

Private Sub Template_List(ByVal pValue As String)
    Dim RS      As New ADODB.Recordset
    Dim strSql  As String
    Dim itmX    As ListItem
    
    strSql = clsSQL.TestMeth_List(pValue)
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    With lvwTemplate
        If RS.BOF = False Then
            .ListItems.Clear
            
            Do Until RS.EOF = True
                
                Set itmX = .ListItems.Add(, , RS.Fields("cdval1").Value & "")
                itmX.SubItems(1) = RS.Fields("text1").Value & ""
                
                RS.MoveNext
            Loop
        End If
    End With
    
    RS.Close
    Set RS = Nothing
    
End Sub

Private Sub cmdDelete_Click()
    Dim strSql  As String
    Dim Message As String
            
    If Value_Check = False Then
        Exit Sub
    End If
    
    On Error GoTo ErrMsg
    
    Message = MsgBox("삭제 하시겠습니까?", vbCritical + vbYesNo, "중간결과 삭제")
    
    If Message = vbNo Then
        Exit Sub
    End If
    
    DBConn.BeginTrans
    
    With HData
        '** Header Delete
        strSql = clsSQL.LAB315_DELETE_SQL(.WorkDt, .WorkDept, .ColDt, .ColTm)
        
        Call DBConn.Execute(strSql)
        
        '** Body Delete
        strSql = clsSQL.LAB316_DELETE_SQL(.WorkDt, .WorkDept, .ColDt, .ColTm)
        
        Call DBConn.Execute(strSql)
    End With
    
    DBConn.CommitTrans
    
    Call ClearData
    
    Call cmdMQuery_Click
    
    Exit Sub
    
ErrMsg:
    MsgBox Err.Description
    DBConn.RollbackTrans
End Sub

Private Sub cmdDept_Click()
    If lstDept.ListCount = 0 Then
        MsgBox "등록된 의뢰부서가 없습니다.", vbCritical
        Exit Sub
    End If
    lstDept.Visible = True
    lstDept.ZOrder 0
    lstDept.SetFocus
End Sub

Private Sub cmDelSelTmp_Click()
    Dim strSql  As String
    Dim strCode As String
    Dim strMsg  As String
    
    If txtCdindex.Text = "" Then
        Exit Sub
    End If
    
    If Trim(txtTemp.Text) = "" Then
        Exit Sub
    End If
    
    strMsg = MsgBox("선택한 항목을 삭제 하시겠습니까?", vbCritical + vbYesNo, "삭제")
    
    If strMsg = vbNo Then
        Exit Sub
    End If
    
    strCode = Trim(txtTemp.Text)
    
    On Error GoTo ErrMsg
    
    DBConn.BeginTrans
    
    strSql = clsSQL.Template_Delete(txtCdindex.Text, strCode)
    
    Call DBConn.Execute(strSql)
    
    DBConn.CommitTrans
    
    Call TempClear
    
    '-- Display
    Call Template_List(txtCdindex.Text)
    
    Exit Sub
    
ErrMsg:
    MsgBox Err.Description
    DBConn.RollbackTrans
End Sub

Private Sub DelSelTemp()
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFDel_Click()
    rtfFootWard.Text = ""
End Sub

Private Sub cmdFootWard_Click()
    If fraTemp.Visible = False Then
        Call TempClear
        
        lblTitle.Caption = fraReport.Caption
        txtCdindex.Text = LC4_FootWard
        
        '-- Template Display
        Call Template_List(LC4_FootWard)
        
        '-- Disp Position
        With fraTemp
            .Left = fraReport.Left
            .Top = fraReport.Top + fraReport.Height
        End With
        
        fraTemp.Visible = True
    Else
        fraTemp.Visible = False
    End If
End Sub

Private Sub cmdFResult_Click()
    Call VerifyResult(StsCd_LIS_FinRst)
End Sub

Private Sub cmdMfy_Click()
    Call ModifyResult(StsCd_LIS_ModRst)
End Sub

Private Sub cmdMQuery_Click()
    Dim strWorkDt   As String
    Dim strDeptCd   As String
    Dim strFDate    As String
    Dim strTDate    As String
    
    strWorkDt = Format(dptWorkDt.Value, "yyyymmdd")
    strDeptCd = Trim(txtDeptCd.Text)
    
    strFDate = Format(dtpMFDate.Value, "yyyymmdd")
    strTDate = Format(dtpMTDate.Value, "yyyymmdd")
    
    Call Mid_DispInfo(strWorkDt, strDeptCd, strFDate, strTDate)
    cmdMfy.Enabled = True
End Sub

Private Sub cmdMResult_Click()
    Call VerifyResult(StsCd_LIS_MidRst)
End Sub

Private Sub cmdQuery_Click()
    Dim RS          As New ADODB.Recordset
    Dim strSql      As String
    Dim strFDate    As String
    Dim strTDate    As String
    Dim strFlag     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim strVfyDt    As String
    Dim strVfyTm    As String
    Dim strStatus   As String
    Dim i           As Integer
    
    '** 조회조건 분류 Flag Value (의뢰일:0, 보고일:1)
    If optQueryKey(0).Value = True Then
        strFlag = "0"
        strFDate = Format(dtpFromDate.Value, "yyyymmdd")
        strTDate = Format(dtpToDate.Value, "yyyymmdd")
    Else
        strFlag = "1"
        strFDate = Format(dtpFromDate.Value, "yyyymmdd")
        strTDate = Format(dtpToDate.Value, "yyyymmdd")
    End If
    
    strSql = clsSQL.Result_Info(strFDate, strTDate, strFlag)
        
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    With tblReview
        .MaxRows = 0: i = 1
        If RS.BOF = False Then
            
            Do Until RS.EOF = True
                
                .Row = i: .MaxRows = i
                
                .Col = 1: .Value = Format(Mid(RS.Fields("workdt").Value & "", 3, 6), "0#/##/##")
                .Col = 2: .Value = RS.Fields("deptnm").Value & ""
                .Col = 3: .Value = RS.Fields("testmeth").Value & ""
                
                strColDt = Format(RS.Fields("coldt").Value & "", "####/##/##")
                strColTm = Format(RS.Fields("coltm").Value & "", "##:##:##")
                .Col = 4: .Value = strColDt & " " & strColTm
                
                strStatus = RS.Fields("stscd").Value & ""
                If strStatus = StsCd_LIS_FinRst Then
                    strVfyDt = Format(RS.Fields("fvfydt").Value & "", "####/##/##")
                    strVfyTm = Format(RS.Fields("fvfytm").Value & "", "##:##:##")
                    .Col = 5: .Value = strVfyDt & " " & strVfyTm
                    .Col = 6: .Value = "최종"
                ElseIf strStatus = StsCd_LIS_ModRst Then
                        strVfyDt = Format(RS.Fields("mfydt").Value & "", "####/##/##")
                        strVfyTm = Format(RS.Fields("mfytm").Value & "", "##:##:##")
                        .Col = 5: .Value = strVfyDt & " " & strVfyTm
                        .Col = 6: .Value = "수정"
                Else
                    strVfyDt = Format(RS.Fields("mvfydt").Value & "", "####/##/##")
                    strVfyTm = Format(RS.Fields("mvfytm").Value & "", "##:##:##")
                    .Col = 5: .Value = strVfyDt & " " & strVfyTm
                    .Col = 6: .Value = "중간"
                End If
                
                strColDt = Format(RS.Fields("coldt").Value & "", "####/##/##")
                strColTm = Format(RS.Fields("coltm").Value & "", "##:##:##")
                .Col = 4: .Value = strColDt & " " & strColTm
                
                .Col = 7: .Value = RS.Fields("workdt").Value & ""
                .Col = 8: .Value = RS.Fields("deptcd").Value & ""
                .Col = 9: .Value = RS.Fields("coldt").Value & ""
                .Col = 10: .Value = RS.Fields("coltm").Value & ""
                
                .Col = 11: .Value = RS.Fields("waterfg").Value & ""
                .Col = 12: .Value = RS.Fields("waterarea").Value & ""
                .Col = 13: .Value = RS.Fields("rstval").Value & ""
                
                .Col = 14: .Value = RS.Fields("stscd").Value & ""
                .Col = 15: .Value = RS.Fields("mvfydt").Value & ""
                .Col = 16: .Value = RS.Fields("mvfytm").Value & ""
                .Col = 17: .Value = RS.Fields("mvfyid").Value & ""
                .Col = 18: .Value = Verify_Name(RS.Fields("mvfyid").Value & "")
                
                .Col = 19: .Value = RS.Fields("fvfydt").Value & ""
                .Col = 20: .Value = RS.Fields("fvfytm").Value & ""
                .Col = 21: .Value = RS.Fields("fvfyid").Value & ""
                .Col = 22: .Value = Verify_Name(RS.Fields("fvfyid").Value & "")
                
                .Col = 23: .Value = RS.Fields("rpttxt").Value & ""
                
                i = i + 1
                RS.MoveNext
            Loop
        End If
    End With
    
    RS.Close
    Set RS = Nothing
    
End Sub

Private Sub cmdReport_Click()
    Dim lngFileNo As Long
    
    lngFileNo = FreeFile
    
    If Printers.Count = 0 Then
        MsgBox "현재 설정된 프린터가 없으므로 출력할 수 없습니다.", vbInformation, "프린터"
        Exit Sub
    End If
    
    Call Print_Init
    Call ReportPrint
    
End Sub

Private Sub Print_Init()
    Printer.Font = "굴림체"
    Printer.FontSize = 10
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
    Twidth = Printer.ScaleWidth
    
    Select Case Printer.PaperSize
        Case 9
            PrtLeft = 10
            'LineSpace = 6
            LineSpace = 5
        Case 7
            PrtLeft = 0
            LineSpace = 4
        Case Else
            PrtLeft = 0
            LineSpace = 4
    End Select
    LastLineYpos = Printer.ScaleHeight - iCm             '마지막라인Y위치
    
End Sub

Private Sub ReportPrint()
    Dim tmpWorkArea     As String
    Dim tmpAccDt        As String
    Dim tmpAccSeq       As String
    Dim tmpWorkAreaNm   As String
    Dim SaveWorkArea    As String
    Dim strBuffer       As String
    Dim strTxt          As String
    Dim strNotice       As String
    Dim strTmp          As String
    Dim strWorkDt       As String
    Dim strSpcNm        As String
    Dim strRstNm        As String
    Dim strRepDt        As String
    Dim aryComment()    As String
    Dim aryFoot()       As String
    
    Dim i       As Integer
    Dim ii      As Integer
    Dim jj      As Integer
    Dim kk      As Integer
    Dim lngCnt  As Integer
    
    Dim objRichText As RichTextBox
    Dim objimage    As Image
    Dim objESign    As clsLISElectronSign
    Dim lngCurYPos  As Long
On Error GoTo Err_Trap
    
    Printer.FontSize = 9
    
    tmpWorkArea = ""
    tmpAccDt = ""
    tmpAccSeq = ""
                
    Call PrtHeader
    
    Set objRichText = frmControls.rtfTextBox
    Set objimage = frmControls.Image1
                            
    Call CheckNewPage
    
    Printer.FontSize = 12
    
    If chkWater.Value = 1 Then
        With tblWater
            
            Call Print_Setting("▶ 정수기 위치 ◀", iCm * 2, LineSpace, Twidth, "L", "C", True)
            Call ChangeLine
            
            Call Print_Setting(txtWaterArea.Text, iCm * 3, LineSpace, Twidth, "L", "C", True)
            Call ChangeLine
            
            Call Print_Setting("▶ 판정 ◀", iCm * 2, LineSpace, Twidth, "L", "C", True)
            Call ChangeLine
            
            Call Print_Setting(cboResult.Text, iCm * 3, LineSpace, Twidth, "L", "C", True)
            Call ChangeLine
            
            Call Print_Setting("▶ 세균명 :", iCm * 2, LineSpace, Twidth, "L", "C", True)
            Call ChangeLine
            
            '-- 세균명
            For i = 1 To .DataRowCnt
                .Row = i: .Col = 1
                
                strSpcNm = i & ". " & .Value
                
                Call Print_Setting(strSpcNm, iCm * 3, LineSpace, Twidth, "L", "C", True)
                Call ChangeLine
                
                Call CheckNewPage
            Next
            
            Call Print_Setting("", PrtLeft, LineSpace * 1, Twidth, "R", "C")
            Call CheckNewPage
    
            Call Print_Setting("▶ Colony/물1ml :", iCm * 2, LineSpace, Twidth, "L", "C", True)
            Call ChangeLine
            
            '-- Colony/물1ml
            For i = 1 To .DataRowCnt
                .Row = i: .Col = 3
                
                strRstNm = i & ". " & .Value
                
                Call Print_Setting(strRstNm, iCm * 3, LineSpace, Twidth, "L", "C", True)
                Call ChangeLine
                
                Call CheckNewPage
            Next
        End With
    Else
        With tblResult
            Call Print_Setting("▶ 의뢰된 검체 :", iCm * 1, LineSpace, Twidth, "L", "C", False)
            Call Print_Setting("▶ 배양된 균 :", iCm * 8, LineSpace, Twidth, "L", "C", True)
            Call ChangeLine
            lngCurYPos = 58
            '-- 의뢰검체
            For i = 1 To .DataRowCnt
                .Row = i: .Col = 1
                strSpcNm = i & ". " & .Value
                .Row = i: .Col = 2
                strRstNm = "" & .Value

                Call Print_Setting(strSpcNm, iCm * 1, LineSpace, Twidth, "L", "C", False)
                Printer.Line (PrtLeft, lngCurYPos + LineSpace * i)-(Twidth - PrtLeft, lngCurYPos + LineSpace * i)
                Call Print_Setting(strRstNm, iCm * 8, LineSpace, Twidth, "L", "C", False)
                
                Call ChangeLine
                
                Call CheckNewPage
            Next
            
            Printer.Line (PrtLeft, lngCurYPos + LineSpace)-(Twidth - PrtLeft, lngCurYPos + LineSpace)
    
            Call Print_Setting("", PrtLeft, LineSpace * 2, Twidth, "R", "C")
            Call CheckNewPage
' 08.09.26 양성현 검체와 배양균이 나누어져출력되던것을 합침.
'            Call Print_Setting("▶ 배양된 균 :", iCm * 3, LineSpace, Twidth, "L", "C", True)
'            Call ChangeLine
'
'            '-- 균종결과
'            For i = 1 To .DataRowCnt
'                .Row = i: .Col = 2
'
'                strRstNm = i & ". " & .Value
'
'                Call Print_Setting(strRstNm, iCm * 5, LineSpace, Twidth, "L", "C", True)
'                Call ChangeLine
'
'                Call CheckNewPage
'            Next
        End With
    End If
    
    Call Print_Setting("", PrtLeft, LineSpace * 2, Twidth, "R", "C")
    Call CheckNewPage
    
'    Call Print_Setting("☞ Comments :", 10 + iCm * 5, LineSpace, Twidth, "L", "C", True)
'    Call ChangeLine
'
'    Call Print_Setting(rtfComment.Text, 0, LineSpace, Twidth, "C", "C", True)
'    Call ChangeLine
'    Call CheckNewPage
    
    '-- 소견결과
    If rtfComment.Text <> "" Then
        aryComment() = Split(rtfComment.Text, vbCrLf)
        Printer.FontBold = True
        
        Call Print_Setting("▣ Comments :", iCm * 2, LineSpace, Twidth, "L", "C", True)
        Call ChangeLine
        Printer.FontBold = False
        For ii = LBound(aryComment) To UBound(aryComment)
            If Trim(aryComment(ii)) <> "" Then
                If LenB(StrConv(aryComment(ii), vbFromUnicode)) > 60 Then
                    lngCnt = LenB(StrConv(aryComment(ii), vbFromUnicode)) \ 60
                    kk = 1
                    For jj = 1 To lngCnt
                        Call Print_Setting(Trim(Mid(aryComment(ii), kk, 60)), iCm * 3, LineSpace, Twidth, "L", "C", True)
                        Call CheckNewPage
                        kk = kk + 60
                    Next
                Else
                    Call Print_Setting(aryComment(ii), 0, LineSpace, Twidth, "C", "C", True)
                    Call CheckNewPage
                End If
            End If
        Next
        Call ChangeLine
    End If
    
    Call Print_Setting("", PrtLeft, LineSpace * 2, Twidth, "R", "C")
    Call CheckNewPage
    
    Printer.FontSize = 12
    Printer.FontBold = True
        
    strRepDt = Format(Now, "yyyy년mm월dd일")
    Call Print_Setting(strRepDt, 0, LineSpace, Twidth, "C", "C", True)
    Call ChangeLine
    Call CheckNewPage
    
    Call Print_Setting("", PrtLeft, LineSpace, Twidth, "R", "C")
    
    '-- FootWard
    If rtfFootWard.Text <> "" Then
        aryFoot() = Split(rtfFootWard.Text, vbCrLf)

        For ii = LBound(aryFoot) To UBound(aryFoot)
            If Trim(aryFoot(ii)) <> "" Then
                If LenB(StrConv(aryFoot(ii), vbFromUnicode)) > 60 Then
                    lngCnt = LenB(StrConv(aryFoot(ii), vbFromUnicode)) \ 60
                    kk = 1
                    For jj = 1 To lngCnt
                        Call Print_Setting(Trim(Mid(aryFoot(ii), kk, 60)), 0, LineSpace, Twidth, "C", "C", True)
                        Call CheckNewPage
                        kk = kk + 60
                    Next
                Else
                    Call Print_Setting(aryFoot(ii), 0, LineSpace, Twidth, "C", "C", True)
                    Call CheckNewPage
                End If
            End If
        Next
        Call ChangeLine
    End If

'    Call Print_Setting("보고자 : 진단검사의학과 전문의    " & lblFvfyID.Caption & "  M.D", iCm * 0.4, LineSpace, Twidth, "L", "C", False)
'    Call ChangeLine

    Printer.FontBold = False
    
    Call Print_Setting("", 0, LineSpace * 2, Twidth, "C", "C", True)
    
    Set objESign = New clsLISElectronSign
    If objESign.LoadElectronSign(lblFvfy.Caption, InstallDir & "LIS\bin") = True Then
        If objESign.ElectronSignPrintOk = True Then
            strImgPath = objESign.ElectronSignPath & "\" & objESign.ElectronSignFileName
            picESign.Picture = LoadPicture(strImgPath)
            Printer.PaintPicture picESign.Picture, 120 - iCm / 2, Printer.CurrentY - 10, 30, 15
        End If
    End If
    
    Call Print_Last
    
    Printer.EndDoc
    
    Exit Sub
    
Err_Trap:
    MsgBox Err.Description
On Error GoTo Err_Trap
    Resume Next

End Sub

Private Sub PrtHeader()
    Dim Header1     As Integer
    Dim Header2     As Integer
    Dim strTmp      As String
    Dim sICSString  As String
    Dim strTitle    As String
    Dim strWorkDt   As String
    Dim strColDt    As String
    
    Header1 = Twidth * (1 / 3) + PrtLeft + iCm / 2
    Header2 = Twidth * (2 / 3) + PrtLeft
    
    lngCurYPos = 10
    Printer.FontSize = 18: Printer.FontBold = True
    
    If chkWater.Value = 1 Then
        strTitle = "정수기 CULTURE"
    Else
        strTitle = lblDeptNm.Caption & " " & "CULTURE"
    End If
    
    Call Print_Setting(strTitle, 0, 12, Twidth, "C", "C")
    Printer.FontSize = 12: Printer.FontBold = False
    Call Print_Setting("", PrtLeft, LineSpace * 2, Twidth, "C", "C")
    
    Printer.DrawStyle = vbSolid
    Printer.DrawWidth = 3
    Printer.Line (PrtLeft, lngCurYPos - 2)-(Twidth - PrtLeft, lngCurYPos - 2)
    
    strWorkDt = "◈ 검체의뢰일 :" & dptWorkDt.Value
    
    Call Print_Setting(strWorkDt, PrtLeft, LineSpace, Twidth, "L", "C", False): Printer.FontBold = True: Printer.FontBold = False
    
    strColDt = "◈ 채취일시 : " & dptColDt.Value
    
    Call Print_Setting(strColDt, Header1, LineSpace, Twidth, "L", "C", False): Printer.FontBold = True: Printer.FontBold = False
    
    Call Print_Setting("", PrtLeft, LineSpace / 5, Twidth, "R", "C")
    Printer.Line (PrtLeft, lngCurYPos + LineSpace)-(Twidth - PrtLeft, lngCurYPos + LineSpace)
    
    Call Print_Setting("", PrtLeft, LineSpace * 3, Twidth, "R", "C")

End Sub

Private Sub CheckNewPage()
    
    If lngCurYPos > LastLineYpos - (1# * iCm) Then  ' newPage일 경우
        PageNumber = PageNumber + 1
        Printer.Line (PrtLeft, LastLineYpos)-(Twidth - PrtLeft, LastLineYpos)
        Call P_FIX(PageNumber, PrtLeft, LastLineYpos + 3, Twidth - PrtLeft, "C", , "C")
        Printer.NewPage
        Call PrtHeader
    End If
            
End Sub

Private Sub Print_Last()
    PageNumber = PageNumber + 1
    Printer.Line (PrtLeft, LastLineYpos)-(Twidth - PrtLeft, LastLineYpos)
    Call P_FIX(PageNumber, PrtLeft, LastLineYpos + 7, Twidth - PrtLeft, "C", , "C")
    
    Printer.FontSize = 13: Printer.FontBold = True
    Call P_FIX("예수병원 진단검사의학과     전북 전주시 완산구 중화산동 1-300", PrtLeft, LastLineYpos + 3, Twidth - PrtLeft, "C", , "C")
    
    Printer.FontSize = 9: Printer.FontBold = False
End Sub

Private Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer, iSpace As Integer)

    Select Case sLineStyle
        Case "solid"
            Printer.DrawStyle = 0
        Case "dash"
            Printer.DrawStyle = 1
        Case "dot"
            Printer.DrawStyle = 2
        Case "dashdot"
            Printer.DrawStyle = 3
        Case "dashdotdot"
            Printer.DrawStyle = 4
    End Select
         
    Printer.DrawWidth = iLinewidth
    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
    iCurY = Printer.CurrentY + iSpace
End Sub

Private Sub ChangeLine()
    Call Print_Setting("", 10 + iCm * 0.4, LineSpace, Twidth, "C", "C")
    
    '추가함
    Call CheckNewPage
End Sub

Private Sub prtPageNum()
    
    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String
    
    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
    
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page
            
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6
    Printer.Print "RUN-DATE : " & sDate
        
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + iCm / 6
    Printer.Print "RUN-TIME : " & sTime
        
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
End Sub

Private Sub cmdSpc_Click()
    If lstSpc.ListCount = 0 Then
        MsgBox "등록된 검체가 없습니다.", vbCritical
        Exit Sub
    End If
    lstSpc.Visible = True
    lstSpc.ZOrder 0
    lstSpc.SetFocus
End Sub

Private Sub cmdTClear_Click()
    txtTemp.Text = ""
    rtfTemp.Text = ""
    txtTemp.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        '-- Template 화면 지움
        If fraTemp.Visible = True Then
            fraTemp.Visible = False
        End If
        
        If lstDept.Visible = True Then
            lstDept.Visible = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Call ShowDeptList
    Call ShowSpcList
    Call ShowItemLit
    
    Call ClearData
    
    dptWorkDt.Value = Now
    dptColDt.Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    dtpMFDate.Value = DateAdd("d", -3, Format(Now, "yyyy-mm-dd"))
    dtpMTDate.Value = Format(Now, "yyyy-mm-dd")
    
End Sub

Private Sub ShowDeptList()
    Dim RS          As New ADODB.Recordset
    Dim strSql      As String
    Dim i           As Long
    Dim strDeptNm   As String
    Dim strDeptCd   As String
    
    On Error GoTo ErrList
    
    strSql = " select * from " & TB_Dept & _
             "  order by deptnm "
' 08.09.26 양성현 부서명으로 정열하게 변경
'             "  order by dpcd "
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.BOF = False Then
        
        lstDept.Clear: i = 0
        
        Do While RS.EOF = False
            
            strDeptCd = RS.Fields("dpcd").Value & ""
            strDeptNm = RS.Fields("deptnm").Value & ""
            
            lstDept.AddItem strDeptCd & vbTab & strDeptNm & vbTab, i
            i = i + 1
            
            RS.MoveNext
        Loop
        
    End If
    
    If lstDept.ListCount = 0 Then
        MsgBox "등록된 부서가 없습니다.", vbCritical
    End If
    
    Exit Sub
    
ErrList:
    MsgBox Err.Description, vbExclamation
    On Error Resume Next
End Sub

Private Sub ShowSpcList()
    Dim i As Integer
    Dim tmpSpcCd As String
    Dim tmpSpcNm As String
    Dim RS       As New ADODB.Recordset
    Dim strSql   As String

    On Error GoTo ErrList

    strSql = clsSQL.Spc_List

    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly

    DoEvents

    With lstSpc
        .Clear

        medLockWindowUpdate (.hwnd)

        RS.MoveFirst

        While (Not RS.EOF)

            tmpSpcNm = Mid(RS.Fields("field1").Value & "", 1, 50)
            tmpSpcNm = tmpSpcNm & Space(50 - Len(tmpSpcNm)) & vbTab  ' 검체명
            tmpSpcCd = Trim(Mid(RS.Fields("cdval1").Value & "", 1, 9))
            tmpSpcCd = tmpSpcCd & Space(9 - Len(tmpSpcCd)) & vbTab   ' 검체코드

            If Trim(tmpSpcCd) <> "" Then

                .AddItem tmpSpcNm & tmpSpcCd & "1"  '검체명기준
                .AddItem tmpSpcCd & tmpSpcNm & "2"  '검체코드기준

            End If

            DoEvents
            RS.MoveNext
        Wend
        .Visible = False

        medLockWindowUpdate (0&)

    End With

    RS.Close
    Set RS = Nothing

    Exit Sub

ErrList:
    MsgBox Err.Description, vbExclamation
    Set RS = Nothing
    On Error Resume Next
End Sub

Private Sub ShowItemLit()
    Dim i As Integer
    Dim tmpMicCd As String
    Dim tmpMicNm As String
    
    Dim RS As Recordset
    Dim strSql As String
    
    On Error GoTo ErrList
    
    strSql = " select * from " & TB_LAB032 & _
             "  where cdindex = " & DBS(CD2_Micro)
    
    Set RS = New Recordset
    
    RS.Open strSql, DBConn
    
    DoEvents
    
    With lstResult
        .Clear
        
        medLockWindowUpdate (.hwnd)
        
        RS.MoveFirst
        
        While (Not RS.EOF)
             
            tmpMicNm = Mid(RS.Fields("field2").Value & "", 1, 50)
            tmpMicNm = tmpMicNm & Space(50 - Len(tmpMicNm)) & vbTab  ' 균명
            tmpMicCd = Trim(Mid(RS.Fields("cdval1").Value & "", 1, 9))
            tmpMicCd = tmpMicCd & Space(9 - Len(tmpMicCd)) & vbTab   ' 균코드
             
            If Trim(tmpMicCd) <> "" Then
            
                .AddItem tmpMicNm & tmpMicCd & "1"  '균명기준
                .AddItem tmpMicCd & tmpMicNm & "2"  '균코드기준
                
            End If
         
            DoEvents
            RS.MoveNext
        Wend
        .Visible = False
        
        medLockWindowUpdate (0&)
        
    End With
    
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ErrList:
    MsgBox Err.Description, vbExclamation
    Set RS = Nothing
    On Error Resume Next
End Sub

Private Sub ClearData()
    Dim i       As Integer
    Dim j       As Integer
    
    txtDeptCd.Text = ""
    lblDeptNm.Caption = ""
    txtSpcCd.Text = ""
    lblSpcNm.Caption = ""
    
    lblMvfyID.Caption = ""
    lblMvfyDt.Caption = ""
    lblFvfyID.Caption = ""
    lblFvfyDt.Caption = ""
    
    rtfComment.Text = ""
    rtfFootWard.Text = ""
    rtfTemp.Text = ""
    
    chkWater.Value = 0
    txtWaterArea.Text = ""
    cboResult.ListIndex = 0
    
    With tblResult
        .MaxRows = 0: .MaxRows = 20
        For i = 1 To .MaxRows
            .Row = i
            For j = 1 To .MaxCols
                .Col = i: .Value = ""
            Next
        Next
    End With
    
    With tblWater
        .MaxRows = 0: .MaxRows = 20
        For i = 1 To .MaxRows
            .Row = i
            For j = 1 To .MaxCols
                .Col = i: .Value = ""
            Next
        Next
    End With
    
    tblReview.MaxRows = 0: tblReview.MaxRows = 20
    tblMidReview.MaxRows = 0: tblMidReview.MaxRows = 20
    
    optQueryKey(0).Value = True
    dtpFromDate.Value = Now
    dtpToDate.Value = Now
    
    Call Enable_Check(True)
    cmdFResult.Enabled = True
    
    cmdDelete.Enabled = False
    
    cmdMfy.Enabled = False
    cmdReport.Enabled = False
    
End Sub

Private Sub Enable_Check(ByVal pFlag As Boolean)
    dptWorkDt.Enabled = pFlag
    dptColDt.Enabled = pFlag
    
    cmdDept.Enabled = pFlag
    cmdMResult.Enabled = pFlag
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsSQL = Nothing
    Set frm208Infection = Nothing
    Set ObjMyUser = Nothing
End Sub


Private Sub lstDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RS          As New ADODB.Recordset
    Dim strSql      As String
    Dim strWorkDt   As String
    Dim strDeptCd   As String
    Dim strColDt    As String
    Dim strColTm    As String
    
    If KeyCode = vbKeyReturn Then
        txtDeptCd.Text = medGetP(lstDept.Text, 1, vbTab)
        lblDeptNm.Caption = medGetP(lstDept.Text, 2, vbTab)
        lstDept.Visible = False
        'txtSpcCd.SetFocus
    End If
    
    '-- Values Set
    strWorkDt = Format(dptWorkDt.Value, "yyyymmdd")
    strDeptCd = Trim(txtDeptCd.Text)
    strColDt = Format(dptColDt.Value, "yyyymmdd")
    strColTm = Format(dptColDt.Value, "hhmmss")
    
    '** 중간결과 상태 조회 루틴
    If clsSQL.Mid_Verify_Check(strWorkDt, strDeptCd, strColDt, strColTm) = True Then
        Call Mid_DispInfo(strWorkDt, strDeptCd, strColDt, strColTm)
    End If
    
End Sub

Private Sub Mid_DispInfo(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                         ByVal pFDate As String, ByVal pTDate As String)
    Dim RS          As New ADODB.Recordset
    Dim strSql      As String
    Dim strMvfyID   As String
    Dim strMvfyDt   As String
    Dim strMvfyTm   As String
    Dim i           As Integer
    
    With tblMidReview
        strSql = clsSQL.Mid_Result_Find(pWorkDt, pDeptCd, pFDate, pTDate, StsCd_LIS_MidRst)
        
        RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
        
        If RS.BOF = False Then
        
            i = 1: .MaxRows = 0: .MaxRows = RS.RecordCount + 2
            Do Until RS.EOF = True
                .MaxRows = i
                
                .Row = i
                
                .Col = 1: .Value = Format(Mid(RS.Fields("workdt").Value & "", 3, 6), "0#-##-##")
                .Col = 2: .Value = DeptName(RS.Fields("deptcd").Value & "")
                
                strMvfyDt = Format(RS.Fields("mvfydt").Value & "", "####-##-##")
                strMvfyTm = Format(RS.Fields("mvfytm").Value & "", "##:##:##")
                
                .Col = 3: .Value = strMvfyDt & " " & strMvfyTm
                
                '-- Hidden
                .Col = 4: .Value = RS.Fields("workdt").Value & ""
                .Col = 5: .Value = RS.Fields("deptcd").Value & ""
                .Col = 6: .Value = RS.Fields("coldt").Value & ""
                .Col = 7: .Value = RS.Fields("coltm").Value & ""
                .Col = 8: .Value = RS.Fields("waterfg").Value & ""
                .Col = 9: .Value = RS.Fields("waterarea").Value & ""
                .Col = 10: .Value = RS.Fields("rstval").Value & ""
                .Col = 11: .Value = RS.Fields("testmeth").Value & ""
                .Col = 12: .Value = RS.Fields("rpttxt").Value & ""
                .Col = 13: .Value = RS.Fields("mvfyid").Value & ""
                .Col = 14: .Value = Verify_Name(RS.Fields("mvfyid").Value & "")
                
                i = i + 1
                RS.MoveNext
            Loop
            
        End If
        
    End With
End Sub

Private Function Verify_Name(ByVal pMfyID As String) As String
    Dim RS          As New ADODB.Recordset
    Dim strSql      As String
    
    strSql = clsSQL.Mid_Result_Name(pMfyID)
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Verify_Name = RS.Fields("empnm").Value & ""
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

Private Sub lstDept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call lstDept_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If lstDept.Enabled Then lstDept.SetFocus
End Sub

Private Sub lstResult_KeyPress(KeyAscii As Integer)
    If chkWater.Value = 1 Then
        Select Case KeyAscii
            Case 13:    'Enter Key 또는 Space
                Call lstResult_MouseDown(1, 0, 0, 0)
            Case 27:  'ESC
                lstResult.Visible = False
                tblWater.SetFocus
            Case Else:   '그 밖에...
                tblWater.SetFocus
                tblWater.Action = ActionActiveCell
                SendKeys Chr(KeyAscii)
        End Select
    Else
        Select Case KeyAscii
            Case 13:    'Enter Key 또는 Space
                Call lstResult_MouseDown(1, 0, 0, 0)
            Case 27:  'ESC
                lstResult.Visible = False
                tblResult.SetFocus
            Case Else:   '그 밖에...
                tblResult.SetFocus
                tblResult.Action = ActionActiveCell
                SendKeys Chr(KeyAscii)
        End Select
    End If
End Sub

Private Sub lstResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpStr As String
    Dim tmpField1 As String
    Dim tmpField2 As String
    
    If Button <> 1 Then Exit Sub
    If lstResult.ListIndex < 0 Then Exit Sub
    
    tmpStr = lstResult.List(lstResult.ListIndex)
    
    If chkWater.Value = 1 Then
        With tblWater
            tmpField1 = Trim(medShift(tmpStr, vbTab))
            tmpField2 = Trim(medShift(tmpStr, vbTab))
            
            If tmpStr = "1" Then
                .Col = 1:  .Value = Trim(tmpField1)    ' 균명
                .Col = 2:  .Value = Trim(tmpField2)    ' 균코드
            Else
                .Col = 1:  .Value = Trim(tmpField2)    ' 균명
                .Col = 2:  .Value = Trim(tmpField1)    ' 균코드
            End If
            
            lstResult.Visible = False
'            Call tblWater_LeaveCell(.Col, .Row, 2, .Row, False)
        End With
    Else
        With tblResult
            tmpField1 = Trim(medShift(tmpStr, vbTab))
            tmpField2 = Trim(medShift(tmpStr, vbTab))
            
            If tmpStr = "1" Then
                .Col = 2:  .Value = Trim(tmpField1)    ' 균명
                .Col = 4:  .Value = Trim(tmpField2)    ' 균코드
            Else
                .Col = 2:  .Value = Trim(tmpField2)    ' 균명
                .Col = 4:  .Value = Trim(tmpField1)    ' 균코드
            End If
            
            lstResult.Visible = False
'            Call tblResult_LeaveCell(.Col, .Row, 2, .Row, False)
        End With
    End If
End Sub

Private Sub lstResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
'    If lstResult.Enabled Then lstResult.SetFocus
End Sub

Private Sub lstSpc_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        txtSpcCd.Text = medGetP(lstSpc.Text, 1, vbTab)
'        lblSpcNm.Caption = medGetP(lstSpc.Text, 2, vbTab)
'        lstSpc.Visible = False
'        'txtSpcCd.SetFocus
'    End If
End Sub

Private Sub lstSpc_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 32:    'Enter Key 또는 Space
            Call lstSpc_MouseDown(1, 0, 0, 0)
        Case 27:  'ESC
            lstSpc.Visible = False
            tblResult.SetFocus
        Case Else:   '그 밖에...
            tblResult.SetFocus
            tblResult.Action = ActionActiveCell
            SendKeys Chr(KeyAscii)
    End Select
End Sub

Private Sub lstSpc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpStr As String
    Dim tmpField1 As String
    Dim tmpField2 As String

    If Button <> 1 Then Exit Sub
    If lstSpc.ListIndex < 0 Then Exit Sub

    tmpStr = lstSpc.List(lstSpc.ListIndex)

    With tblResult
        tmpField1 = Trim(medShift(tmpStr, vbTab))
        tmpField2 = Trim(medShift(tmpStr, vbTab))

        If tmpStr = "1" Then
            .Col = 1:  .Value = Trim(tmpField1)    ' 균명
            .Col = 3:  .Value = Trim(tmpField2)    ' 균코드
        Else
            .Col = 1:  .Value = Trim(tmpField2)    ' 균명
            .Col = 3:  .Value = Trim(tmpField1)    ' 균코드
        End If

        lstSpc.Visible = False
        Call tblResult_LeaveCell(.Col, .Row, 2, .Row, False)

    End With
End Sub

Private Sub lstSpc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
'    If lstSpc.Enabled Then lstSpc.SetFocus
End Sub

Private Sub lvwTemplate_DblClick()
    Select Case txtCdindex.Text
        Case LC4_Infection
            rtfComment.Text = rtfTemp.Text
            
        Case LC4_FootWard
            rtfFootWard.Text = rtfTemp.Text
        
    End Select
    
    fraTemp.Visible = False
End Sub

Private Sub lvwTemplate_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvwTemplate
        If .ListItems.Count = 0 Then
            Exit Sub
        End If
        
        txtTemp.Text = lvwTemplate.SelectedItem.Text
        rtfTemp.Text = lvwTemplate.SelectedItem.SubItems(1)
    End With
End Sub

Private Sub lvwTemplate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lvwTemplate_DblClick
    End If
End Sub

Private Sub tblMidReview_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strWorkDt       As String
    Dim strDeptCd       As String
    Dim strColDt        As String
    Dim strColTm        As String
    
    With tblMidReview
        
        If Row < 1 Or .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        .Row = Row
        
        .Col = 4
        dptWorkDt.Value = Format(.Value, "####-##-##"): strWorkDt = .Value
        
        .Col = 6: strColDt = .Value
        .Col = 7: strColTm = .Value
        dptColDt.Value = Format(strColDt & strColTm, "####-##-## 0#:##:##")
        
        .Col = 5: txtDeptCd.Text = .Value: strDeptCd = .Value
        .Col = 2: lblDeptNm.Caption = .Value
        
        .Col = 8
        If .Value = "1" Then
            chkWater.Value = 1
            .Col = 9: txtWaterArea.Text = .Value
            .Col = 10: cboResult.Text = .Value
            
            '-- Body Disp
            Call Mid_Water_Result(strWorkDt, strDeptCd, strColDt, strColTm, StsCd_LIS_MidRst)
            
        Else
            txtWaterArea.Text = ""
            cboResult.Text = "적합"
            chkWater.Value = 0
            
            '-- Body Disp
            Call Mid_Body_Result(strWorkDt, strDeptCd, strColDt, strColTm, StsCd_LIS_MidRst)
            
        End If
        
        .Col = 11: rtfComment.Text = .Value
        .Col = 12: rtfFootWard.Text = .Value
        .Col = 14: lblMvfyID.Caption = .Value
        
        .Col = 3: lblMvfyDt.Caption = .Value
        
    End With
    
    
    Call Enable_Check(False)
    cmdFResult.Enabled = True
    
    cmdDelete.Enabled = True
    
    cmdReport.Enabled = False
    
End Sub

Private Sub Mid_Body_Result(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                            ByVal pColDt As String, ByVal pColTm As String, _
                            Optional ByVal pStatus As String)
                            
    Dim RS      As New ADODB.Recordset
    Dim strSql  As String
    Dim i       As Integer
    
    strSql = clsSQL.Mid_Body_Result_Find(pWorkDt, pDeptCd, pColDt, pColTm, pStatus)
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    With tblResult
        .MaxRows = 0: .MaxRows = 20
        i = 1
        
        If RS.BOF = False Then
            Do Until RS.EOF = True
                .Row = i: .MaxRows = i + 1
                
                .Col = 1
                If RS.Fields("spcnm").Value & "" <> "" Then
                    .Value = RS.Fields("spcnm").Value & ""
                Else
                    .Value = RS.Fields("spccd").Value & ""
                End If
                .Col = 2: .Value = RS.Fields("rstnm").Value & ""
                .Col = 3: .Value = RS.Fields("spccd").Value & ""
                .Col = 4: .Value = RS.Fields("rstcd").Value & ""
                
                i = i + 1
                RS.MoveNext
            Loop
        End If
    End With
    
    RS.Close
    Set RS = Nothing
    
End Sub

Private Sub Mid_Water_Result(ByVal pWorkDt As String, ByVal pDeptCd As String, _
                             ByVal pColDt As String, ByVal pColTm As String, _
                             Optional ByVal pStatus As String)
                             
    Dim RS      As New ADODB.Recordset
    Dim strSql  As String
    Dim i       As Integer
    
    strSql = clsSQL.Mid_Body_Result_Find(pWorkDt, pDeptCd, pColDt, pColTm, pStatus)
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    With tblWater
        .MaxRows = 0: .MaxRows = 20
        i = 1
        
        If RS.BOF = False Then
            
            Do Until RS.EOF = True
                .Row = i: .MaxRows = i + 1
                
                .Col = 1: .Value = RS.Fields("rstnm").Value & ""
                .Col = 2: .Value = RS.Fields("rstcd").Value & ""
                .Col = 3: .Value = RS.Fields("rstcount").Value & ""
                
                i = i + 1
                RS.MoveNext
            Loop
        End If
    End With
    
    RS.Close
    Set RS = Nothing
    
End Sub

Private Sub tblResult_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim tmpIndex As Integer
    Dim tmpStr As String
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean

    'If Col <> 1 Then Exit Sub

    With tblResult
        .Col = Col
        .Row = Row
        
'        If Trim(.Value) = "" Then
'            Exit Sub
'        End If
        
        Select Case Col
            Case 1:
                tmpIndex = medListFind(lstSpc, tblResult.Value)
                tmpStr = lstSpc.List(tmpIndex)

                ' 문자가 입력될때마다 유사어 찾기

                If tmpIndex <> -1 Or UCase(tmpStr) = UCase(.Value) Then
                    Ret = .GetCellPos(Col, Row + 1, X, Y, Wdt, Hgt)
                    If .Height - Y < lstSpc.Height Or Y < 0 Then
                        Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                        lstSpc.Top = .Top + Y - lstSpc.Height
                    Else
                        lstSpc.Top = .Top + Y
                    End If
                    If tmpIndex >= 0 Then
                        medLockWindowUpdate (lstSpc.hwnd)

                        lstSpc.ListIndex = tmpIndex
                        medLockWindowUpdate (0&)
                        If tmpIndex - lstSpc.TopIndex > 10 Then lstSpc.TopIndex = tmpIndex
                    End If
                    lstSpc.Visible = True
'                    lstSpc.SetFocus
                    lstSpc.ZOrder 0
                Else
                    medLockWindowUpdate (lstSpc.hwnd)

                    lstSpc.ListIndex = tmpIndex
                    medLockWindowUpdate (0&)
                    Call lstSpc_MouseDown(1, 0, 0, 0)
                    lstSpc.Visible = False
                End If

                .Row = Row: .Col = Col
                If Trim(.Value) = "" Then
                    .Col = 3: .Value = ""
                    lstSpc.Visible = False
                End If

                If Row = .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                
            Case 2:
                tmpIndex = medListFind(lstResult, .Value)
                tmpStr = lstResult.List(tmpIndex)
                
                ' 문자가 입력될때마다 유사어 찾기
                If tmpIndex <> -1 Or UCase(tmpStr) = UCase(.Value) Then
                    Ret = .GetCellPos(Col, Row + 1, X, Y, Wdt, Hgt)
                    If .Height - Y < lstResult.Height Or Y < 0 Then
                        Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                        lstResult.Top = .Top + Y - lstResult.Height
                    Else
                        lstResult.Top = .Top + Y
                    End If
                    If tmpIndex >= 0 Then
                        medLockWindowUpdate (lstResult.hwnd)
        
                        lstResult.ListIndex = tmpIndex
                        medLockWindowUpdate (0&)
                        If tmpIndex - lstResult.TopIndex > 10 Then lstResult.TopIndex = tmpIndex
                    End If
                    lstResult.Visible = True
'                    lstResult.SetFocus
                    lstResult.ZOrder 0
                Else
                    medLockWindowUpdate (lstResult.hwnd)
        
                    lstResult.ListIndex = tmpIndex
                    medLockWindowUpdate (0&)
                    Call lstResult_MouseDown(1, 0, 0, 0)
                    lstResult.Visible = False
                End If
        
                .Row = Row: .Col = Col
                If Trim(.Value) = "" Then
                    .Col = 4: .Value = ""
                    lstResult.Visible = False
                End If
        
                If Row = .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                
        End Select
    End With
End Sub

Private Sub tblResult_KeyDown(KeyCode As Integer, Shift As Integer)
    With lstSpc
        If .Visible Then
            Select Case KeyCode
                Case vbKeyReturn
'                    Call lstResult_KeyPress(vbKeyReturn)
'                Case vbKeyDown, vbKeyPageDown:
'                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
'                    .SetFocus
'                Case vbKeyUp, vbKeyPageUp:
'                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
'                    .SetFocus
                Case vbKeyEscape:
                    .Visible = False
            End Select
        End If
    End With
    
    With lstResult
        If .Visible Then
            Select Case KeyCode
                Case vbKeyReturn
'                    Call lstResult_KeyPress(vbKeyReturn)
'                Case vbKeyDown, vbKeyPageDown:
'                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
'                    .SetFocus
'                Case vbKeyUp, vbKeyPageUp:
'                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
'                    .SetFocus
                Case vbKeyEscape:
                    .Visible = False
            End Select
        End If
    End With
End Sub

Private Sub tblResult_KeyPress(KeyAscii As Integer)
    With tblResult
        If KeyAscii = vbKeyReturn And lstSpc.Visible Then
            DoEvents
'            Call lstSpc_MouseDown(1, 0, 0, 0)
            lstSpc.Visible = False
        End If
        
        If KeyAscii = vbKeyReturn And lstResult.Visible Then
            DoEvents
'            Call lstResult_MouseDown(1, 0, 0, 0)
            lstResult.Visible = False
        End If
    End With
End Sub

'** 결과등록(중간/최종)
Private Sub VerifyResult(ByVal strStatusCd As String)
    Dim strSql As String
    Dim i      As Long
    
    If Value_Check(strStatusCd) = False Then
        Exit Sub
    End If
    
    On Error GoTo ErrMsg
    
    '-------------------------------------------------------------------------------------------
    '전자서명 Validation Check

    Dim objESign        As clsLISElectronSign

    If strStatusCd = StsCd_LIS_FinRst Then
'        If P_MicElectronicSign Then
            Set objESign = New clsLISElectronSign
            If objESign.LoadElectronSign(HData.FvfyID, InstallDir & "LIS\bin") = False Then
                '전자서명 인증 에러
'                medBeep 20
                MsgBox objESign.ErrMsg, vbCritical, "전자서명 확인"
                Exit Sub
            Else
                '전자서명 인증
                objESign.ShowESignForm
                If objESign.ElectronSingOk = True Then
                Else
                    MsgBox "전자서명 인증을 하지 않으셨습니다.", vbInformation, "전자서명 인증"
                    Exit Sub
                End If
            End If
'        End If
        Set objESign = Nothing
    End If

    'bEsign = True

    '-------------------------------------------------------------------------------------------
    
    DBConn.BeginTrans
    
    Select Case strStatusCd
        Case StsCd_LIS_MidRst
            With HData
                '** Header Entry
                If clsSQL.LAB315_INSERT_UPDATE_Status(.WorkDt, .WorkDept, .ColDt, .ColTm) = False Then
                    '-- Insert
                    strSql = clsSQL.LAB315_INSERT_SQL(.WorkDt, .WorkDept, .ColDt, .ColTm, .TestMeth, _
                                                      .WaterFg, .WaterArea, .RstVal, .StsCd, _
                                                      .MvfyDt, .MvfyTm, .MvfyID, .FvfyDt, .FvfyTm, _
                                                      .FvfyID, .MfyDt, .MfyTm, .MfyID, .RptFg, _
                                                      .RptTxt, .RptDt, .RptTm, .RptID)
                     
                Else
                    '-- Update
                    strSql = clsSQL.LAB315_MUPDATE_SQL(.WorkDt, .WorkDept, .ColDt, .ColTm, .TestMeth, _
                                                       .WaterFg, .WaterArea, .RstVal, .StsCd, .MvfyDt, _
                                                       .MvfyTm, .MvfyID, .RptTxt)
                    
                End If
                
                Call DBConn.Execute(strSql)
                
            End With
            
        Case StsCd_LIS_FinRst
            With HData
                '** Header Entry
                If clsSQL.LAB315_INSERT_UPDATE_Status(.WorkDt, .WorkDept, .ColDt, .ColTm) = False Then
                    '-- Insert
                    strSql = clsSQL.LAB315_INSERT_SQL(.WorkDt, .WorkDept, .ColDt, .ColTm, .TestMeth, _
                                                      .WaterFg, .WaterArea, .RstVal, .StsCd, _
                                                      .FvfyDt, .FvfyTm, .FvfyID, .FvfyDt, .FvfyTm, _
                                                      .FvfyID, .MfyDt, .MfyTm, .MfyID, .RptFg, _
                                                      .RptTxt, .RptDt, .RptTm, .RptID)
                     
                Else
                    '-- Update
                    strSql = clsSQL.LAB315_FUPDATE_SQL(.WorkDt, .WorkDept, .ColDt, .ColTm, .TestMeth, _
                                                       .WaterFg, .WaterArea, .RstVal, .StsCd, .FvfyDt, _
                                                       .FvfyTm, .FvfyID, .RptTxt)
                End If
                
                Call DBConn.Execute(strSql)
                
            End With
        
    End Select
    
    '** Body Entry
    Dim strSpcCd    As String
    Dim strRstCd    As String
    Dim strRstCnt   As String
    
    If chkWater.Value = 1 Then
        With tblWater
            
            strSpcCd = txtSpcCd.Text
            
            For i = 1 To .DataRowCnt
                .Row = i
                
                .Col = 1: strRstCd = .Value
                .Col = 3: strRstCnt = .Value
                
                If clsSQL.LAB316_INSERT_UPDATE_Status(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd) = False Then
                    '-- Insert
                    strSql = clsSQL.LAB316_INSERT_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, strRstCnt, "", "")
                    
                Else
                    '-- Update
                    strSql = clsSQL.LAB316_UPDATE_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, strRstCnt, "", "")
                                                    
                End If
                
                Call DBConn.Execute(strSql)
                
            Next
        End With
    Else
        With tblResult
            For i = 1 To .DataRowCnt
                .Row = i
                
                .Col = 1: strSpcCd = .Value
                .Col = 2: strRstCd = .Value
                
                If clsSQL.LAB316_INSERT_UPDATE_Status(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd) = False Then
                    '-- Insert
                    strSql = clsSQL.LAB316_INSERT_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, "", "", "")
                    
                Else
                    '-- Update
                    strSql = clsSQL.LAB316_UPDATE_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, "", "", "")
                                                    
                End If
                
                Call DBConn.Execute(strSql)
                
            Next
        End With
    End If

Skip:

    DBConn.CommitTrans
    
    Call ClearData
    
    Exit Sub
    
ErrMsg:
    MsgBox "결과등록 시 오류가 발생하였습니다.", vbCritical
    DBConn.RollbackTrans
End Sub

'** 결과수정
Private Sub ModifyResult(ByVal strStatusCd As String)
    Dim strSql As String
    Dim i      As Long
    
    If Value_Check(strStatusCd) = False Then
        Exit Sub
    End If
    
    On Error GoTo ErrMsg
    
    '-------------------------------------------------------------------------------------------
    '전자서명 Validation Check

    Dim objESign        As clsLISElectronSign

    Set objESign = New clsLISElectronSign
    If objESign.LoadElectronSign(frmLogOn.EmpId, InstallDir & "LIS\bin") = False Then
        '전자서명 인증 에러
'                medBeep 20
        MsgBox objESign.ErrMsg, vbCritical, "전자서명 확인"
        Exit Sub
    Else
        '전자서명 인증
        objESign.ShowESignForm
        If objESign.ElectronSingOk = True Then
        Else
            MsgBox "전자서명 인증을 하지 않으셨습니다.", vbInformation, "전자서명 인증"
            Exit Sub
        End If
    End If
    
    Set objESign = Nothing
    '-------------------------------------------------------------------------------------------
    
    DBConn.BeginTrans
    
    With HData
        '** Header Entry
        If clsSQL.LAB315_INSERT_UPDATE_Status(.WorkDt, .WorkDept, .ColDt, .ColTm) = True Then
            '-- Update
            strSql = clsSQL.LAB315_Modify_SQL(.WorkDt, .WorkDept, .ColDt, .ColTm, .TestMeth, _
                                              .WaterFg, .WaterArea, .RstVal, .StsCd, .MfyDt, _
                                              .MfyTm, .MfyID, .RptTxt)
        Else
            GoTo ErrMsg
        End If
        
        Call DBConn.Execute(strSql)
        
    End With
    
    '** Body Entry
    Dim strSpcCd    As String
    Dim strRstCd    As String
    Dim strRstCnt   As String
    
    '-- Body Delete
    strSql = clsSQL.LAB316_DELETE_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                    HData.ColTm)
                                    
    Call DBConn.Execute(strSql)
    
    If chkWater.Value = 1 Then
        With tblWater
            
            strSpcCd = txtSpcCd.Text
            
            For i = 1 To .DataRowCnt
                .Row = i
                
                .Col = 2: strRstCd = .Value
                .Col = 3: strRstCnt = .Value
                
                If clsSQL.LAB316_INSERT_UPDATE_Status(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd) = False Then
                    '-- Insert
                    strSql = clsSQL.LAB316_INSERT_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, strRstCnt, "", "")
                    
                Else
                    '-- Update
                    strSql = clsSQL.LAB316_UPDATE_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, strRstCnt, "", "")
                                                    
                End If
                
                Call DBConn.Execute(strSql)
                
            Next
        End With
    Else
        With tblResult
            For i = 1 To .DataRowCnt
                .Row = i
                
                .Col = 1: strSpcCd = .Value
                .Col = 4: strRstCd = .Value
                
                If clsSQL.LAB316_INSERT_UPDATE_Status(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd) = False Then
                    '-- Insert
                    strSql = clsSQL.LAB316_INSERT_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, "", "", "")
                    
                Else
                    '-- Update
                    strSql = clsSQL.LAB316_UPDATE_SQL(HData.WorkDt, HData.WorkDept, HData.ColDt, _
                                                    HData.ColTm, strSpcCd, strRstCd, "", "", "")
                                                    
                End If
                
                Call DBConn.Execute(strSql)
                
            Next
        End With
    End If

Skip:

    DBConn.CommitTrans
    
    Call ClearData
    
    Exit Sub
    
ErrMsg:
    MsgBox "결과수정 시 오류가 발생하였습니다.", vbCritical
    DBConn.RollbackTrans
End Sub

'-- 입력오류Check
Private Function Value_Check(Optional ByVal strStatusCd As String = "") As Boolean
    Dim strHopeDt As String
    Dim strColDt  As String
    
    strHopeDt = Format(dptWorkDt.Value, "yyyymmdd")
    strColDt = Format(dptColDt.Value, "yyyymmdd")
    
    '-- 의뢰부서 Check
    If Trim(txtDeptCd.Text) = "" Then
        MsgBox "등록된 부서가 없습니다.", vbCritical
        lstDept.Visible = True
        lstDept.SetFocus
        Value_Check = False
        Exit Function
    End If
    
    If chkWater.Value = 0 Then
        '-- 의뢰검체 Check
        With tblResult
            If .DataRowCnt = 0 Then
                MsgBox "등록된 검체가 없습니다.", vbCritical
                'lstSpc.Visible = True
                'lstSpc.SetFocus
                Value_Check = False
                Exit Function
            End If
        End With
    End If
        
    '-- 채취일시와 의뢰일시 Check
    If strColDt < strHopeDt Then
        MsgBox "채취일시와 의뢰일시를 확인 하십시오.", vbCritical
        Value_Check = False
        Exit Function
    End If
    
    '-- Header 결과 값 Set
    With HData
        .WorkDt = Format(dptWorkDt.Value, "yyyymmdd")
        .WorkDept = Trim(txtDeptCd.Text)
        .ColDt = Format(dptColDt.Value, "yyyymmdd")
        .ColTm = Format(dptColDt.Value, "hhmmss")
        .TestMeth = rtfComment.Text
        If chkWater.Value = 1 Then
            .WaterFg = "1"
            .WaterArea = Trim(txtWaterArea.Text)
            .RstVal = cboResult.Text
        End If
        
        .StsCd = strStatusCd
        .RptTxt = rtfFootWard.Text
        
        Select Case strStatusCd
            Case StsCd_LIS_MidRst
                .MvfyID = frmLogOn.EmpId
                .MvfyDt = Format(Now, "yyyymmdd")
                .MvfyTm = Format(Now, "hhmmss")
                
            Case StsCd_LIS_FinRst
                .FvfyID = frmLogOn.EmpId
                .FvfyDt = Format(Now, "yyyymmdd")
                .FvfyTm = Format(Now, "hhmmss")
                
            Case StsCd_LIS_ModRst
                .MfyID = frmLogOn.EmpId
                .MfyDt = Format(Now, "yyyymmdd")
                .MfyTm = Format(Now, "hhmmss")
                
        End Select
    End With
    
    Value_Check = True
    
End Function

Private Sub tblResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim tmpTestCd As Variant
    Dim tmpSpcCd As Variant
    Dim tmpDate As Variant
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean

    If NewCol = 1 And lstSpc.Visible Then
        Cancel = True
        lstSpc.SetFocus
        Exit Sub
    End If
    If Col = 1 And lstSpc.ListIndex < 0 And lstSpc.Visible Then
        Cancel = True
        lstSpc.SetFocus
        Exit Sub
    End If

    If NewCol = 2 And lstResult.Visible Then
        Cancel = True
        lstResult.SetFocus
        Exit Sub
    End If
    If Col = 2 And lstResult.ListIndex < 0 And lstResult.Visible Then
        Cancel = True
        lstResult.SetFocus
        Exit Sub
    End If
    
'    If ActiveControl.Name = lstResult.Name Then Exit Sub

'    If Col = 2 Then lstResult.Visible = False
'
'    Select Case NewCol
'        Case 2:
'            If lstResult.Visible Then Exit Sub
'            With tblResult
'                .Row = NewRow: .Col = NewCol
'                Ret = .GetText(3, NewRow, tmpTestCd)
'                If tmpTestCd = "" Then Cancel = True: Exit Sub
'                'Ret = .GetText(4, NewRow, tmpSpcCd)
'                Ret = .GetCellPos(NewCol, NewRow + 1, X, Y, Wdt, Hgt)
'                If Y > 0 Then
'                    lstResult.Top = .Top + Y
'                Else
'                    Ret = .GetCellPos(NewCol, NewRow, X, Y, Wdt, Hgt)
'                    lstResult.Top = .Top + Y - lstResult.Height
'                End If
'                'Call objOrder.SpcList(tmpTestCd, lstResult)
'                lstResult.Visible = True
'                lstResult.ZOrder 0
'                lstResult.SetFocus
'                If lstResult.ListCount > 0 Then lstResult.ListIndex = 0
'                DoEvents
'            End With
'
'    End Select
End Sub

Private Sub tblReview_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strWorkDt       As String
    Dim strDeptCd       As String
    Dim strColDt        As String
    Dim strColTm        As String
    Dim strMvfyDt       As String
    Dim strMvfyTm       As String
    Dim strFvfyDt       As String
    Dim strFvfyTm       As String
    
    With tblReview
        If Row < 1 Or .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        .Row = Row
        
        .Col = 7
        dptWorkDt.Value = Format(.Value, "####-##-##"): strWorkDt = .Value
        
        .Col = 9: strColDt = .Value
        .Col = 10: strColTm = .Value
        dptColDt.Value = Format(strColDt & strColTm, "####-##-## 0#:##:##")
        
        .Col = 8: txtDeptCd.Text = .Value: strDeptCd = .Value
        .Col = 2: lblDeptNm.Caption = .Value
        
        .Col = 11
        If .Value = "1" Then
            chkWater.Value = 1
            .Col = 12: txtWaterArea.Text = .Value
            .Col = 13: cboResult.Text = .Value
            
            '-- Body Disp
            Call Mid_Water_Result(strWorkDt, strDeptCd, strColDt, strColTm, StsCd_LIS_MidRst)
            
        Else
            txtWaterArea.Text = ""
            cboResult.Text = "적합"
            chkWater.Value = 0
            
            '-- Body Disp
            Call Mid_Body_Result(strWorkDt, strDeptCd, strColDt, strColTm, StsCd_LIS_MidRst)
            
        End If
        
        .Col = 3: rtfComment.Text = .Value
        .Col = 23: rtfFootWard.Text = .Value
        
        .Col = 18: lblMvfyID.Caption = .Value
        
        .Col = 15: strMvfyDt = .Value
        .Col = 16: strMvfyTm = .Value
        lblMvfyDt.Caption = Format(strMvfyDt, "####-##-##") & " " & Format(strMvfyTm, "0#:##:##")
        
        .Col = 21: lblFvfy.Caption = .Value
        .Col = 22: lblFvfyID.Caption = .Value
        
        .Col = 19: strFvfyDt = .Value
        .Col = 20: strFvfyTm = .Value
        lblFvfyDt.Caption = Format(strFvfyDt, "####-##-##") & " " & Format(strFvfyTm, "0#:##:##")
        
        
    End With
    
    Call Enable_Check(False)
    cmdFResult.Enabled = False
    
    cmdDelete.Enabled = True
    
    cmdReport.Enabled = True
    cmdMfy.Enabled = True
    
End Sub

Private Sub tblWater_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim tmpIndex As Integer
    Dim tmpStr As String
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean

    With tblWater
        .Col = Col
        .Row = Row
        
        Select Case Col
            Case 1:
                tmpIndex = medListFind(lstResult, .Value)
                tmpStr = lstResult.List(tmpIndex)
                
                ' 문자가 입력될때마다 유사어 찾기
                
                If tmpIndex <> -1 Or UCase(tmpStr) = UCase(.Value) Then
                    Ret = .GetCellPos(Col, Row + 1, X, Y, Wdt, Hgt)
                    If .Height - Y < lstResult.Height Or Y < 0 Then
                        Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                        lstResult.Top = .Top + Y - lstResult.Height
                    Else
                        lstResult.Top = .Top + Y
                    End If
                    If tmpIndex >= 0 Then
                        medLockWindowUpdate (lstResult.hwnd)
                        
                        lstResult.ListIndex = tmpIndex
                        medLockWindowUpdate (0&)
                        If tmpIndex - lstResult.TopIndex > 10 Then lstResult.TopIndex = tmpIndex
                    End If
                    lstResult.Visible = True
                    'lstResult.SetFocus
                    lstResult.ZOrder 0
                Else
                    medLockWindowUpdate (lstResult.hwnd)
        
                    lstResult.ListIndex = tmpIndex
                    medLockWindowUpdate (0&)
                    Call lstResult_MouseDown(1, 0, 0, 0)
                    lstResult.Visible = False
                End If
        
                .Row = Row: .Col = Col
                If Trim(.Value) = "" Then
                    .Col = 3: .Value = ""
                    lstResult.Visible = False
                End If
        
                If Row = .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                
        End Select
    End With
End Sub

Private Sub tblWater_KeyDown(KeyCode As Integer, Shift As Integer)
    With lstResult
        If .Visible Then
            Select Case KeyCode
                Case vbKeyReturn
                    Call lstResult_KeyPress(vbKeyReturn)
'                Case vbKeyDown, vbKeyPageDown:
'                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
'                    .SetFocus
'                Case vbKeyUp, vbKeyPageUp:
'                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
'                    .SetFocus
                Case vbKeyEscape:
                    .Visible = False
            End Select
        End If
    End With
End Sub

Private Sub tblWater_KeyPress(KeyAscii As Integer)
    With tblResult
'        If KeyAscii = vbKeyReturn And lstSpc.Visible Then
'            DoEvents
'            Call lstSpc_MouseDown(1, 0, 0, 0)
'        End If
        
        If KeyAscii = vbKeyReturn And lstResult.Visible Then
            DoEvents
'            Call lstResult_MouseDown(1, 0, 0, 0)
            lstResult.Visible = False
        End If
    End With
End Sub

Private Sub tblWater_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim tmpTestCd As Variant
    Dim tmpSpcCd As Variant
    Dim tmpDate As Variant
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean

    If NewCol = 1 And lstResult.Visible Then
        Cancel = True
        lstResult.SetFocus
        Exit Sub
    End If
    
    If Col = 1 And lstResult.ListIndex < 0 And lstResult.Visible Then
        Cancel = True
        lstResult.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
    Dim strDeptCd   As String
    
    strDeptCd = Trim(txtDeptCd.Text)
    
    If strDeptCd = "" Then
        Exit Sub
    End If
    
    lblDeptNm.Caption = DeptName(UCase(strDeptCd))
    
End Sub

Private Sub txtDeptCd_LostFocus()
    Dim strDeptCd   As String
    
    strDeptCd = Trim(txtDeptCd.Text)
    
    If strDeptCd = "" Then
        Exit Sub
    End If
    
    lblDeptNm.Caption = DeptName(UCase(strDeptCd))
    
End Sub

Private Function DeptName(ByVal pDeptCd As String) As String
    Dim RS      As New ADODB.Recordset
    Dim strSql  As String
    
    strSql = clsSQL.Dept_Name(pDeptCd)
    
    RS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        DeptName = RS.Fields("deptnm").Value & ""
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

Private Sub txtTemp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        '-- 기존 데이터 확인
        If lvwTemplate.ListItems.Count <> 0 Then
            Call Template_Disp_Find
        End If
        
        rtfTemp.SetFocus
    End If
End Sub

Private Sub Template_Disp_Find()
    Dim i       As Integer
    Dim strTemp As String
    
    strTemp = Trim(txtTemp.Text)
    rtfTemp.Text = ""
    
    With lvwTemplate
        For i = 1 To .ListItems.Count
            strTemp = .ListItems.Item(i).Text
            
            If strTemp = Trim(txtTemp.Text) Then
                rtfTemp.Text = .ListItems.Item(i).ListSubItems(1).Text
            End If
        Next
    End With
End Sub
