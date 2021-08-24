VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm159RoundSchedule 
   BackColor       =   &H00DBE6E6&
   Caption         =   "병동 아침채혈 스케줄 작성"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "frm159.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   11400
   Visible         =   0   'False
   WindowState     =   2  '최대화
   Begin VB.PictureBox picPop 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   11760
      ScaleHeight     =   2370
      ScaleWidth      =   2415
      TabIndex        =   59
      Top             =   3900
      Visible         =   0   'False
      Width           =   2445
      Begin VB.CommandButton cmdApplyBypass 
         BackColor       =   &H00DBE6E6&
         Caption         =   "적용"
         Height          =   330
         Left            =   -15
         Style           =   1  '그래픽
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2490
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   2085
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "병동코드"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "병동명"
            Object.Width           =   1940
         EndProperty
      End
   End
   Begin VB.HScrollBar spnDay 
      Height          =   315
      Left            =   10050
      Max             =   2
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   300
      Width           =   480
   End
   Begin VB.HScrollBar spnMonth 
      Height          =   705
      Left            =   7080
      Max             =   2
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   15
      Width           =   480
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   53
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraSave 
      BackColor       =   &H00DBE6E6&
      Height          =   3315
      Left            =   7620
      TabIndex        =   10
      Top             =   690
      Width           =   6840
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "삭제(&D)"
         Height          =   420
         Left            =   5670
         Style           =   1  '그래픽
         TabIndex        =   67
         Top             =   765
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00DBE6E6&
         Caption         =   "추가(&N)"
         Height          =   420
         Left            =   4575
         Style           =   1  '그래픽
         TabIndex        =   66
         Top             =   345
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수정(&M)"
         Height          =   420
         Left            =   4575
         Style           =   1  '그래픽
         TabIndex        =   63
         Top             =   765
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저장(&S)"
         Height          =   420
         Left            =   5670
         Style           =   1  '그래픽
         TabIndex        =   62
         Top             =   345
         Width           =   1095
      End
      Begin VB.PictureBox picWard 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   4
         Left            =   4335
         ScaleHeight     =   450
         ScaleWidth      =   2430
         TabIndex        =   11
         Top             =   2820
         Width           =   2430
         Begin VB.CommandButton cmdWardPop 
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
            Height          =   345
            Index           =   4
            Left            =   2115
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":000C
            Style           =   1  '그래픽
            TabIndex        =   12
            Top             =   30
            Width           =   300
         End
         Begin MedControls1.LisLabel lblWard 
            Height          =   345
            Index           =   4
            Left            =   0
            TabIndex        =   13
            Top             =   30
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picWard 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   4335
         ScaleHeight     =   390
         ScaleWidth      =   2430
         TabIndex        =   14
         Top             =   2430
         Width           =   2430
         Begin VB.CommandButton cmdWardPop 
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
            Height          =   345
            Index           =   3
            Left            =   2115
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":0596
            Style           =   1  '그래픽
            TabIndex        =   15
            Top             =   30
            Width           =   300
         End
         Begin MedControls1.LisLabel lblWard 
            Height          =   345
            Index           =   3
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picWard 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   2
         Left            =   4335
         ScaleHeight     =   360
         ScaleWidth      =   2430
         TabIndex        =   17
         Top             =   2070
         Width           =   2430
         Begin VB.CommandButton cmdWardPop 
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
            Height          =   345
            Index           =   2
            Left            =   2115
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":0B20
            Style           =   1  '그래픽
            TabIndex        =   18
            Top             =   15
            Width           =   300
         End
         Begin MedControls1.LisLabel lblWard 
            Height          =   345
            Index           =   2
            Left            =   0
            TabIndex        =   19
            Top             =   15
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picWard 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   4335
         ScaleHeight     =   360
         ScaleWidth      =   2430
         TabIndex        =   20
         Top             =   1680
         Width           =   2430
         Begin VB.CommandButton cmdWardPop 
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
            Height          =   345
            Index           =   1
            Left            =   2115
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":10AA
            Style           =   1  '그래픽
            TabIndex        =   21
            Top             =   15
            Width           =   300
         End
         Begin MedControls1.LisLabel lblWard 
            Height          =   345
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Top             =   15
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picWard 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   4335
         ScaleHeight     =   360
         ScaleWidth      =   2430
         TabIndex        =   23
         Top             =   1305
         Width           =   2430
         Begin VB.CommandButton cmdWardPop 
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
            Height          =   345
            Index           =   0
            Left            =   2115
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":1634
            Style           =   1  '그래픽
            TabIndex        =   24
            Top             =   15
            Width           =   300
         End
         Begin MedControls1.LisLabel lblWard 
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Top             =   15
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1500
         ScaleHeight     =   315
         ScaleWidth      =   2565
         TabIndex        =   41
         Top             =   945
         Width           =   2595
         Begin VB.OptionButton optDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "병동분담"
            Height          =   300
            Index           =   1
            Left            =   1320
            TabIndex        =   42
            Top             =   30
            Width           =   1365
         End
         Begin VB.OptionButton optDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈자수"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   43
            Top             =   30
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin VB.TextBox txtCnt 
         Height          =   345
         Left            =   1500
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   1500
         ScaleHeight     =   375
         ScaleWidth      =   2715
         TabIndex        =   38
         Top             =   1320
         Width           =   2715
         Begin VB.TextBox txtEmpID 
            Height          =   345
            Index           =   0
            Left            =   -15
            TabIndex        =   2
            Top             =   0
            Width           =   1005
         End
         Begin VB.CommandButton cmdPopup 
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
            Height          =   345
            Index           =   0
            Left            =   990
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":1BBE
            Style           =   1  '그래픽
            TabIndex        =   39
            Top             =   0
            Width           =   300
         End
         Begin MedControls1.LisLabel lblEmpNm 
            Height          =   345
            Index           =   0
            Left            =   1305
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1500
         ScaleHeight     =   375
         ScaleWidth      =   2715
         TabIndex        =   35
         Top             =   1695
         Width           =   2715
         Begin VB.CommandButton cmdPopup 
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
            Height          =   345
            Index           =   1
            Left            =   990
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":2148
            Style           =   1  '그래픽
            TabIndex        =   36
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtEmpID 
            Height          =   345
            Index           =   1
            Left            =   -15
            TabIndex        =   3
            Top             =   0
            Width           =   1005
         End
         Begin MedControls1.LisLabel lblEmpNm 
            Height          =   345
            Index           =   1
            Left            =   1305
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   1500
         ScaleHeight     =   375
         ScaleWidth      =   2715
         TabIndex        =   32
         Top             =   2085
         Width           =   2715
         Begin VB.CommandButton cmdPopup 
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
            Height          =   345
            Index           =   2
            Left            =   990
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":26D2
            Style           =   1  '그래픽
            TabIndex        =   33
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtEmpID 
            Height          =   345
            Index           =   2
            Left            =   -15
            TabIndex        =   4
            Top             =   0
            Width           =   1005
         End
         Begin MedControls1.LisLabel lblEmpNm 
            Height          =   345
            Index           =   2
            Left            =   1305
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   1500
         ScaleHeight     =   375
         ScaleWidth      =   2715
         TabIndex        =   29
         Top             =   2460
         Width           =   2715
         Begin VB.CommandButton cmdPopup 
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
            Height          =   345
            Index           =   3
            Left            =   990
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":2C5C
            Style           =   1  '그래픽
            TabIndex        =   30
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtEmpID 
            Height          =   345
            Index           =   3
            Left            =   -15
            TabIndex        =   5
            Top             =   0
            Width           =   1005
         End
         Begin MedControls1.LisLabel lblEmpNm 
            Height          =   345
            Index           =   3
            Left            =   1305
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   4
         Left            =   1500
         ScaleHeight     =   405
         ScaleWidth      =   2715
         TabIndex        =   26
         Top             =   2850
         Width           =   2715
         Begin VB.CommandButton cmdPopup 
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
            Height          =   345
            Index           =   4
            Left            =   990
            MousePointer    =   14  '화살표와 물음표
            Picture         =   "frm159.frx":31E6
            Style           =   1  '그래픽
            TabIndex        =   27
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtEmpID 
            Height          =   345
            Index           =   4
            Left            =   -15
            TabIndex        =   6
            Top             =   0
            Width           =   1005
         End
         Begin MedControls1.LisLabel lblEmpNm 
            Height          =   345
            Index           =   4
            Left            =   1305
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   1515
         TabIndex        =   1
         Text            =   "채혈시간대"
         Top             =   555
         Width           =   2595
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "채혈자수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1323
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "채혈자1"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   3
         Left            =   90
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2085
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "채혈자3"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   4
         Left            =   90
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1704
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "채혈자2"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2466
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "채혈자4"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   5
         Left            =   90
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2850
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "채혈자5"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   7
         Left            =   90
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   942
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "업무분담"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Index           =   6
         Left            =   90
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   561
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "채혈시간"
         Appearance      =   0
      End
      Begin VB.ComboBox cboSaveTime 
         Height          =   300
         Left            =   1515
         TabIndex        =   64
         Top             =   555
         Width           =   2580
      End
      Begin VB.Label lblcnt 
         BackColor       =   &H00DBE6E6&
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   4125
         TabIndex        =   65
         Top             =   585
         Width           =   240
      End
   End
   Begin VB.TextBox txtSchdule 
      Appearance      =   0  '평면
      BackColor       =   &H00D1ECFC&
      Enabled         =   0   'False
      Height          =   2430
      Left            =   7620
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "frm159.frx":3770
      Top             =   4005
      Width           =   6825
   End
   Begin VB.TextBox txtmesg 
      Appearance      =   0  '평면
      BackColor       =   &H00D1ECFC&
      Height          =   1635
      Left            =   7620
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "frm159.frx":3776
      Top             =   6810
      Width           =   6825
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   345
      Index           =   8
      Left            =   7620
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6450
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   609
      BackColor       =   10392451
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
      Alignment       =   1
      Caption         =   "비고등록"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblMonth 
      Height          =   7680
      Left            =   75
      TabIndex        =   52
      Top             =   765
      Width           =   7470
      _Version        =   196608
      _ExtentX        =   13176
      _ExtentY        =   13547
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   16777215
      MaxCols         =   7
      MaxRows         =   12
      ScrollBars      =   0
      ShadowColor     =   12582911
      SpreadDesigner  =   "frm159.frx":377C
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "아침채혈 스케줄작성(             )  "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   750
      TabIndex        =   58
      Top             =   120
      Width           =   6510
   End
   Begin VB.Label lblDay 
      BackColor       =   &H00DBE6E6&
      Caption         =   "1998년 12월 28일 월요일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7695
      TabIndex        =   57
      Top             =   315
      Width           =   2310
   End
   Begin VB.Label lblMonth 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "1998년 12월"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4950
      TabIndex        =   56
      Top             =   165
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frm159.frx":4198
      Top             =   105
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   630
      Index           =   0
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   15
      Width           =   6960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  '단색
      Height          =   630
      Index           =   1
      Left            =   105
      Shape           =   4  '둥근 사각형
      Top             =   75
      Width           =   6975
   End
End
Attribute VB_Name = "frm159RoundSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents objMyList    As clspopuplist
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private ScheduleCnt(1 To 31)    As String
Private Today                   As Date
Private tmpDate1
Private RealDate

Private Sub ClearData()
    Dim ii As Integer
    
    txtCnt.Text = ""
    txtSchdule.Text = ""
    txtmesg.Text = ""
    For ii = 0 To 4
        txtEmpID(ii).Text = "": lblEmpNm(ii).Caption = "": lblWard(ii).Caption = ""
        pic(ii).Enabled = False: picWard(ii).Enabled = False
    Next
    lblCnt.Caption = ""
    cboTime.Visible = True
    cboSaveTime.Visible = False
    cmdSave.Enabled = True
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End Sub


Private Sub cboTime_LostFocus()
    Dim ii As Integer
    
    For ii = 0 To 4
        txtEmpID(ii).Text = "": lblEmpNm(ii).Caption = ""
        lblWard(ii).Caption = "":
        pic(ii).Enabled = False: picWard(ii).Enabled = False
    Next
    For ii = 0 To Val(txtCnt.Text) - 1
        txtEmpID(ii).Text = "": lblEmpNm(ii).Caption = ""
        lblWard(ii).Caption = "":
        pic(ii).Enabled = True
    Next
    
    optDiv(0).Value = True
End Sub

Private Sub cmdApplyBypass_Click()
    Dim iTmx    As ListItem
    Dim strTmp  As String
    
    If cmdApplyBypass.Tag < 4 Then
        If cmdWardPop(cmdApplyBypass.Tag + 1).Enabled Then
            cmdWardPop(cmdApplyBypass.Tag + 1).SetFocus
        Else
            txtmesg.SetFocus
        End If
    Else
        txtmesg.SetFocus
    End If
    
    
    For Each iTmx In lvw.ListItems
        If iTmx.Checked = True Then
            strTmp = strTmp & iTmx.Text & ","
        End If
    Next
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        lblWard(cmdApplyBypass.Tag).Caption = strTmp
    End If
    picPop.Visible = False
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub GetRoundTime()
    Dim SSQL As String
    Dim RS   As Recordset
    
    
    fraSave.Enabled = False
    SSQL = " SELECT * FROM " & T_LAB032 & " WHERE " & DBW("cdindex=", LC3_RoundTime)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    cboTime.Clear
    
    If Not RS.EOF Then
        Do Until RS.EOF
            cboTime.AddItem Format(RS.Fields("cdval1").Value & "", "0#:##") & " [ " & RS.Fields("field1").Value & "" & " ]"
            RS.MoveNext
        Loop
        cboTime.ListIndex = 0
        fraSave.Enabled = True
    Else
        MsgBox "아침채혈 시간대 설정을 하세요.", vbInformation + vbOKOnly, "Info        "
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub cmdModify_Click()
    Dim SSQL        As String
    Dim strColTm    As String
    Dim strColdt    As String
    Dim strBuss     As String
    Dim strCnt      As String
    Dim ii          As Integer

    strColdt = Format(tmpDate1, "YYYYMMDD")
    strColTm = Replace(medGetP(cboTime.Text, 1, " "), ":", "")
    
    strCnt = txtCnt.Text
    strBuss = IIf(optDiv(0).Value, "1", "2")
    If strCnt = "" Then Exit Sub

    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    SSQL = DeleteSQL(strColdt, strColTm)
    DBConn.Execute SSQL
    
    For ii = 0 To Val(strCnt) - 1
        SSQL = InsertSQL(strColdt, strColTm, txtEmpID(ii).Text, lblEmpNm(ii).Caption, strCnt, strBuss, lblWard(ii).Caption, txtmesg.Text)
        DBConn.Execute SSQL
    Next
    
    DBConn.CommitTrans
    Call ClearData
    Call GetCalendar(Today)
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub cmdNew_Click()
    Call ClearData
    cboTime.Visible = True: cboTime.ZOrder 0
    cboSaveTime.Visible = False
    cmdSave.Enabled = True
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
    Dim strColdt As String
    Dim strColTm As String
    Dim SSQL     As String
    
    strColdt = Format(tmpDate1, "YYYYMMDD")
    strColTm = Replace(medGetP(cboSaveTime.Text, 1, " "), ":", "")
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    SSQL = DeleteSQL(strColdt, strColTm)
    
    DBConn.Execute SSQL
    DBConn.CommitTrans
    Call GetCalendar(Today)
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox Err.Description
    
End Sub
Private Sub cmdSave_Click()
    Dim SSQL        As String
    Dim strColTm    As String
    Dim strColdt    As String
    Dim strBuss     As String
    Dim strCnt      As String
    Dim ii          As Integer
    
    
    If lblDay.Caption = "" Then
        MsgBox "스케줄 작성일을 선택하세요.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    strColdt = Format(tmpDate1, "YYYYMMDD")
    strColTm = Replace(medGetP(cboTime.Text, 1, " "), ":", "")
    
    strCnt = txtCnt.Text
    strBuss = IIf(optDiv(0).Value, "1", "2")
    If strCnt = "" Then Exit Sub

    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    For ii = 0 To Val(strCnt) - 1
        SSQL = InsertSQL(strColdt, strColTm, txtEmpID(ii).Text, lblEmpNm(ii).Caption, strCnt, strBuss, lblWard(ii).Caption, txtmesg.Text)
        DBConn.Execute SSQL
    Next
    DBConn.CommitTrans
    Call ClearData
    Call GetCalendar(Today)
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox Err.Description
    
End Sub
Private Function DeleteSQL(ByVal strColdt As String, ByVal strColTm As String) As String
    DeleteSQL = "delete FROM " & T_LAB901 & " WHERE " & DBW("coldt=", strColdt) & " AND " & DBW("coltm=", strColTm)
End Function
Private Function InsertSQL(ByVal strColdt As String, strColTm As String, ByVal sEmpID As String, _
                           ByVal sEmpNm As String, ByVal strCnt As String, _
                           ByVal strBuss As String, ByVal strWard As String, ByVal strMesg As String) As String
    Dim SSQL    As String
    
    SSQL = "insert into " & T_LAB901 & " (coldt,coltm,colid,empnm,entdt,entid,cnt,bussdiv,wardid,mesg) values (" & _
            DBV("coldt", strColdt, 1) & DBV("coltm", strColTm, 1) & DBV("colid", sEmpID, 1) & _
            DBV("empnm", sEmpNm, 1) & DBV("entdt", Format(GetSystemDate, "YYYYMMDD"), 1) & _
            DBV("entid", ObjSysInfo.EmpId, 1) & DBV("cnt", strCnt, 1) & DBV("bussdiv", strBuss, 1) & _
            DBV("wardid", strWard, 1) & DBV("mesg", strMesg) & ")"
    InsertSQL = SSQL
End Function

Private Sub cmdWardPop_Click(Index As Integer)
    Dim iTmx        As ListItem
    Dim ObjDic      As clsDictionary
    Dim aryTmp()    As String
    Dim strTmp      As String
    Dim lngTop      As Long
    Dim lngLeft     As Long
    Dim ii          As Long
'    Dim objData As clsBasisData
    Dim RS As Recordset
    
    
    lngTop = fraSave.Top + picWard(Index).Top + lblWard(Index).Top + picWard(Index).Height
    lngLeft = fraSave.Left + picWard(Index).Left + lblWard(Index).Left
    
    picPop.Visible = True
    picPop.Left = lngLeft
    picPop.Top = lngTop
    
    lvw.ListItems.Clear
    For ii = 0 To 4
        If ii <> Index Then strTmp = strTmp & lblWard(ii).Caption & ","
    Next
    
'    Set objData = New clsBasisData
    Set RS = New Recordset
    
    RS.Open GetSQLWardList, DBConn
    
    If strTmp <> "" Then
        Set ObjDic = New clsDictionary
        
        ObjDic.Clear
        ObjDic.FieldInialize "wardid", "wardnm,fg"
        
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        aryTmp() = Split(strTmp, ",")
        ObjDic.Sort = False
        
'        With objLisComCode.WardId
        With RS
'            .MoveFirst
            Do Until .EOF
                If Not ObjDic.Exists(.Fields("wardid").Value & "") Then
                    ObjDic.AddNew .Fields("wardid").Value & "", .Fields("wardnm").Value & "" & COL_DIV & "0"
                End If
                .MoveNext
            Loop
        End With
        
        
        With ObjDic
            For ii = LBound(aryTmp) To UBound(aryTmp)
                If .Exists(aryTmp(ii)) Then
                    .KeyChange aryTmp(ii)
                    .Fields("fg") = "1"
                End If
            Next
            
            .MoveFirst
            Do Until .EOF
                If .Fields("fg") <> "1" Then
                    Set iTmx = lvw.ListItems.Add(, , .Fields("wardid"))
                    iTmx.SubItems(1) = .Fields("wardnm")
                End If
                .MoveNext
            Loop
            
        End With
        Set ObjDic = Nothing
    Else
        With RS
'        With objLisComCode.WardId
            .MoveFirst
            Do Until .EOF
                Set iTmx = lvw.ListItems.Add(, , .Fields("wardid").Value & "")
                iTmx.SubItems(1) = .Fields("wardnm").Value & ""
                .MoveNext
            Loop
        End With
    End If
    
    cmdApplyBypass.Tag = Index
        
    Set RS = Nothing
'    Set objData = Nothing
End Sub


Private Sub cmdPopup_Click(Index As Integer)
    Dim SSQL    As String
'    Dim lngTop  As Long
'    Dim lngLeft As Long
    
'    lngTop = 2000 + fraSave.Top + pic(Index).Top + cmdPopup(Index).Top
'    lngLeft = fraSave.Left + pic(Index).Left + cmdPopup(Index).Left
    SSQL = " SELECT empid,empnm FROM " & T_COM006
    
'    Set objMyList = New clspopuplist
    Set objMyList = New clsPopUpList

    With objMyList
        .Connection = DBConn
        .FormCaption = "채혈자 조회"
        .ColumnHeaderText = "채혈자코드;채혈자명"
        .Tag = Index
        .LoadPopUp SSQL
'        .Caption = "채혈자 조회"
'        .HeadName = "채혈자코드,채혈자명"
'        .Tag = Index
'         Call .ListPop(SSQL, lngTop, lngLeft)
    End With
    Set objMyList = Nothing
End Sub

Private Sub Form_Load()
    Dim WeekDayKor  As String
    
    
    Call ClearData
    Erase ScheduleCnt
    Today = Format(GetSystemDate, "YY-MM-DD")
    RealDate = Today
    Call GetCalendar(Today)
    Call DisplayDate(Today)
    Call FindDay(Today)
    
    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", _
                          "화요일", "수요일", "목요일", "금요일", "토요일")
    spnMonth.Value = 1
    spnDay.Value = 1
    Call GetRoundTime
    
    
'    Dim iTmx As ListItem
'
'    lvw.ListItems.Clear
'    With ObjLISComCode.WardID
'        .MoveFirst
'        Do Until .EOF
'            Set iTmx = lvw.ListItems.Add(, , .Fields("wardid"))
'            iTmx.SubItems(1) = .Fields("wardnm")
'            .MoveNext
'        Loop
'    End With
    
End Sub

Private Sub spnDay_SpinDown()
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")
    
    
    ThisMonth = Month(Today)
    Today = DateAdd("y", -1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor

    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If

End Sub

Private Sub spnDay_SpinUp()
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")
    
    ThisMonth = Month(Today)
    Today = DateAdd("y", 1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor
    
    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If

End Sub

Private Sub spnMonth_SpinDown()
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")

    ThisMonth = Month(Today)
    Today = DateAdd("m", -1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor

    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If

End Sub

Private Sub spnMonth_SpinUp()
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")

    ThisMonth = Month(Today)
    Today = DateAdd("m", 1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor

    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If
End Sub

Private Sub objMyList_SendCode(ByVal SelString As String)
    Dim ii  As Integer
    
    If SelString <> "" Then
        txtEmpID(objMyList.Tag).Text = medGetP(SelString, 1, ";")
        lblEmpNm(objMyList.Tag).Caption = medGetP(SelString, 2, ";")
    Else
        txtEmpID(objMyList.Tag).Text = ""
        lblEmpNm(objMyList.Tag).Caption = ""
        txtEmpID(objMyList.Tag).SetFocus
    End If
    
    For ii = 0 To 4
        If ii <> objMyList.Tag Then
            If txtEmpID(ii).Text = txtEmpID(objMyList.Tag) Then
                MsgBox "중복되었습니다.확인하세요.", vbInformation + vbOKOnly, "Info"
                txtEmpID(objMyList.Tag).Text = ""
                lblEmpNm(objMyList.Tag).Caption = ""
                txtEmpID(objMyList.Tag).SetFocus
                Exit Sub
            End If
        End If
    Next
    If objMyList.Tag < 4 Then
        If txtEmpID(objMyList.Tag + 1).Enabled Then txtEmpID(objMyList.Tag + 1).SetFocus
    End If
'    SendKeys "{TAB}"
    
End Sub

Private Sub objMyList_SelectedItem(ByVal pSelectedItem As String)
    Dim i As Long
    
    If pSelectedItem <> "" Then
        txtEmpID(objMyList.Tag).Text = objMyList.SelectedItems(0)
        lblEmpNm(objMyList.Tag).Caption = objMyList.SelectedItems(1)
    Else
        txtEmpID(objMyList.Tag).Text = ""
        lblEmpNm(objMyList.Tag).Caption = ""
        txtEmpID(objMyList.Tag).SetFocus
    End If
    
    For i = 0 To 4
        If i <> objMyList.Tag Then
            If txtEmpID(i).Text = txtEmpID(objMyList.Tag) Then
                MsgBox "중복되었습니다.확인하세요.", vbInformation + vbOKOnly
                txtEmpID(objMyList.Tag).Text = ""
                lblEmpNm(objMyList.Tag).Caption = ""
                txtEmpID(objMyList.Tag).SetFocus
                Exit Sub
            End If
        End If
    Next
    If objMyList.Tag < 4 Then
        If txtEmpID(objMyList.Tag + 1).Enabled Then txtEmpID(objMyList.Tag + 1).SetFocus
    End If
End Sub

Private Sub optDiv_Click(Index As Integer)
    Dim ii As Integer
    
    If txtCnt.Text = "" Then Exit Sub
    If Not IsNumeric(txtCnt.Text) Then Exit Sub
    
    If Index = 0 Then
        For ii = 0 To Val(txtCnt.Text) - 1
            lblWard(ii).Caption = ""
            picWard(ii).Enabled = False
        Next
    Else
        For ii = 0 To Val(txtCnt.Text) - 1
            lblWard(ii).Caption = ""
            picWard(ii).Enabled = True
        Next
    End If
End Sub

Private Sub spnDay_Change()
   If spnDay.Value = 2 Then
      Call spnDay_SpinUp
   ElseIf spnDay.Value = 0 Then
      Call spnDay_SpinDown
   End If
   spnDay.Value = 1
End Sub

Private Sub spnMonth_Change()
   If spnMonth.Value = 2 Then
      Call spnMonth_SpinUp
   ElseIf spnMonth.Value = 0 Then
      Call spnMonth_SpinDown
   End If
   spnMonth.Value = 1
End Sub

Private Sub tblMonth_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim tmpDay
'
'    If Row = 0 Then Exit Sub
'    If Row Mod 2 = 1 Then Exit Sub
'
'    lblDay.Caption = ""
'    With tblMonth
'        If Row Mod 2 = 0 Then .Row = Row - 1
'        If Row Mod 2 <> 0 Then .Row = Row
'        .Col = Col: If .Value = "" Then Exit Sub
'
'        tmpDay = Format(Today, "MM") & " " & .Value & "," & Format(Today, "YYYY")
'        tmpDate1 = CDate(tmpDay)
''        tmpDay = .Value
'        lblDay.Caption = Format(tmpDate1, "Long Date")
'        If lblDay.Caption <> "" Then GetScheduleQuery (.Value)
'    End With
End Sub

Private Sub tblMonth_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim tmpDay
    If NewRow < 0 Then Exit Sub
    If NewRow = 0 Then Exit Sub
    If NewRow Mod 2 = 1 Then Exit Sub
    
    lblDay.Caption = ""
    With tblMonth
        If NewRow Mod 2 = 0 Then .Row = NewRow - 1
        If NewRow Mod 2 <> 0 Then .Row = NewRow
        .Col = NewCol: If .Value = "" Then Exit Sub
        tmpDay = Format(Today, "MM") & " " & .Value & "," & Format(Today, "YYYY")
        tmpDate1 = CDate(tmpDay)
'        tmpDay = .Value
        lblDay.Caption = Format(tmpDate1, "Long Date")
        
        If lblDay.Caption <> "" Then GetScheduleQuery (.Value)
    End With
    
End Sub

Private Sub GetScheduleQuery(ByVal sValue As String)
    Dim aryTmp()    As String
    Dim ii          As Integer
    
    cboTime.Visible = True
    Call ClearData
    If ScheduleCnt(sValue) = "" Then Exit Sub
    
    cmdSave.Enabled = False
    cmdModify.Enabled = True
    cmdDelete.Enabled = True
    aryTmp = Split(ScheduleCnt(sValue), COL_DIV)
    lblCnt.Caption = UBound(aryTmp) + 1
    cboTime.Visible = False: cboSaveTime.Visible = True
    cboSaveTime.Clear
    
    For ii = LBound(aryTmp) To UBound(aryTmp)
        cboSaveTime.AddItem aryTmp(ii)
    Next
    cboSaveTime.ListIndex = 0

End Sub

Private Sub cboSaveTime_Click()
    Dim RS       As Recordset
    Dim strColdt As String
    Dim strColTm As String
    Dim strTmp   As String
    Dim SSQL     As String
    Dim ii       As Integer
    
    strColdt = Format(tmpDate1, "YYYYMMDD")
    strColTm = Replace(medGetP(cboSaveTime.Text, 1, " "), ":", "")
    
    
    txtCnt.Text = "" ': lblcnt.Caption = ""
    
    For ii = 0 To 4
        lblWard(ii).Caption = "": lblEmpNm(ii).Caption = "": txtEmpID(ii).Text = ""
        pic(ii).Enabled = False:  picWard(ii).Enabled = False
    Next
    
    SSQL = " SELECT * FROM " & T_LAB901 & _
           " WHERE " & DBW("coldt=", strColdt) & _
           " AND " & DBW("coltm=", strColTm) & _
           " ORDER BY coldt,coltm"
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        txtCnt.Text = RS.Fields("cnt").Value & ""
        If RS.Fields("bussdiv").Value & "" = "1" Then
            optDiv(0).Value = True
        Else
            optDiv(1).Value = True
        End If
        
        txtmesg.Text = RS.Fields("mesg").Value & ""
        RS.MoveFirst
        For ii = 0 To Val(txtCnt.Text) - 1
            pic(ii).Enabled = True: picWard(ii).Enabled = True
        Next
        ii = 0
        Do Until RS.EOF
            txtEmpID(ii).Text = RS.Fields("colid").Value & ""
            lblEmpNm(ii).Caption = RS.Fields("empnm").Value & ""
            lblWard(ii).Caption = RS.Fields("wardid").Value & ""
            strTmp = strTmp & " 채혈자" & ii + 1 & "  : " & RS.Fields("empnm").Value & "" & vbTab & _
                              " [채혈병동 :" & RS.Fields("wardid").Value & "" & "]" & vbCRLF
            ii = ii + 1
            RS.MoveNext
        Loop
    End If
    
    
    txtSchdule.Text = " 채혈일자 : " & tmpDate1 & vbCRLF & _
                      " 채혈시간 : " & cboSaveTime.Text & vbCRLF & _
                    strTmp
    
    Set RS = Nothing
End Sub


Private Sub GetCalendar(ByVal datToday As Date)
    Dim tmpDate         As Date
    Dim ThisMonth       As Integer
    Dim ThisDay         As Integer
    Dim FirstWeekDay    As Integer
    Dim i               As Integer
    
    Call ClearData

    tmpDate = CDate(Format(datToday, "YYYY-MM-") & "01")
    ThisMonth = Month(tmpDate)
    FirstWeekDay = Weekday(tmpDate)
    
    With tblMonth
        .Row = -1: .Col = -1
        .BackColor = &HFFFFFF
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        Do While Month(tmpDate) = ThisMonth
            ThisDay = Day(tmpDate)
            .Row = (((ThisDay + FirstWeekDay - 2) \ 7) * 2) + 1
            .Col = (ThisDay + FirstWeekDay - 2) Mod 7 + 1
            Select Case .Col
                Case 6:     .ForeColor = &HFF0000
                Case 7:     .ForeColor = &HFF&
                Case Else:  .ForeColor = &H0&
            End Select
            .Value = ThisDay
            .Row = .Row + 1
            Call WriteSchedule(Format(tmpDate, "YYYYMMDD"), ThisDay, .Row, .Col)
            
'            If ScheduleCnt(ThisDay) <> 0 Then
'                .Text = ScheduleCnt(ThisDay)
'            End If
            
            tmpDate = DateAdd("d", 1, tmpDate)
        Loop
    End With
    
    Call FindDay(Today)
End Sub

Private Sub WriteSchedule(ByVal sColDt As String, ByVal ThisDay As String, ByVal lngRow As Long, ByVal lngCol As Long)
    Dim SSQL    As String
    Dim RS      As Recordset
    
    ScheduleCnt(ThisDay) = ""
    SSQL = " SELECT distinct a.coltm ,b.field1 FROM " & T_LAB032 & " b," & T_LAB901 & " a " & _
           " WHERE " & DBW("a.coldt=", sColDt) & _
           " AND " & DBW("b.cdindex=", LC3_RoundTime) & _
           " AND a.coltm=b.cdval1 " & _
           " ORDER BY coltm"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        With tblMonth
            .Row = lngRow
            .Col = lngCol
            Do Until RS.EOF
                ScheduleCnt(ThisDay) = ScheduleCnt(ThisDay) & Format(RS.Fields("coltm").Value & "", "0#:##") & " [ " & _
                                                              RS.Fields("field1").Value & "" & " ]" & COL_DIV
                .Value = "작성완료": .ForeColor = DCM_LightRed
                
                RS.MoveNext
            Loop
            ScheduleCnt(ThisDay) = Mid(ScheduleCnt(ThisDay), 1, Len(ScheduleCnt(ThisDay)) - 1)
        End With
    End If
    Set RS = Nothing
End Sub



Private Sub DisplayDate(ByVal datToday As Date)
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(datToday), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")
    lblMonth = Format(datToday, "YYYY년 MM월")
    lblDay = Format(datToday, "YYYY년 MM월 DD일 ") & WeekDayKor

End Sub

Private Sub FindDay(ByVal NewDate As Date)
    Dim ii As Integer
    Dim jj  As Integer
    
    With tblMonth
        For ii = 1 To 9 Step 2
            .Row = ii
            For jj = 1 To 7
                .Col = jj
                If Format(Today, "yyMM") & Format(.Value, "00") = Format(RealDate, "yymmdd") Then
                    .BackColor = &HFFFFC0
                End If
            Next jj
        Next ii
    
    End With
End Sub


Private Sub txtCnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCnt_LostFocus()
    Dim ii As Integer
    
    
    If txtCnt.Text = "" Then Exit Sub
    If Not IsNumeric(txtCnt.Text) Then
        txtCnt.Text = ""
        Exit Sub
    End If
    
    If Val(txtCnt.Text) > 5 Then
        MsgBox "채혈자는 5인 이내로 작성하셔야 합니다.", vbInformation + vbOKOnly, "Info"
        txtCnt.Text = ""
        Exit Sub
    End If
End Sub


Private Sub txtEmpID_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEmpID_LostFocus(Index As Integer)
'    Dim objData As clsBasisData
    
    If txtEmpID(Index).Text = "" Then Exit Sub
    
'    Set objData = New clsBasisData
    
    lblEmpNm(Index).Caption = GetEmpNm(txtEmpID(Index).Text) ' GetEmpName(txtEmpID(Index).Text)
'    Set objData = Nothing
    
    If lblEmpNm(Index).Caption = "" Then
        txtEmpID(Index).Text = "": txtEmpID(Index).SetFocus
    End If
End Sub
