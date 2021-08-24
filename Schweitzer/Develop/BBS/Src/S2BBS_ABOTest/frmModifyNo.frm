VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmModifyNo 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "ABO결과 수정"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel13 
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6600
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Remark"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2655
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  검사 결과"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel12 
      Height          =   315
      Left            =   1890
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4470
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Comment by Accession No"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   1890
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1155
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  환자 정보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   1890
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   225
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  접수 번호(검체 번호)"
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   645
      Left            =   1890
      TabIndex        =   45
      Top             =   450
      Width           =   9945
      Begin VB.CheckBox chkBarcode 
         BackColor       =   &H00FFC0C0&
         Caption         =   "바코드리더로 읽기"
         Height          =   225
         Left            =   8040
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   255
         Width           =   1800
      End
      Begin VB.TextBox txtAccNo 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   1515
         TabIndex        =   46
         Text            =   "123456789012"
         Top             =   180
         Width           =   1590
      End
      Begin MedControls1.LisLabel lblBarNo 
         Height          =   360
         Left            =   4485
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   635
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
         Caption         =   "01-031028-1234"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAccNo 
         Height          =   360
         Left            =   315
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
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
      Begin MedControls1.LisLabel lblBarcode 
         Height          =   360
         Left            =   3285
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "바코드 번호"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1530
      Left            =   1890
      TabIndex        =   25
      Top             =   2880
      Width           =   9945
      Begin VB.OptionButton optType 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Back Typing"
         Height          =   420
         Index           =   1
         Left            =   1875
         Style           =   1  '그래픽
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   165
         Width           =   1695
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Front Typing"
         Height          =   420
         Index           =   0
         Left            =   180
         Style           =   1  '그래픽
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txtABO 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   1800
         TabIndex        =   33
         Text            =   "123456789012"
         Top             =   675
         Width           =   1095
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   0
         Left            =   2910
         Picture         =   "frmModifyNo.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   675
         Width           =   330
      End
      Begin VB.TextBox txtABOSub 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   1800
         TabIndex        =   31
         Text            =   "123456789012"
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   2
         Left            =   2910
         Picture         =   "frmModifyNo.frx":00B2
         Style           =   1  '그래픽
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1110
         Width           =   330
      End
      Begin VB.TextBox txtRh 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   6705
         TabIndex        =   29
         Text            =   "123456789012"
         Top             =   675
         Width           =   1095
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   1
         Left            =   7815
         Picture         =   "frmModifyNo.frx":0164
         Style           =   1  '그래픽
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   675
         Width           =   330
      End
      Begin VB.TextBox txtRhSub 
         Alignment       =   2  '가운데 맞춤
         Height          =   315
         Left            =   6705
         TabIndex        =   27
         Text            =   "123456789012"
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   3
         Left            =   7815
         Picture         =   "frmModifyNo.frx":0216
         Style           =   1  '그래픽
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1110
         Width           =   330
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   180
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   675
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
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
         Caption         =   "ABO 결과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   180
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
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
         Caption         =   "ABO SubType"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   5100
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
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
         Caption         =   "Rh SubType"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   315
         Left            =   5115
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   675
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
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
         Caption         =   "Rh 결과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   315
         Left            =   3255
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   675
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblABOSub 
         Height          =   315
         Left            =   3255
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRh 
         Height          =   315
         Left            =   8160
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   675
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRhSub 
         Height          =   315
         Left            =   8160
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoubleCheck 
         Height          =   315
         Left            =   6930
         TabIndex        =   51
         Top             =   225
         Visible         =   0   'False
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         BackColor       =   4194304
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "Double Check Complete"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   24
      Tag             =   "128"
      Top             =   7785
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   23
      Tag             =   "124"
      Top             =   7785
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "수정(&S)"
      Height          =   510
      Left            =   7830
      Style           =   1  '그래픽
      TabIndex        =   22
      Tag             =   "15101"
      Top             =   7785
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   1890
      TabIndex        =   9
      Top             =   1380
      Width           =   9945
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1290
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   4545
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   330
         Left            =   7920
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDept 
         Height          =   330
         Left            =   1290
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWard 
         Height          =   330
         Left            =   4545
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoct 
         Height          =   330
         Left            =   7920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   6900
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "성/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   6900
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "주치의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3525
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "환자명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   3525
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "병동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   270
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   1890
      TabIndex        =   5
      Tag             =   "20003"
      Top             =   4695
      Width           =   9945
      Begin VB.TextBox txtComment 
         Height          =   1485
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   7
         Top             =   225
         Width           =   9270
      End
      Begin VB.CommandButton cmdCommentTemplete 
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
         Left            =   9495
         Picture         =   "frmModifyNo.frx":02C8
         Style           =   1  '그래픽
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1380
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   705
      Left            =   1890
      TabIndex        =   3
      Top             =   6825
      Width           =   9945
      Begin VB.ComboBox cboRemark 
         Height          =   300
         Left            =   150
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   240
         Width           =   9690
      End
   End
End
Attribute VB_Name = "frmModifyNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum KindOfTyping
    NoUse = -1
    AnyUse = 0
    FrontUse = 1
    BackUse = 2
End Enum

Public Event FormClose()

Private WithEvents objPopList As clsPopUpList
Attribute objPopList.VB_VarHelpID = -1
Private WithEvents objComment As frmTempSearch
Attribute objComment.VB_VarHelpID = -1

Private Sub chkBarcode_Click()
    If chkBarcode.Value = 1 Then
        lblAccNo.Caption = "바코드 번호"
        lblBarcode.Caption = "접수 번호"
    Else
        lblAccNo.Caption = "접수 번호" '
        lblBarcode.Caption = "바코드 번호"
    End If

    txtAccNo.Text = ""
    lblBarNo.Caption = ""

    Call ClearPtInfo
    Call ClearOthers

    On Error Resume Next
    txtAccNo.SetFocus
End Sub

Private Sub cmdClear_Click()
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    Call ClearPtInfo
    Call ClearOthers
    On Error Resume Next
    txtAccNo.SetFocus
End Sub

Private Sub cmdCommentTemplete_Click()
    Set objComment = New frmTempSearch
    
    With objComment
        .CurrentComment = txtComment.Text
        .Show vbModal
    End With
            
    Set objComment = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub

Private Sub cmdSave_Click()
    Dim objABOSql As clsABOSql
    Dim strRemark As String
    Dim strComment As String
'    Dim strVfyDt As String
'    Dim strVfyTm As String
    Dim blnSave As Boolean
    Dim strWorkarea As String
    Dim strAccdt As String
    Dim strAccseq As String
    
    If CheckValidation = False Then Exit Sub

    If cboRemark.ListIndex = 0 Or cboRemark.ListIndex = -1 Then
        strRemark = ""
    Else
        strRemark = medGetP(cboRemark.Text, 1, vbTab)
    End If
    strComment = txtComment.Text
   
    If chkBarcode.Value = 0 Then
        strWorkarea = medGetP(Trim(txtAccNo.Text), 1, "-")
        strAccdt = IIf(medGetP(Trim(txtAccNo.Text), 2, "-") Like "99*", _
                      "19" & medGetP(Trim(txtAccNo.Text), 2, "-"), _
                      "20" & medGetP(Trim(txtAccNo.Text), 2, "-"))
        strAccseq = medGetP(Trim(txtAccNo.Text), 3, "-")
    Else
        strWorkarea = medGetP(Trim(lblBarNo.Caption), 1, "-")
        strAccdt = IIf(medGetP(Trim(lblBarNo.Caption), 2, "-") Like "99*", _
                      "19" & medGetP(Trim(lblBarNo.Caption), 2, "-"), _
                      "20" & medGetP(Trim(lblBarNo.Caption), 2, "-"))
        strAccseq = medGetP(Trim(lblBarNo.Caption), 3, "-")
    End If

    Set objABOSql = New clsABOSql
    blnSave = objABOSql.ModifyABOResult(strWorkarea, strAccdt, strAccseq, _
            IIf(optType(0).Value, 0, 1), _
            txtABO.Text, txtRh.Text, txtABOSub.Text, txtRhSub.Text)
    Set objABOSql = Nothing
    
    If blnSave Then
        MsgBox "정상적으로 처리되었습니다.", vbInformation
        Call cmdClear_Click
    Else
        MsgBox "처리도중 오류가 발생하였습니다.", vbExclamation
        Call cmdClear_Click
    End If
End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = False
    
    If txtAccNo.Text = "" Then
        MsgBox "접수번호(검체번호)를 입력하십시오.", vbExclamation
        Exit Function
    End If
    
    If optType(0).Value = False And optType(1).Value = False Then
        MsgBox "결과입력 타입을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If txtABO.Text = "" And txtABOSub.Text = "" Then
        MsgBox "ABO결과를 입력하십시오.", vbExclamation
        Exit Function
    End If

    If txtRh.Text = "" And txtRhSub.Text = "" Then
        MsgBox "RH결과를 입력하십시오.", vbExclamation
        Exit Function
    End If

    If txtABO.Text <> "" And txtABOSub.Text <> "" Then
        MsgBox "ABO와 Subtype중 하나만 입력하십시오.", vbExclamation
        Exit Function
    End If

    If txtRh.Text <> "" And txtRhSub.Text <> "" Then
        MsgBox "RH와 Subgroup중 하나만 입력하십시오.", vbExclamation
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Sub Form_Activate()
    On Error Resume Next
    txtAccNo.SetFocus
End Sub

Private Sub Form_Load()
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    Call ClearPtInfo
    Call ClearOthers
    
'    Call LoadTestCd
    Call LoadRemark(cboRemark)
End Sub

Private Sub ClearPtInfo()
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblDept.Caption = ""
    lblWard.Caption = ""
    lblDoct.Caption = ""
End Sub

Private Sub ClearOthers()
    optType(0).Value = True
    optType(1).Value = False
    optType(0).Enabled = True
    optType(1).Enabled = True
    lblDoubleCheck.Visible = False
    txtABO.Text = "": lblABO.Caption = ""
    txtRh.Text = "": lblRh.Caption = ""
    txtABOSub.Text = "": lblABOSub.Caption = ""
    txtRhSub.Text = "": lblRHSub.Caption = ""
    txtComment.Text = ""
    cboRemark.ListIndex = -1
End Sub

Private Sub objComment_Selected(ByVal vSelectedComment As String)
    If vSelectedComment <> "" Then txtComment.Text = vSelectedComment
End Sub

Private Sub txtAccNo_Change()
'화면 지움
    Dim lngLen As Long
    Static lngAccDt As Long

    On Error Resume Next
    If Screen.ActiveControl.Name <> txtAccNo.Name Then Exit Sub

    If chkBarcode.Value = 0 Then    '접수번호로 입력할때만 유효
        With txtAccNo
            lngLen = Len(Trim(.Text))

            If lngLen < 2 Then
                lngAccDt = 0
            End If

            If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)

                lngAccDt = GetLenOfAccDt(Mid(txtAccNo.Text, 1, 2))
            End If

            If lngLen > 2 And lngLen = lngAccDt + 3 And lngAccDt <> 0 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    End If

    If lblPtId.Caption <> "" Then
        lblBarNo.Caption = ""
        Call ClearPtInfo
        Call ClearOthers
    End If
End Sub

Private Sub txtAccNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtAccNo.Text) = "" Then Exit Sub

    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    If chkBarcode.Value = 0 Then '접수번호를 입력할때만 유효
        If KeyAscii = vbKeyBack Then
            With txtAccNo
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    If Len(.Text) > 2 Then
                        .Text = Mid(.Text, 1, Len(.Text) - 2)
                        .SelStart = Len(.Text)
                        KeyAscii = 0
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub txtAccNo_Validate(Cancel As Boolean)
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    Dim strAccNo As String

    If txtAccNo.Text = "" Then Exit Sub

    If chkBarcode.Value = 1 Then
        strAccNo = GetAccNo
        lblBarNo.Caption = strAccNo
    Else
        strAccNo = Trim(txtAccNo.Text)
        lblBarNo.Caption = GetSpcNo
    End If

    If strAccNo = "" Then
        MsgBox "유효한 바코드 정보가 아닙니다.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If

    If lblBarNo.Caption = "" Then
        MsgBox "유효한 접수번호가 아닙니다.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If

    Dim strWorkarea As String
    Dim strAccdt As String
    Dim strAccseq As String

    strWorkarea = medGetP(strAccNo, 1, "-")
    strAccdt = medGetP(strAccNo, 2, "-")
    strAccdt = IIf(strAccdt Like "99*", "19" & strAccdt, "20" & strAccdt)
    strAccseq = medGetP(strAccNo, 3, "-")

    Set objABOSql = New clsABOSql

    '접수내역 읽기
    Set Rs = New Recordset
    Set Rs = objABOSql.GetAccessInfo(strWorkarea, strAccdt, strAccseq)

    '값을 Display-------------------------------------------------------
    lblPtId.Caption = Rs.Fields("ptid").Value & ""
    lblPtNm.Caption = GetPtNm(Rs.Fields("ptid").Value & "")
    lblSexAge.Caption = Rs.Fields("sex").Value & "" & "/" & (Val(Rs.Fields("ageday").Value & "") \ 365)
    lblDept.Caption = GetDeptNm(Rs.Fields("deptcd").Value & "")
    lblWard.Caption = GetWardNm(Rs.Fields("wardid").Value & "")
    lblDoct.Caption = GetDoctNm(Rs.Fields("majdoct").Value & "")

    '상태 읽기
    If Val(Rs.Fields("stscd").Value & "") < 2 Then
        MsgBox "접수되지 않은 접수번호(검체번호)입니다.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If

    If Val(Rs.Fields("stscd").Value & "") < 3 Then
        MsgBox "결과가 등록되지 않았습니다.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If

    '리마크-------------------------------------------------------------
    cboRemark.ListIndex = -1
    If Rs.Fields("rmkcd").Value & "" & "" <> "" Then
        cboRemark.ListIndex = medComboFind(cboRemark, Rs.Fields("rmkcd").Value & "" & "")
    End If
    Set Rs = Nothing

    '코멘트 읽기
    Set Rs = New Recordset
    Set Rs = objABOSql.GetAccComment(strWorkarea, strAccdt, strAccseq)

    If Rs.EOF = False Then
        txtComment.Text = Rs.Fields("rsttxt").Value & ""
    End If
    Set Rs = Nothing

    '혈액형 검사가 있는 지 여부 체크
    If objABOSql.IsExistABO(strWorkarea, strAccdt, strAccseq) = False Then '검사가 없는 경우..
        MsgBox "혈액형 검사가 없습니다.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If

    '이미 검사가 완료되었는지 검사한다.--------------------------------
    Select Case CanModifyABOResult(strWorkarea, strAccdt, strAccseq)
        Case KindOfTyping.NoUse 'BBS303 결과내역이 없는 경우
            Cancel = True
            GoTo SetMe
'        Case KindOfTyping.AnyUse
'            optType(0).Enabled = True
'            optType(1).Enabled = True
        Case KindOfTyping.FrontUse
            optType(0).Value = True
            optType(0).Enabled = False
            optType(1).Enabled = False
        Case KindOfTyping.BackUse
            optType(1).Value = True
            optType(0).Enabled = False
            optType(1).Enabled = False
    End Select

    Set objABOSql = Nothing

    txtABO.SetFocus
    Exit Sub

SetMe:
    Set Rs = Nothing
    Set objABOSql = Nothing
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function GetAccNo() As String
    Dim Rs As Recordset
    Dim strSQL As String
    Dim strSpcYY As String
    Dim strSpcNo As Long

    strSpcYY = Mid(Trim(txtAccNo.Text), 1, P_SpcYyLength)
    If Mid(Trim(txtAccNo.Text), P_SpcYyLength + 1, P_SpcNoLength) <> "" Then
        strSpcNo = Format(Mid(Trim(txtAccNo.Text), P_SpcYyLength + 1, P_SpcNoLength), "#0")
    End If

    strSQL = " select vWorkarea,vAccdt,vAccseq from " & T_LAB201 & _
             " where " & DBW("spcyy=", strSpcYY) & _
             " and " & DBW("spcno=", strSpcNo)

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn

    If Rs.EOF Then
        GetAccNo = ""
    Else
        GetAccNo = Rs.Fields("workarea").Value & "" & "-" & _
                   Mid(Rs.Fields("accdt").Value, 3) & "" & "-" & _
                   Rs.Fields("accseq").Value & ""
    End If

    Set Rs = Nothing
End Function

Private Function GetSpcNo() As String
    Dim Rs As Recordset
    Dim strSQL As String
    Dim strWorkarea As String
    Dim strAccdt As String
    Dim strAccseq As String

    strWorkarea = medGetP(Trim(txtAccNo.Text), 1, "-")
    strAccdt = IIf(medGetP(Trim(txtAccNo.Text), 2, "-") Like "99*", _
                  "19" & medGetP(Trim(txtAccNo.Text), 2, "-"), _
                  "20" & medGetP(Trim(txtAccNo.Text), 2, "-"))
    strAccseq = medGetP(Trim(txtAccNo.Text), 3, "-")

    strSQL = " select spcyy,spcno from " & T_LAB201 & _
             " where " & DBW("workarea=", strWorkarea) & _
             " and " & DBW("accdt=", strAccdt) & _
             " and " & DBW("accseq=", strAccseq)

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn

    If Rs.EOF Then
        GetSpcNo = ""
    Else
        GetSpcNo = Rs.Fields("spcyy").Value & "" & Format(Rs.Fields("spcno").Value & "", LIS_BarFormat)
    End If

    Set Rs = Nothing
End Function

Private Function CanModifyABOResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Long
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset

    Dim vfydt  As String    '최종 Verify되면 set된다.
    Dim strVfyid1 As String    'Front Typing 사용자
    Dim strVfydt1 As String    'Front Typing 일자
    Dim strVfytm1 As String    'Front Typing 시간
    Dim strVfyid2 As String    'Back  Typing 사용자
    Dim strVfydt2 As String    'Back  Typing 일자
    Dim strVfytm2 As String    'Back  Typing 시간

    Set objABOSql = New clsABOSql
    Set Rs = objABOSql.GetABOResultInfo(vWorkarea, vAccdt, vAccseq)

    If Rs.EOF Then
        Set Rs = Nothing
        Set objABOSql = Nothing
        CanModifyABOResult = False
        MsgBox "입력할 혈액형 검사가 없습니다.", vbExclamation
        Exit Function
    End If

    With Rs
        vfydt = .Fields("vfydt").Value & ""

        strVfyid1 = .Fields("vfyid1").Value & ""
        strVfydt1 = .Fields("vfydt1").Value & ""
        strVfytm1 = .Fields("vfytm1").Value & ""

        strVfyid2 = .Fields("vfyid2").Value & ""
        strVfydt2 = .Fields("vfydt2").Value & ""
        strVfytm2 = .Fields("vfytm2").Value & ""

        If strVfydt1 <> "" And strVfydt2 <> "" Then '결과를 모두 등록한 경우
            lblDoubleCheck.Visible = True
            
            If strVfyid1 = ObjMyUser.EmpID Then 'empid가 Front 수정 가능
                CanModifyABOResult = KindOfTyping.FrontUse
                
                txtABO.Text = .Fields("abo1").Value & ""
                lblABO.Caption = GetABOBackNm(.Fields("abo1").Value & "")
                
                txtRh.Text = .Fields("rh1").Value & ""
                lblRh.Caption = GetRHNM(.Fields("rh1").Value & "")
                
                txtABOSub.Text = .Fields("abosub").Value & ""
                lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                
                txtRhSub.Text = .Fields("rhsub").Value & ""
                lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
            ElseIf strVfyid2 = ObjMyUser.EmpID Then 'empid가 Back 수정 가능
                CanModifyABOResult = KindOfTyping.BackUse
                
                txtABO.Text = .Fields("abo2").Value & ""
                lblABO.Caption = GetABOBackNm(.Fields("abo2").Value & "")
                
                txtRh.Text = .Fields("rh2").Value & ""
                lblRh.Caption = GetRHNM(.Fields("rh2").Value & "")
                
                txtABOSub.Text = .Fields("abosub").Value & ""
                lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                
                txtRhSub.Text = .Fields("rhsub").Value & ""
                lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
            Else '수정 불가능
                Set Rs = Nothing
                Set objABOSql = Nothing
                CanModifyABOResult = KindOfTyping.NoUse
                MsgBox "다른 사용자에 의해 결과가 등록되었습니다.", vbExclamation
                Exit Function
            End If
        ElseIf strVfydt1 <> "" And strVfydt2 = "" Then 'Front 만 결과가 등록된 경우
            If strVfyid1 <> ObjMyUser.EmpID Then   '결과등록자와  empid가 다른 경우 결과수정 불가처리
                Set Rs = Nothing
                Set objABOSql = Nothing
                CanModifyABOResult = KindOfTyping.NoUse
                MsgBox "다른 사용자에 의해 Front Type결과가 등록되었습니다. 결과 등록: " & GetEmpNm(strVfyid1), vbExclamation
            Else '이미 등록한 결과 표시
                CanModifyABOResult = KindOfTyping.FrontUse
                
                txtABO.Text = .Fields("abo1").Value & ""
                lblABO.Caption = GetABOBackNm(.Fields("abo1").Value & "")
                
                txtRh.Text = .Fields("rh1").Value & ""
                lblRh.Caption = GetRHNM(.Fields("rh1").Value & "")
                
                txtABOSub.Text = .Fields("abosub").Value & ""
                lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                
                txtRhSub.Text = .Fields("rhsub").Value & ""
                lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
            End If
        ElseIf strVfydt1 = "" And strVfydt2 <> "" Then 'Back만 결과가 등록된 경우
            If strVfyid2 <> ObjMyUser.EmpID Then '결과등록자와 empid가 다른 경우 결과수정 불가처리
                Set Rs = Nothing
                Set objABOSql = Nothing
                CanModifyABOResult = KindOfTyping.NoUse
                MsgBox "다른 사용자에 의해 Back Type결과가 등록되었습니다. 결과 등록: " & GetEmpNm(strVfyid2), vbExclamation
            Else '이미 등록한 결과표시
                CanModifyABOResult = KindOfTyping.BackUse
                
                txtABO.Text = .Fields("abo12").Value & ""
                lblABO.Caption = GetABOBackNm(.Fields("abo2").Value & "")
                
                txtRh.Text = .Fields("rh2").Value & ""
                lblRh.Caption = GetRHNM(.Fields("rh2").Value & "")
                
                txtABOSub.Text = .Fields("abosub").Value & ""
                lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                
                txtRhSub.Text = .Fields("rhsub").Value & ""
                lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
            End If
        ElseIf strVfydt1 = "" And strVfydt2 = "" Then '결과를 등록하지 않은 경우(입력했는지는 나중에 판단)
            Set Rs = Nothing
            Set objABOSql = Nothing
            CanModifyABOResult = KindOfTyping.NoUse
            MsgBox "Front/Back Type결과가 등록되어 있지 않습니다. 결과등록 화면을 사용하십시오.", vbExclamation
            Exit Function
        End If
    End With
'
'        If strVfyid1 = "" And strVfyid2 = "" Then
'            MsgBox "아직 결과가 등록된 적이 없습니다. 결과등록화면을 사용하십시오.", vbCritical, Me.Caption
'            CanModifyABOResult = False
'            cmdSave.Enabled = False
'        ElseIf strVfyid1 <> "" And strVfyid1 = ObjMyUser.EmpID Then
'            optType(0).Value = True
'
'            CanModifyABOResult = True
'            cmdSave.Enabled = True
'
'            typing = 0
'        ElseIf strVfyid2 <> "" And strVfyid2 = ObjMyUser.EmpID Then
'            optType(1).Value = True
'
'            CanModifyABOResult = True
'            cmdSave.Enabled = True
'
'            typing = 1
'        Else
'            MsgBox "결과를 입력하지 않았읍니다. 수정을 사용할 수 없습니다.", vbCritical, Me.Caption
'            CanModifyABOResult = False
'            cmdSave.Enabled = False
'        End If
'        '유성은
'        If strVfydt1 <> "" And strVfydt2 <> "" Then lblDoubleCheck.Visible = True
'
'
'        If CanModifyABOResult = True Then
'            '--------저장해놓았던 결과를 Display한다.
'            tblResult.Row = 1
'
'            tblResult.Col = TblColumn.tcABOCD
'            tblResult.Value = IIf(typing = 0, .Fields("abo1").Value & "", .Fields("abo2").Value & "")
'            tblResult.Col = TblColumn.tcABO
'            tblResult.Value = GetABONM(IIf(typing = 0, .Fields("abo1").Value & "", .Fields("abo2").Value & ""))
'
'            tblResult.Col = TblColumn.tcRHCD
'            tblResult.Value = IIf(typing = 0, .Fields("rh1").Value & "", .Fields("rh2").Value & "")
'            tblResult.Col = TblColumn.tcRH
'            tblResult.Value = GetRHNM(IIf(typing = 0, .Fields("rh1").Value & "", .Fields("rh2").Value & ""))
'
'            tblResult.Col = TblColumn.tcABOSUBCD
'            tblResult.Value = .Fields("abosub").Value & ""
'            tblResult.Col = TblColumn.tcABOSUB
'            tblResult.Value = GetABOSUBNM(.Fields("abosub").Value & "")
'
'            tblResult.Col = TblColumn.tcRHSUBCD
'            tblResult.Value = .Fields("rhsub").Value & ""
'            tblResult.Col = TblColumn.tcRHSUB
'            tblResult.Value = GetRHSUBNM(.Fields("rhsub").Value & "")
'            '-----------------------결과 Display완료
'        End If
'    End With

    Set Rs = Nothing
    Set objABOSql = Nothing
End Function

'Private Sub Query(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String)
'    Dim objABOSql As clsABOSql
'    Dim Rs As Recordset
'    Dim i As Long
'
'    Dim deptnm As String
'    Dim wardnm As String
'    Dim doctnm As String
'    Dim ptnm As String
'
'    Dim canModify As Boolean
'
'    '입력된 접수번호 검사----------------------------------------------
'    If vWorkarea = "" Or vAccdt = "" Or vAccseq = "" Then
'        onPgm = True
'        MsgBox "접수번호가 완전하지 않습니다.", vbCritical, Me.Caption
'        onPgm = False
'        Exit Sub
'    End If
'
'    Set objABOSql = New clsABOSql
'
'    Set Rs = objABOSql.GetAccessInfo(vWorkarea, vAccdt, vAccseq)
'
'    If Not (Rs Is Nothing) Then
'        With Rs
'            If .RecordCount < 1 Then
'                MsgBox "존재하지 않는 접수번호입니다.", vbCritical, Me.Caption
'                cmdSave.Enabled = False
'            Else
'                '이 접수번호에 ABO검사항목이 있는지 검사한다.----------------------
'                If objABOSql.IsExistABO(vWorkarea, vAccdt, vAccseq) = False Then
'                    MsgBox "이 접수번호에서 혈액형검사를 찾을 수 없습니다", vbCritical, Me.Caption
'                    Set Rs = Nothing
'                    Set objABOSql = Nothing
'                    cmdSave.Enabled = False
'                    Exit Sub
'                End If
'
'                '이미 검사가 완료되었는지 검사한다.--------------------------------
'                canModify = CanModifyABOResult(vWorkarea, vAccdt, vAccseq)
'                If canModify Then
'
'                    '코드에 대한 명칭을 불러온다.--------------------------------------
'                    ptnm = GetPtNm(.Fields("ptid").Value & "")
'
'                    deptnm = GetDeptNm(.Fields("deptcd").Value & "")
'                    wardnm = GetWardNm(.Fields("wardid").Value & "")
'                    doctnm = GetDoctNm(.Fields("majdoct").Value & "")
'
'                    '값을 Display-------------------------------------------------------
'                    lblPtId.Caption = .Fields("ptid").Value & ""
'                    lblPtNm.Caption = ptnm
'                    lblSexAge.Caption = .Fields("sex").Value & "" & "/" & (Val(.Fields("ageday").Value & "") \ 365)
'                    lblDept.Caption = deptnm
'                    lblWard.Caption = wardnm
'                    lblDoct.Caption = doctnm
'
'                    '리마크-------------------------------------------------------------
'                    cboRemark.ListIndex = -1
'                    If .Fields("rmkcd").Value & "" <> "" Then
'                        For i = 1 To cboRemark.ListCount
'                            If .Fields("rmkcd").Value & "" = medGetP(cboRemark.List(i), 1, vbTab) Then
'                                cboRemark.ListIndex = i
'                                Exit For
'                            End If
'                        Next i
'                    End If
'                End If
'            End If
'        End With
'
'        Set Rs = Nothing
'
'        If canModify Then
'            '코멘트-------------------------------------------------------------
'            Set Rs = objABOSql.GetAccComment(vWorkarea, vAccdt, vAccseq)
'            If Not (Rs Is Nothing) Then
'                With Rs
'                    If .RecordCount > 0 Then
'                        txtComment = .Fields("rsttxt").Value & ""
'                    End If
'                End With
'                Set Rs = Nothing
'            End If
'        End If
'    End If
'
'
'    Set objABOSql = Nothing
'End Sub

'Private Function Save(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Boolean
'    '------------------------------------------------------
'    '1. bbs303에 저장시킨다.
'    '2. 만일, double check완료시점이면 lab302에도 저장한다.
'    '
'    '* lab302에 저장시 결과유형(rsttype)은 F로 변경시킨다.
'    '------------------------------------------------------
'    Dim ABO As String
'    Dim Rh As String
'    Dim ABOSub As String
'    Dim RhSub As String
'    Dim comment As String
'    Dim remark As String
'
'    Dim isVerify As Boolean
'    Dim typing As Long
'
'    Dim vfydt As String
'    Dim vfytm As String
'
'    Dim objABOSql As clsABOSql
'
'    '입력된 접수번호 검사----------------------------------------------
'    If vWorkarea = "" Or vAccdt = "" Or vAccseq = "" Then
'        MsgBox "접수번호가 완전하지 않습니다.", vbCritical, Me.Caption
'        Exit Function
'    End If
'
'    If optType(0).Value = True Then
'        typing = 0
'    ElseIf optType(1).Value = True Then
'        typing = 1
'    Else
'        MsgBox "결과 Type을 선택하십시오.", vbCritical, Me.Caption
'        Save = False
'        Exit Function
'    End If
'
'    With tblResult
'        .Row = 1
'        .Col = TblColumn.tcSEL: isVerify = IIf(.Value = 1, False, True)
'        .Col = TblColumn.tcABOCD: ABO = .Value
'        .Col = TblColumn.tcRHCD:  Rh = .Value
'        .Col = TblColumn.tcABOSUBCD: ABOSub = .Value
'        .Col = TblColumn.tcRHSUBCD: RhSub = .Value
'    End With
'
'
'    '결과값 검사-----------------------------------------------------------
'    If isVerify Then
'        If ABO = "" And ABOSub = "" Then
'            MsgBox "ABO결과가 없습니다.", vbCritical, Me.Caption
'            Save = False
'            Exit Function
'        End If
'
'        If Rh = "" And RhSub = "" Then
'            MsgBox "RH결과가 없습니다.", vbCritical, Me.Caption
'            Save = False
'            Exit Function
'        End If
'
'        If ABO <> "" And ABOSub <> "" Then
'            MsgBox "ABO와 Subtype중 하나만 입력하십시오.", vbCritical, Me.Caption
'            Save = False
'            Exit Function
'        End If
'
'        If Rh <> "" And RhSub <> "" Then
'            MsgBox "RH와 Subgroup중 하나만 입력하십시오.", vbCritical, Me.Caption
'            Save = False
'            Exit Function
'        End If
'    End If
'
'
'    vfydt = Format(GetSystemDate, "YYYYMMDD")
'    vfytm = Format(GetSystemDate, "HHMMSS")
'
'    Set objABOSql = New clsABOSql
'    Save = objABOSql.ModifyABOResult(vWorkarea, vAccdt, vAccseq, typing, ABO, Rh, ABOSub, RhSub, ObjMyUser.EmpID, vfydt, vfytm)
'End Function
