VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBBS312 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액조회"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14520
   Icon            =   "frmBBS312.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14520
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdBMS 
      BackColor       =   &H00F4F0F2&
      Caption         =   "BMS(출고)"
      Height          =   510
      Left            =   60
      Style           =   1  '그래픽
      TabIndex        =   64
      Tag             =   "15101"
      Top             =   8520
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BMS(입고)"
      Height          =   510
      Left            =   1410
      TabIndex        =   63
      Top             =   8520
      Width           =   1320
   End
   Begin VB.CommandButton cmdexcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "To Excel(&E)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   54
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin DRcontrol1.DrFrame fraUpdate 
      Height          =   2040
      Left            =   5295
      TabIndex        =   40
      Top             =   2625
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   3598
      Title           =   "혈액입출고현황"
      TitlePos        =   1
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtVol 
         Height          =   300
         Left            =   210
         TabIndex        =   43
         Top             =   1230
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저장"
         Height          =   315
         Left            =   135
         Style           =   1  '그래픽
         TabIndex        =   42
         Top             =   1620
         Width           =   690
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00DBE6E6&
         Caption         =   "닫기"
         Height          =   315
         Left            =   900
         Style           =   1  '그래픽
         TabIndex        =   41
         Top             =   1605
         Width           =   630
      End
      Begin VB.Label lblCompo 
         Caption         =   "Label13"
         Height          =   240
         Left            =   195
         TabIndex        =   52
         Top             =   3045
         Width           =   855
      End
      Begin VB.Label lblBldno 
         Caption         =   "Label12"
         Height          =   225
         Left            =   195
         TabIndex        =   51
         Top             =   2775
         Width           =   810
      End
      Begin VB.Label lblBldYY 
         Caption         =   "Label11"
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   2445
         Width           =   825
      End
      Begin VB.Label lblBldSrc 
         Caption         =   "Label10"
         Height          =   285
         Left            =   180
         TabIndex        =   49
         Top             =   2085
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "혈액제제 :"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   195
         TabIndex        =   48
         Tag             =   "103"
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "혈액제제 :"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   195
         TabIndex        =   47
         Tag             =   "103"
         Top             =   990
         Width           =   900
      End
      Begin VB.Label lblVol 
         BackColor       =   &H00DBE6E6&
         Height          =   225
         Left            =   285
         TabIndex        =   46
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label8 
         BackColor       =   &H00DBE6E6&
         Caption         =   "cc"
         Height          =   210
         Left            =   1245
         TabIndex        =   45
         Top             =   735
         Width           =   270
      End
      Begin VB.Label Label9 
         BackColor       =   &H00DBE6E6&
         Caption         =   "cc"
         Height          =   210
         Left            =   1230
         TabIndex        =   44
         Top             =   1320
         Width           =   270
      End
   End
   Begin DRcontrol1.DrFrame fraQuery 
      Height          =   5295
      Left            =   8910
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   9340
      Title           =   "혈액입출고현황"
      TitlePos        =   1
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton optBlood 
         BackColor       =   &H00DBE6E6&
         Caption         =   "출고현황"
         Height          =   270
         Index           =   1
         Left            =   4320
         TabIndex        =   39
         Top             =   60
         Width           =   1110
      End
      Begin VB.OptionButton optBlood 
         BackColor       =   &H00DBE6E6&
         Caption         =   "입고현황"
         Height          =   270
         Index           =   0
         Left            =   3105
         TabIndex        =   38
         Top             =   60
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.CommandButton cmdBexit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "닫기"
         Height          =   480
         Left            =   4635
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   465
         Width           =   810
      End
      Begin VB.CommandButton cmdBQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회"
         Height          =   480
         Left            =   3810
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   465
         Width           =   810
      End
      Begin MSComCtl2.DTPicker dtpBEntDt 
         Height          =   315
         Left            =   1020
         TabIndex        =   31
         Top             =   555
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   83820547
         CurrentDate     =   36803
      End
      Begin MSComctlLib.ListView lvwHosB 
         Height          =   4170
         Left            =   105
         TabIndex        =   34
         Top             =   990
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   7355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "일자"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "혈액제제"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "혈액형"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "용량"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "갯수"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpBEntdtTo 
         Height          =   315
         Left            =   2490
         TabIndex        =   36
         Top             =   555
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   83820547
         CurrentDate     =   36803
      End
      Begin MedControls1.LisLabel lblCondition 
         Height          =   315
         Left            =   120
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   555
         Width           =   885
         _ExtentX        =   1561
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
         Caption         =   "입고일자"
         Appearance      =   0
      End
      Begin VB.Label lblTot 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2820
         TabIndex        =   60
         Top             =   45
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2355
         TabIndex        =   37
         Tag             =   "103"
         Top             =   615
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1755
      Left            =   75
      TabIndex        =   1
      Top             =   45
      Width           =   14400
      Begin VB.CheckBox chkDE 
         BackColor       =   &H00DBE6E6&
         Caption         =   "BMS파일(Local출고/폐기혈액포함)"
         Height          =   315
         Left            =   8520
         TabIndex        =   65
         Top             =   960
         Visible         =   0   'False
         Width           =   3390
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00F4F0F2&
         Caption         =   "일별현황"
         Height          =   510
         Left            =   12855
         Style           =   1  '그래픽
         TabIndex        =   35
         Tag             =   "15101"
         Top             =   375
         Width           =   1320
      End
      Begin VB.TextBox txtptid 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   10440
         MaxLength       =   14
         TabIndex        =   28
         Text            =   "123456-1234567"
         Top             =   615
         Width           =   945
      End
      Begin VB.CheckBox chkDelivery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "출고된 혈액만 조회"
         Height          =   315
         Left            =   8535
         TabIndex        =   27
         Top             =   630
         Width           =   1995
      End
      Begin VB.CheckBox optall 
         BackColor       =   &H00FDF7F8&
         Caption         =   "전체"
         Height          =   1455
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkUse 
         BackColor       =   &H00DBE6E6&
         Caption         =   "사용가능한 혈액만"
         Height          =   315
         Left            =   8535
         TabIndex        =   15
         Top             =   285
         Width           =   1995
      End
      Begin VB.Frame fraABO 
         BackColor       =   &H00DBE6E6&
         Height          =   1575
         Left            =   1500
         TabIndex        =   3
         Top             =   120
         Width           =   2880
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00C0FFFF&
            Height          =   900
            Left            =   1035
            ScaleHeight     =   840
            ScaleWidth      =   1650
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   615
            Width           =   1710
            Begin VB.Label lblABO 
               Alignment       =   2  '가운데 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "AB+"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   27.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   555
               Left            =   135
               TabIndex        =   5
               Top             =   105
               Width           =   1425
            End
         End
         Begin VB.PictureBox Picture3 
            Height          =   420
            Left            =   1020
            ScaleHeight     =   360
            ScaleWidth      =   1695
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   195
            Width           =   1755
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "A"
               Height          =   375
               Index           =   0
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   13
               Top             =   0
               Width           =   435
            End
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "B"
               Height          =   375
               Index           =   1
               Left            =   420
               Style           =   1  '그래픽
               TabIndex        =   12
               Top             =   0
               Width           =   435
            End
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "O"
               Height          =   375
               Index           =   2
               Left            =   840
               Style           =   1  '그래픽
               TabIndex        =   11
               Top             =   0
               Width           =   435
            End
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "AB"
               Height          =   375
               Index           =   3
               Left            =   1260
               Style           =   1  '그래픽
               TabIndex        =   10
               Top             =   0
               Width           =   435
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   915
            Left            =   135
            ScaleHeight     =   855
            ScaleWidth      =   870
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   600
            Width           =   930
            Begin VB.OptionButton optRh 
               BackColor       =   &H00DBE6E6&
               Caption         =   "-"
               Height          =   435
               Index           =   1
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   8
               Top             =   420
               Width           =   855
            End
            Begin VB.OptionButton optRh 
               BackColor       =   &H00DBE6E6&
               Caption         =   "+"
               Height          =   435
               Index           =   0
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   7
               Top             =   0
               Width           =   855
            End
         End
         Begin MedControls1.LisLabel LisLabel1 
            Height          =   405
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   714
            BackColor       =   8421504
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
            Caption         =   "혈액형"
            Appearance      =   0
         End
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   12855
         Style           =   1  '그래픽
         TabIndex        =   2
         Tag             =   "15101"
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Frame fraCondition 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   1515
         Left            =   4500
         TabIndex        =   19
         Top             =   120
         Width           =   4005
         Begin MedControls1.LisLabel lbldt 
            Height          =   315
            Left            =   15
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   180
            Width           =   1140
            _ExtentX        =   2011
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
            Caption         =   "입고일자"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   4
            Left            =   15
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   645
            Width           =   1140
            _ExtentX        =   2011
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
            Caption         =   "혈액제제"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   6
            Left            =   15
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1140
            _ExtentX        =   2011
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
            Caption         =   "Center"
            Appearance      =   0
         End
         Begin VB.ComboBox cboCompo 
            Height          =   300
            ItemData        =   "frmBBS312.frx":076A
            Left            =   1200
            List            =   "frmBBS312.frx":076C
            Style           =   2  '드롭다운 목록
            TabIndex        =   21
            Top             =   660
            Width           =   2760
         End
         Begin VB.ComboBox cboCenter 
            Height          =   300
            Left            =   1200
            Style           =   2  '드롭다운 목록
            TabIndex        =   20
            Top             =   1140
            Width           =   2760
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   1200
            TabIndex        =   22
            Top             =   180
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   83820547
            CurrentDate     =   36803
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   2715
            TabIndex        =   23
            Top             =   180
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   83820547
            CurrentDate     =   36803
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2550
            TabIndex        =   24
            Tag             =   "103"
            Top             =   240
            Width           =   90
         End
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   11415
         TabIndex        =   29
         Top             =   615
         Width           =   1020
         _ExtentX        =   1799
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   8535
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "조 회 건 수"
         Appearance      =   0
      End
      Begin VB.Label lblCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9945
         TabIndex        =   26
         Top             =   1260
         Width           =   495
      End
   End
   Begin FPSpread.vaSpread tblBldList 
      Height          =   6645
      Left            =   75
      TabIndex        =   0
      Top             =   1800
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   11721
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      ColsFrozen      =   3
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
      MaxCols         =   36
      MaxRows         =   27
      OperationMode   =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS312.frx":076E
      TextTip         =   4
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   675
      _Version        =   196608
      _ExtentX        =   1191
      _ExtentY        =   1191
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmBBS312.frx":13D3
   End
   Begin FPSpread.vaSpread tblBMSList 
      Height          =   3615
      Left            =   3690
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   5235
      _Version        =   196608
      _ExtentX        =   9234
      _ExtentY        =   6376
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      SpreadDesigner  =   "frmBBS312.frx":157C
   End
   Begin FPSpread.vaSpread tblBMSList1 
      Height          =   2130
      Left            =   0
      TabIndex        =   61
      Top             =   3915
      Visible         =   0   'False
      Width           =   9915
      _Version        =   196608
      _ExtentX        =   17489
      _ExtentY        =   3757
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      SpreadDesigner  =   "frmBBS312.frx":1733
   End
End
Attribute VB_Name = "frmBBS312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TblColumn
    tcCOMPONM = 1
    tcABO
    tcBldNo
    tcColDt
    tcAVAIL
    
    tcENTDT
    tcVol
    tcENTNM
    tcIRR
    tcSTATUS
    tcLARC
    tcSMLC
    tcLARE
    tcSMLE
    
    tcDELDT
    tcPTID
    tcPTNM
    tcSEX
    tcAGE
    
    tcDEPT
    tcABO_P
    tcDELNM
    tcDELRCVNM
    tcVFYDT
    
    tcVFYNM
    tcSTAT
    tcRSTV
    tcRETDT
    tcRETNMC
    
    tcRETNM
    tcExpDt
    tcEXPNMC
    tcEXPNM
    tcEXPRSN
    
End Enum
Private objSql                  As New clsBloodQuery
'Private WithEvents mnuPopup     As Menu
'Private WithEvents mnuDelete    As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1

Private Sub chkDelivery_Click()
    txtPtId.Visible = False: lblPtNm.Visible = False
    txtPtId.Text = "": lblPtNm.Caption = ""
    If chkDelivery.value = 1 Then
'        If chkUse.value = 1 Then
'            MsgBox "사용가능한 혈액과 동시에 조회하실수 없습니다.", vbInformation + vbOKOnly, "출고혈액조회"
'            chkDelivery.value = 0
'            lbldt.Caption = "입고일자"
'            Exit Sub
'        End If
    
        lbldt.Caption = "출고일자"
        txtPtId.Visible = True: lblPtNm.Visible = True
    Else
        lbldt.Caption = "입고일자"
    End If
    
    chkUse.value = IIf(chkDelivery.value = 1, 0, 1)
    If tblBldList.MaxRows <> 0 Then tblBldList.MaxRows = 0
    lblCnt.Caption = ""
End Sub


Private Sub chkUse_Click()
    If chkUse.value = 1 Then
'        If chkDelivery.value = 1 Then
'            MsgBox "출고된혈액과 동시에 조회하실수 없습니다.", vbInformation + vbOKOnly, "사용가능한 혈액조회"
'            chkUse.value = 0
'            lbldt.Caption = "출고일자"
'        End If
        lbldt.Caption = "입고일자"
    Else
        lbldt.Caption = "출고일자"
        txtPtId.Visible = True: lblPtNm.Visible = True
    End If
    
    chkDelivery.value = IIf(chkUse.value = 1, 0, 1)
    If tblBldList.MaxRows <> 0 Then tblBldList.MaxRows = 0
    lblCnt.Caption = ""
End Sub

Private Sub cmdBexit_Click()
    fraQuery.Visible = False
End Sub

Private Sub cmdBMS_Click()
    Dim iCnt, iCol, iRow As Integer
    Dim varTmp
    Dim strBMS As String
    Dim Resp
    
    Resp = MsgBox(lblCnt.Caption & " 건 BMS 전송하시겠습니까 ? ", vbQuestion + vbOKCancel, "확인")
    
    If Resp = vbCancel Then Exit Sub
    
    tblBMSList.MaxRows = 0
    strBMS = ""
    
    With tblBMSList
        .MaxRows = .MaxRows + 1
        .SetText 1, 1, "혈액번호"
        .SetText 2, 1, "혈액제제코드"
        .SetText 3, 1, "혈액제제명"
        .SetText 4, 1, "출고구분"
        .SetText 5, 1, "출고일자"
        .SetText 6, 1, "출고시간"
        .SetText 7, 1, "혈액형코드"
        .SetText 8, 1, "혈액형명"
        .SetText 9, 1, "출고자"
        '추가 수혈자성별 수혈자출생연도 수혈자혈액형   수혈자거주지역 수혈자진료과목
        .SetText 10, 1, "성별"
        .SetText 11, 1, "출생연도"
        .SetText 12, 1, "혈액형"
        .SetText 13, 1, "거주지역"
        .SetText 14, 1, "진료과목"
    End With
    
    With tblBldList
        For iCnt = 1 To .MaxRows
            tblBMSList.MaxRows = tblBMSList.MaxRows + 1
            .GetText 3, iCnt, varTmp: tblBMSList.SetText 1, iCnt + 1, varTmp
            .GetText 1, iCnt, varTmp
            
            Select Case varTmp
                Case "FFP400": tblBMSList.SetText 2, iCnt + 1, "60": tblBMSList.SetText 3, iCnt + 1, "신선동결혈장"
                Case "FFP320": tblBMSList.SetText 2, iCnt + 1, "10": tblBMSList.SetText 3, iCnt + 1, "신선동결혈장"
                
                Case "PC400": tblBMSList.SetText 2, iCnt + 1, "56": tblBMSList.SetText 3, iCnt + 1, "농축혈소판"
                Case "PC320": tblBMSList.SetText 2, iCnt + 1, "06": tblBMSList.SetText 3, iCnt + 1, "농축혈소판"
                
                Case "PRC400": tblBMSList.SetText 2, iCnt + 1, "51": tblBMSList.SetText 3, iCnt + 1, "농축적혈구"
                Case "PRC320": tblBMSList.SetText 2, iCnt + 1, "01": tblBMSList.SetText 3, iCnt + 1, "농축적혈구"
                
                
                Case "FRB400": tblBMSList.SetText 2, iCnt + 1, "54": tblBMSList.SetText 3, iCnt + 1, "백혈구여과제거적혈구"
                Case "FRB320": tblBMSList.SetText 2, iCnt + 1, "04": tblBMSList.SetText 3, iCnt + 1, "백혈구여과제거적혈구"
                
                Case "FRBC400": tblBMSList.SetText 2, iCnt + 1, "54": tblBMSList.SetText 3, iCnt + 1, "백혈구여과제거적혈구"
                Case "FRBC320": tblBMSList.SetText 2, iCnt + 1, "04": tblBMSList.SetText 3, iCnt + 1, "백혈구여과제거적혈구"
                
                Case "CRY400": tblBMSList.SetText 2, iCnt + 1, "62": tblBMSList.SetText 3, iCnt + 1, "동결침전제제"
                Case "CRY320": tblBMSList.SetText 2, iCnt + 1, "12": tblBMSList.SetText 3, iCnt + 1, "동결침전제제"
                
                Case "PLT-Phere(혈)": tblBMSList.SetText 2, iCnt + 1, "83": tblBMSList.SetText 3, iCnt + 1, "성분채혈혈소판[M]"
                Case "LFPLT-P": tblBMSList.SetText 2, iCnt + 1, "72": tblBMSList.SetText 3, iCnt + 1, "백혈구여과제거성분채혈혈소판"
            End Select
            
'            If varTmp = "320cc" Then
'                tblBMSList.SetText 2, iCnt + 1, "01"
'            Else
'                tblBMSList.SetText 2, iCnt + 1, "51"
'            End If
'            tblBMSList.SetText 3, iCnt + 1, "농축적혈구"
            .GetText 10, iCnt, varTmp
            If varTmp = "출고" Then
                tblBMSList.SetText 4, iCnt + 1, 1
            Else
                tblBMSList.SetText 4, iCnt + 1, 2
            End If
            .GetText 15, iCnt, varTmp
            tblBMSList.SetText 5, iCnt + 1, Mid(varTmp, 1, 10)
            tblBMSList.SetText 6, iCnt + 1, Trim(Mid(varTmp, 11))
            .GetText 2, iCnt, varTmp
            Select Case varTmp
                Case "O+"
                    tblBMSList.SetText 7, iCnt + 1, "1"
                    tblBMSList.SetText 8, iCnt + 1, "O (+)"
                Case "A+"
                    tblBMSList.SetText 7, iCnt + 1, "2"
                    tblBMSList.SetText 8, iCnt + 1, "A (+)"
                Case "B+"
                    tblBMSList.SetText 7, iCnt + 1, "3"
                    tblBMSList.SetText 8, iCnt + 1, "B (+)"
'                Case "O+"
'                    tblBMSList.SetText 7, iCnt + 1, 1
'                    tblBMSList.SetText 8, iCnt + 1, "O (+)"
                Case "AB+"
                    tblBMSList.SetText 7, iCnt + 1, "4"
                    tblBMSList.SetText 8, iCnt + 1, "AB (+)"
                Case "O-"
                    tblBMSList.SetText 7, iCnt + 1, "5"
                    tblBMSList.SetText 8, iCnt + 1, "O (-)"
                Case "A-"
                    tblBMSList.SetText 7, iCnt + 1, "6"
                    tblBMSList.SetText 8, iCnt + 1, "A (-)"
                Case "B-+"
                    tblBMSList.SetText 7, iCnt + 1, "7"
                    tblBMSList.SetText 8, iCnt + 1, "B (-)"
                Case "AB-"
                    tblBMSList.SetText 7, iCnt + 1, "8"
                    tblBMSList.SetText 8, iCnt + 1, "AB (-)"
            End Select
            .GetText 22, iCnt, varTmp: tblBMSList.SetText 9, iCnt + 1, varTmp
            .GetText 18, iCnt, varTmp:
            If varTmp = "남" Then
                tblBMSList.SetText 10, iCnt + 1, "M"
            Else
                tblBMSList.SetText 10, iCnt + 1, "W"
            End If
            .GetText 19, iCnt, varTmp: tblBMSList.SetText 11, iCnt + 1, varTmp
            .GetText 21, iCnt, varTmp
            Select Case varTmp
                Case "O(O)+"
                    tblBMSList.SetText 12, iCnt + 1, "1"
                Case "A(A)+"
                    tblBMSList.SetText 12, iCnt + 1, "2"
                Case "B(B)+"
                    tblBMSList.SetText 12, iCnt + 1, "3"
'                Case "O+"
'                    tblBMSList.SetText 12, iCnt + 1, 1
                Case "AB(AB)+"
                    tblBMSList.SetText 12, iCnt + 1, "4"
                Case "O(O)-"
                    tblBMSList.SetText 12, iCnt + 1, "5"
                Case "A(A)-"
                    tblBMSList.SetText 12, iCnt + 1, "6"
                Case "B(B)-+"
                    tblBMSList.SetText 12, iCnt + 1, "7"
                Case "AB(AB)-"
                    tblBMSList.SetText 12, iCnt + 1, "8"
            End Select
            tblBMSList.SetText 13, iCnt + 1, "112"
            
            .GetText 20, iCnt, varTmp
            Dim strDeptNm As String
            strDeptNm = GetDeptNm(varTmp)
            Select Case strDeptNm
                Case "소화기내과":        tblBMSList.SetText 14, iCnt + 1, "011"
                Case "신경외과":          tblBMSList.SetText 14, iCnt + 1, "027"
                Case "응급의학과":        tblBMSList.SetText 14, iCnt + 1, "048"
                Case "정형외과":          tblBMSList.SetText 14, iCnt + 1, "026"
                Case "호흡기.알레르기내과": tblBMSList.SetText 14, iCnt + 1, "013"
                Case "산부인과":          tblBMSList.SetText 14, iCnt + 1, "033"
                Case "재활의학과":        tblBMSList.SetText 14, iCnt + 1, "021"
                Case "간담췌외과":        tblBMSList.SetText 14, iCnt + 1, "048"
                Case "일반내과":          tblBMSList.SetText 14, iCnt + 1, "010"
                Case "순환기내과":        tblBMSList.SetText 14, iCnt + 1, "012"
                Case "호흡기내과":        tblBMSList.SetText 14, iCnt + 1, "013"
                Case "내분비대사내과":    tblBMSList.SetText 14, iCnt + 1, "014"
                Case "신장내과":          tblBMSList.SetText 14, iCnt + 1, "015"
                Case "혈액종양내과":      tblBMSList.SetText 14, iCnt + 1, "016"
                Case "감염내과":          tblBMSList.SetText 14, iCnt + 1, "017"
                Case "알레르기내과":      tblBMSList.SetText 14, iCnt + 1, "018"
                Case "류마티스내과":      tblBMSList.SetText 14, iCnt + 1, "019"
                Case "일반외과":          tblBMSList.SetText 14, iCnt + 1, "020"
                Case "대장항문외과":      tblBMSList.SetText 14, iCnt + 1, "022"
                Case "소아외과":          tblBMSList.SetText 14, iCnt + 1, "023"
                Case "위장관외과":        tblBMSList.SetText 14, iCnt + 1, "024"
                Case "유방질환외과":      tblBMSList.SetText 14, iCnt + 1, "025"
                Case "흉부외과":          tblBMSList.SetText 14, iCnt + 1, "028"
                Case "성형외과":          tblBMSList.SetText 14, iCnt + 1, "029"
                Case "신경과":            tblBMSList.SetText 14, iCnt + 1, "030"
                Case "정신건강의학과":    tblBMSList.SetText 14, iCnt + 1, "031"
                Case "마취통증의학과":    tblBMSList.SetText 14, iCnt + 1, "032"
                Case "소아청소년과":      tblBMSList.SetText 14, iCnt + 1, "034"
                Case "안과":              tblBMSList.SetText 14, iCnt + 1, "035"
                Case "이비인후과":        tblBMSList.SetText 14, iCnt + 1, "036"
                Case "피부과":            tblBMSList.SetText 14, iCnt + 1, "037"
                Case "비뇨기과":          tblBMSList.SetText 14, iCnt + 1, "038"
                Case "영상의학과":        tblBMSList.SetText 14, iCnt + 1, "039"
                Case "방사선종양학과":    tblBMSList.SetText 14, iCnt + 1, "040"
                Case "병리과":            tblBMSList.SetText 14, iCnt + 1, "041"
                Case "진단검사의학과":    tblBMSList.SetText 14, iCnt + 1, "042"
                Case "결핵과":            tblBMSList.SetText 14, iCnt + 1, "044"
                Case "가정의학과":        tblBMSList.SetText 14, iCnt + 1, "045"
                Case "핵의학과":          tblBMSList.SetText 14, iCnt + 1, "046"
                Case "직업환경의학과":    tblBMSList.SetText 14, iCnt + 1, "047"
                Case "구강악안면외과":    tblBMSList.SetText 14, iCnt + 1, "050"
                Case "치과보철과":        tblBMSList.SetText 14, iCnt + 1, "051"
                Case "치과교정과":        tblBMSList.SetText 14, iCnt + 1, "052"
                Case "소아치과":          tblBMSList.SetText 14, iCnt + 1, "053"
                Case "치주과":            tblBMSList.SetText 14, iCnt + 1, "054"
                Case "치과보존과":        tblBMSList.SetText 14, iCnt + 1, "055"
                Case "구강내과":          tblBMSList.SetText 14, iCnt + 1, "056"
                Case "영상치의학과":      tblBMSList.SetText 14, iCnt + 1, "057"
                Case "구강병리과":        tblBMSList.SetText 14, iCnt + 1, "058"
                Case "예방치과":          tblBMSList.SetText 14, iCnt + 1, "059"
                Case "일반치과":          tblBMSList.SetText 14, iCnt + 1, "060"
                Case Else: tblBMSList.SetText 14, iCnt + 1, "099"
            End Select
            
        Next
    End With
    
    With tblBMSList
        For iCnt = 1 To .MaxRows
            .GetText 1, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 2, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 3, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 4, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 5, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 6, iCnt, varTmp:
            
            Debug.Print varTmp
            If InStr(varTmp, ":") = 2 Then
                strBMS = strBMS & "0" & varTmp & ","
            Else
                If Left((Trim(varTmp & "")), 1) = ":" And Len(Trim(varTmp & "")) = 3 Then
                   strBMS = strBMS & "00" & varTmp & ","
                Else
                strBMS = strBMS & varTmp & ","
                End If
            End If
            .GetText 7, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 8, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 9, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 10, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 11, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 12, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 13, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 14, iCnt, varTmp: strBMS = strBMS & varTmp & vbCrLf
        Next
    End With
    
    Dim strPath As String
    Dim FreeFileNum As String

    strPath = "C:\BMS출고\BMS출고.csv"

    If Dir("C:\BMS출고\BMS출고.csv") <> "" Then
        Kill "C:\BMS출고\BMS출고.csv"
    End If

    FreeFileNum = FreeFile
    Open strPath For Append As #FreeFileNum
    Print #FreeFileNum, strBMS
    Close #FreeFileNum
    
    Call MsgBox("BMS전송이 완료되었습니다.", vbExclamation, "BMS전송완료")
End Sub

Private Sub cmdBQuery_Click()
    Call BloodQueryPre
End Sub

Private Sub BloodQueryPre()
    Dim objBSQL As New clsHospital05
    Dim RS      As Recordset
    Dim strEntdt As String
    Dim strEntTo As String
    Dim itmX     As ListItem
    Dim i        As Long
    
    Dim objPro As clsProgress
    
    Screen.MousePointer = vbHourglass
    
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Left = fraQuery.Left + lvwHosB.Left
        .Top = fraQuery.Top + lvwHosB.Top
        .Width = lvwHosB.Width
        .DisplayPercent = False
        .Message = "자료를 읽기 위해 준비중입니다..."
    End With
    
    strEntdt = Format(dtpBEntDt.value, "yyyymmdd")
    strEntTo = Format(dtpBEntdtTo.value, "yyyymmdd")
    
    strEntdt = strEntdt & COL_DIV & strEntTo
    
    lvwHosB.ListItems.Clear
    lblTot.Caption = ""
    If optBlood(0).value = True Then
        Set RS = New Recordset
        RS.Open objBSQL.GetBloodDetailQuery(strEntdt), DBConn
    Else
        Set RS = New Recordset
        RS.Open objBSQL.GetBloodDeliveryQuery(strEntdt), DBConn
    End If
    
    If Not RS.EOF Then
        objPro.Max = RS.RecordCount
        objPro.DisplayPercent = True
        objPro.Message = "자료를 읽고 있습니다..."
        
        With RS
            .MoveFirst
            For i = 1 To .RecordCount
                
                objPro.value = objPro.value + 1
                
                Set itmX = lvwHosB.ListItems.Add()
                itmX.Text = Format(.Fields("dt").value & "", "####-##-##")
                itmX.SubItems(1) = .Fields("abbrnm").value & ""
                itmX.SubItems(2) = .Fields("abo").value & "" & .Fields("rh").value & ""
                itmX.SubItems(3) = .Fields("volumn").value & "" & "cc"
                itmX.SubItems(4) = .Fields("cnt").value & ""
                
                lblTot.Caption = Val(lblTot.Caption) + Val(.Fields("cnt").value & "")
                .MoveNext
            Next i
        End With
        
        Set itmX = lvwHosB.ListItems.Add()
            itmX.Text = "총 건수"
            itmX.SubItems(4) = lblTot.Caption
            itmX.ForeColor = vbBlue
            itmX.ListSubItems(4).ForeColor = vbBlue
            itmX.Bold = True
            itmX.ListSubItems(4).Bold = True
            itmX.EnsureVisible
            
    End If
    
    If Val(lblTot.Caption) > 0 Then lvwHosB.ToolTipText = "총 건수 : " & lblTot.Caption
    
    Screen.MousePointer = vbDefault
    
    Set RS = Nothing
    Set objBSQL = Nothing
    Set objPro = Nothing
End Sub
Private Sub cmdClear_Click()
    Clear
End Sub
Private Sub Clear()
    tblBldList.MaxRows = 0
    cboCompo.ListIndex = 0
    lblCnt.Caption = ""
    optall.value = 0
    lblABO.Caption = ""
    chkDelivery.value = 0
End Sub

Private Sub cmdClose_Click()
    fraUpdate.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPre_Click()
    fraQuery.Visible = True
    optBlood(0).value = True
    dtpBEntDt.value = Format(GetSystemDate, "yyyy-mm-dd")
    dtpBEntdtTo.value = Format(GetSystemDate, "yyyy-mm-dd")
    Call BloodQueryPre
End Sub

Private Sub cmdQuery_Click()
    '조회조건----------------------------------------
    Dim CenterCd    As String
    Dim lngRowcnt   As Integer
    Dim ADt         As String
    
'    Dim objProBar   As New clsProgressBar
    Dim RS          As Recordset
    Dim RsS         As Recordset
    Dim SexTmp      As String
    
    If MsgBox("혈액정보를 조회하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    If cboCenter.ListIndex = 0 Then
        CenterCd = ""
    Else
        CenterCd = medGetP(cboCenter.Text, 1, " ")
    End If

    objSql.CenterCd = CenterCd
    objSql.EntdtF = Format(dtpFrom.value, PRESENTDATE_FORMAT)
    objSql.EntdtL = Format(dtpTo.value, PRESENTDATE_FORMAT)
    
    If optall.value = 1 Then
        objSql.ABO = "": objSql.Rh = ""
    Else
        If optABO(0).value Then objSql.ABO = "A"
        If optABO(1).value Then objSql.ABO = "B"
        If optABO(2).value Then objSql.ABO = "O"
        If optABO(3).value Then objSql.ABO = "AB"
        If optRh(0).value Then objSql.Rh = "+"
        If optRh(1).value Then objSql.Rh = "-"
    End If
    If chkDelivery.value = 1 Then
        objSql.Ptid = txtPtId.Text
    Else
        objSql.Ptid = ""
    End If
    
    If cboCompo.ListIndex = 0 Then
        objSql.CompoCd = ""
    Else
        objSql.CompoCd = medGetP(cboCompo.Text, 1, " ")
    End If
    
    Screen.MousePointer = vbHourglass
    If chkUse.value = 1 Then
    '사용가능한 혈액만 조회한다.
    '입고일자/혈액제제에 상관없이 조회한다.
        Call QUERY_USED
    Else
        Call ALL_QUERY
    End If
    
    '* 2014-09-16 BMS파일선택시 LOCAL혈액 출고데이타/ 혈액폐기데이타 출력한다... PSK
    '/=> START --------------------------------------------------------------------------
    If chkDE.value = 1 Then
       Set RS = objSql.GetBlood_BMS_LocalOutPut
       
       If RS.RecordCount < 1 Then
          Set RS = Nothing
          Me.MousePointer = vbDefault
          Exit Sub
       End If
       
       lngRowcnt = tblBldList.MaxRows
       lblCnt.Caption = lblCnt.Caption + RS.RecordCount
       
'       Set objProBar.StatusBar = medMain.stsBar
'       objProBar.Min = lngRowcnt
'       objProBar.Max = lngRowcnt + RS.RecordCount
    
       With tblBldList
        .ReDraw = False
        .MaxRows = .MaxRows + RS.RecordCount
        Do Until RS.EOF
            lngRowcnt = lngRowcnt + 1
            .Row = lngRowcnt
            .Col = TblColumn.tcCOMPONM:  .value = "PRC"
'            .Col = TblColumn.tcABO:      .value = GetABO(Trim(RS.Fields("bldsrc").value & ""), Trim(RS.Fields("bldyy").value & ""), Format(Trim(RS.Fields("bldno").value & ""), "00000#"), Trim(RS.Fields("compocd").value & ""))
            .Col = TblColumn.tcBldNo::   .value = RS.Fields("bldsrc").value & "" & "-" & RS.Fields("bldyy").value & "" & "-" & Format(RS.Fields("bldno").value & "", "00000#")
            .Col = TblColumn.tcColDt:    .value = Format(RS.Fields("coldt").value & "", "####-##-##")        '채혈일자
             ADt = Format(DateAdd("d", Val(RS.Fields("available").value) - 1, .value), "YYYY-MM-dd")
            .Col = TblColumn.tcAVAIL:   .value = ADt: .ForeColor = DCM_LightBlue
            
            .Col = TblColumn.tcENTDT:    .value = Format(RS.Fields("entdt").value & "", "####-##-##")        '유효일자
            .Col = TblColumn.tcENTNM:    .value = GetEmpNm(RS.Fields("entid").value & "")                   '입고자
            .Col = TblColumn.tcVol:     .value = RS.Fields("volumn").value & "" & "cc"
            
            Select Case RS.Fields("STSCD").value & ""
             Case "3"   '출고
                .Col = TblColumn.tcSTATUS:   .value = "출고"
                .ForeColor = BBSOrderStatusColor.clrINPROCESS
             Case "4"   '폐기
                .Col = TblColumn.tcSTATUS: .value = "폐기": .ForeColor = BBSOrderStatusColor.cIrEND
            End Select
            .Col = TblColumn.tcDELDT:    .value = Format(RS.Fields("deliverydt").value & "", "####-##-##") & " " & Format$(Mid$(RS.Fields("deliverytm"), 1, 4), "00:00") '출고일자/시간
            
            If Len(Trim(RS.Fields("localcd").value & "")) > 0 Then
                .Col = TblColumn.tcPTID:     .value = RS.Fields("localcd").value & ""                               '환자번호
                .Col = TblColumn.tcPTNM:     .value = RS.Fields("localnm").value & ""
                .Col = TblColumn.tcSEX:      .value = ""
                .Col = TblColumn.tcAGE:      .value = ""
                .Col = TblColumn.tcDEPT:     .value = ""
            Else
                Set RsS = objSql.AssignBloodRs(RS.Fields("bldsrc").value & "", RS.Fields("bldyy").value & "", Format(Trim(RS.Fields("bldno").value & ""), "00000#"), Trim(RS.Fields("compocd").value & ""))
                If Not RsS.EOF Then
                    SexTmp = SDA_String(RsS.Fields("ssn").value & "")
                    .Col = TblColumn.tcPTID: .value = RsS.Fields("ptid").value & ""
                    .Col = TblColumn.tcPTNM: .value = RsS.Fields("ptnm").value & ""
                    .Col = TblColumn.tcSEX:  .value = medGetP(SexTmp, 1, COL_DIV)
                    .Col = TblColumn.tcAGE:  .value = Mid(medGetP(SexTmp, 2, COL_DIV), 1, 4) '.value = medGetP(SexTmp, 3, COL_DIV)
                    .Col = TblColumn.tcDEPT: .value = RsS.Fields("deptcd").value & ""
                        If RsS.Fields("wardid") & "" <> "" Then
                            .value = .value & "/" & RsS.Fields("wardid").value & ""
                        End If
                    
                    .Col = TblColumn.tcVFYDT:     .value = Format(RsS.Fields("vfydt").value & "", "####-##-##")
                    .Col = TblColumn.tcVFYNM:     .value = GetEmpNm(RsS.Fields("vfyid").value & "")
                    .Col = TblColumn.tcSTAT:      .value = IIf(RsS.Fields("stat").value & "" = "1", "Y", "")
                    .Col = TblColumn.tcRSTV:
                    
                    Select Case RsS.Fields("rstv").value & ""
                        Case "1": .value = "OK"
                        Case "2": .value = "NOT"
                        Case Else: .value = ""
                    End Select
                End If
                Set RsS = Nothing
            End If
'            .Col = TblColumn.tcABO_P::   .value = GetABO(Trim(RS.Fields("bldsrc").value & ""), Trim(RS.Fields("bldyy").value & ""), Format(Trim(RS.Fields("bldno").value & ""), "00000#"), Trim(RS.Fields("compocd").value & ""))
            .Col = TblColumn.tcDELNM:    .value = GetEmpNm(RS.Fields("deliveryid").value & "")
            .Col = TblColumn.tcDELRCVNM: .value = GetEmpNm(RS.Fields("rcvid").value & "")
'            objProBar.value = lngRowcnt
            RS.MoveNext
        Loop
         RS.Close
        .ReDraw = True
       End With
    End If
    '/<= END --------------------------------------------------------------------------
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub cmdSave_Click()
    If chkUse.value = 1 Then
        Call print_use
    Else
        Call Print_all
    End If
End Sub

Private Sub Print_all()
'출력하자.....크리스탈
    Dim strTmp As String
    Dim strRfile As String
    Dim strRptPath As String
    Dim strDisease As String
    Dim intFNum As Integer
    Dim strEntdt As String
    
    Dim ii As Integer
    Dim jj As Integer

    If tblBldList.DataRowCnt = 0 Then Exit Sub
    Me.MousePointer = 11
    With tblBldList
        For ii = 1 To .DataRowCnt
            .Row = ii
            For jj = TblColumn.tcCOMPONM To TblColumn.tcDELRCVNM
                .Col = jj
                strTmp = strTmp & Trim(.value) & vbTab
            Next
            strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
            strTmp = strTmp & vbCr
        Next
    End With
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)

    strRfile = InstallDir & "BBS\Rpt" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\Rpt" & "\frmBBS303N_2.rpt"

    If chkDelivery.value = 1 Then
        strEntdt = "출고일자 : " & Format(dtpFrom, "yyyy-mm-dd") & " ~ " & Format(dtpTo.value, "YYYY_MM_DD")
    Else
        strEntdt = "조회일자 : " & Format(dtpFrom, "yyyy-mm-dd") & " ~ " & Format(dtpTo.value, "YYYY_MM_DD")
    End If
    
    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    '
    With CReport
        .ParameterFields(0) = "entdt;" & strEntdt & ";TRUE"
        .ParameterFields(1) = "hosnm;" & HOSPITAL_NAME & ";TRUE"
        If chkDelivery.value = 1 Then
            .ParameterFields(2) = "Title;" & " 출고혈액 리스트" & ";TRUE"
        Else
            .ParameterFields(2) = "Title;" & " 혈액조회 리스트" & ";TRUE"
        End If
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        
        .WindowState = 0
        .WindowTitle = "혈액 List"
        
        .Action = 1
        .Reset
    End With
    Me.MousePointer = 0
    
End Sub
Private Sub print_use()
'출력하자.....크리스탈
    Dim strTmp As String
    Dim strRfile As String
    Dim strRptPath As String
    Dim strDisease As String
    Dim intFNum As Integer
    Dim strEntdt As String
    
    Dim ii As Integer
    Dim jj As Integer

    If tblBldList.DataRowCnt = 0 Then Exit Sub
    Me.MousePointer = 11
    With tblBldList
        For ii = 1 To .DataRowCnt
            .Row = ii
            For jj = TblColumn.tcCOMPONM To TblColumn.tcIRR
                .Col = jj
                strTmp = strTmp & Trim(.value) & vbTab
            Next
            strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
            strTmp = strTmp & vbCr
        Next
    End With
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)

    strRfile = InstallDir & "BBS\Rpt" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\Rpt" & "\frmBBS303N_1.rpt"

    
    strEntdt = Format(dtpFrom, "yyyy-mm-dd") & " ~ " & Format(dtpTo.value, "YYYY_MM_DD")
    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    '
    With CReport
        .ParameterFields(0) = "entdt;" & strEntdt & ";TRUE"
        .ParameterFields(1) = "hosnm;" & HOSPITAL_NAME & ";TRUE"
        
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        
        .WindowState = 0
        .WindowTitle = "입고혈액List"
        
        .Action = 1
        .Reset
    End With
    Me.MousePointer = 0
End Sub

Private Sub cmdUpdate_Click()
    Dim strSQL As String
    
    If Trim(txtVol.Text) = "" Then Exit Sub
    
    If MsgBox("혈액용량을 변경하시겠습니까?", vbExclamation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
    
    '입고자와 현재 로그인 사용자가 같은 경우에만 변경할 수 있도록 한다.
    tblBldList.Row = tblBldList.ActiveRow
    tblBldList.Col = TblColumn.tcENTNM
    
    If tblBldList.value <> GetEmpNm(ObjSysInfo.EmpId) Then
        MsgBox "이 혈액을 입고한 사용자만이 혈액용량을 변경할 수 있습니다.", vbCritical
        fraUpdate.Visible = False
        Exit Sub
    End If
    
    strSQL = objSql.UpdateVolumn401(txtVol.Text, lblBldSrc.Caption, lblBldYY.Caption, lblBldNo.Caption, lblCompo.Caption)
        
On Error GoTo ErrUpdateData
    With DBConn
        .BeginTrans
        
        .Execute strSQL
    
        .CommitTrans
    End With
    
    With tblBldList
        .Row = .ActiveRow
        .Col = TblColumn.tcVol: .value = txtVol.Text & "cc"
    End With
    
    fraUpdate.Visible = False
    
    Exit Sub
ErrUpdateData:
     With DBConn
        .RollbackTrans
        .DisplayErrors
    End With
End Sub

Private Sub Command1_Click()
    Dim iCnt, iCol, iRow As Integer
    Dim varTmp
    Dim strBMS As String
    Dim Resp
    
    Resp = MsgBox(lblCnt.Caption & " 건 BMS 전송하시겠습니까 ? ", vbQuestion + vbOKCancel, "확인")
    
    If Resp = vbCancel Then Exit Sub
    
    tblBMSList1.MaxRows = 0
    strBMS = ""
    
    With tblBMSList1
        .MaxRows = .MaxRows + 1
        .SetText 1, 1, "혈액번호"
        .SetText 2, 1, "혈액제제코드"
        .SetText 3, 1, "혈액제제명"
        .SetText 4, 1, "입고구분"
        .SetText 5, 1, "입고일자"
        .SetText 6, 1, "입고시간"
        .SetText 7, 1, "채혈일자"
        .SetText 8, 1, "혈액형코드"
        .SetText 9, 1, "혈액형명"
        .SetText 10, 1, "입고자"
    End With
    
    With tblBldList
        For iCnt = 1 To .MaxRows
            tblBMSList1.MaxRows = tblBMSList1.MaxRows + 1
            .GetText 3, iCnt, varTmp: tblBMSList1.SetText 1, iCnt + 1, varTmp
            .GetText 7, iCnt, varTmp
            If varTmp = "320cc" Then
                tblBMSList1.SetText 2, iCnt + 1, "01"
            Else
                tblBMSList1.SetText 2, iCnt + 1, "51"
            End If
            tblBMSList1.SetText 3, iCnt + 1, "농축적혈구"
            .GetText 10, iCnt, varTmp
            If varTmp = "입고" Then
                tblBMSList1.SetText 4, iCnt + 1, "Y"
            Else
                tblBMSList1.SetText 4, iCnt + 1, "R"
            End If
            .GetText 6, iCnt, varTmp
            tblBMSList1.SetText 5, iCnt + 1, Mid(varTmp, 1, 10)
            .GetText 34, iCnt, varTmp
            tblBMSList1.SetText 6, iCnt + 1, Trim(varTmp)
            .GetText 4, iCnt, varTmp
            tblBMSList1.SetText 7, iCnt + 1, Trim(varTmp)
            .GetText 2, iCnt, varTmp
            Select Case varTmp
                Case "O+"
                    tblBMSList1.SetText 8, iCnt + 1, "01"
                    tblBMSList1.SetText 9, iCnt + 1, "O (+)"
                Case "A+"
                    tblBMSList1.SetText 8, iCnt + 1, "02"
                    tblBMSList1.SetText 9, iCnt + 1, "A (+)"
                Case "B+"
                    tblBMSList1.SetText 8, iCnt + 1, "03"
                    tblBMSList1.SetText 9, iCnt + 1, "B (+)"
'                Case "O+"
'                    tblBMSList.SetText 7, iCnt + 1, 1
'                    tblBMSList.SetText 8, iCnt + 1, "O (+)"
                Case "AB+"
                    tblBMSList1.SetText 8, iCnt + 1, "04"
                    tblBMSList1.SetText 9, iCnt + 1, "AB (+)"
                Case "O-"
                    tblBMSList1.SetText 8, iCnt + 1, "05"
                    tblBMSList1.SetText 9, iCnt + 1, "O (-)"
                Case "A-"
                    tblBMSList1.SetText 8, iCnt + 1, "06"
                    tblBMSList1.SetText 9, iCnt + 1, "A (-)"
                Case "B-+"
                    tblBMSList1.SetText 8, iCnt + 1, "07"
                    tblBMSList1.SetText 9, iCnt + 1, "B (-)"
                Case "AB-"
                    tblBMSList1.SetText 8, iCnt + 1, "08"
                    tblBMSList1.SetText 9, iCnt + 1, "AB (-)"
            End Select
            .GetText 8, iCnt, varTmp: tblBMSList1.SetText 10, iCnt + 1, varTmp
        Next
    End With
    
    With tblBMSList1
        For iCnt = 1 To .MaxRows
            .GetText 1, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 2, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 3, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 4, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 5, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 6, iCnt, varTmp:
            
            Debug.Print varTmp
            If InStr(varTmp, ":") = 2 Then
                strBMS = strBMS & "0" & varTmp & ","
            Else
                If Left((Trim(varTmp & "")), 1) = ":" And Len(Trim(varTmp & "")) = 3 Then
                   strBMS = strBMS & "00" & varTmp & ","
                Else
                strBMS = strBMS & varTmp & ","
                End If
            End If
            .GetText 7, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 8, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 9, iCnt, varTmp: strBMS = strBMS & varTmp & ","
            .GetText 10, iCnt, varTmp: strBMS = strBMS & varTmp & vbCrLf
        Next
    End With
    
    Dim strPath As String
    Dim FreeFileNum As String

    strPath = "C:\BMS_File.csv"

    If Dir("C:\BMS_File.csv") <> "" Then
        Kill "C:\BMS_File.csv"
    End If

    FreeFileNum = FreeFile
    Open strPath For Append As #FreeFileNum
    Print #FreeFileNum, strBMS
    Close #FreeFileNum
    
    Call MsgBox("BMS전송이 완료되었습니다.", vbExclamation, "BMS전송완료")
End Sub

Private Sub Form_Load()
    Dim objBBSsql   As clsGetSqlStatement
    Dim RS          As Recordset
    Dim ii          As Long
    
    Set objBBSsql = New clsGetSqlStatement
    Set RS = objBBSsql.Get_CompoRecordSet
    
    '혈액제제
    With RS
        cboCompo.Clear
        cboCompo.AddItem "전체혈액제제"
        For ii = 1 To .RecordCount
             cboCompo.AddItem .Fields("compocd").value & "" & Space(2) & .Fields("abbrnm").value & ""
            .MoveNext
        Next ii
        cboCompo.ListIndex = 0
    End With

    Call SetCenterCombo

    dtpFrom = DateAdd("d", -15, GetSystemDate)
    dtpTo = GetSystemDate

    chkUse.value = False
    txtPtId.Text = ""
    lblPtNm.Caption = ""
    txtPtId.Visible = False: lblPtNm.Visible = False
    Clear
    

    Set RS = Nothing
    Set objBBSsql = Nothing
End Sub
Private Sub SetCenterCombo()
    Dim objcom003 As clsCom003
    Dim i As Long

    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter, True)
    Set objcom003 = Nothing

    cboCenter.ListIndex = -1

    For i = 0 To cboCenter.ListCount - 1
        If ObjSysInfo.BuildingCd = medGetP(cboCenter.List(i), 1, " ") Then
            cboCenter.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub


Private Sub lvwHosB_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'소트
    Static blnToggle() As Boolean
    Static blnFirst As Boolean
    Dim i As Long
    
    If blnFirst = False Then
        ReDim blnToggle(lvwHosB.ColumnHeaders.Count - 1)
        blnFirst = True
    End If
    
    '▲▼
    
    For i = 1 To lvwHosB.ColumnHeaders.Count
        lvwHosB.ColumnHeaders(i).Text = Trim(Replace(lvwHosB.ColumnHeaders(i).Text, "▲", ""))
        lvwHosB.ColumnHeaders(i).Text = Trim(Replace(lvwHosB.ColumnHeaders(i).Text, "▼", ""))
    Next
    
    With lvwHosB
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(blnToggle(ColumnHeader.Index - 1), lvwDescending, lvwAscending)
        .Sorted = True
        
        ColumnHeader.Text = ColumnHeader.Text & " " & IIf(.SortOrder = lvwAscending, "▲", "▼")
        
        blnToggle(ColumnHeader.Index - 1) = IIf(blnToggle(ColumnHeader.Index - 1), False, True)
    End With
    
    If lvwHosB.ListItems.Count <> 0 Then
        lvwHosB.ListItems(1).Selected = True
        lvwHosB.ListItems(1).EnsureVisible
    End If
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            Dim strBldSrc As String
            Dim strBldYY  As String
            Dim strBldNo  As String
            Dim strCompo  As String
            
            Dim strTmp    As String
            Dim SSQL      As String
            
            With tblBldList
                .Row = .ActiveRow
                .Col = TblColumn.tcBldNo
                strBldSrc = medGetP(.value, 1, "-")
                strBldYY = medGetP(.value, 2, "-")
                strBldNo = Format(medGetP(.value, 3, "-"), "######")
                .Col = 35: strCompo = .value
            
                'If objSql.IpGoPossibleFg(strBldSrc, strBldYY, strBldNo, strCompo) = False Then Exit Sub
                
                strTmp = MsgBox("혈액을 입고 취소하시겠습니까?", vbExclamation + vbDefaultButton2 + vbYesNo)
                If strTmp = vbNo Then Exit Sub
                
                '입고자와 현재 로그인 사용자를 비교해서 다른 경우에는 입고취소 할 수 없도록 한다.
                .Col = TblColumn.tcENTNM
                If .value <> GetEmpNm(ObjSysInfo.EmpId) Then
                    MsgBox "이 혈액을 입고한 사용자만이 입고 취소할 수 있습니다.", vbCritical
                    Exit Sub
                End If
                
                
            On Error GoTo SAVE_ERROR
                DBConn.BeginTrans
                
                SSQL = objSql.IpGoCancel(strBldSrc, strBldYY, strBldNo, strCompo)
                
                Debug.Print SSQL
                DBConn.Execute SSQL
                
                DBConn.CommitTrans
                .Action = ActionDeleteRow
                
                lblCnt.Caption = CLng(lblCnt.Caption) - 1
                MsgBox "입고 취소되었습니다.", vbInformation + vbOKOnly, "혈액입고취소"
                Exit Sub
            End With
SAVE_ERROR:
            DBConn.RollbackTrans
            MsgBox Err.Description, vbExclamation
    End Select
End Sub

Private Sub optABO_Click(Index As Integer)
    lblABO = optABO(Index).Caption

    If optRh(0).value = True Then
        lblABO = lblABO & "+"
    ElseIf optRh(1).value = True Then
        lblABO = lblABO & "-"
    End If
End Sub

Private Sub optall_Click()
    Dim ii As Integer
    
    If optall.value = 1 Then
        For ii = 0 To 3
            optABO(ii).value = False
        Next
        For ii = 0 To 1
            optRh(ii).value = False
        Next
        fraABO.Enabled = False
        lblABO.Caption = ""
        optall.Caption = "전체"
    Else
        fraABO.Enabled = True
        optall.Caption = "선택"
    End If
End Sub

Private Sub optBlood_Click(Index As Integer)
    If Index = 0 Then
        lblCondition.Caption = "입고일자"
    Else
        lblCondition.Caption = "출고일자"
    End If
    
    lvwHosB.ListItems.Clear
    lvwHosB.ToolTipText = ""
    
    Call BloodQueryPre
End Sub

Private Sub optRh_Click(Index As Integer)
    Dim i As Long
    Dim TF As Boolean
    
    For i = 0 To 3
        If optABO(i).value = True Then
            lblABO = optABO(i).Caption
            TF = True
            Exit For
        End If
    Next i
    If TF = False Then
        lblABO = optRh(Index).Caption
    Else
        lblABO = lblABO & optRh(Index).Caption
    End If
End Sub

Private Sub QUERY_USED()
    Dim objProBar As New clsProgress
    Dim RS As Recordset
    Dim ADt As String
    
    Dim ii As Integer
    
    tblBldList.MaxRows = 0
    lblCnt.Caption = 0
    Set RS = objSql.GetBlood_UsedList
    
    If RS.RecordCount < 1 Then GoTo Skip
    DoEvents
    lblCnt.Caption = RS.RecordCount
    
'    Set objProBar.StatusBar = medMain.stsBar
    objProBar.Container = MainFrm.stsBar
    objProBar.Min = 1
    objProBar.Max = RS.RecordCount
    
    With tblBldList
        .ReDraw = False
        .MaxRows = RS.RecordCount
        
        Do Until RS.EOF
            ii = ii + 1
            .Row = ii
            .Col = TblColumn.tcCOMPONM: .value = RS.Fields("componm").value & ""
            .Col = TblColumn.tcABO:     .value = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
            .Col = TblColumn.tcBldNo:   .value = RS.Fields("bldsrc").value & "" & "-" & _
                                                 RS.Fields("bldyy").value & "" & "-" & _
                                                 Format(RS.Fields("bldno").value & "", "00000#")
            '병원헌혈입고일경우 붉은색
            If RS.Fields("hosfg").value & "" = "1" Then .ForeColor = vbRed
            
            .Col = TblColumn.tcColDt:   .value = Format(RS.Fields("coldt").value & "", "####-##-##")
            ADt = Format(DateAdd("d", Val(RS.Fields("available").value & "") - 1, .value), "YYYY-MM-dd")
            .Col = TblColumn.tcAVAIL:   .value = ADt: .ForeColor = DCM_LightBlue
            
            .Col = TblColumn.tcENTDT:   .value = Format(RS.Fields("entdt").value & "", "####-##-##")
            .Col = TblColumn.tcVol:     .value = RS.Fields("volumn").value & "" & "cc"
            .Col = TblColumn.tcENTNM:   .value = GetEmpNm(RS.Fields("entid").value & "")
            .Col = TblColumn.tcIRR: .value = IIf(RS.Fields("irrfg").value & "" = "1", "Y", ""): .ForeColor = DCM_LightRed
            
'''            .Col = 31: .value = RS.Fields("compocd").value & ""
'''            .Col = 32: .value = RS.Fields("stscd").value & ""
            
            '==> 2014-03-03 칼럼추가
            .Col = 35: .value = RS.Fields("compocd").value & ""
            .Col = 36: .value = RS.Fields("stscd").value & ""
            '<==
            .Col = TblColumn.tcSTATUS
            Select Case RS.Fields("stscd").value & ""
                Case BBSBloodStatus.stsENTER:  .value = "입고": .ForeColor = BBSOrderStatusColor.clrOrder
                Case BBSBloodStatus.stsRETURN: .value = "반환": .ForeColor = BBSOrderStatusColor.clrCOLLECT
            End Select
            
            .Col = TblColumn.tcLARC: .value = RS.Fields("larc").value & ""
            .Col = TblColumn.tcSMLC: .value = RS.Fields("smlc").value & ""
            .Col = TblColumn.tcLARE: .value = RS.Fields("lare").value & ""
            .Col = TblColumn.tcSMLE: .value = RS.Fields("smle").value & ""
                        
            '.value = rs.Fields("stscd").value & ""
            
            
            objProBar.value = ii
            RS.MoveNext
        Loop
        .ReDraw = True
    End With
Skip:
    
    Set RS = Nothing
    Set objProBar = Nothing
End Sub
Private Sub ALL_QUERY()
    Dim objProBar As New clsProgress
    Dim ObjABO    As New clsABO
    Dim RS        As Recordset
    Dim RsS       As Recordset
    Dim Ret       As Recordset
    Dim ADt       As String
    Dim BldSrc    As String
    Dim BldYY     As String
    Dim BldNo     As String
    Dim CompoCd   As String
    Dim SexTmp    As String
    Dim donorid   As String
    Dim DonorDt   As String
    
    Dim ii As Integer
    
    tblBldList.MaxRows = 0
    
    
    If chkDelivery.value = 0 Then
        Set RS = objSql.GetBlood_AllList
    Else
        Set RS = objSql.GetBlood_DeliveryList   '출고현황
    End If
    
    lblCnt.Caption = 0
    If RS.RecordCount < 1 Then GoTo Skip
    
    DoEvents
    lblCnt.Caption = RS.RecordCount
    
'    Set objProBar.StatusBar = medMain.stsBar
    objProBar.Container = MainFrm.stsBar
    objProBar.Min = 1
    objProBar.Max = RS.RecordCount
    
    With tblBldList
        .ReDraw = False
        .MaxRows = RS.RecordCount
        
        Do Until RS.EOF
                        
            ii = ii + 1
            .Row = ii
            BldSrc = RS.Fields("bldsrc").value & ""
            BldYY = RS.Fields("bldyy").value & ""
            BldNo = RS.Fields("bldno").value & ""
            CompoCd = RS.Fields("compocd").value & ""
            donorid = RS.Fields("donorid").value & ""
            DonorDt = RS.Fields("donoraccdt").value & ""
            
            .Col = TblColumn.tcCOMPONM: .value = RS.Fields("componm").value & ""
            .Col = TblColumn.tcABO:     .value = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
            .Col = TblColumn.tcBldNo:   .value = BldSrc & "-" & _
                                                 BldYY & "-" & _
                                                 Format(BldNo, "00000#")
            '헌혈입고여부
            If RS.Fields("hosfg").value & "" = "1" Then .ForeColor = vbRed
            
            .Col = TblColumn.tcColDt:   .value = Format(RS.Fields("coldt").value & "", "####-##-##")
                        
            ADt = Format(DateAdd("d", Val(RS.Fields("available").value & "") - 1, .value), "YYYY-MM-dd")
            
            .Col = TblColumn.tcAVAIL:   .value = ADt: .ForeColor = DCM_LightBlue
            .Col = TblColumn.tcENTDT:   .value = Format(RS.Fields("entdt").value & "", "####-##-##")
            .Col = TblColumn.tcVol:     .value = RS.Fields("volumn").value & "" & "cc"
            .Col = TblColumn.tcENTNM:   .value = GetEmpNm(RS.Fields("entid").value & "")
            .Col = TblColumn.tcIRR:     .value = IIf(RS.Fields("irrfg").value & "" = "1", "Y", ""): .ForeColor = DCM_LightRed
            
            .Col = TblColumn.tcLARC: .value = RS.Fields("larc").value & ""
            .Col = TblColumn.tcSMLC: .value = RS.Fields("smlc").value & ""
            .Col = TblColumn.tcLARE: .value = RS.Fields("lare").value & ""
            .Col = TblColumn.tcSMLE: .value = RS.Fields("smle").value & ""
            
            '.Col = 31: .value = CompoCd
            '.Col = 32: .value = RS.Fields("stscd").value & ""
            
            '==> 2014-03-03 칼럼추가
            .Col = 35: .value = CompoCd
            .Col = 36: .value = RS.Fields("stscd").value & ""
            '<==
            
            Select Case RS.Fields("stscd").value & ""
                Case BBSBloodStatus.stsENTER
                        .Col = TblColumn.tcSTATUS: .value = "입고": .ForeColor = BBSOrderStatusColor.clrOrder
                Case BBSBloodStatus.stsRETURN
                        .Col = TblColumn.tcSTATUS: .value = "반환": .ForeColor = BBSOrderStatusColor.clrCOLLECT
                        '** 항목추가요청 By M.G.Choi 2008.04.17 --------------------------------------------------
                        Set Ret = New Recordset
                        Ret.Open objSql.GetBlood_SubList(BldSrc, BldYY, BldNo, CompoCd), DBConn
                        
                        If Not Ret.EOF Then
                            '.Col = 24: .value = Format(Ret.Fields("retdt").value & "", "####-##-##")
                            '.Col = 25: .value = GetEmpNm(Ret.Fields("retid").value & "")
                            '.Col = 26: .value = GetEmpNm(Ret.Fields("retrcvid").value & "")
                            
                            '==> 2014-03-03 칼럼추가
                            .Col = 28: .value = Format(Ret.Fields("retdt").value & "", "####-##-##")
                            .Col = 29: .value = GetEmpNm(Ret.Fields("retid").value & "")
                            .Col = 30: .value = GetEmpNm(Ret.Fields("retrcvid").value & "")
                            '<==
                        End If
                        
                        Ret.Close: Set Ret = Nothing
                        '-----------------------------------------------------------------------------------------
                        
                        '출고된 환자(있음)
                Case BBSBloodStatus.stsASSIGN
                        .Col = TblColumn.tcSTATUS: .value = "ASSIGN": .ForeColor = BBSOrderStatusColor.clrACCESS
                        Set RsS = objSql.AssignBloodRs(BldSrc, BldYY, BldNo, CompoCd)
                        If Not RsS.EOF Then
                            SexTmp = SDA_String(RsS.Fields("ssn").value & "")
                            .Col = TblColumn.tcPTID: .value = RsS.Fields("ptid").value & ""
                            .Col = TblColumn.tcPTNM: .value = RsS.Fields("ptnm").value & ""
                            .Col = TblColumn.tcSEX:  .value = medGetP(SexTmp, 1, COL_DIV)
                            .Col = TblColumn.tcAGE:  .value = Mid(medGetP(SexTmp, 2, COL_DIV), 1, 4) '.value = medGetP(SexTmp, 3, COL_DIV)
                            .Col = TblColumn.tcDEPT: .value = RsS.Fields("deptcd").value & ""
                            ObjABO.Ptid = RsS.Fields("ptid").value & ""
                            If ObjABO.GetABO = False Then
                                .Col = TblColumn.tcABO_P: .value = ""
                            Else
                                .Col = TblColumn.tcABO_P: .value = ObjABO.ABO & ObjABO.Rh
                            End If
                            
                            .Col = TblColumn.tcVFYDT:     .value = Format(RsS.Fields("vfydt").value & "", "####-##-##")
                            .Col = TblColumn.tcVFYNM:     .value = GetEmpNm(RsS.Fields("vfyid").value & "")
                            .Col = TblColumn.tcSTAT:      .value = IIf(RsS.Fields("stat").value & "" = "1", "Y", "")
                            .Col = TblColumn.tcRSTV:
                            
                            Select Case RsS.Fields("rstv").value & ""
                                Case "1": .value = "OK"
                                Case "2": .value = "NOT"
                                Case Else: .value = ""
                            End Select
                        End If
                        Set RsS = Nothing
                Case BBSBloodStatus.stsDELIVERY, BBSBloodStatus.stsEXPIRE
                        If RS.Fields("stscd").value & "" = BBSBloodStatus.stsDELIVERY Then
                            If RS.Fields("splitoutfg").value & "" = "1" Then
                                .Col = TblColumn.tcSTATUS: .value = "분획": .ForeColor = BBSOrderStatusColor.clrINPROCESS
                            Else
                                .Col = TblColumn.tcSTATUS: .value = "출고": .ForeColor = BBSOrderStatusColor.clrINPROCESS
                            End If
                        Else
                            .Col = TblColumn.tcSTATUS: .value = "폐기": .ForeColor = BBSOrderStatusColor.cIrEND
                        End If
                        '출고된 환자(있음)
                        Set RsS = objSql.DeLBloodRs(BldSrc, BldYY, BldNo, CompoCd)
                        If Not RsS.EOF Then
                            SexTmp = SDA_String(RsS.Fields("ssn").value & "")
                            .Col = TblColumn.tcPTID: .value = RsS.Fields("ptid").value & ""
                            .Col = TblColumn.tcPTNM: .value = RsS.Fields("ptnm").value & ""
                            .Col = TblColumn.tcSEX: .value = medGetP(SexTmp, 1, COL_DIV)
                            .Col = TblColumn.tcAGE: .value = Mid(medGetP(SexTmp, 2, COL_DIV), 1, 4) '.value = medGetP(SexTmp, 3, COL_DIV)
                            .Col = TblColumn.tcDEPT: .value = RsS.Fields("deptcd").value & ""
                            ObjABO.Ptid = RsS.Fields("ptid").value & ""
                            If ObjABO.GetABO = False Then
                                .Col = TblColumn.tcABO_P: .value = ""
                            Else
                                .Col = TblColumn.tcABO_P: .value = ObjABO.ABO & ObjABO.Rh
                            End If
                            .Col = TblColumn.tcDELDT: .value = Format(Mid(RsS.Fields("deliverydt").value & "", 1, 12), "####-##-## ##:##")
                            .Col = TblColumn.tcDELNM: .value = GetEmpNm(RsS.Fields("deliveryid").value & "")
                            .Col = TblColumn.tcDELRCVNM: .value = GetEmpNm(RsS.Fields("rcvid").value & "")
                            If Trim(RsS.Fields("localnm").value & "") <> "" Then .value = RsS.Fields("localnm").value & ""
                        End If
                        Set RsS = Nothing
                        
                        '** 항목추가요청 By M.G.Choi 2008.04.17 --------------------------------------------------
'''                        .Col = 27:
'''                        If RS.Fields("realexpdt").value & "" <> "" Then
'''                            .value = Format(RS.Fields("realexpdt").value & "", "####-##-##")
'''                        End If
'''                        .Col = 28: .value = GetEmpNm(RS.Fields("exprcvid").value & "")
'''                        .Col = 29: .value = GetEmpNm(RS.Fields("expid").value & "")
'''                        .Col = 30: .value = RS.Fields("exprsncd").value & ""
                        '-----------------------------------------------------------------------------------------
                        
                        '** 항목추가요청 2014.03.03 칼럼추가 --------------------------------------------------
                        .Col = 31:
                        If RS.Fields("realexpdt").value & "" <> "" Then
                            .value = Format(RS.Fields("realexpdt").value & "", "####-##-##")
                        End If
                        .Col = 32: .value = GetEmpNm(RS.Fields("exprcvid").value & "")
                        .Col = 33: .value = GetEmpNm(RS.Fields("expid").value & "")
                        .Col = 34: .value = RS.Fields("exprsncd").value & ""
                        '-----------------------------------------------------------------------------------------
            
                Case BBSBloodStatus.stsBAG
                        .Col = TblColumn.tcSTATUS: .value = "회수":: .ForeColor = DCM_LightGray
            End Select
            
            Set RsS = Nothing
            Set RsS = New Recordset
            RsS.Open objSql.GetBloodOkNotFg(donorid, DonorDt), DBConn
            If Not RsS.EOF Then
                '검사결과 적격/부적격(헌혈)
                .Col = TblColumn.tcSTATUS: .value = IIf(RsS.Fields("okdiv3").value & "" = "0", "부적격", .value)
                If .value = "부적격" Then
                    .ForeColor = vbRed
                    .Col = TblColumn.tcBldNo: .ForeColor = vbRed
                End If
            End If
            Set RsS = Nothing
            
            objProBar.value = ii
            RS.MoveNext
        Loop
        .ReDraw = True
    End With
Skip:
    
    Set RS = Nothing
    Set ObjABO = Nothing
    Set objProBar = Nothing
    
End Sub




'Private Sub mnuDelete_Click()
'    Dim strBldSrc As String
'    Dim strBldYY  As String
'    Dim strBldno  As String
'    Dim strCompo  As String
'
'    Dim strTmp    As String
'    Dim SSQL      As String
'
'    With tblBldList
'        .Row = .ActiveRow
'        .Col = TblColumn.tcBLDNO
'        strBldSrc = medGetP(.value, 1, "-")
'        strBldYY = medGetP(.value, 2, "-")
'        strBldno = Format(medGetP(.value, 3, "-"), "######")
'        .Col = 31: strCompo = .value
'
'        'If objSql.IpGoPossibleFg(strBldSrc, strBldYY, strBldNo, strCompo) = False Then Exit Sub
'
'        strTmp = MsgBox("혈액을 입고 취소하시겠습니까?", vbInformation + vbYesNo, "혈액입고취소")
'        If strTmp = vbNo Then Exit Sub
'
'    On Error GoTo SAVE_ERROR
'        dbconn.BeginTrans
'
'        SSQL = objSql.IpGoCancel(strBldSrc, strBldYY, strBldno, strCompo)
'        dbconn.Execute SSQL
'
'        dbconn.CommitTrans
'        .Action = ActionDeleteRow
'
'        lblCnt.Caption = CLng(lblCnt.Caption) - 1
'        MsgBox "입고 취소되었습니다.", vbInformation + vbOKOnly, "혈액입고취소"
'        Exit Sub
'    End With
'SAVE_ERROR:
'    dbconn.RollbackTrans
'    'dbconn.DisplayErrors
'
'End Sub

Private Sub tblBldList_DblClick(ByVal Col As Long, ByVal Row As Long)
    '혈액용량변경
    If tblBldList.DataRowCnt < 1 Then Exit Sub
    
    If Row < 1 Then Exit Sub
       
    With tblBldList
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
        .Col = 32:
        If .value > BBSBloodStatus.stsASSIGN Then Exit Sub
        .Col = TblColumn.tcBldNo
        lblBldSrc.Caption = medGetP(.value, 1, "-")
        lblBldYY.Caption = medGetP(.value, 2, "-")
        lblBldNo.Caption = Format(medGetP(.value, 3, "-"), "######")
        .Col = 31: lblCompo.Caption = .value
        .Col = TblColumn.tcVol:     lblVol.Caption = .value
        fraUpdate.Visible = True
    End With
   
End Sub

Private Sub tblBldList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If tblBldList.DataRowCnt < 1 Then Exit Sub
    If Row < 1 Then Exit Sub
    
    Dim strBldSrc As String
    Dim strBldYY  As String
    Dim strBldNo  As String
    Dim strCompo  As String
    
    With tblBldList
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
    
        .Col = TblColumn.tcBldNo
        strBldSrc = medGetP(.value, 1, "-")
        strBldYY = medGetP(.value, 2, "-")
        strBldNo = Format(medGetP(.value, 3, "-"), "######")
        .Col = 35: strCompo = .value
    
    End With
    
    If objSql.IpGoPossibleFg(strBldSrc, strBldYY, strBldNo, strCompo) = False Then Exit Sub
    
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "입고취소"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
End Sub

Private Sub txtPtId_Change()
    If lblPtNm.Caption <> "" Then
        lblPtNm.Caption = ""
    End If
End Sub


Private Sub txtPtId_LostFocus()
    If txtPtId = "" Then Exit Sub

    Call Query_Pt(txtPtId)

End Sub
Private Function Query_Pt(ByVal Ptid As String) As Boolean
    Dim objMeSql As clsGetSqlStatement
    Dim strTmp   As String
    
    Set objMeSql = New clsGetSqlStatement
    
    strTmp = objMeSql.TransPtidHistory(Ptid, Format(dtpFrom.value, PRESENTDATE_FORMAT), Format(dtpTo.value, PRESENTDATE_FORMAT))
    If strTmp <> "" Then
        txtPtId.Text = Ptid
        lblPtNm.Caption = medGetP(strTmp, 1, COL_DIV)
        tblBldList.MaxRows = 0
        cmdQuery.SetFocus
    Else
        MsgBox "해당환자가 존재하지 않습니다.확인후 조회하세요.", vbInformation + vbOKOnly, "환자조회"
        txtPtId = ""
        lblPtNm.Caption = ""
    End If
    
    Set objMeSql = Nothing
End Function
Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblBldList.DataRowCnt = 0 And tblBldList.DataRowCnt = 0 Then Exit Sub
    
    With tblBldList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblBldList.MaxRows + 1
        .MaxCols = tblBldList.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = tblBldList.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    
    If chkUse.value = 1 Then
        DlgSave.FileName = "사용가능혈액대장"
    Else
        DlgSave.FileName = "혈액대장"
    End If
    If chkDelivery.value = 1 Then
        DlgSave.FileName = "출고대장"
    Else
        DlgSave.FileName = "혈액대장"
    End If
    
    
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub
