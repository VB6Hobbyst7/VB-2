VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm101Order 
   BackColor       =   &H00DBE6E6&
   Caption         =   "처방등록"
   ClientHeight    =   9225
   ClientLeft      =   -420
   ClientTop       =   825
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14625
   WindowState     =   2  '최대화
   Begin VB.ListBox lstTestList 
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
      Height          =   2565
      Left            =   7200
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   3165
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.PictureBox fraDatePicker 
      BackColor       =   &H00FFFEEE&
      Height          =   450
      Left            =   11925
      ScaleHeight     =   390
      ScaleWidth      =   1875
      TabIndex        =   52
      Top             =   3165
      Visible         =   0   'False
      Width           =   1935
      Begin MSComCtl2.DTPicker txtDatePicker 
         Height          =   360
         Left            =   15
         TabIndex        =   53
         Top             =   15
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyy-MM-dd HH:mm"
         Format          =   83558403
         CurrentDate     =   36328
      End
   End
   Begin VB.TextBox txtReceptNo 
      Height          =   330
      Left            =   6465
      TabIndex        =   44
      Text            =   "영수증번호"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  '평면
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7890
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   40
      ToolTipText     =   "검사 리마크를 입력하세요."
      Top             =   7650
      Width           =   6345
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Caption         =   " ◈ 다빈도 처방 리스트"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7605
      Left            =   75
      TabIndex        =   37
      Top             =   1425
      Width           =   6420
      Begin FPSpread.vaSpread tblTestList 
         Height          =   7140
         Left            =   165
         TabIndex        =   38
         Top             =   270
         Width           =   6060
         _Version        =   196608
         _ExtentX        =   10689
         _ExtentY        =   12594
         _StockProps     =   64
         AllowUserFormulas=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         AutoSize        =   -1  'True
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
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   28
         MaxRows         =   38
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "Lis101.frx":0000
         VisibleCols     =   4
         VisibleRows     =   25
      End
   End
   Begin VB.ListBox lstSpcList 
      Appearance      =   0  '평면
      BackColor       =   &H00FCE9F7&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11100
      TabIndex        =   33
      Top             =   3180
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00FDDDF0&
      Caption         =   "Item Move ▼"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   6570
      Style           =   1  '그래픽
      TabIndex        =   35
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Frame fraRecept 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Repeat검사여부"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9105
      TabIndex        =   34
      Tag             =   "10105"
      Top             =   1425
      Width           =   2400
      Begin VB.CheckBox Check2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Repeat 검사"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   510
         TabIndex        =   42
         Tag             =   "10101"
         Top             =   270
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "처방 && 채혈(&B)"
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
      Index           =   1
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "131"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "처방 && 접수(&R)"
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
      Index           =   2
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "131"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame framPtInfo 
      BackColor       =   &H00DBE6E6&
      Caption         =   " ◈ 환자 기본정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   60
      TabIndex        =   18
      Tag             =   "104"
      Top             =   60
      Width           =   14400
      Begin VB.TextBox txtSpecNm 
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
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   72
         Top             =   690
         Width           =   1875
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   8850
         TabIndex        =   49
         Top             =   330
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
         Index           =   5
         Left            =   8850
         TabIndex        =   50
         Top             =   720
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   3315
         TabIndex        =   47
         Top             =   690
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
         Caption         =   "검체번호"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   2670
         Style           =   1  '그래픽
         TabIndex        =   43
         Tag             =   "PtID"
         Top             =   315
         Width           =   300
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   10965
         MaskColor       =   &H00F4F0F2&
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   32
         Tag             =   "DeptCd"
         Top             =   705
         Width           =   285
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   10965
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   31
         Tag             =   "DoctId"
         Top             =   330
         Width           =   285
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   12060
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   30
         Tag             =   "WardId"
         Top             =   705
         Width           =   315
      End
      Begin VB.TextBox txtBedId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
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
         Height          =   225
         Left            =   13815
         TabIndex        =   5
         Top             =   795
         Width           =   390
      End
      Begin VB.TextBox txtRoomId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13155
         TabIndex        =   4
         Top             =   795
         Width           =   525
      End
      Begin VB.TextBox txtWardId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   12420
         TabIndex        =   3
         Top             =   795
         Width           =   600
      End
      Begin VB.TextBox txtDeptCd 
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
         Height          =   315
         Left            =   9795
         TabIndex        =   2
         Top             =   720
         Width           =   1155
      End
      Begin MedControls1.LisLabel lblDob 
         Height          =   300
         Left            =   7170
         TabIndex        =   27
         Top             =   735
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   529
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   300
         Left            =   1185
         TabIndex        =   26
         Top             =   705
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4275
         ScaleHeight     =   315
         ScaleWidth      =   4365
         TabIndex        =   23
         Top             =   330
         Width           =   4395
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "외래"
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
            Height          =   240
            Index           =   3
            Left            =   3540
            TabIndex        =   70
            Tag             =   "10110"
            Top             =   60
            Width           =   720
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "응급"
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
            Height          =   240
            Index           =   2
            Left            =   2625
            TabIndex        =   25
            Tag             =   "10110"
            Top             =   60
            Width           =   690
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "외부정도관리"
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
            Left            =   975
            TabIndex        =   24
            Tag             =   "10109"
            Top             =   60
            Width           =   1440
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "외래"
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
            Left            =   75
            TabIndex        =   13
            Tag             =   "10108"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.TextBox txtDoctorId 
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
         Height          =   315
         Left            =   9795
         TabIndex        =   1
         Top             =   330
         Width           =   1155
      End
      Begin VB.TextBox txtPtId 
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
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   0
         Top             =   330
         Width           =   1485
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   330
         Left            =   11295
         TabIndex        =   28
         Top             =   330
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
         BackColor       =   15857140
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
         LeftGab         =   150
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   225
         TabIndex        =   45
         Top             =   330
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
         Index           =   0
         Left            =   225
         TabIndex        =   46
         Top             =   690
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
         Caption         =   "성    명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   6225
         TabIndex        =   48
         Top             =   720
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
         Index           =   7
         Left            =   11280
         TabIndex        =   51
         Top             =   720
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "병 실"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   3315
         TabIndex        =   71
         Top             =   330
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
         Caption         =   "처방구분"
         Appearance      =   0
      End
      Begin VB.Label lblWard 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Caption         =   "           -          -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12390
         TabIndex        =   29
         Tag             =   "107"
         Top             =   705
         Width           =   1860
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4365
         TabIndex        =   21
         Top             =   1110
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5280
         TabIndex        =   20
         Top             =   1140
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5700
         TabIndex        =   19
         Top             =   765
         Width           =   60
      End
      Begin VB.Label Label8 
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
         Height          =   315
         Left            =   4275
         TabIndex        =   22
         Top             =   1185
         Visible         =   0   'False
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "처방(&O)"
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
      Index           =   0
      Left            =   7860
      Style           =   1  '그래픽
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "131"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraColTm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "희망채취일시"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   11535
      TabIndex        =   16
      Tag             =   "10104"
      Top             =   1425
      Width           =   2940
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "아침채혈"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   105
         TabIndex        =   39
         Tag             =   "10101"
         Top             =   255
         Width           =   660
      End
      Begin MSComCtl2.DTPicker dtpColTime 
         Height          =   300
         Left            =   2010
         TabIndex        =   7
         Top             =   300
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
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
         CustomFormat    =   "HH:mm"
         Format          =   83558403
         UpDown          =   -1  'True
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpColDate 
         Height          =   300
         Left            =   780
         TabIndex        =   6
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
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
         Format          =   83558403
         CurrentDate     =   36328
      End
   End
   Begin VB.Frame fraPriority 
      BackColor       =   &H00DBE6E6&
      Caption         =   "응급여부"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7755
      TabIndex        =   14
      Tag             =   "10105"
      Top             =   1425
      Width           =   1290
      Begin VB.CheckBox chkStat 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Stat"
         Height          =   360
         Left            =   210
         TabIndex        =   15
         Tag             =   "10101"
         Top             =   270
         Width           =   945
      End
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   5310
      Left            =   6525
      TabIndex        =   36
      Tag             =   "10114"
      Top             =   2190
      Width           =   7935
      _Version        =   196608
      _ExtentX        =   13996
      _ExtentY        =   9366
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModePermanent=   -1  'True
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
      MaxCols         =   23
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis101.frx":22C5
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin VB.PictureBox fraStndOrder 
      BackColor       =   &H00F7F3F8&
      Height          =   4335
      Left            =   6600
      ScaleHeight     =   4275
      ScaleWidth      =   5835
      TabIndex        =   54
      Top             =   2760
      Width           =   5895
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DEDBDD&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3165
         Style           =   1  '그래픽
         TabIndex        =   64
         Top             =   2520
         Width           =   525
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00DEDBDD&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3165
         Style           =   1  '그래픽
         TabIndex        =   63
         Top             =   2145
         Width           =   525
      End
      Begin VB.ListBox lstHour2 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   2730
         ItemData        =   "Lis101.frx":4F35
         Left            =   2100
         List            =   "Lis101.frx":4F5D
         TabIndex        =   62
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DEDBDD&
         Caption         =   "Cancel"
         Height          =   405
         Left            =   3750
         Style           =   1  '그래픽
         TabIndex        =   61
         Tag             =   "123"
         Top             =   3750
         Width           =   810
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00DEDBDD&
         Caption         =   "OK"
         Height          =   405
         Left            =   4620
         Style           =   1  '그래픽
         TabIndex        =   60
         Tag             =   "130"
         Top             =   3750
         Width           =   810
      End
      Begin VB.ListBox lstColTime 
         Appearance      =   0  '평면
         Height          =   2280
         ItemData        =   "Lis101.frx":4F91
         Left            =   3780
         List            =   "Lis101.frx":4F98
         TabIndex        =   59
         Top             =   1380
         Width           =   1650
      End
      Begin VB.ListBox lstMinute 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   2730
         ItemData        =   "Lis101.frx":4FAC
         Left            =   2670
         List            =   "Lis101.frx":4FC2
         TabIndex        =   58
         Top             =   1380
         Width           =   420
      End
      Begin VB.ListBox lstHour1 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   2730
         ItemData        =   "Lis101.frx":4FDE
         Left            =   1650
         List            =   "Lis101.frx":5006
         TabIndex        =   57
         Top             =   1380
         Width           =   435
      End
      Begin VB.ListBox lstDate 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   2730
         ItemData        =   "Lis101.frx":503A
         Left            =   375
         List            =   "Lis101.frx":5068
         TabIndex        =   56
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   5715
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00F7F3F8&
         Caption         =   "Standing Order"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2010
         TabIndex        =   69
         Top             =   105
         Width           =   1845
      End
      Begin VB.Label lblSpecimen 
         BackStyle       =   0  '투명
         Caption         =   "Whole Blood"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   3780
         TabIndex        =   68
         Top             =   810
         Width           =   1545
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Date                     Hour               Minute"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   375
         TabIndex        =   67
         Tag             =   "10106"
         Top             =   1155
         Width           =   2760
      End
      Begin VB.Label lblSpecimen1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체명 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3765
         TabIndex        =   66
         Tag             =   "157"
         Top             =   585
         Width           =   750
      End
      Begin VB.Label lblTestName 
         BackStyle       =   0  '투명
         Caption         =   "Routine Chemistry"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   345
         TabIndex        =   65
         Top             =   810
         Width           =   2610
      End
      Begin VB.Label lblTest 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사명 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   345
         TabIndex        =   55
         Tag             =   "159"
         Top             =   600
         Width           =   750
      End
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ 처방Remark"
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
      Left            =   6645
      TabIndex        =   41
      Top             =   7680
      Width           =   1170
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E8FFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   900
      Index           =   0
      Left            =   6555
      Shape           =   4  '둥근 사각형
      Top             =   7560
      Width           =   7920
   End
End
Attribute VB_Name = "frm101Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents frmPtFind As frmPtInfo
Private WithEvents objSearchPt As clsPatient
Attribute objSearchPt.VB_VarHelpID = -1
Private WithEvents frmPtTemp As frmTmpPt
Attribute frmPtTemp.VB_VarHelpID = -1
Private WithEvents frmPtTempB As frmTmpPtB
Attribute frmPtTempB.VB_VarHelpID = -1
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private Const TAG_WARD& = 1
Private Const TAG_DEPT& = 2
Private Const TAG_DOCT& = 3
'Private WithEvents mnuPopup As menu
'Private WithEvents mnuDelete As menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DELETE& = 1

Private objSqlStmt As New clsLISSqlStatement     ' SQL 클래스
Private objPatient As New clsPatient
Private objOrder   As New clsLISOrder
Private objCollect As New clsLISCollectioin
Private objAccess  As New clsLISAccession

Private blnClearFg As Boolean
Private Const DabindoCols = 14

Private Sub cmdClear_Click()
   Call ClearRtn
   blnClearFg = True
   txtPtId.SetFocus
End Sub

Private Sub cmdMove_Click()

    Dim i As Integer
    Dim j As Integer

    With tblTestList
        For j = 0 To 1
            .Col = j * 14 + 1
            For i = 1 To .MaxRows
                .Row = i
                If .Value = 1 Then Call Item_Move(.Col, .Row)
            Next
        Next

        .Row = 1: .Row2 = .MaxRows
        .Col = enORDSHEET.tcORDNO
        .COL2 = enORDSHEET.tcORDNO
        .BlockMode = True
        .Value = 0
        .BlockMode = False
        .Col = enORDSHEET.tcORDSEQ
        .COL2 = enORDSHEET.tcORDSEQ
        .BlockMode = True
        .Value = 0
        .BlockMode = False
    End With

End Sub

Private Sub cmdHelpList_Click(Index As Integer)
'    Dim objData As clsBasisData
    
'    Set objData = New clsBasisData
'    Set objMyList = New clspopuplist
    Set objMyList = New clsPopUpList
    
    With objMyList
        Select Case Index
            Case 0
'                Set frmPtFind = frmPtInfo
'                frmPtFind.Show 1
                Set objSearchPt = New clsPatient
                objSearchPt.LoadSearchForm
            Case 1
                .Connection = DBConn
                .FormCaption = "처방의 조회"
                .Delimiter = COL_DIV
                .ColumnHeaderText = "처방의ID" & COL_DIV & "처방의명"
                .Tag = TAG_DOCT
                .LoadPopUp GetSQLDoctList
                
'                 .Caption = "처방의 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "처방의ID, 처방의명"
'                 Call .ListPop(getsqldoct) ', 1640, framPtInfo.Left + cmdHelpList(Index).Left)  ', objLisComCode.Doct)
                 
            Case 2
                .Connection = DBConn
                .FormCaption = "진료과 조회"
                .Delimiter = COL_DIV
                .ColumnHeaderText = "코 드" & COL_DIV & "진료과명"
                .Tag = TAG_DEPT
                .LoadPopUp GetSQLDeptList
                
'                 .Caption = "진료과 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "진료과코드, 진료과명"
'                 Call .ListPop(GetSQLDept) ', 1640, framPtInfo.Left + cmdHelpList(Index).Left)  ', objLisComCode.DeptCd)
                 If txtPtId <> "" Then
'                    PtInfoEnable False
'                    cmdHelpList(0).Enabled = False
                 End If
                 
            Case 3
                .Connection = DBConn
                .FormCaption = "병동 조회"
                .Delimiter = COL_DIV
                .ColumnHeaderText = "코 드" & COL_DIV & "병동명"
                .Tag = TAG_WARD
                .LoadPopUp GetSQLWardList
'                 .Caption = "병동 조회"
'                 .Tag = cmdHelpList(Index).Tag
'                 .HeadName = "병동코드,병동명"
'                 Call .ListPop(GetSQLWard) ', 1640, 10550)  ', objLisComCode.WardId)
                                
        End Select
    End With
'    Set objData = Nothing
    Set objMyList = Nothing
End Sub

Private Sub dtpColDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpColTime.SetFocus
    End If

End Sub

Private Sub dtpColTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tblOrdSheet.SetFocus
    End If

End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Deactivate()
    Set objMyList = Nothing
End Sub

'% 폼 로드...
Private Sub Form_Load()

    Dim tmpDate As Date
    Dim i As Integer
    Dim objProgress As clsProgress
    
    Set objProgress = New clsProgress
    
    Me.Show
    Call ClearRtn
    
    DoEvents
    
    With objProgress
        .Container = MainFrm.stsBar
        .Message = "검사항목 리스트를 로드하고 있습니다..."
    End With
    
'    objProgress.CaptionOn = False
'    objProgress.MSG = "검사항목 리스트를 로드하고 있습니다..."
'    objProgress.mode = 0
'    objProgress.Visible = True
'    objProgress.Value = 0
    
'    medMain.Enabled = False
'    Call LockPtInfo(False)
'    fraPriority.Enabled = False
'    fraColTm.Enabled = False
'    fraRecept.Enabled = False
'    tblOrdSheet.Enabled = False
    
    tblTestList.AutoSize = False
    tblTestList.AutoSize = True
    
    '다빈도 항목 / 검사항목 로드...
    
    MouseRunning
    
'    Call objOrder.SetDatabase(DbConn)
    Call objOrder.ItemList(lstTestList, objProgress)
    Call objOrder.DabindoList(tblTestList)
    objProgress.Value = objProgress.Max
'    objProgress.Visible = False
    Set objProgress = Nothing
    
    MouseDefault
    
    txtDoctorId = "001201"
    lblDoctNm.Caption = "이춘희"
    txtDeptCd.Text = "LABM"

    With objPatient
        .ptnm = lblPtNm.Caption
        .DeptCd = "LABM"
        .DoctNm = "이춘희"
        .MajDoct = "001201"
    End With
    
'    medMain.Enabled = True
'    Call LockPtInfo(True)
'    fraPriority.Enabled = True
'    fraColTm.Enabled = True
'    fraRecept.Enabled = True
'    tblOrdSheet.Enabled = True
'    txtBedId.Enabled = False
    txtPtId.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call objOrder.RemoveParameters
    '감염관리표시
    Call ICSPatientMark
    Set objOrder = Nothing
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
    Set objMyList = Nothing
    Set objSqlStmt = Nothing
    Set objPatient = Nothing
End Sub

Private Sub LockPtInfo(ByVal CtlEnabled As Boolean)
    framPtInfo.Enabled = CtlEnabled
    txtPtId.Enabled = CtlEnabled
    txtDoctorId.Enabled = CtlEnabled
    txtWardId.Enabled = CtlEnabled
    txtRoomId.Enabled = CtlEnabled
    txtBedId.Enabled = CtlEnabled

    If CtlEnabled Then
        txtPtId.BackColor = vbWhite
        txtDoctorId.BackColor = vbWhite
        txtDeptCd.BackColor = vbWhite
        lblWard.BackColor = vbWhite
        txtWardId.BackColor = vbWhite
        txtRoomId.BackColor = vbWhite
        txtBedId.BackColor = vbWhite
    Else
        txtPtId.BackColor = DCM_LightGray   '회색
        txtDoctorId.BackColor = DCM_LightGray
        txtDeptCd.BackColor = DCM_LightGray
        lblWard.BackColor = DCM_LightGray
        txtWardId.BackColor = DCM_LightGray
        txtRoomId.BackColor = DCM_LightGray
        txtBedId.BackColor = DCM_LightGray
    End If

End Sub

'Private Sub frmPtFind_Click(ByVal isSELECT As Boolean, ByVal ptInfo As S2LIS_CollectionLib.clsPtInformation)
'
'
'    If isSELECT Then
'        With ptInfo
'            txtPtId.Text = .PtId
'            lblPtNm.Caption = .ptnm
'            lblSex.Caption = .Sex
'            lblAge.Caption = .Age
'            lblAgeDiv.Caption = "Y"
'            lblDob.Caption = .DOB
'        End With
'        Call objPatient.getpatient(txtPtId.Text)
'        Call EnableButton(True)
'    End If
'
'End Sub

Private Sub frmPtTempB_OKButtonClick(ByVal strPtInfo As String)
    lblDob.Caption = medGetP(strPtInfo, 4, COL_DIV)
    lblAge.Caption = medGetP(strPtInfo, 5, COL_DIV)
    lblSex.Caption = medGetP(strPtInfo, 3, COL_DIV)
    lblPtNm.Caption = medGetP(strPtInfo, 2, COL_DIV)
    txtPtId.Text = medGetP(strPtInfo, 1, COL_DIV)
    
    Call EnableButton(True)
    
    '이하는 성모자애병원용
    
'    txtDoctorId = "001201"
'    lblDoctNm.Caption = "이춘희"
'    txtDeptCd.Text = "LABM"
'
'    With objPatient
'        .ptnm = lblPtNm.Caption
'        .DeptCd = "LABM"
'        .DoctNm = "이춘희"
'        .MajDoct = "001201"
'    End With
    
End Sub

Private Sub frmPtTemp_OKButtonClick(ByVal strPtInfo As String)
    lblDob.Caption = medGetP(strPtInfo, 3, COL_DIV)
    lblAge.Caption = medGetP(strPtInfo, 4, COL_DIV)
    lblSex.Caption = medGetP(strPtInfo, 2, COL_DIV)
    lblPtNm.Caption = medGetP(strPtInfo, 1, COL_DIV)
    txtDoctorId.Text = "99999"
    lblDoctNm.Caption = "건진의사"
    txtDeptCd.Text = "THC"
    
    With objPatient
        .DeptCd = "THC"
        .ptnm = lblPtNm.Caption
        .DoctNm = "건진의사"
        .MajDoct = "99999"
    End With
End Sub

'% 검체리스트에서 항목 선택을 키보드로 했을 경우...
Private Sub lstSpcList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 32:    'Enter Key 또는 Space
            Call lstSpcList_MouseDown(1, 0, 0, 0)
        Case 27:  'ESC
            lstSpcList.Visible = False
            tblOrdSheet.SetFocus
        Case Else:   '그 밖에...
            tblOrdSheet.SetFocus
            tblOrdSheet.Action = ActionActiveCell
    End Select
End Sub

Private Sub lstSpcList_LostFocus()
    If lstSpcList.Visible Then
        lstSpcList.SetFocus
        Exit Sub
    End If
    tblOrdSheet.SetFocus
End Sub

'% 검체리스트에서 항목 선택을 마우스로 했을 경우...
Private Sub lstSpcList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim tmpStr As String
    
    If Button <> 1 Then Exit Sub
    
    tmpStr = lstSpcList.List(lstSpcList.ListIndex)
    
    With tblOrdSheet
        .Col = enORDSHEET.tcSPCCD:       .Value = Trim(medShift(tmpStr, vbTab))       ' 검체코드
        
        Call medShift(tmpStr, vbTab)
        
        .Col = enORDSHEET.tcSTATFG:      .Value = Trim(medShift(tmpStr, vbTab))       ' **응급여부(해당건물)
        If .Value = "1" Then
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeStaticText
        End If
        .Col = enORDSHEET.tcSTORECD:     .Value = Trim(medShift(tmpStr, vbTab))       ' 보관구분
        .Col = enORDSHEET.tcMULTIFG:     .Value = Trim(medShift(tmpStr, vbTab))       ' 복수검체여부
        .Col = enORDSHEET.tcSPCGRP:      .Value = Trim(medShift(tmpStr, vbTab))       ' 검체군
        .Col = enORDSHEET.tcBARCNT:      .Value = Trim(medShift(tmpStr, vbTab))       ' 라벨출력장수
        .Col = enORDSHEET.tcTESTFLAG:    .Value = Trim(medShift(tmpStr, vbTab))       ' **검사가능여부(해당건물)
    '***건물정보 사용
        If ObjSysInfo.UseBuildingInfo = "1" Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd    ' ** 해당건물에서 일반검사 가능함
                .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab    ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
    '***건물정보 사용하지 않음
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
            .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd    ' ** 해당건물에서 일반검사 가능함
            .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
        End If
        .Col = enORDSHEET.tcSPCABBR:     .Value = Trim(medShift(tmpStr, vbTab))       ' 검체약어명
        .Col = enORDSHEET.tcLABDIV:      .Value = Trim(medShift(tmpStr, vbTab))       ' 접수번호 부여기준
        .Col = enORDSHEET.tcLABRANGE:    .Value = Trim(medShift(tmpStr, vbTab))       ' 미생물 접수번호 범위
        
        lstSpcList.Visible = False
        .SetFocus
        .Col = enORDSHEET.tcSTATCHK
        .Action = ActionActiveCell
    End With

End Sub

'% 처방항목 리스트에서 항목 선택을 키보드로 했을 경우...
Private Sub lstTestList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 32:    'Enter Key 또는 Space
            Call lstTestList_MouseDown(1, 0, 0, 0)
        Case 27:  'ESC
            lstTestList.Visible = False
            tblOrdSheet.SetFocus
        Case Else:   '그 밖에...
            tblOrdSheet.SetFocus
            tblOrdSheet.Action = ActionActiveCell
            SendKeys Chr(KeyAscii)
   End Select
End Sub

'% 처방항목 리스트에서 항목 선택을 마우스로 했을 경우...
Private Sub lstTestList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim tmpStr As String
    Dim tmpField1 As String
    Dim tmpField2 As String
    Dim strFlag1 As String, strFlag2 As String
    Dim objSQL As clsLISSqlStatement
    Dim RS As Recordset

    If Button <> 1 Then Exit Sub
    If lstTestList.ListIndex < 0 Then Exit Sub
    
    Set objSQL = New clsLISSqlStatement
    Set RS = New Recordset
    
    tmpStr = lstTestList.List(lstTestList.ListIndex)
    
    With tblOrdSheet
        tmpField1 = Trim(medShift(tmpStr, vbTab))
        tmpField2 = Trim(medShift(tmpStr, vbTab))
        
        If tmpStr = "1" Then
            .Col = enORDSHEET.tcTESTNM:  .Value = Trim(tmpField1)    ' 처방명
            .Col = enORDSHEET.tcTESTCD:  .Value = Trim(tmpField2)    ' 처방코드
'            Call rs.KeyChange(Trim$(tmpField2))
            RS.Open objSQL.GetItemInfo(Trim(tmpField2)), DBConn
        Else
            .Col = enORDSHEET.tcTESTNM:  .Value = Trim(tmpField2)    ' 처방명
            .Col = enORDSHEET.tcTESTCD:  .Value = Trim(tmpField1)    ' 처방코드
'            Call Rs.KeyChange(Trim$(tmpField1))
            RS.Open objSQL.GetItemInfo(Trim(tmpField1)), DBConn
        End If
        
        .Col = enORDSHEET.tcINSURFG:     .Value = RS.Fields("insurfg").Value & ""       ' 급여구분
        .Col = enORDSHEET.tcREQDTTM:     .Value = Format(dtpColDate.Value & " " & dtpColTime.Value, _
                                                  CS_DateLongFormat & " " & CS_TimeShortFormat)     ' 희망채취시간
        .Col = enORDSHEET.tcSPCCD:       .Value = RS.Fields("spccd").Value & ""         ' 검체코드
        .Col = enORDSHEET.tcWORKAREA:    .Value = RS.Fields("workarea").Value & ""      ' WorkArea
        .Col = enORDSHEET.tcSTORECD:     .Value = RS.Fields("storecd").Value & ""       ' StoreCd
        .Col = enORDSHEET.tcRNDFG:       .Value = RS.Fields("rndfg").Value & ""         ' 아침채혈여부
        .Col = enORDSHEET.tcTESTDIV:     .Value = RS.Fields("testdiv").Value & ""       ' 검사구분
        .Col = enORDSHEET.tcMULTIFG:     .Value = RS.Fields("multifg").Value & ""       ' 복수검체여부
        .Col = enORDSHEET.tcSPCGRP:      .Value = RS.Fields("spcgrp").Value & ""        ' 검체군
        .Col = enORDSHEET.tcABBRNM:      .Value = RS.Fields("abbrnm5").Value & ""       ' 약어명
        .Col = enORDSHEET.tcBARCNT:      .Value = RS.Fields("labelcnt").Value & ""      ' 라벨출력장수
        .Col = enORDSHEET.tcSPCABBR:     .Value = RS.Fields("spcnm").Value & ""         ' 검체약어명
        .Col = enORDSHEET.tcLABDIV:      .Value = RS.Fields("labdiv").Value & ""        ' 접수번호 부여기준
        .Col = enORDSHEET.tcLABRANGE:    .Value = RS.Fields("labrange").Value & ""      ' 미생물 접수번호 범위
        
        tmpStr = RS.Fields("statflags").Value & ""
        strFlag1 = medGetP(tmpStr, 1, ";")
        strFlag2 = medGetP(tmpStr, 2, ";")
        Dim strStatFg As String
        Dim strTestFg As String
        strStatFg = Mid(strFlag1, ObjSysInfo.BuildingNo, 1)
'        RS.Fields("statfg") = Mid(strFlag1, ObjSysInfo.BuildingNo, 1)
        strTestFg = Mid(strFlag2, ObjSysInfo.BuildingNo, 1)
'        RS.Fields("testfg") = Mid(strFlag2, ObjSysInfo.BuildingNo, 1)
        
        .Col = enORDSHEET.tcSTATFG:      .Value = strStatFg        ' **응급여부(해당건물)
        If .Value = "1" Then
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeStaticText
        End If
        .Col = enORDSHEET.tcTESTFLAG:    .Value = strTestFg       ' **검사가능여부(해당건물)
    '***건물정보 사용
        If ObjSysInfo.UseBuildingInfo = "1" Then
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd     ' **해당건물에서 일반검사 가능함
                .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab                ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
    '***건물정보 사용하지 않음
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
                                         .TypeCheckCenter = True
            .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd     ' **해당건물에서 일반검사 가능함
            .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
        End If
        lstTestList.Visible = False
        Call tblOrdSheet_LeaveCell(.Col, .Row, enORDSHEET.tcSPCCD, .Row, False)
    
    End With
    Set RS = Nothing
    Set objSQL = Nothing
End Sub


'Private Sub objMyList_SendCode(ByVal SelString As String)
'
'    Dim strCD As String
'    Dim strNm As String
'
'    Select Case objMyList.Tag
'        Case "PtID"
'             txtPtId.Text = medGetP(SelString, 1, ";")
'             lblPtNm.Caption = medGetP(SelString, 2, ";")
''             Set_PtntHistory
'             Call EnableButton(True)
'
'        Case "DoctId"
'             txtDoctorId.Text = Trim(medGetP(SelString, 1, ";"))
'             lblDoctNm.Caption = Trim(medGetP(SelString, 2, ";"))
'
'        Case "DeptCd"
'             txtDeptCd.Text = Trim(medGetP(SelString, 1, ";"))
'
'        Case "WardId"
'             txtWardId.Text = Trim(medGetP(SelString, 1, ";"))
''             txtHosilID.text = Trim(medGetP(SelString, 2, ";"))
'
'             '==> 나중에 병상마스터가 제공되면 사용하자.. 2001.1.17 kmk
''             With objRef
''                  If .WardInfo(Trim(txtWardID)) = True Then
''                     txtHosilID.text = .RoomId
''                     txtBedID.text = .BedId
''                  End If
''             End With
'
'    End Select
'
'End Sub

Private Sub objMyList_SelectedItem(ByVal pSelectedItem As String)
    Select Case objMyList.Tag
        Case TAG_DEPT
            txtDeptCd.Text = objMyList.SelectedItems(0)
        Case TAG_DOCT
            txtDoctorId.Text = objMyList.SelectedItems(0)
            lblDoctNm.Caption = objMyList.SelectedItems(1)
        Case TAG_WARD
            txtWardId.Text = objMyList.SelectedItems(0)
    End Select
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DELETE
            tblOrdSheet.Col = -1
            tblOrdSheet.Action = ActionDeleteRow
    End Select
End Sub

'Private Sub objSearchPt_Selected(ByVal vPtInfo As S2CON_HOS.clsPatient)
'    If vPtInfo Is Nothing Then Exit Sub
'
'    With vPtInfo
'        txtPtId.Text = .PtId
'        lblPtNm.Caption = .ptnm
'        lblSex.Caption = .SEXNM
'        lblAge.Caption = .Age
'        lblAgeDiv.Caption = .AGEDIV ' "Y"
'        lblDob.Caption = Format(.DOB, CS_DateMask)
'
'        Set objPatient = vPtInfo
'        Call EnableButton(True)
'    End With
'
'    blnClearFg = False
'    Set objSearchPt = Nothing
'End Sub

Private Sub objSearchPt_SelectedId(ByVal vPtID As String)
    If vPtID = "" Then Exit Sub
    
    With objPatient
        .GETPatient (vPtID)
        
        txtPtId.Text = .PtId
        lblPtNm.Caption = .ptnm
        lblSex.Caption = .SEXNM
        lblAge.Caption = .Age
        lblAgeDiv.Caption = .AGEDIV ' "Y"
        lblDob.Caption = Format(.DOB, CS_DateMask)

        Call EnableButton(True)
    End With
    
    Set objSearchPt = Nothing
End Sub

Private Sub optOption_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtPtId.Text = "" Then
            txtPtId.SetFocus
            Exit Sub
        End If
        Call optOption_Click(Index)
        txtDoctorId.SetFocus
    End If
End Sub

Private Sub optOption_LostFocus(Index As Integer)
    If txtPtId.Text = "" Then
        txtPtId.SetFocus
        Exit Sub
    End If
    'txtDoctorId.SetFocus
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If Col <> enORDSHEET.tcSTATCHK Then Exit Sub   '응급여부CheckBox

    With tblOrdSheet
        
        .Row = Row
        .Col = Col
        .CellType = CellTypeStaticText
        
    '***건물정보 사용
        If ObjSysInfo.UseBuildingInfo = "1" Then
        
            .Col = enORDSHEET.tcSTATCHK
            If .Value = 1 Then   '응급선택
                .Col = enORDSHEET.tcSTATFG
                .CellType = CellTypeCheckBox
                If .Value = "1" Then
                    
                    ' ** 중앙/안이검사실에서 응급검사가 발생하면 --> 응급센터로...
                    If ObjSysInfo.BuildingCd = CentralLab Or ObjSysInfo.BuildingCd = AneLab Then
                        .Col = enORDSHEET.tcBUILDCD: .Value = EmergencyLab
                        .Col = enORDSHEET.tcBUILDNM: .Value = EmergencyLabNm
                    
                    ' ** 해당건물에서 응급검사 가능함
                    Else
                        .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd
                        .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
                    End If
                    Exit Sub
                    
                Else
                    ' ** 해당건물에서 응급검사 불가능...
                    .Col = enORDSHEET.tcSTATCHK
                    .CellType = CellTypeStaticText
                    .Text = ""
                End If
            End If
    
            '** 일반검사 가능여부
            .Col = enORDSHEET.tcTESTFLAG
            
            ' ** 해당건물에서 일반검사 가능함
            If .Value = "1" Then
                .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd
                .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
                
            ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
            Else
                .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab
                .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
            End If
    '***건물정보 사용하지 않음
        Else
            .Col = enORDSHEET.tcSTATCHK: .CellType = CellTypeCheckBox
            .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd
            .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
        End If
        
    End With

End Sub

Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
    If Col = enORDSHEET.tcREQDTTM Then
        If fraDatePicker.Visible Then
            fraDatePicker.Visible = False
        Else
            fraDatePicker.Visible = True
        End If
    End If
End Sub

'% 처방명 또는 처방코드 입력
Private Sub tblOrdSheet_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim tmpIndex As Integer
    Dim tmpStr As String
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean

    If Col <> enORDSHEET.tcTESTNM Then Exit Sub

    With tblOrdSheet
        .Col = Col
        .Row = Row

        tmpIndex = medListFind(lstTestList, tblOrdSheet.Value)
        tmpStr = lstTestList.List(tmpIndex)


        ' 문자가 입력될때마다 유사어 찾기

        If tmpIndex = -1 Or UCase(tmpStr) <> UCase(.Value) Then
            Ret = .GetCellPos(Col, Row + 1, X, Y, Wdt, Hgt)
            If .Height - Y < lstTestList.Height Or Y < 0 Then
                Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                lstTestList.Top = .Top + Y - lstTestList.Height
            Else
                lstTestList.Top = .Top + Y
            End If
            If tmpIndex >= 0 Then
                medLockWindowUpdate (lstTestList.hwnd)

                lstTestList.ListIndex = tmpIndex
                medLockWindowUpdate (0&)
                If tmpIndex - lstTestList.TopIndex > 10 Then lstTestList.TopIndex = tmpIndex
            End If
            lstTestList.Visible = True
            lstTestList.ZOrder 0
        Else
            medLockWindowUpdate (lstTestList.hwnd)

            lstTestList.ListIndex = tmpIndex
            medLockWindowUpdate (0&)
            Call lstTestList_MouseDown(1, 0, 0, 0)
            lstTestList.Visible = False
        End If
    End With
End Sub

'% 처방항목 리스트가 떠 있고 아래화살표키를 눌렀을 경우 포커스 이동
Private Sub tblOrdSheet_KeyDown(KeyCode As Integer, Shift As Integer)

    With lstTestList
        If .Visible Then
            Select Case KeyCode
                Case vbKeyDown, vbKeyPageDown:
                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
                    .SetFocus
                Case vbKeyUp, vbKeyPageUp:
                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
                    .SetFocus
                Case vbKeyEscape:
                    .Visible = False
                    'tblOrdSheet.SetFocus
            End Select
        End If
    End With

End Sub

'% 처방항목 리스트가 떠 있고 엔터키를 눌렀을 경우 항목 선택
Private Sub tblOrdSheet_KeyPress(KeyAscii As Integer)
    With tblOrdSheet
        If KeyAscii = vbKeyReturn And lstTestList.Visible Then
            DoEvents
            Call lstTestList_MouseDown(1, 0, 0, 0)
        End If
    End With
End Sub

'% 검체코드/희망채취일시 필드로 커서가 옮겨지면 검체리스트/날짜변경box 팝업
Private Sub tblOrdSheet_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim tmpTestCd As Variant
    Dim tmpSpcCd As Variant
    Dim tmpDate As Variant
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean

    If NewCol = enORDSHEET.tcTESTNM And lstTestList.Visible Then
        Cancel = True
        lstTestList.SetFocus
        Exit Sub
    End If
    If Col = enORDSHEET.tcTESTNM And lstTestList.ListIndex < 0 And lstTestList.Visible Then
        Cancel = True
        lstTestList.SetFocus
        Exit Sub
    End If

    If ActiveControl.Name = lstSpcList.Name Then Exit Sub
    
    If Col = enORDSHEET.tcSPCCD Then lstSpcList.Visible = False
    If Col = enORDSHEET.tcREQDTTM Then fraDatePicker.Visible = False

    Select Case NewCol
    Case enORDSHEET.tcSPCCD:    ' 검체리스트
        If lstSpcList.Visible Then Exit Sub
        With tblOrdSheet
            .Row = NewRow: .Col = NewCol
            Ret = .GetText(enORDSHEET.tcTESTCD, NewRow, tmpTestCd)
            If tmpTestCd = "" Then Cancel = True: Exit Sub
            'Ret = .GetText(4, NewRow, tmpSpcCd)
            Ret = .GetCellPos(NewCol, NewRow + 1, X, Y, Wdt, Hgt)
            If Y > 0 Then
                lstSpcList.Top = .Top + Y
            Else
                Ret = .GetCellPos(NewCol, NewRow, X, Y, Wdt, Hgt)
                lstSpcList.Top = .Top + Y - lstSpcList.Height
            End If
            Call objOrder.SpcList(tmpTestCd, lstSpcList)
            lstSpcList.Visible = True
            lstSpcList.ZOrder 0
            lstSpcList.SetFocus
            If lstSpcList.ListCount > 0 Then lstSpcList.ListIndex = 0
            DoEvents
        End With
    Case 7:   ' 희망채혈일시 입력
        If fraDatePicker.Visible Then Exit Sub
        With tblOrdSheet
            .Row = NewRow: .Col = NewCol
            Ret = .GetText(enORDSHEET.tcTESTCD, NewRow, tmpTestCd)
            If tmpTestCd = "" Then Cancel = True: Exit Sub
            Ret = .GetText(enORDSHEET.tcREQDTTM, NewRow, tmpDate)
            Ret = .GetCellPos(NewCol, NewRow + 1, X, Y, Wdt, Hgt)
            If Y > 0 Then
                fraDatePicker.Top = .Top + Y
            Else
                Ret = .GetCellPos(NewCol, NewRow, X, Y, Wdt, Hgt)
                fraDatePicker.Top = .Top + Y - fraDatePicker.Height
            End If
            fraDatePicker.Visible = True
            If tmpDate = "" Then
                txtDatePicker.Value = GetSystemDate
            Else
                txtDatePicker.Value = tmpDate
            End If
            txtDatePicker.SetFocus
            DoEvents
        End With
    End Select

End Sub

'% 부서코드
Private Sub txtDeptCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then Call cmdHelpList_Click(2)
End Sub

Private Sub txtDeptCd_LostFocus()
    If txtDeptCd.Text <> "" Then Call txtDeptCd_KeyPress(vbKeyReturn)
End Sub

Private Sub txtDoctorId_LostFocus()
    If txtDoctorId.Text <> "" Then Call txtDoctorId_KeyPress(vbKeyReturn)
End Sub

Private Sub txtReceptNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tblOrdSheet.Row = 1
        tblOrdSheet.Col = enORDSHEET.tcTESTNM
        tblOrdSheet.Action = ActionActiveCell
        tblOrdSheet.SetFocus
    End If
End Sub

'% 병동ID
Private Sub txtWardId_GotFocus()
    With txtWardId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWardId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then Call cmdHelpList_Click(3)
End Sub

'% SetFocus : 병동ID --> 병실ID
Private Sub txtWardId_KeyPress(KeyAscii As Integer)
'    Dim objData As clsBasisData
    Dim strData As String
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If txtWardId.Text = "" Then
            txtWardId.SetFocus
            Exit Sub
        Else
'            Set objData = New clsBasisData
            strData = GetWardNm(txtWardId.Text)
'            Set objData = Nothing
            
            If strData = "" Then
'            If Not objLisComCode.WardId.Exists(txtWardId.Text) Then
                MsgBox "병동 코드를 확인하세요.."
                txtWardId.Text = ""
                Call cmdHelpList_Click(3)
                Exit Sub
            End If
        End If
        txtRoomId.SetFocus
    End If

End Sub

'% 병실 ID
Private Sub txtRoomId_GotFocus()
    With txtRoomId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% SetFocus : 병실ID --> 침상ID
Private Sub txtRoomId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtBedId.Enabled Then
        txtBedId.SetFocus
'        Call LockPtInfo(False)
    End If
End Sub

'% 침상ID
Private Sub txtBedId_GotFocus()
    With txtBedId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% SetFocus : 침상ID --> Order sheet
Private Sub txtBedId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And dtpColDate.Enabled Then
        dtpColDate.SetFocus
    End If
End Sub

'% SetFocus : 희망채취일시 변경 Box --> Order sheet
Private Sub txtDatePicker_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With tblOrdSheet
            .Col = enORDSHEET.tcREQDTTM
            .Value = Format(txtDatePicker.Value, CS_DateLongFormat & " " & CS_TimeShortFormat)
            fraDatePicker.Visible = False
            
            DoEvents
            
            .SetFocus
            .Row = .DataRowCnt + 1:  .Col = enORDSHEET.tcTESTNM
            .Action = ActionActiveCell
        End With
    End If
End Sub

Private Sub txtDatePicker_Change()

   'If KeyCode = vbKeyReturn Then
    With tblOrdSheet
        .Col = enORDSHEET.tcREQDTTM
        .Value = Format(txtDatePicker.Value, CS_DateLongFormat & " " & CS_TimeShortFormat)
    End With
   'End If

End Sub

Private Sub txtDatePicker_LostFocus()
    If fraDatePicker.Visible Then
        txtDatePicker.SetFocus
        Exit Sub
    End If
'   tblOrdSheet.SetFocus
End Sub

'% 부서코드
Private Sub txtDeptCd_GotFocus()
    With txtDeptCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% SetFocus : 부서코드 --> Ward ID / ReceptNo
Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
'    Dim objData As clsBasisData
    Dim strData As String
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If ActiveControl.Name = txtPtId.Name Then Exit Sub
    If ActiveControl.Name = optOption(0).Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub

    If KeyAscii = vbKeyReturn Then
        If txtDeptCd.Text = "" Then
            txtDeptCd.SetFocus
            Exit Sub
        Else
'            Set objData = New clsBasisData
            strData = GetDeptNm(txtDeptCd.Text)
'            Set objData = Nothing
            
            If strData = "" Then
'            If Not objLisComCode.DeptCd.Exists(txtDeptCd.Text) Then
                MsgBox "부서 코드를 확인하세요.."
                txtDeptCd.Text = ""
                Call cmdHelpList_Click(2)
                Exit Sub
            End If
        End If
        If optOption(1).Value Then
            If txtWardId.Enabled Then txtWardId.SetFocus
        Else
            If txtPtId.Text = "" Then
                txtPtId.SetFocus
                Exit Sub
            End If
            If txtDoctorId.Text = "" Then
                txtDoctorId.SetFocus
                Exit Sub
            End If
            If optOption(1).Value And (txtWardId.Text = "") Then
                txtWardId.SetFocus
                Exit Sub
            End If

'            txtReceptNo.SetFocus
'            Call LockPtInfo(False)
        End If
    End If
    
End Sub

'% 의사ID
Private Sub txtDoctorId_GotFocus()
    With txtDoctorId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDoctorId_Change()
    lblDoctNm.Caption = ""
End Sub

'% Arrow Down --> 의사ID 리스트 팝업
Private Sub txtDoctorId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then Call cmdHelpList_Click(1)
End Sub

'% SetFocus : 의사ID --> 부서코드
Private Sub txtDoctorId_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        If txtDoctorId.Text = "" Then
'            lblDoctNm.Caption = ""
'            txtDoctorId.SetFocus
'            Exit Sub
'        Else
'            lblDoctNm.Caption = GetEmpName(txtDoctorId.Text)
'
'            If lblDoctNm.Caption = "" Then
'                MsgBox "처방의 코드를 확인하세요.."
'                txtDoctorId.Text = ""
'                Call cmdHelpList_Click(1)
'                Exit Sub
'            End If
'        End If
'        txtDeptCd.SetFocus
'    End If

'    Dim objData As clsBasisData
    Dim strData As String
    
    If KeyAscii = vbKeyReturn Then
        If txtDoctorId.Text = "" Then
            lblDoctNm.Caption = ""
            txtDoctorId.SetFocus
            Exit Sub
        Else
'            Set objData = New clsBasisData
            strData = GetDoctNm(txtDoctorId.Text)
'            Set objData = Nothing
            
            If strData = "" Then
'            If Not objLisComCode.Doct.Exists(txtDoctorId.Text) Then
                MsgBox "처방의 코드를 확인하세요.."
                txtDoctorId.Text = ""
                Call cmdHelpList_Click(1)
                Exit Sub
            Else
'                Call objLisComCode.Doct.KeyChange(txtDoctorId.Text)
                lblDoctNm.Caption = strData 'objLisComCode.Doct.Fields("doctnm")
            End If
        End If
        txtDeptCd.SetFocus
    End If
End Sub

'% 환자ID
Private Sub txtPtId_GotFocus()
    With txtPtId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPtId_Change()
    If Not blnClearFg Then Call ClearRtn(False)
End Sub

'% 환자ID를 Key로 데이타 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then optOption(0).SetFocus

End Sub

'% 환자ID를 Key로 데이타 검색
Private Sub txtPtId_LostFocus()

    Dim blnRst As Boolean

    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub

    If txtPtId.Text = "" Then Exit Sub
    If Not blnClearFg Then Exit Sub
    

    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    

    txtPtId.Text = UCase(txtPtId.Text)
    blnRst = objPatient.GETPatient(UCase((txtPtId.Text)))
    
    If Not blnRst Then
        MsgBox "등록되지 않은 환자입니다. ID를 확인하세요 ! ", vbExclamation + vbOKOnly, "처방등록"
        txtPtId.Text = ""
        DoEvents
        txtPtId.SetFocus
        Exit Sub
    End If
    Call DisplayPtInfo
    Call EnableButton(True)
    
    '감염관리 표시
    Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)
    
    blnClearFg = False

    txtDoctorId = "001201"
    lblDoctNm.Caption = "이춘희"
    txtDeptCd.Text = "LABM"

    With objPatient
        .ptnm = lblPtNm.Caption
        .DeptCd = "LABM"
        .DoctNm = "이춘희"
        .MajDoct = "001201"
    End With
    
End Sub

'% 환자정보 클래스 objPatient 로부터 기본정보를 Screen에 Display한다.
Sub DisplayPtInfo()

    With objPatient
        lblPtNm.Caption = .ptnm
        lblAgeDiv.Caption = .AGEDIV
        

        txtDeptCd.Text = .DeptCd
        txtDoctorId.Text = .MajDoct
        lblDoctNm.Caption = .DoctNm
        txtWardId.Text = .WardId
        txtRoomId.Text = .RoomId
        txtBedId.Text = .BedID

'        If .INADMISSION Then
'            optOption(1).Value = True
'        Else
'            optOption(0).Value = True
'        End If
        
        optOption(1).Value = True

        ''''''''''''''''''''''
        '건진환자...PopUp
        '''''''''''''''''''''
        If .TmpDiv = "1" Then
            '건진환자...PopUp
            Set frmPtTemp = frmTmpPt
            frmPtTemp.PtId = txtPtId
            frmPtTemp.ptnm = lblPtNm.Caption
            frmPtTemp.RTop = framPtInfo.Top + lblPtNm.Top + 1600
            frmPtTemp.Rleft = framPtInfo.Left + lblPtNm.Left
            frmPtTemp.ssn = .ssn
            frmPtTemp.Show 1
        Else
            lblSex.Caption = .SEXNM
            lblAge.Caption = .Age
            lblDob.Caption = Format(.DOB, CS_DateMask)
        End If
'
    End With

End Sub


' 다빈도 처방 리스트 클릭
Private Sub tblTestList_Click(ByVal Col As Long, ByVal Row As Long)

    Dim tmpValue As Variant

    If (Col \ DabindoCols) > 3 Then Exit Sub

    With tblTestList
        .Row = Row
        .Col = Col + ((Col Mod DabindoCols) Mod 2) - 1

        Call .GetText(.Col + 1, Row, tmpValue)  ' 처방명이 없는 경우...
        If tmpValue = "" Then Exit Sub

        If .Value = 1 Then
            .Value = 0
        Else
            .Value = 1
        End If
    End With

End Sub

Private Sub Item_Move(ByVal Col As Long, ByVal Row As Long)
    Call objOrder.DabindoListClick(Col, Row, tblTestList, tblOrdSheet, _
                                 Format(dtpColDate.Value & " " & dtpColTime.Value, "YYYY/MM/DD HH:MM")) ', _
                                 objLisComCode.LisItem)
    tblOrdSheet.Col = enORDSHEET.tcTESTNM
    tblOrdSheet.Row = tblOrdSheet.DataRowCnt + 1
    tblOrdSheet.SetFocus
    tblOrdSheet.Action = ActionActiveCell
End Sub


'% 일괄채혈 일시 선택
'Private Sub cboColTime_Click()
''   dtpColTime.Value = cboColTime.Text
'End Sub

'% 응급여부 Check 해서 Order sheet에 표시...
Private Sub chkStat_Click()

    Dim i As Long

    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            
        '***건물정보사용
            If ObjSysInfo.UseBuildingInfo = "1" Then
                .Col = enORDSHEET.tcSTATFG
                If chkStat.Value = 1 And .Value = "1" Then
                    .Col = enORDSHEET.tcSTATCHK: .Value = 1
                    If ObjSysInfo.BuildingCd = CentralLab Or ObjSysInfo.BuildingCd = AneLab Then
                        ' ** 중앙/안이검사실에서 응급검사가 발생하면..
                        .Col = enORDSHEET.tcBUILDCD: .Value = EmergencyLab   ' --> 응급센터로...
                        .Col = enORDSHEET.tcBUILDNM: .Value = EmergencyLabNm
                    Else
                        ' ** 해당건물에서 응급검사 가능함
                        .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd
                        .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
                    End If
                Else
                    .Col = enORDSHEET.tcSTATCHK: .Value = ""
                    .Col = enORDSHEET.tcTESTFLAG
                    If .Value = "1" Then
                        '** 해당건물에서 일반검사 가능
                        .Col = enORDSHEET.tcBUILDCD: .Value = ObjSysInfo.BuildingCd
                        .Col = enORDSHEET.tcBUILDNM: .Value = ObjSysInfo.BuildingNm
                    Else
                        '** 해당건물에서 일반검사 불가능 --> 중앙검사실로...
                        .Col = enORDSHEET.tcBUILDCD: .Value = CentralLab
                        .Col = enORDSHEET.tcBUILDNM: .Value = CentralLabNm
                    End If
                End If
            
        '***건물정보사용하지 않음
            Else
                .Col = enORDSHEET.tcSTATCHK: .Value = chkStat.Value
                .Col = enORDSHEET.tcBUILDCD: .Value = APS_BUILDCD   '10
                .Col = enORDSHEET.tcBUILDNM: .Value = APS_BUILDNM   '본원
            End If
            
            .Col = enORDSHEET.tcSTATCHK
            .CellType = CellTypeCheckBox
        Next
    End With

End Sub


'% Order sheet에서 Row 삭제
'Private Sub mnuDelete_Click()
'    tblOrdSheet.Col = -1
'    tblOrdSheet.Action = ActionDeleteRow
'End Sub

'% 처방유형에 따라 희망채혈시간 컨트롤 변경
Private Sub optOption_Click(Index As Integer)

    objOrder.OrdDiv = lis_orddiv   'Choose(Index + 1, "S", "W", "L")
    If Index = 2 Then
'        cboColTime.Visible = True
'        dtpColTime.Visible = False
    Else
'        cboColTime.Visible = False
'        dtpColTime.Visible = True
        dtpColTime.Value = Format(Now, CS_TimeShortFormat)
        If Index = 0 Then
            txtWardId.Enabled = False
            txtRoomId.Enabled = False
            cmdHelpList(3).Enabled = False
            txtBedId.Enabled = False
        ElseIf Index = 1 Then
            txtWardId.Enabled = False
            txtRoomId.Enabled = False
            cmdHelpList(3).Enabled = False
            txtBedId.Enabled = False
            
            
'            txtWardId.Enabled = True
'            txtRoomId.Enabled = True
'            cmdHelpList(3).Enabled = True
'            txtBedId.Enabled = True
        End If
    End If

End Sub

'% 오른쪽 버튼 클릭 --> Delete Row
Private Sub tblOrdSheet_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    Dim lngOldColor As Long
    
    tblOrdSheet.OperationMode = OperationModeRead
    tblOrdSheet.Row = Row
    tblOrdSheet.Col = -1
    lngOldColor = tblOrdSheet.BackColor
    tblOrdSheet.BackColor = DCM_LightGray

'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Delete"
'    frmControls.mnuSub1.Visible = False
'    frmControls.mnuSub2.Visible = False
'    PopupMenu mnuPopup

    Set objPop = Nothing
    Set objPop = New clsPopupMenu

    With objPop
        .AddMenu MENU_DELETE, "Delete"
        
        .PopupMenus Me.hwnd
    End With
    
    Set objPop = Nothing

    tblOrdSheet.Row = Row
    tblOrdSheet.Col = -1
    tblOrdSheet.BackColor = lngOldColor
    tblOrdSheet.OperationMode = OperationModeNormal
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
End Sub

'% 종료
Private Sub cmdExit_Click()
    Unload Me
    Set objSqlStmt = Nothing
    Set objPatient = Nothing
    Set objOrder = Nothing
    Set objCollect = Nothing
    Set objAccess = Nothing
    Set frm101Order = Nothing
End Sub

Private Function ValidationCheck() As Boolean

    Dim CheckOrder As Long

    ValidationCheck = True
    
    If tblOrdSheet.DataRowCnt = 0 Then GoTo Err_Trap
    
    If txtDoctorId.Text = "" Then
        MsgBox "처방의ID를 반드시 입력하세요..", vbExclamation, "입력오류"
        txtDoctorId.SetFocus
        GoTo Err_Trap
    End If
    If txtDeptCd.Text = "" Then
        MsgBox "진료과를 반드시 입력하세요..", vbExclamation, "입력오류"
        txtDeptCd.SetFocus
        GoTo Err_Trap
    End If
    
    ' 외부정도관리용 검체명입력확인
    If (txtSpecNm.Text = "") And txtSpecNm.Enabled Then
        MsgBox "검체명을 반드시 입력하세요..", vbExclamation, "입력오류"
        txtSpecNm.SetFocus
        GoTo Err_Trap
    End If
    
    If (txtWardId.Text = "") And txtWardId.Enabled Then
        MsgBox "병동ID를 반드시 입력하세요..", vbExclamation, "입력오류"
        txtWardId.SetFocus
        GoTo Err_Trap
    End If

    CheckOrder = objOrder.CheckSameOrder(tblOrdSheet)
    If CheckOrder > 0 Then
        tblOrdSheet.Row = CheckOrder
        tblOrdSheet.Col = enORDSHEET.tcTESTNM
        tblOrdSheet.Action = ActionActiveCell
        MsgBox "중복처방입니다. 희망채혈시간을 조정하세요..", vbExclamation, "입력오류"
        tblOrdSheet.SetFocus
        GoTo Err_Trap
    End If

    Exit Function
    
Err_Trap:
    ValidationCheck = False

End Function

'% 처방 Save 버튼
Private Sub cmdOrder_Click(Index As Integer)

    Dim i As Integer
    Dim Success As Boolean
    
    If tblOrdSheet.DataRowCnt = 0 Then Exit Sub
    If Not ValidationCheck Then Exit Sub
    
'    Dim objProgress As New S2Progress.clsProgress
    Dim objProgress As New clsProgress

    MouseRunning  '13
    
    objProgress.Container = MainFrm.stsBar
    objProgress.Message = "환자의 처방내역을 저장하고 있습니다..."
    objProgress.Max = tblOrdSheet.DataRowCnt * (Index + 1)
''    objProgress.CaptionOn = False
''    objProgress.Visible = True
''    objProgress.MSG = "환자의 처방내역을 저장하고 있습니다."
''    objProgress.Min = 0
'    objProgress.Max = tblOrdSheet.DataRowCnt * (Index + 1)
'    objProgress.Value = 0

    DoEvents

    '데이타베이스의 날짜/시간으로 System Date/Time을 셋팅...
'    Date = Format(GetSystemdate, CS_DateMask)
'    Time = Format(GetSystemdate, CS_TimeLongMask)

    With tblOrdSheet
        .SortBy = SortByRow
        .SortKey(1) = enORDSHEET.tcBUILDCD      'DeliveryLocation
        '.SortKey(2) = 12  '검사구분  --> 제외 1999.10.08 by 김미경
        .SortKey(2) = enORDSHEET.tcREQDTTM      '희망채취시간
        .SortKey(3) = enORDSHEET.tcWORKAREA     'WorkArea
        .SortKey(4) = enORDSHEET.tcSPCCD        '검체코드
        .SortKey(5) = enORDSHEET.tcSTORECD      '보관구분
        .SortKey(6) = enORDSHEET.tcSTATCHK      '응급여부
        .SortKey(7) = enORDSHEET.tcTESTCD       '검사코드
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKeyOrder(3) = SortKeyOrderAscending
        .SortKeyOrder(4) = SortKeyOrderAscending
        .SortKeyOrder(5) = SortKeyOrderAscending
        .SortKeyOrder(6) = SortKeyOrderAscending
        .SortKeyOrder(7) = SortKeyOrderAscending
        .Col = 1:   .COL2 = .MaxCols
        .Row = 0:  .Row2 = .MaxRows
        .Action = ActionSort
    End With

    'Button 1  : 처방

    Success = SaveOrder(objProgress)                    '처방내역 저장

    If Success Then
        If Index = 0 Then
            objProgress.Value = objProgress.Max
            Set objProgress = Nothing
            MsgBox "처방내역이 저장되었습니다.", vbInformation, "처방등록"
            GoTo Exit_Routine
        End If
    Else
        Set objProgress = Nothing
        MsgBox "처방내역 저장중 오류가 발생했습니다. ", vbCritical, "오류발생"
        GoTo Exit_Routine
    End If


    'Button 2  : 채혈수행

    objProgress.Message = "채혈 Procedure를 수행하고 있습니다."

    Call ReadyToCollect                             '채혈준비
    Success = objCollect.DoCollection(objProgress)    '채혈수행

    If Success Then
        If Index = 1 Then
            objProgress.Value = objProgress.Max
            Set objProgress = Nothing
            MsgBox "정상적으로 채혈처리 되었습니다." & vbCRLF & _
                   "Barcode " & CStr(objCollect.ColCount) & " 장 발행..." & vbCRLF, vbInformation, "채혈"
            'Call Print_BarcodeLabel   '바코드출력
            GoTo Exit_Routine
        End If
    Else
        Set objProgress = Nothing
        MsgBox "채혈내역 생성중 오류가 발생했습니다. ", vbCritical, "오류발생"
        GoTo Exit_Routine
    End If


    'Button 3:  접수수행
    objProgress.Message = "접수 Procedure를 수행하고 있습니다."

    With objCollect
        If .CollectDone Then
            Dim pWorkArea As String
            Dim pAccDt As String
            Dim pAccSeq As Long

'            Call objAccess.SetDatabase(DbConn)
            For i = 1 To .ColCount
                objProgress.Message = "접수 Procedure를 수행하고 있습니다. (" & CStr(i) & "/" & CStr(.ColCount) & ")"
                Call .GetLabNumbers(i, pWorkArea, pAccDt, pAccSeq)
                Success = objAccess.DoAccession_New(pWorkArea, pAccDt, pAccSeq, ObjMyUser.EmpId)
                If Not Success Then Exit For
                If objProgress.Value = objProgress.Max Then objProgress.Max = objProgress.Max + 1
                objProgress.Value = objProgress.Value + 1
                DoEvents
            Next
        End If
    End With

    If Success Then
        objProgress.Value = objProgress.Max
        Set objProgress = Nothing
        Dim strMsg As String
        
        strMsg = "정상적으로 접수되었습니다."
        If P_UseBarcodeSystem Then strMsg = strMsg & vbCRLF & "Barcode " & CStr(objCollect.ColCount) & " 장 발행..." & vbCRLF
        
        MsgBox strMsg, vbInformation, "접수완료"
        
    Else
        Set objProgress = Nothing
        MsgBox "접수처리중 오류가 발생했습니다.", vbCritical, "오류발생"
    End If

Exit_Routine:
    MouseDefault
    Call cmdClear_Click

End Sub

'% 클래스 objOrder의 데이타 속성을 채우고 SaveData 메쏘드를 Call하여
'% 데이타베이스로 저장한다.
'% 채혈/접수 클래스를 통하여 채혈/접수 내역을 생성하고 Barcode를 출력한다.
Private Function SaveOrder(Optional ByRef objPrgBar As clsProgress = Nothing) As Boolean

    Dim i As Integer
    Dim StartOrdNo As Integer
    Dim OrderSuccess As Boolean
    
    With objOrder

        .PtId = txtPtId.Text
        .ordDt = Format(dtpColDate.Value, CS_DateDbFormat)
        .OrdTm = Format(dtpColDate.Value, CS_TimeDbFormat)
        If optOption(1).Value Then
            .Bussdiv = enBussDiv.BussDiv_OutPatient   '외래
            .BedinDt = ""
            .DeptCd = txtDeptCd.Text
            .MajDoct = txtDoctorId.Text
            .WardId = txtSpecNm.Text
            .HosilId = ""
            .RoomId = ""
        
'            .Bussdiv = enBussDiv.BussDiv_InPatient    '병동
'            '## 병동ID(HIS002)는 임시로 현재일 Setting ...
'            .BedInDt = objPatient.BedInDt
'            .DeptCd = txtDeptCd.Text
'            .MajDoct = objPatient.MajDoct
'            .WardId = txtWardId.Text
'            .HosilId = txtRoomId.Text
'            .ROOMID = txtBedId.Text
        ElseIf optOption(0).Value Then
            .Bussdiv = enBussDiv.BussDiv_OutPatient   '외래
            .BedinDt = ""
            .DeptCd = txtDeptCd.Text
            .MajDoct = txtDoctorId.Text
            .WardId = txtSpecNm.Text
            .HosilId = ""
            .RoomId = ""
        ElseIf optOption(3).Value = True Then
            .Bussdiv = enBussDiv.BussDiv_OutPatient   '외래
            .BedinDt = ""
            .DeptCd = txtDeptCd.Text
            .MajDoct = txtDoctorId.Text
            .WardId = txtSpecNm.Text
            .HosilId = ""
            .RoomId = ""
        End If
        .OrdDoct = txtDoctorId.Text
        .ReceptNo = txtReceptNo.Text
        .EntId = ObjMyUser.EmpId
        .EntDt = Format(GetSystemDate, CS_DateDbFormat)
        .EntTm = Format(GetSystemDate, CS_TimeDbFormat)
        .DoneFg = enStsCd.StsCd_LIS_Order
        .RepeatFg = ""
        .OrgAccNo = ""
        .SpOrdDiv = ""
        .OrdDiv = lis_orddiv
        
        Call .MoveData(tblOrdSheet)                     '클래스로 데이타 Move
        OrderSuccess = .SaveData(StartOrdNo, objPrgBar) 'Database로 저장
        
    End With

    If Not OrderSuccess Then
        SaveOrder = False
        Exit Function
    End If

    '처방번호 Display
    With tblOrdSheet
        .Col = 1
        For i = 1 To .DataRowCnt
            .Row = i
'            .Value = Val(.Value) + StartOrdNo
            .Value = i + StartOrdNo
        Next
    End With
    SaveOrder = True

End Function

'% 발생한 처방데이타를 기준으로 채혈접수내역을 생성하기 위해
'% 모든 데이타를 Array로 Assign한다.
Sub ReadyToCollect()

    Dim i As Integer
    Dim tmpData() As String
    Dim strDOB    As String
    
    With objCollect
'        Call .SetDatabase(DbConn)

        ReDim tmpData(0 To 16)
        
        tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)  '검체년도
        tmpData(1) = objPatient.PtId                            '환자ID
        tmpData(2) = objPatient.ptnm
        tmpData(3) = objPatient.Sex                             '성별
        If IsDate(lblDob.Caption) Then                          '환자일령
            tmpData(4) = DateDiff("y", lblDob.Caption, GetSystemDate)
        Else
            tmpData(4) = DateDiff("y", Mid(lblDob.Caption, 1, 4) & "-01-01", GetSystemDate)
        End If
        
        tmpData(5) = objPatient.BedinDt                         '입원일
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat) '입력일
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat) '입력시간
        tmpData(8) = ObjMyUser.EmpId                            '입력자
        tmpData(9) = ""                                         '원접수번호
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat) '채혈일
        .ColTm = Format(GetSystemDate, "hhmmss")
        tmpData(11) = ObjMyUser.EmpId                           '채혈자
        tmpData(12) = txtSpecNm.Text 'txtWardId.Text                            '병동ID
        tmpData(13) = txtRoomId.Text                            '병실ID
        tmpData(14) = ""                                        '병실ID
        tmpData(15) = txtBedId.Text                             '침상ID
        tmpData(16) = ObjSysInfo.BuildingCd                                '** 채혈이 수행되는 건물코드
        
        
        Call .SetColData(tmpData)
    End With
    
    With tblOrdSheet
        ReDim tmpData(0 To .MaxCols)
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = enORDSHEET.tcBUILDCD:    tmpData(0) = .Value     'Delivery Location
            .Col = enORDSHEET.tcWORKAREA:   tmpData(1) = .Value     'WorkArea
            .Col = enORDSHEET.tcSPCCD:      tmpData(2) = .Value     'SpcCd
            .Col = enORDSHEET.tcSTORECD:    tmpData(3) = .Value     'StoreCd
            .Col = enORDSHEET.tcSTATCHK:    tmpData(4) = .Value     'StatFg
            .Col = enORDSHEET.tcREQDTTM:    tmpData(5) = .Value     'ReqColDate
            
            .Col = enORDSHEET.tcTESTDIV:    tmpData(6) = .Value     'TestDiv
            .Col = enORDSHEET.tcMULTIFG:    tmpData(7) = .Value     'MultiFg
            .Col = enORDSHEET.tcSPCGRP:     tmpData(8) = .Value     'SpcGrp
            tmpData(9) = Format(dtpColDate.Value, CS_DateDbFormat)  '처방일을 희망채혈일로.. 2000/04/03 by 정미경
            .Col = enORDSHEET.tcORDNO:      tmpData(10) = .Value    'OrdNo
            .Col = enORDSHEET.tcORDSEQ:     tmpData(11) = .Value    'OrdSeq
            .Col = enORDSHEET.tcTESTCD:     tmpData(12) = .Value    'OrdCd
            tmpData(13) = txtDeptCd.Text                            '진료과
            tmpData(14) = txtDoctorId.Text                          '처방의
            tmpData(15) = objPatient.MajDoct                        '주치의
            .Col = enORDSHEET.tcABBRNM:     tmpData(16) = .Value    '약어명
            .Col = enORDSHEET.tcBARCNT:     tmpData(17) = .Value    '라벨출력장수
            .Col = enORDSHEET.tcLABDIV:     tmpData(18) = .Value    '접수번호부여기준
            .Col = enORDSHEET.tcSPCABBR:    tmpData(19) = .Value    '검체약어명
            .Col = enORDSHEET.tcLABRANGE:   tmpData(20) = .Value    '미생물접수번호범위
            
            Call objCollect.AddLabCollect(tmpData)
        Next
    End With

End Sub

'% 채혈수행후 바코드라벨 출력...
Sub Print_BarcodeLabel(Optional ByVal AccFg As Boolean = False)

    Dim LabelCount As Integer
    Dim BarcodeBuffer As String
    Dim i As Integer

    LabelCount = objCollect.ColCount

'    Call Label_PortOpen      '바코드프린터 포트 오픈

    For i = 1 To LabelCount
        Call objCollect.GetBarcodeLabel(i, AccFg)
    Next
    objCollect.BarFormFeed   '마지막에 폼피드...

'    Call Label_PortClose     '바코드프린터 포트 클로즈

End Sub



'% Clear Routine : 환자정보 클래스, 처방 클래스 및 각 콘트롤들 초기화
Sub ClearRtn(Optional ByVal blnAll As Boolean = True)

'    Call LockPtInfo(True)
    'Call medMsg("")
    If blnAll Then txtPtId.Text = ""
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDob.Caption = ""
    optOption(3).Value = True
    txtDoctorId.Text = ""
    lblDoctNm.Caption = ""
    txtDeptCd.Text = ""
    txtWardId.Text = ""
    txtRoomId.Text = ""
    txtBedId.Text = ""
    chkStat.Value = 0
    txtReceptNo.Text = ""
    'optPriority(0).Value = True
    dtpColDate.Value = Format(GetSystemDate, CS_DateLongFormat)
    dtpColTime.Value = Format(GetSystemDate, CS_TimeShortFormat)
    Call EnableButton(False)
    With tblTestList
        .Row = -1
        .Col = 1: .COL2 = 1
        .BlockMode = True
        .Value = 0
        .BlockMode = False
        .Col = 15: .COL2 = 15
        .BlockMode = True
        .Value = 0
        .BlockMode = False
    End With
    With tblOrdSheet
        .MaxRows = 0
        .MaxRows = 50
        .Row = -1
        .Col = enORDSHEET.tcORDNO: .COL2 = enORDSHEET.tcORDNO
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
    End With

    Set objPatient = Nothing
    Set objPatient = New clsPatient
    
    Set objOrder = Nothing
    Set objOrder = New clsLISOrder
    objOrder.BuildingCd = ObjSysInfo.BuildingCd
    objOrder.BuildingNm = ObjSysInfo.BuildingNm
    objOrder.BuildingNo = ObjSysInfo.BuildingNo
    
    Set objCollect = Nothing
    Set objCollect = New clsLISCollectioin
    
    Set objAccess = Nothing
    Set objAccess = New clsLISAccession

    blnClearFg = True

End Sub


Private Sub EnableButton(ByVal ValFg As Boolean)

    Dim i As Integer

'    For i = 1 To cmdOrder.Count
'        cmdOrder(i - 1).Enabled = ValFg
'    Next
    
    For i = 1 To 1
        cmdOrder(i - 1).Enabled = False
    Next
        
    tblTestList.Enabled = ValFg
    tblOrdSheet.Enabled = ValFg

End Sub

Private Sub txtWardId_LostFocus()
    If txtWardId.Text <> "" Then Call txtWardId_KeyPress(vbKeyReturn)
End Sub
