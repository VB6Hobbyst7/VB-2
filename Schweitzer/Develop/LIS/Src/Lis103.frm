VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm103AddOrder 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Edit Order"
   ClientHeight    =   9285
   ClientLeft      =   435
   ClientTop       =   1065
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis103.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   5910
      TabIndex        =   6
      Top             =   45
      Width           =   8550
      _ExtentX        =   15081
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
      Caption         =   "환자기본정보"
      LeftGab         =   100
   End
   Begin VB.Frame fraPtInfo 
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
      Height          =   1470
      Left            =   5910
      TabIndex        =   14
      Tag             =   "104"
      Top             =   270
      Width           =   8565
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   330
         Left            =   3900
         TabIndex        =   30
         Top             =   600
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDob 
         Height          =   330
         Left            =   1050
         TabIndex        =   31
         Top             =   600
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   3900
         TabIndex        =   32
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtId 
         Height          =   330
         Left            =   1050
         TabIndex        =   33
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   330
         Left            =   6660
         TabIndex        =   34
         Top             =   600
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   330
         Left            =   1050
         TabIndex        =   35
         Top             =   1035
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BackColor       =   15857140
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   90
         TabIndex        =   36
         Top             =   180
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
         Index           =   7
         Left            =   2925
         TabIndex        =   37
         Top             =   180
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
         Caption         =   "성 명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   5715
         TabIndex        =   38
         Top             =   195
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   90
         TabIndex        =   39
         Top             =   600
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
         Index           =   10
         Left            =   2925
         TabIndex        =   40
         Top             =   600
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
         Index           =   11
         Left            =   5715
         TabIndex        =   41
         Top             =   600
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
         Caption         =   "병실"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   90
         TabIndex        =   42
         Top             =   1035
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
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   6660
         TabIndex        =   46
         Top             =   195
         Width           =   1725
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6705
         TabIndex        =   45
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   7905
         TabIndex        =   44
         Top             =   270
         Width           =   345
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7590
         TabIndex        =   43
         Top             =   240
         Width           =   60
      End
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   300
      Left            =   75
      TabIndex        =   7
      Top             =   45
      Width           =   5805
      _ExtentX        =   10239
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
      Caption         =   "접수번호"
      LeftGab         =   100
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
      Height          =   1470
      Left            =   75
      TabIndex        =   8
      Top             =   270
      Width           =   5820
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1425
         ScaleHeight     =   300
         ScaleWidth      =   2355
         TabIndex        =   9
         Top             =   570
         Width           =   2415
         Begin VB.TextBox txtWorkArea 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   45
            MaxLength       =   2
            TabIndex        =   1
            Top             =   30
            Width           =   465
         End
         Begin VB.TextBox txtAccDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            MaxLength       =   6
            TabIndex        =   2
            Top             =   30
            Width           =   810
         End
         Begin VB.TextBox txtAccSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1725
            TabIndex        =   3
            Top             =   30
            Width           =   600
         End
         Begin VB.Label Label2 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   525
            TabIndex        =   11
            Top             =   -30
            Width           =   210
         End
         Begin VB.Label Label17 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1530
            TabIndex        =   10
            Top             =   -45
            Width           =   210
         End
      End
      Begin VB.Label lblAccNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   375
         TabIndex        =   13
         Tag             =   "151"
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "상태"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4470
         TabIndex        =   12
         Top             =   660
         Width           =   450
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00DF6A3E&
         FillStyle       =   0  '단색
         Height          =   360
         Left            =   4080
         Shape           =   4  '둥근 사각형
         Top             =   570
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1035
      Left            =   75
      TabIndex        =   15
      Top             =   1650
      Width           =   14400
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   330
         Left            =   180
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   555
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblColDtTm 
         Height          =   330
         Left            =   2070
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   570
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   4710
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   555
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblRcvDtTm 
         Height          =   330
         Left            =   7785
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   555
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblRcvNm 
         Height          =   330
         Left            =   10500
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   555
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   210
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
         Caption         =   "검체"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   26
         Top             =   210
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
         Caption         =   "채혈일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   4710
         TabIndex        =   27
         Top             =   210
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
         Caption         =   "채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   7785
         TabIndex        =   28
         Top             =   210
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
         Caption         =   "접수일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   10500
         TabIndex        =   29
         Top             =   210
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
         Caption         =   "접수자"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료 (&X)"
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
      TabIndex        =   5
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
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
      TabIndex        =   4
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장 (&S)"
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
      TabIndex        =   0
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   75
      TabIndex        =   22
      Top             =   2685
      Width           =   7380
      _ExtentX        =   13018
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
      Caption         =   "접수번호별 처방내역"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   5325
      Left            =   75
      TabIndex        =   21
      Tag             =   "10114"
      Top             =   3000
      Width           =   7380
      _Version        =   196608
      _ExtentX        =   13017
      _ExtentY        =   9393
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
      MaxCols         =   6
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis103.frx":000C
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   300
      Left            =   7500
      TabIndex        =   23
      Top             =   2685
      Width           =   6960
      _ExtentX        =   12277
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
      Caption         =   "추가처방조회"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblAddOrder 
      Height          =   5325
      Left            =   7500
      TabIndex        =   24
      Tag             =   "10114"
      Top             =   3000
      Width           =   6960
      _Version        =   196608
      _ExtentX        =   12277
      _ExtentY        =   9393
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
      MaxCols         =   12
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis103.frx":2712
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
End
Attribute VB_Name = "frm103AddOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSQL      As New clsLISAccCancel
Private objAccData  As New clsDictionary
Private ClearFg     As Boolean
Private tmpAccDt    As String

Private Type LabInfo
    sRcvDt As String
    sRcvTm As String
    sRcvId As String
    sStoreCd As String
End Type

Private Enum TblCol
    tcChk = 1
    tcORDDT
    tcORDNO
    tcOrdCd
    tcTESTNM
    tcReqdt
    tcSTAT
    tcORDSEQ
End Enum
Private objType As LabInfo

Private Sub cmdClear_Click()
    Call ClearRtn
    DoEvents
    txtWorkArea.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If txtWorkArea.Enabled Then
        txtWorkArea.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ClearFg = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSQL = Nothing
    Set objAccData = Nothing
End Sub

Private Sub txtWorkArea_Change()
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then
        If txtAccDt.Enabled Then txtAccDt.SetFocus
    End If
End Sub

Private Sub txtWorkArea_GotFocus()
    
    With txtWorkArea
       .SelStart = 0
       .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtWorkArea_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If txtWorkArea = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccDt.SetFocus
 End Sub

Private Sub txtAccDt_Change()
    
    If Mid(txtAccDt.Text, 1, 1) = "9" Then
        tmpAccDt = "19" & txtAccDt.Text
    ElseIf Mid(txtAccDt.Text, 1, 1) = "0" Then
        tmpAccDt = "20" & txtAccDt.Text
    Else
        tmpAccDt = ""
    End If
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then
       If txtAccSeq.Enabled Then txtAccSeq.SetFocus
    End If

End Sub

Private Sub txtAccDt_GotFocus()
    With txtAccDt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    If txtAccDt.Text = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccSeq.SetFocus
End Sub

Private Sub txtAccSeq_GotFocus()
    
    With txtAccSeq
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'% 접수번호를 입력했을 경우
Private Sub txtAccSeq_KeyPress(KeyAscii As Integer)
    Dim objPatient  As clsPatient
    Dim strStsNm    As String
    Dim strMultiFg  As String
    Dim tmpStatus   As String
    Dim AccFg       As Boolean
    
    Dim strTmp      As String
    
'    Dim objData As clsBasisData
    Dim strData As String
    
    Call medClearTable(tblOrdSheet)
    
    If txtAccSeq.Text = "" Then Exit Sub
    
'    Set objData = New clsBasisData
    
    If KeyAscii = vbKeyReturn Then
        
        objAccData.FieldInialize "workarea,accdt,accseq", "stscd,ptid,orddoct,deptcd,wardid,roomid,bedid,hosilid," & _
                                 "spccd,coldt,coltm,colid,rcvdt,rcvtm,rcvid,multifg,spcnm"
      
        tmpStatus = objSQL.CheckStatus(txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, objAccData)
        
        If tmpStatus = "0" Then
            MsgBox "입력하신 접수번호는 정상적인 데이타가 아닙니다", vbOKOnly + vbExclamation, "Info"
            Call ClearRtn
            txtWorkArea.SetFocus
            Exit Sub
        ElseIf Val(objAccData.Fields("stscd")) = enStsCd.StsCd_LIS_Cancel Then
            MsgBox "입력하신 접수번호는  취소되었습니다.", vbOKOnly + vbCritical, "Info"
            Call ClearRtn
            txtWorkArea.SetFocus
            Exit Sub
        ElseIf Val(objAccData.Fields("stscd")) > enStsCd.StsCd_LIS_Accession Then
            
            MsgBox "추가처방접수는 접수상태 까지의 검체에 대해서만 할수 있습니다.", vbInformation + vbOKOnly, "Info"
            Call ClearRtn
            txtWorkArea.SetFocus
            Exit Sub
        End If
        
        lblStatus.Caption = tmpStatus
        lblStatus.Tag = objAccData.Fields("stscd")
        '처방의
        lblDoctNm.Caption = GetEmpNm(objAccData.Fields("ORDDOCT")) 'GetEmpName(objAccData.Fields("ORDDOCT"))
        
        '진료과
'        objLisComCode.DeptCd.KeyChange objAccData.Fields("deptcd")
        lblDeptNm.Caption = GetDeptNm(objAccData.Fields("deptcd")) 'objLisComCode.DeptCd.Fields("deptnm")
        
        '환자정보
        Set objPatient = New clsPatient
        lblPtid.Caption = objAccData.Fields("PtId")
        With objPatient
            If .getpatient(objAccData.Fields("PtId")) Then
                lblPtNm.Caption = .ptnm
                lblSex.Caption = .SexNm
                lblAge.Caption = .Age
                lblAgeDiv.Caption = .AgeDiv
                lblDob.Caption = Format(.DOB, "####-##-##")
            End If
        End With
        Set objPatient = Nothing
        
        Call ICSPatientMark(lblPtid.Caption, enICSNum.LIS_ALL)
        
        lblLocation.Caption = objAccData.Fields("WardId") & "-" & objAccData.Fields("RoomId")
        lblSpcNm.Caption = objAccData.Fields("SpcNm")
        lblSpcNm.Tag = objAccData.Fields("spccd")
        
        lblColDtTm.Caption = Format(objAccData.Fields("ColDt"), CS_DateMask) & "  " & _
                             Format(objAccData.Fields("ColTm"), CS_TimeLongMask)
        
        lblColNm.Caption = GetEmpNm(objAccData.Fields("ColId")) 'GetEmpName(objAccData.Fields("ColId"))
        
        If Val(objAccData.Fields("stscd")) > enStsCd.StsCd_LIS_Collection Then
            lblRcvDtTm.Caption = Format(objAccData.Fields("RcvDt"), CS_DateMask) & "  " & _
                                 Format(objAccData.Fields("RcvTm"), CS_TimeLongMask)
            lblRcvNm.Caption = GetEmpNm(objAccData.Fields("RcvId")) 'GetEmpName(objAccData.Fields("RcvId"))
        End If
        
        Call DisplayOrder(txtWorkArea.Text, tmpAccDt, txtAccSeq.Text)
        
        Call DetailQuery
        
        txtWorkArea.Enabled = False
        txtAccDt.Enabled = False
        txtAccSeq.Enabled = False
        cmdSave.Enabled = True
        
        strTmp = MsgBox("처방을 조회하시겠습니까?", vbYesNo + vbInformation, "Info")
        If strTmp = vbYes Then
            Call AddOrder
        End If
    End If
    
'    Set objData = Nothing
End Sub
Private Sub DetailQuery()
    Dim SSQL    As String
    
    Dim RS      As Recordset
    
    objType.sRcvDt = "": objType.sRcvId = ""
    objType.sRcvTm = "": objType.sStoreCd = ""
    
    SSQL = " SELECT * " & _
           " FROM  " & T_LAB201 & " a " & _
           " WHERE " & DBW("a.workarea", txtWorkArea.Text, 2) & _
           " AND   " & DBW("a.accdt", tmpAccDt, 2) & _
           " AND   " & DBW("a.accseq", txtAccSeq.Text, 2)
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    
    If Not RS.EOF Then
        If lblStatus.Caption = STS_LIS_Access Then
            objType.sRcvDt = RS.Fields("rcvdt").Value & ""
            objType.sRcvTm = RS.Fields("rcvtm").Value & ""
            objType.sRcvId = RS.Fields("rcvid").Value & ""
        Else
            objType.sRcvDt = RS.Fields("coldt").Value & ""
        End If
        objType.sStoreCd = RS.Fields("storecd").Value & ""
    End If
    Set RS = Nothing
End Sub
Private Sub DisplayOrder(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = " SELECT a.ptid, a.orddt, a.ordno, a.ordseq, a.ordcd as testcd, a.statfg, b.testnm " & _
                " FROM  " & T_LAB102 & " a, " & t_lab001 & " b " & _
                " WHERE " & DBW("a.workarea", pWorkArea, 2) & _
                " AND   " & DBW("a.accdt", pAccDt, 2) & _
                " AND   " & DBW("a.accseq", pAccSeq, 2) & _
                " AND   b.testcd = a.ordcd " & _
                " ORDER BY orddt, ordno, ordseq"
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    If Not RS.EOF Then
        With tblOrdSheet
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = Format(RS.Fields("orddt").Value & "", "####-##-##")
                .Col = 2: .Value = RS.Fields("ordno").Value & ""
                .Col = 3: .Value = RS.Fields("testcd").Value & ""
                .Col = 4: .Value = RS.Fields("testnm").Value & ""
                .Col = 5: .Value = IIf(RS.Fields("statfg").Value & "" = "1", "Y", ""):
                          .ForeColor = DCM_LightRed:
                          .TypeHAlign = TypeHAlignCenter
                .Col = 6: .Value = RS.Fields("ordseq").Value & ""
                RS.MoveNext
            Loop
        End With
    End If
    Set RS = Nothing
End Sub

Private Sub AddOrder()
    Dim RS   As Recordset
    Dim SSQL As String
    
    Call TableIni
    
    SSQL = "SELECT a.ptid ,a.orddt ,a.ordno,a.ordseq,a.ordcd as testcd,b.reqdt,b.reqtm,a.statfg ,c.abbrnm10 as testnm" & _
           " FROM " & T_LAB102 & " a," & T_LAB101 & " b," & t_lab001 & " c" & _
           " WHERE " & _
           "      " & DBW("a.ptid=", lblPtid.Caption) & _
           " AND  " & DBW("a.orddt<=", objType.sRcvDt) & _
           " AND  " & DBW("b.orddiv=", lis_orddiv) & _
           " AND  " & DBW("b.donefg=", enStsCd.StsCd_LIS_Order) & _
           " AND   a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno" & _
           " AND   (a.dcfg<>'' or a.dcfg is null)" & _
           " AND   not exists(SELECT * FROM " & T_LAB031 & " z" & _
           "                 WHERE " & DBW("z.cdindex=", LC2_MultiSpc) & _
           "                 AND  z.cdval1=a.spccd)" & _
           " AND  a.ordcd=c.testcd" & _
           " AND  " & DBW("c.workarea=", txtWorkArea.Text) & _
           " AND  " & DBW("c.testdiv=", enTestDiv.TST_RouTest) '& _
           " AND  a.storecd=(SELECT storecd FROM " & t_lab201 & _
           "                 WHERE " & _
                              DBW("workarea=", txtWorkArea.Text) & _
           " AND          " & DBW("accdt=", tmpAccDt) & _
           " AND          " &dbw("accseq=",txtaccseq.Text) & _
           " AND  a.spccd=d.spccd"
    Set RS = New Recordset
    RS.Open SSQL, dbconn
    
    With tblAddOrder
        .ReDraw = False
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows
                End If
                .Row = .DataRowCnt + 1
                .Col = TblCol.tcChk:    .CellType = CellTypeCheckBox: .TypeCheckCenter = True
                .Col = TblCol.tcOrdCd:  .Value = RS.Fields("testcd").Value & ""
                .Col = TblCol.tcORDDT:  .Value = Format(RS.Fields("orddt").Value & "", "####-##-##")
                .Col = TblCol.tcORDNO:  .Value = RS.Fields("ordno").Value & ""
                .Col = TblCol.tcORDSEQ: .Value = RS.Fields("ordseq").Value & ""
                .Col = TblCol.tcReqdt:  .Value = Format(RS.Fields("reqdt").Value & "", "####-##-##")
                                        .Value = .Value & " " & Format(RS.Fields("reqtm").Value & "", "0#:##:##")
                .Col = TblCol.tcSTAT:   .Value = RS.Fields("statfg").Value & ""
                                        .Value = IIf(.Value = "1", "Y", ""): .ForeColor = IIf(.Value <> "", DCM_LightRed, vbBlack)
                .Col = TblCol.tcTESTNM: .Value = RS.Fields("testnm").Value & ""
                RS.MoveNext
            Loop
        End If
        .ReDraw = True
    End With
    Set RS = Nothing

End Sub


Private Sub ClearRtn()
    
    txtWorkArea.Enabled = True
    txtAccDt.Enabled = True
    txtAccSeq.Enabled = True
    cmdSave.Enabled = False
    
    txtWorkArea.Text = ""
    txtAccDt.Text = ""
    txtAccSeq.Text = ""

    lblStatus.Caption = ""
   
    lblPtid.Caption = ""
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDoctNm.Caption = ""
    lblDeptNm.Caption = ""
    lblLocation.Caption = ""
    
    lblSpcNm.Caption = ""
    lblColDtTm.Caption = ""
    lblColNm.Caption = ""
    lblRcvDtTm.Caption = ""
    lblRcvNm.Caption = ""
    lblDob.Caption = ""
    
    Call medClearTable(tblOrdSheet)
    Call TableIni
    ClearFg = True
    Call ICSPatientMark
End Sub

Private Sub TableIni()
    With tblAddOrder
        .Row = -1
        .Col = 1: .COL2 = .MaxRows
        .BlockMode = True
        .CellType = CellTypeStaticText
        .Value = ""
        .BlockMode = False
    End With
End Sub
Private Sub cmdSave_Click()
'상태에 따라 2가지로 구분
'1:채혈
'2:접수

'1=====================================
'lab101,lab102의 상태변경및 102의 접수번호를 업데이트 한다
'업데이트 후 접수처리여부 확인

'2=====================================
'접수취소(채혈상태까지만)
'1번상태와 동일처리
'무조건 접수처리
'접수시 기존 접수시간,접수자를 업데이트 한다.
'추가 오더의 경우 접수자,접수시간이 틀릴수 있으므로
    
    If lblStatus.Caption = STS_LIS_HaveSpc Then
        If CollectStatusAddOrder Then
            Dim sTmp As String
            
            sTmp = MsgBox("접수하시겠습니까?", vbYesNo + vbInformation, "Info")
            If sTmp = vbYes Then
                Call DoAccession
            End If
        End If
    End If
    
    If lblStatus.Caption = STS_LIS_Access Then Call AccessStatueAddOrder
    Call ClearRtn
End Sub

Private Function CollectStatusAddOrder() As Boolean

    Dim SSQL    As String
    Dim sPtid   As String
    Dim sOrdDt  As String
    Dim sOrdNo  As String
    Dim sOrdSeq As String
    Dim ii      As Integer
    
    On Error GoTo SAVE_ERROR
    
    dbconn.BeginTrans
    sPtid = lblPtid.Caption
    
    With tblAddOrder
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblCol.tcChk
            If .CellType = CellTypeCheckBox Then
                If .Value = 1 Then
                    .Col = TblCol.tcORDDT: sOrdDt = Replace(.Value, "-", "")
                    .Col = TblCol.tcORDNO: sOrdNo = .Value
                    .Col = TblCol.tcORDSEQ: sOrdSeq = .Value
                    
                    dbconn.Execute HeaderBodyUpdate(sPtid, sOrdDt, sOrdNo)
                    dbconn.Execute HeaderBodyUpdate(sPtid, sOrdDt, sOrdNo, sOrdSeq)
                End If
            End If
        Next
    End With
    dbconn.CommitTrans
    CollectStatusAddOrder = True
    Exit Function
SAVE_ERROR:
    dbconn.RollbackTrans
    
End Function

Private Function AccessStatueAddOrder() As Boolean
    '접수취소
    If DoCancelAccession = True Then
        '추가 처방채혈
        If CollectStatusAddOrder = True Then
            Call DoAccession
        End If
    End If
End Function
Private Sub DoAccession()
    Dim objAccess  As New clsLISAccession
    Dim blnSuccess As Boolean
    
    MouseRunning  '13
      
    With objAccess
        blnSuccess = .DoAccession(txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, ObjSysInfo.EmpId)
        If blnSuccess Then
            Call RcvDataChange
        End If
    End With
    Set objAccess = Nothing
End Sub

Private Sub RcvDataChange()
    Dim SSQL    As String
    
    If objType.sRcvDt = "" Then Exit Sub
On Error GoTo SAVE_ERROR
    dbconn.BeginTrans
    
    SSQL = "update " & T_LAB201 & " set " & DBW("rcvdt", objType.sRcvDt, 3) & DBW("rcvtm", objType.sRcvTm, 3) & _
                              DBW("rcvid", objType.sRcvId, 2) & _
           "  WHERE " & DBW("workarea", txtWorkArea.Text, 2) & _
           " AND   " & DBW("accdt", tmpAccDt, 2) & _
           " AND   " & DBW("accseq", txtAccSeq.Text, 2)
         
    dbconn.Execute SSQL
    dbconn.CommitTrans
    Exit Sub
SAVE_ERROR:
    dbconn.RollbackTrans
End Sub

Private Function DoCancelAccession() As Boolean
    Dim objAccSql   As New clsLISSqlAccession
    Dim objOrdDic   As New clsDictionary
    Dim sOrdDt      As String
    Dim sOrdNo      As String
    Dim sOrdSeq     As String
    Dim sFlag       As String
    Dim sPtid       As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim sStsCd      As String
    
    Dim SqlStmt As String
    Dim Resp    As Boolean
    Dim i       As Long
    Dim TblNames(0 To 14)
    
    sFlag = "2"
    sPtid = lblPtid.Caption
    sWorkArea = txtWorkArea.Text
    sAccDt = tmpAccDt
    sAccSeq = txtAccSeq.Text
    sStsCd = enStsCd.StsCd_LIS_Collection
    
    
    TblNames(0) = T_LAB203
    TblNames(1) = T_LAB205
    TblNames(2) = T_LAB302
    TblNames(3) = T_LAB303
    TblNames(4) = T_LAB304
    TblNames(5) = T_LAB305
    TblNames(6) = T_LAB308
    TblNames(7) = T_LAB351
    TblNames(8) = T_LAB353
    TblNames(9) = T_LAB354
    TblNames(10) = T_LAB404
    TblNames(11) = T_LAB405
    TblNames(12) = T_LAB407
    TblNames(13) = T_LAB360
    TblNames(14) = T_LAB361
   
    On Error GoTo Err_Trap
    
    dbconn.BeginTrans
    
    '결과내역 Delete
    For i = 0 To 14
        If IsDataExists(TblNames(i), sWorkArea, sAccDt, sAccSeq) Then
            SqlStmt = objAccSql.SqlDelRstTable(TblNames(i), sWorkArea, sAccDt, sAccSeq)
            dbconn.Execute SqlStmt
        End If
    Next

    '접수내역 Update : 처방상태로.. status는 'D'(취소), '채혈상태로.. status는 '1'(채혈)
    SqlStmt = objAccSql.SqlCancel201(sWorkArea, sAccDt, sAccSeq, sStsCd)
    Call dbconn.Execute(SqlStmt)

    objOrdDic.Clear
    objOrdDic.FieldInialize "orddt,ordno", "updfg"
    
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1: sOrdDt = Format(.Value, CS_DateDbFormat)  '처방일
            .Col = 2: sOrdNo = .Value   '처방번호
            .Col = 6: sOrdSeq = .Value  '처방Seq
            
            '처방Body Update
            SqlStmt = objAccSql.SqlUpdateOrdB(sPtid, sOrdDt, sOrdNo, sOrdSeq, sFlag)
            Call dbconn.Execute(SqlStmt)
            
            '처방Header Update
            If Not objOrdDic.Exists(sOrdDt & COL_DIV & sOrdNo) Then
                SqlStmt = objAccSql.SqlUpdateOrdH(sPtid, sOrdDt, sOrdNo, sFlag)
                Call dbconn.Execute(SqlStmt)
                
                objOrdDic.AddNew sOrdDt & COL_DIV & sOrdNo, "Y"
            End If
        Next
    End With
    
    dbconn.CommitTrans
    DoCancelAccession = True
    Set objAccSql = Nothing
    Set objOrdDic = Nothing
    Exit Function
    
Err_Trap:
    dbconn.RollbackTrans
    DoCancelAccession = False
    Set objAccSql = Nothing
    Set objOrdDic = Nothing
End Function
Private Function IsDataExists(ByVal pTblNm As String, ByVal pWorkArea As String, _
                              ByVal pAccDt As String, ByVal pAccSeq As String) As Boolean

    Dim tmpRs As Recordset
    Dim objAccSql   As New clsLISSqlAccession
    Set tmpRs = New Recordset
    tmpRs.Open objAccSql.SqlDataExists(pTblNm, pWorkArea, pAccDt, pAccSeq), dbconn
    IsDataExists = Not tmpRs.EOF
    Set tmpRs = Nothing
    Set objAccSql = Nothing
End Function
Private Function HeaderBodyUpdate(ByVal sPtid As String, ByVal sOrdDt As String, ByVal sOrdNo As String, _
                                  Optional ByVal sOrdSeq As String = "") As String
    Dim SSQL As String
    
    If sOrdSeq = "" Then
        SSQL = "UPDATE " & T_LAB101 & " SET " & _
                           DBW("donefg", enStsCd.StsCd_LIS_Collection, 2) & _
               " WHERE " & DBW("ptid=", sPtid) & _
               " AND " & DBW("orddt=", sOrdDt) & _
               " AND " & DBW("ordno=", sOrdNo)
               
    Else
        SSQL = "UPDATE " & T_LAB102 & " SET " & _
                           DBW("donefg", enStsCd.StsCd_LIS_Collection, 3) & _
                           DBW("stscd", enStsCd.StsCd_LIS_Collection, 3) & _
                           DBW("workarea", txtWorkArea.Text, 3) & _
                           DBW("accdt", tmpAccDt, 3) & _
                           DBW("accseq", txtAccSeq.Text, 2) & _
               " WHERE " & DBW("ptid=", sPtid) & _
               " AND " & DBW("orddt=", sOrdDt) & _
               " AND " & DBW("ordno=", sOrdNo) & _
               " AND " & DBW("ordseq=", sOrdSeq)
    End If
    HeaderBodyUpdate = SSQL


End Function

