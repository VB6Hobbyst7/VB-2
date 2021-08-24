VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "검사설정"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16245
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   16245
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraTypeChange 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 결과 변환 "
      Height          =   3165
      Left            =   5010
      TabIndex        =   107
      Top             =   5160
      Visible         =   0   'False
      Width           =   6105
      Begin VB.CommandButton cmdUnView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "숨김 ▶"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4560
         Style           =   1  '그래픽
         TabIndex        =   60
         Top             =   2580
         Width           =   1275
      End
      Begin VB.CommandButton cmdConfirm 
         BackColor       =   &H00C0FFFF&
         Caption         =   "전체코드적용"
         Height          =   435
         Index           =   5
         Left            =   1830
         Style           =   1  '그래픽
         TabIndex        =   58
         Top             =   2580
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "현재코드적용"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   3210
         TabIndex        =   59
         Top             =   2580
         Width           =   1305
      End
      Begin VB.TextBox txtAMRResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   3540
         TabIndex        =   57
         Top             =   2190
         Width           =   1845
      End
      Begin VB.TextBox txtAMRLimit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   1170
         TabIndex        =   56
         Top             =   2190
         Width           =   1845
      End
      Begin VB.TextBox txtAMRResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   3540
         TabIndex        =   55
         Top             =   1860
         Width           =   1845
      End
      Begin VB.TextBox txtAMRLimit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   1170
         TabIndex        =   54
         Top             =   1860
         Width           =   1845
      End
      Begin VB.TextBox txtAMRResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   3540
         TabIndex        =   53
         Top             =   1530
         Width           =   1845
      End
      Begin VB.TextBox txtAMRLimit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   1170
         TabIndex        =   52
         Top             =   1530
         Width           =   1845
      End
      Begin VB.TextBox txtAMRResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   3540
         TabIndex        =   51
         Top             =   1200
         Width           =   1845
      End
      Begin VB.TextBox txtAMRLimit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   1170
         TabIndex        =   50
         Top             =   1200
         Width           =   1845
      End
      Begin VB.TextBox txtAMRResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   3540
         TabIndex        =   49
         Top             =   870
         Width           =   1845
      End
      Begin VB.TextBox txtAMRLimit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   1170
         TabIndex        =   48
         Top             =   870
         Width           =   1845
      End
      Begin VB.TextBox txtAMRResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   3540
         TabIndex        =   47
         Top             =   540
         Width           =   1845
      End
      Begin VB.TextBox txtAMRLimit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   1170
         TabIndex        =   46
         Top             =   540
         Width           =   1845
      End
      Begin VB.TextBox txtAMRResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   3540
         TabIndex        =   45
         Top             =   210
         Width           =   1845
      End
      Begin VB.TextBox txtAMRLimit 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   1170
         TabIndex        =   44
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "변환"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   56
         Left            =   5490
         TabIndex        =   128
         Top             =   2250
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경우"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   55
         Left            =   3090
         TabIndex        =   127
         Top             =   2250
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사결과"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   54
         Left            =   330
         TabIndex        =   126
         Top             =   2250
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "변환"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   53
         Left            =   5490
         TabIndex        =   125
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경우"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   52
         Left            =   3090
         TabIndex        =   124
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사결과"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   51
         Left            =   330
         TabIndex        =   123
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "변환"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   50
         Left            =   5490
         TabIndex        =   122
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경우"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   49
         Left            =   3090
         TabIndex        =   121
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사결과"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   48
         Left            =   330
         TabIndex        =   120
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "변환"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   47
         Left            =   5490
         TabIndex        =   119
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경우"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   46
         Left            =   3090
         TabIndex        =   118
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사결과"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   45
         Left            =   330
         TabIndex        =   117
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "변환"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   44
         Left            =   5490
         TabIndex        =   116
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경우"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   43
         Left            =   3090
         TabIndex        =   115
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사결과"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   42
         Left            =   330
         TabIndex        =   114
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "변환"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   41
         Left            =   5490
         TabIndex        =   113
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경우"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   39
         Left            =   3090
         TabIndex        =   112
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사결과"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   38
         Left            =   330
         TabIndex        =   111
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "변환"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   37
         Left            =   5490
         TabIndex        =   110
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "경우"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   36
         Left            =   3090
         TabIndex        =   109
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사결과"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   35
         Left            =   330
         TabIndex        =   108
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   16245
      TabIndex        =   62
      Top             =   0
      Width           =   16245
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "검사정보 설정"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   210
         TabIndex        =   63
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.Frame frameTestSet 
      BackColor       =   &H00FFFFFF&
      Height          =   9675
      Left            =   10950
      TabIndex        =   64
      Top             =   540
      Width           =   5205
      Begin VB.CommandButton cmdTypeChange 
         BackColor       =   &H00FFFFFF&
         Caption         =   "◀ 결과변환 보임"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   210
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   7350
         Width           =   4755
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   210
         TabIndex        =   129
         Top             =   3600
         Width           =   4755
         Begin VB.OptionButton optResType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "수치결과"
            Height          =   195
            Index           =   0
            Left            =   1170
            TabIndex        =   16
            Top             =   210
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optResType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "판정결과"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   17
            Top             =   210
            Width           =   1035
         End
         Begin VB.OptionButton optResType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "수치/판정"
            Height          =   195
            Index           =   2
            Left            =   3450
            TabIndex        =   18
            Top             =   210
            Width           =   1125
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과사용"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   34
            Left            =   150
            TabIndex        =   130
            Top             =   210
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 결과포함표시 "
         Height          =   945
         Left            =   210
         TabIndex        =   106
         Top             =   7860
         Width           =   4755
         Begin VB.OptionButton optINQuant 
            BackColor       =   &H00FFFFFF&
            Caption         =   "수치_판정"
            Height          =   255
            Index           =   4
            Left            =   3030
            TabIndex        =   38
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton optINQuant 
            BackColor       =   &H00FFFFFF&
            Caption         =   "판정_수치"
            Height          =   255
            Index           =   3
            Left            =   1530
            TabIndex        =   37
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton optINQuant 
            BackColor       =   &H00FFFFFF&
            Caption         =   "수치(판정)"
            Height          =   255
            Index           =   2
            Left            =   3030
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optINQuant 
            BackColor       =   &H00FFFFFF&
            Caption         =   "판정(수치)"
            Height          =   255
            Index           =   1
            Left            =   1530
            TabIndex        =   35
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optINQuant 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용안함"
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.CheckBox chkResSpec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   180
         Left            =   1440
         TabIndex        =   8
         Top             =   2940
         Width           =   705
      End
      Begin VB.TextBox txtRChannel 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   1410
         Width           =   3435
      End
      Begin VB.TextBox txtEqpCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txtTestCd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   5
         Top             =   1770
         Width           =   2145
      End
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   2130
         Width           =   2145
      End
      Begin VB.TextBox txtOChannel 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   1050
         Width           =   3435
      End
      Begin VB.TextBox txtAbbrNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   2490
         Width           =   2145
      End
      Begin VB.TextBox txtResSpec 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         TabIndex        =   9
         Top             =   2880
         Width           =   405
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtRefMLow 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   3270
         Width           =   585
      End
      Begin VB.TextBox txtRefMHigh 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2490
         TabIndex        =   13
         Top             =   3270
         Width           =   585
      End
      Begin VB.CommandButton cmdSeqDown 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3180
         TabIndex        =   2
         Top             =   660
         Width           =   405
      End
      Begin VB.CommandButton cmdSeqUp 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2760
         TabIndex        =   1
         Top             =   660
         Width           =   405
      End
      Begin VB.CommandButton cmdSpecDown 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3150
         TabIndex        =   11
         Top             =   2850
         Width           =   435
      End
      Begin VB.CommandButton cmdSpecUP 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2700
         TabIndex        =   10
         Top             =   2850
         Width           =   435
      End
      Begin VB.TextBox txtRefFLow 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3390
         TabIndex        =   14
         Top             =   3240
         Width           =   585
      End
      Begin VB.TextBox txtRefFHigh 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4200
         TabIndex        =   15
         Top             =   3240
         Width           =   585
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4020
         TabIndex        =   43
         Top             =   8970
         Width           =   915
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   2160
         TabIndex        =   41
         Top             =   8970
         Width           =   915
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   3090
         TabIndex        =   42
         Top             =   8970
         Width           =   915
      End
      Begin VB.CommandButton cmdConfirm 
         BackColor       =   &H00C0FFFF&
         Caption         =   "전체저장"
         Height          =   465
         Index           =   2
         Left            =   210
         Style           =   1  '그래픽
         TabIndex        =   39
         Top             =   8970
         Width           =   945
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 결과 변환 (수치형) "
         Height          =   1695
         Left            =   210
         TabIndex        =   75
         Top             =   4200
         Width           =   4755
         Begin VB.TextBox txtAMRLimit 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1410
            TabIndex        =   19
            Top             =   270
            Width           =   1035
         End
         Begin VB.TextBox txtAMRResult 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   2970
            TabIndex        =   20
            Top             =   270
            Width           =   1125
         End
         Begin VB.TextBox txtAMRLimit 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1410
            TabIndex        =   21
            Top             =   600
            Width           =   1035
         End
         Begin VB.TextBox txtAMRResult 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2970
            TabIndex        =   22
            Top             =   600
            Width           =   1125
         End
         Begin VB.TextBox txtAMRLimit 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1410
            TabIndex        =   23
            Top             =   930
            Width           =   1035
         End
         Begin VB.TextBox txtAMRResult 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   2970
            TabIndex        =   24
            Top             =   930
            Width           =   1125
         End
         Begin VB.TextBox txtAMRLimit 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1410
            TabIndex        =   25
            Top             =   1260
            Width           =   1035
         End
         Begin VB.TextBox txtAMRResult 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   2970
            TabIndex        =   26
            Top             =   1260
            Width           =   1125
         End
         Begin VB.TextBox TtxtCmp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   79
            Text            =   "<"
            Top             =   270
            Width           =   315
         End
         Begin VB.TextBox TtxtCmp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   78
            Text            =   "<="
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox TtxtCmp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   77
            Text            =   ">"
            Top             =   930
            Width           =   315
         End
         Begin VB.TextBox TtxtCmp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   76
            Text            =   ">="
            Top             =   1260
            Width           =   315
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사결과"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   210
            TabIndex        =   91
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사결과"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   210
            TabIndex        =   90
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사결과"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   7
            Left            =   210
            TabIndex        =   89
            Top             =   990
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사결과"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   9
            Left            =   210
            TabIndex        =   88
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "경우"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   2520
            TabIndex        =   87
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "경우"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   18
            Left            =   2520
            TabIndex        =   86
            Top             =   660
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "경우"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   20
            Left            =   2520
            TabIndex        =   85
            Top             =   990
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "경우"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   21
            Left            =   2520
            TabIndex        =   84
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "변환"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   24
            Left            =   4170
            TabIndex        =   83
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "변환"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   25
            Left            =   4170
            TabIndex        =   82
            Top             =   990
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "변환"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   26
            Left            =   4170
            TabIndex        =   81
            Top             =   660
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "변환"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   27
            Left            =   4170
            TabIndex        =   80
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 결과 변환 (판정형) "
         Height          =   1365
         Left            =   210
         TabIndex        =   65
         Top             =   5970
         Width           =   4755
         Begin VB.TextBox txtAMRResult 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   2970
            TabIndex        =   28
            Top             =   270
            Width           =   1125
         End
         Begin VB.TextBox txtAMRLimit 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1080
            TabIndex        =   27
            Top             =   270
            Width           =   1335
         End
         Begin VB.TextBox txtAMRResult 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   2970
            TabIndex        =   30
            Top             =   600
            Width           =   1125
         End
         Begin VB.TextBox txtAMRLimit 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   1080
            TabIndex        =   29
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtAMRResult 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   2970
            TabIndex        =   32
            Top             =   930
            Width           =   1125
         End
         Begin VB.TextBox txtAMRLimit 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   1080
            TabIndex        =   31
            Top             =   930
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "변환"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   23
            Left            =   4230
            TabIndex        =   74
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "경우"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   22
            Left            =   2520
            TabIndex        =   73
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Neg결과"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   17
            Left            =   210
            TabIndex        =   72
            Top             =   330
            Width           =   705
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "변환"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   28
            Left            =   4230
            TabIndex        =   71
            Top             =   660
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "경우"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   29
            Left            =   2520
            TabIndex        =   70
            Top             =   660
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Pos결과"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   30
            Left            =   210
            TabIndex        =   69
            Top             =   660
            Width           =   690
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "변환"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   31
            Left            =   4230
            TabIndex        =   68
            Top             =   990
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "경우"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   32
            Left            =   2520
            TabIndex        =   67
            Top             =   990
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Bod결과"
            ForeColor       =   &H00FF00FF&
            Height          =   180
            Index           =   33
            Left            =   210
            TabIndex        =   66
            Top             =   990
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdConfirm 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   1170
         Style           =   1  '그래픽
         TabIndex        =   40
         Top             =   8970
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   1260
         Left            =   3630
         Picture         =   "frmTestSet.frx":1272
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "남"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   40
         Left            =   1440
         TabIndex        =   105
         Top             =   3330
         Width           =   180
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "순번"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   15
         Left            =   390
         TabIndex        =   104
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과채널"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   19
         Left            =   390
         TabIndex        =   103
         Top             =   1485
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   390
         TabIndex        =   102
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "오더채널"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   10
         Left            =   390
         TabIndex        =   101
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   11
         Left            =   390
         TabIndex        =   100
         Top             =   1845
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사명"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   12
         Left            =   390
         TabIndex        =   99
         Top             =   2205
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사약어"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   13
         Left            =   390
         TabIndex        =   98
         Top             =   2565
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "소수점"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   390
         TabIndex        =   97
         Top             =   2925
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "참고치"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   16
         Left            =   390
         TabIndex        =   96
         Top             =   3330
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "여"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3150
         TabIndex        =   95
         Top             =   3330
         Width           =   180
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   2310
         TabIndex        =   94
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   4020
         TabIndex        =   93
         Top             =   3330
         Width           =   135
      End
   End
   Begin FPSpread.vaSpread spdTest 
      Height          =   9555
      Left            =   60
      TabIndex        =   61
      Top             =   630
      Width           =   10845
      _Version        =   393216
      _ExtentX        =   19129
      _ExtentY        =   16854
      _StockProps     =   64
      BackColorStyle  =   3
      ColsFrozen      =   6
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕코딩"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   22
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestSet.frx":2AE4
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmTestSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click(Index As Integer)
    Dim Test_Property       As Scripting.Dictionary
    Dim objTest_Property    As clsCommon
    Dim i                   As Integer
    Dim strTmp              As String
    Dim intINQuant          As Integer
    Dim intResUse           As Integer
    
    If optINQuant(0).Value = True Then
        intINQuant = 0
    ElseIf optINQuant(1).Value = True Then
        intINQuant = 1      '정성(정량)
    ElseIf optINQuant(2).Value = True Then
        intINQuant = 2      '정량(정성)
    ElseIf optINQuant(3).Value = True Then
        intINQuant = 3      '정성_정량
    ElseIf optINQuant(4).Value = True Then
        intINQuant = 4      '정량_정성
    End If
    
    If optResType(0).Value = True Then
        intResUse = 0
    ElseIf optResType(1).Value = True Then
        intResUse = 1
    ElseIf optResType(2).Value = True Then
        intResUse = 2
    End If
    
    If Index = 1 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If


        If MsgBox(txtTestNm.Text & "를 삭제하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
             Exit Sub
        End If
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .DelTestInfo(Test_Property) Then
                '-- 삭제 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)

    ElseIf Index = 0 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If

'        If Trim(txtOChannel.Text) = "" Then
'            MsgBox "오더채널을 입력하세요", vbCritical, Me.Caption
'            txtOChannel.SetFocus
'            Exit Sub
'        End If
'
        If Trim(txtRChannel.Text) = "" Then
            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
            txtRChannel.SetFocus
            Exit Sub
        End If

        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If

        If Trim(txtTestNm.Text) = "" Then
            MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
            txtTestNm.SetFocus
            Exit Sub
        End If

        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
            If chkResSpec.Value = "0" Then
                .Add "RESUSE", "0"
            Else
                .Add "RESUSE", "1"
            End If
            .Add "RES", txtResSpec.Text
            .Add "REFML", txtRefMLow.Text
            .Add "REFMH", txtRefMHigh.Text
            .Add "REFFL", txtRefFLow.Text
            .Add "REFFH", txtRefFHigh.Text
            
            .Add "USERESULT", intResUse
            
            '-- 결과변환 : 수치형
            .Add "AMRLIMIT1", txtAMRLimit(0).Text
            .Add "AMRLIMIT2", txtAMRLimit(1).Text
            .Add "AMRLIMIT3", txtAMRLimit(2).Text
            .Add "AMRLIMIT4", txtAMRLimit(3).Text
            '-- 결과변환 : 문자형
            .Add "AMRLIMIT5", txtAMRLimit(4).Text
            .Add "AMRLIMIT6", txtAMRLimit(5).Text
            .Add "AMRLIMIT7", txtAMRLimit(6).Text
        
            .Add "AMRRESULT1", txtAMRResult(0).Text
            .Add "AMRRESULT2", txtAMRResult(1).Text
            .Add "AMRRESULT3", txtAMRResult(2).Text
            .Add "AMRRESULT4", txtAMRResult(3).Text
            .Add "AMRRESULT5", txtAMRResult(4).Text
            .Add "AMRRESULT6", txtAMRResult(5).Text
            .Add "AMRRESULT7", txtAMRResult(6).Text
        
            .Add "AMRINRESULT", intINQuant
            
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetTestInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)

    ElseIf Index = 2 Then
        SQL = ""
        SQL = SQL & "DELETE FROM EQPMASTER"
                
        Call DBExec(AdoCn_Local, SQL)
        
        With spdTest
            For i = 1 To .MaxRows
                SQL = ""
                SQL = SQL & "INSERT INTO EQPMASTER " & vbCrLf
                SQL = SQL & "(EQUIPCD, SEQNO, SENDCHANNEL, RSLTCHANNEL"
                SQL = SQL & " , TESTCODE, TESTNAME, ABBRNAME, RESPRECUSE, RESPREC "
                SQL = SQL & " , REFMLOW, REFMHIGH, REFFLOW, REFFHIGH,RESTYPE" & vbCrLf
                '-- AMR
                SQL = SQL & " , AMRLimit1, AMRResult1, AMRLimit2, AMRResult2, AMRLimit3, AMRResult3 " & vbCrLf
                SQL = SQL & " , AMRLimit4, AMRResult4, AMRLimit5, AMRResult5, AMRLimit6, AMRResult6 " & vbCrLf
                SQL = SQL & " , AMRLimit7, AMRResult7, AMRINResult)                                 " & vbCrLf
                SQL = SQL & " VALUES (" & vbCrLf
                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
                SQL = SQL & "," & STS(GetText(spdTest, i, colLOCHANNEL))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTNM))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLABBRNM))
                SQL = SQL & "," & GetText(spdTest, i, colLRESSPECUSE)
                SQL = SQL & "," & GetText(spdTest, i, colLRESSPEC)
                SQL = SQL & "," & STS(GetText(spdTest, i, colLMLOW))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLMHIGH))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLFLOW))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLFHIGH))
                SQL = SQL & "," & STS(GetText(spdTest, i, colRESTYPE))
                '-- AMR
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 1))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 2))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 3))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 4))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 5))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 6))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 7))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                SQL = SQL & "," & STS(GetText(spdTest, i, colRESTYPE + 8))
                SQL = SQL & ")" & vbCrLf
                
                Call DBExec(AdoCn_Local, SQL)
            
            Next
        End With
        
        Call GetTestList
        Call GetTestMaster(spdTest)
        
    ElseIf Index = 3 Then
        
        Call GetTestList
        Call GetTestMaster(spdTest)
    
    ElseIf Index = 4 Then
        If Trim(txtEqpCD.Text) = "" Then
            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
            Exit Sub
        End If

        If Trim(txtRChannel.Text) = "" Then
            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
            txtRChannel.SetFocus
            Exit Sub
        End If
        
        If Trim(txtTestCd.Text) = "" Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If
        
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTCD", txtTestCd.Text
            
            '-- 결과변환 : 문자형
            .Add "AMRLIMIT8", txtAMRLimit(7).Text
            .Add "AMRLIMIT9", txtAMRLimit(8).Text
            .Add "AMRLIMIT10", txtAMRLimit(9).Text
            .Add "AMRLIMIT11", txtAMRLimit(10).Text
            .Add "AMRLIMIT12", txtAMRLimit(11).Text
            .Add "AMRLIMIT13", txtAMRLimit(12).Text
            .Add "AMRLIMIT14", txtAMRLimit(13).Text
        
            .Add "AMRRESULT8", txtAMRResult(7).Text
            .Add "AMRRESULT9", txtAMRResult(8).Text
            .Add "AMRRESULT10", txtAMRResult(9).Text
            .Add "AMRRESULT11", txtAMRResult(10).Text
            .Add "AMRRESULT12", txtAMRResult(11).Text
            .Add "AMRRESULT13", txtAMRResult(12).Text
            .Add "AMRRESULT14", txtAMRResult(13).Text
        
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetAMRInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With

        Call GetTestList
        Call GetTestMaster(spdTest)
        
    ElseIf Index = 5 Then
        SQL = ""
        SQL = SQL & "DELETE FROM AMRMASTER"
                
        Call DBExec(AdoCn_Local, SQL)
        
        With spdTest
            For i = 1 To .MaxRows
                SQL = ""
                SQL = SQL & "INSERT INTO AMRMASTER " & vbCrLf
                SQL = SQL & "(EQUIPCD, SEQNO, RSLTCHANNEL, TESTCODE"
                SQL = SQL & " , AMRLimit8, AMRLimit9, AMRLimit10, AMRLimit11, AMRLimit12, AMRLimit13, AMRLimit14 " & vbCrLf
                SQL = SQL & " , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14 ) " & vbCrLf
                SQL = SQL & " VALUES (" & vbCrLf
                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
                SQL = SQL & "," & STS(txtAMRLimit(7).Text)
                SQL = SQL & "," & STS(txtAMRLimit(8).Text)
                SQL = SQL & "," & STS(txtAMRLimit(9).Text)
                SQL = SQL & "," & STS(txtAMRLimit(10).Text)
                SQL = SQL & "," & STS(txtAMRLimit(11).Text)
                SQL = SQL & "," & STS(txtAMRLimit(12).Text)
                SQL = SQL & "," & STS(txtAMRLimit(13).Text)
                SQL = SQL & "," & STS(txtAMRResult(7).Text)
                SQL = SQL & "," & STS(txtAMRResult(8).Text)
                SQL = SQL & "," & STS(txtAMRResult(9).Text)
                SQL = SQL & "," & STS(txtAMRResult(10).Text)
                SQL = SQL & "," & STS(txtAMRResult(11).Text)
                SQL = SQL & "," & STS(txtAMRResult(12).Text)
                SQL = SQL & "," & STS(txtAMRResult(13).Text)
                SQL = SQL & ")" & vbCrLf
                
                Call DBExec(AdoCn_Local, SQL)
            
            Next
        End With
        
        Call GetTestList
        Call GetTestMaster(spdTest)
    
    End If
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSeqDown_Click()

    txtSeq.Text = txtSeq.Text - 1

End Sub

Private Sub cmdSeqUp_Click()
    
    txtSeq.Text = txtSeq.Text + 1

End Sub

Private Sub cmdSpecDown_Click()

    txtResSpec.Text = txtResSpec.Text - 1

End Sub

Private Sub cmdSpecUP_Click()
    
    If txtResSpec.Text <> "" Then
        txtResSpec.Text = txtResSpec.Text + 1
    End If
    
End Sub


Private Sub cmdTypeChange_Click()
    
    If fraTypeChange.Visible = True Then
        fraTypeChange.Visible = False
        cmdTypeChange.Caption = "◀ 결과변환 보임"
    Else
        fraTypeChange.Visible = True
        cmdTypeChange.Caption = "▶ 결과변환 숨김"
    End If

End Sub


Private Sub cmdUnView_Click()
    
    fraTypeChange.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call frmClear

    Call GetTestMaster(spdTest)
    
End Sub

Private Sub frmClear()
    Dim i As Integer
    
    For i = 7 To 13
        txtAMRLimit(i).Text = ""
        txtAMRResult(i).Text = ""
    Next
        
End Sub


Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strResUse   As String
    
    If Row = 0 Then
        cmdTypeChange.Enabled = False
        Exit Sub
    End If

    With spdTest
        cmdTypeChange.Enabled = True
        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
        txtSeq.Text = GetText(spdTest, Row, colLSEQNO)
        txtOChannel.Text = GetText(spdTest, Row, colLOCHANNEL)
        txtRChannel.Text = GetText(spdTest, Row, colLRCHANNEL)
        txtTestCd.Text = GetText(spdTest, Row, colLTESTCD)
        txtTestNm.Text = GetText(spdTest, Row, colLTESTNM)
        txtAbbrNm.Text = GetText(spdTest, Row, colLABBRNM)
        
        If GetText(spdTest, Row, colLRESSPECUSE) = "0" Then
            chkResSpec.Value = "0"
        Else
            chkResSpec.Value = "1"
        End If
        txtResSpec.Text = GetText(spdTest, Row, colLRESSPEC)
        txtRefMLow.Text = GetText(spdTest, Row, colLMLOW)
        txtRefMHigh.Text = GetText(spdTest, Row, colLMHIGH)
        txtRefFLow.Text = GetText(spdTest, Row, colLFLOW)
        txtRefFHigh.Text = GetText(spdTest, Row, colLFHIGH)
        
        strResUse = GetText(spdTest, Row, colRESTYPE)
        
        If strResUse = "" Or strResUse = "0" Then
            optResType(0).Value = True
        ElseIf strResUse = "1" Then
            optResType(1).Value = True
        ElseIf strResUse = "2" Then
            optResType(2).Value = True
        End If
        
        'AMR
        txtAMRLimit(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 1, "|")
        txtAMRResult(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 2, "|")
    
        txtAMRLimit(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 1, "|")
        txtAMRResult(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 2, "|")
    
        txtAMRLimit(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 1, "|")
        txtAMRResult(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 2, "|")
    
        txtAMRLimit(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 1, "|")
        txtAMRResult(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 2, "|")
    
        txtAMRLimit(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 1, "|")
        txtAMRResult(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 2, "|")
    
        txtAMRLimit(5).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 1, "|")
        txtAMRResult(5).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 2, "|")
    
        txtAMRLimit(6).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 1, "|")
        txtAMRResult(6).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 2, "|")
    
        If GetText(spdTest, Row, colRESTYPE + 8) = "" Then
            optINQuant(0).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "0" Then
            optINQuant(0).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "1" Then
            optINQuant(1).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "2" Then
            optINQuant(2).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "3" Then
            optINQuant(3).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "4" Then
            optINQuant(4).Value = True
        End If
        
        Call frmClear
        Call GetAMRMaster(txtSeq.Text, txtRChannel.Text, txtTestCd.Text)
        
    End With
    
    'txtTestCd.SetFocus
End Sub

