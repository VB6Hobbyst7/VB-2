VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "검사설정"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16245
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   16245
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   16245
      TabIndex        =   1
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
         TabIndex        =   2
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.Frame frameTestSet 
      BackColor       =   &H00FFFFFF&
      Height          =   8865
      Left            =   10950
      TabIndex        =   3
      Top             =   540
      Width           =   5205
      Begin VB.CheckBox chkResSpec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   180
         Left            =   1650
         TabIndex        =   67
         Top             =   3000
         Width           =   705
      End
      Begin VB.TextBox txtRChannel 
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
         Left            =   1650
         TabIndex        =   66
         Top             =   1500
         Width           =   2145
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
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   420
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
         Left            =   1650
         TabIndex        =   64
         Top             =   1860
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
         Left            =   1650
         TabIndex        =   63
         Top             =   2220
         Width           =   2145
      End
      Begin VB.TextBox txtOChannel 
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
         Left            =   1650
         TabIndex        =   62
         Top             =   1140
         Width           =   2145
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
         Left            =   1650
         TabIndex        =   61
         Top             =   2580
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
         Left            =   2460
         TabIndex        =   60
         Top             =   2940
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
         Left            =   1650
         TabIndex        =   59
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txtRefMLow 
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
         Left            =   2100
         TabIndex        =   58
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtRefMHigh 
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
         Left            =   3060
         TabIndex        =   57
         Top             =   3360
         Width           =   735
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
         Left            =   3390
         TabIndex        =   56
         Top             =   750
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
         Left            =   2970
         TabIndex        =   55
         Top             =   750
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
         Left            =   3360
         TabIndex        =   54
         Top             =   2910
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
         Left            =   2910
         TabIndex        =   53
         Top             =   2910
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
         Left            =   2100
         TabIndex        =   52
         Top             =   3720
         Width           =   735
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
         Left            =   3060
         TabIndex        =   51
         Top             =   3720
         Width           =   735
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
         Height          =   375
         Left            =   4020
         TabIndex        =   50
         Top             =   8070
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
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   49
         Top             =   8070
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
         Height          =   375
         Index           =   1
         Left            =   3090
         TabIndex        =   48
         Top             =   8070
         Width           =   915
      End
      Begin VB.CommandButton cmdConfirm 
         BackColor       =   &H00C0FFFF&
         Caption         =   "전체저장"
         Height          =   375
         Index           =   2
         Left            =   210
         Style           =   1  '그래픽
         TabIndex        =   47
         Top             =   8070
         Width           =   945
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 결과 변환 (수치형) "
         Height          =   1815
         Left            =   210
         TabIndex        =   22
         Top             =   4110
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
            TabIndex        =   34
            Top             =   360
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
            TabIndex        =   33
            Top             =   360
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
            TabIndex        =   32
            Top             =   690
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
            TabIndex        =   31
            Top             =   690
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
            TabIndex        =   30
            Top             =   1020
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
            TabIndex        =   29
            Top             =   1020
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
            TabIndex        =   28
            Top             =   1350
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
            TabIndex        =   27
            Top             =   1350
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
            TabIndex        =   26
            Text            =   "<"
            Top             =   360
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
            TabIndex        =   25
            Text            =   "<="
            Top             =   690
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
            TabIndex        =   24
            Text            =   ">"
            Top             =   1020
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
            TabIndex        =   23
            Text            =   ">="
            Top             =   1350
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
            TabIndex        =   46
            Top             =   420
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
            TabIndex        =   45
            Top             =   750
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
            TabIndex        =   44
            Top             =   1080
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
            TabIndex        =   43
            Top             =   1410
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
            TabIndex        =   42
            Top             =   420
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
            TabIndex        =   41
            Top             =   750
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
            TabIndex        =   40
            Top             =   1080
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
            TabIndex        =   39
            Top             =   1410
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
            TabIndex        =   38
            Top             =   1410
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
            TabIndex        =   37
            Top             =   1080
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
            TabIndex        =   36
            Top             =   750
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
            TabIndex        =   35
            Top             =   420
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 결과 변환 (문자형) "
         Height          =   1875
         Left            =   210
         TabIndex        =   5
         Top             =   6030
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
            Left            =   2760
            TabIndex        =   12
            Top             =   300
            Width           =   1395
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
            Left            =   930
            TabIndex        =   11
            Top             =   300
            Width           =   1335
         End
         Begin VB.CheckBox chkAMR 
            BackColor       =   &H00FFFFFF&
            Caption         =   "결과포함표시 [ 예제:Pos 12.0 ]"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   180
            TabIndex        =   10
            Top             =   1410
            Width           =   3465
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
            Left            =   2760
            TabIndex        =   9
            Top             =   630
            Width           =   1395
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
            Left            =   930
            TabIndex        =   8
            Top             =   630
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
            Left            =   2760
            TabIndex        =   7
            Top             =   960
            Width           =   1395
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
            Left            =   930
            TabIndex        =   6
            Top             =   960
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
            TabIndex        =   21
            Top             =   360
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
            Left            =   2340
            TabIndex        =   20
            Top             =   360
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
            Left            =   150
            TabIndex        =   19
            Top             =   360
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
            TabIndex        =   18
            Top             =   690
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
            Left            =   2340
            TabIndex        =   17
            Top             =   690
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
            Left            =   150
            TabIndex        =   16
            Top             =   690
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
            TabIndex        =   15
            Top             =   1020
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
            Left            =   2340
            TabIndex        =   14
            Top             =   1020
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
            Left            =   150
            TabIndex        =   13
            Top             =   1020
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
         Height          =   375
         Index           =   3
         Left            =   1170
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   8070
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   1260
         Left            =   3840
         Picture         =   "frmTestSet.frx":1272
         Top             =   1890
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "남자"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   40
         Left            =   1650
         TabIndex        =   80
         Top             =   3420
         Width           =   360
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
         Left            =   600
         TabIndex        =   79
         Top             =   840
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
         Left            =   600
         TabIndex        =   78
         Top             =   1575
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
         Left            =   600
         TabIndex        =   77
         Top             =   480
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
         Left            =   600
         TabIndex        =   76
         Top             =   1200
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
         Left            =   600
         TabIndex        =   75
         Top             =   1935
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
         Left            =   600
         TabIndex        =   74
         Top             =   2295
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
         Left            =   600
         TabIndex        =   73
         Top             =   2655
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
         Left            =   600
         TabIndex        =   72
         Top             =   3015
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
         Left            =   600
         TabIndex        =   71
         Top             =   3420
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "여자"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1650
         TabIndex        =   70
         Top             =   3810
         Width           =   360
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
         Left            =   2880
         TabIndex        =   69
         Top             =   3450
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
         Left            =   2880
         TabIndex        =   68
         Top             =   3780
         Width           =   135
      End
   End
   Begin FPSpread.vaSpread spdTest 
      Height          =   8745
      Left            =   60
      TabIndex        =   0
      Top             =   630
      Width           =   10845
      _Version        =   393216
      _ExtentX        =   19129
      _ExtentY        =   15425
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
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   21
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
        
            .Add "AMRINRESULT", chkAMR.Value
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
                SQL = SQL & " ,TESTCODE, TESTNAME, ABBRNAME, RESPRECUSE, RESPREC"
                SQL = SQL & " , REFMLOW, REFMHIGH, REFFLOW, REFFHIGH" & vbCrLf
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
                '-- AMR
                strTmp = Trim(GetText(spdTest, i, colLFHIGH + 1))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colLFHIGH + 2))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colLFHIGH + 3))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colLFHIGH + 4))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colLFHIGH + 5))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colLFHIGH + 6))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                strTmp = Trim(GetText(spdTest, i, colLFHIGH + 7))
                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
                SQL = SQL & "," & STS(GetText(spdTest, i, colLFHIGH + 8))
                SQL = SQL & ")" & vbCrLf
                
                Call DBExec(AdoCn_Local, SQL)
            
            Next
        End With
        
        Call GetTestList
        Call GetTestMaster(spdTest)
        
    ElseIf Index = 3 Then
        
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call GetTestMaster(spdTest)
    
End Sub


Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        Exit Sub
    End If

    With spdTest
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
        
        'AMR
        txtAMRLimit(0).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 1), 1, "|")
        txtAMRResult(0).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 1), 2, "|")
    
        txtAMRLimit(1).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 2), 1, "|")
        txtAMRResult(1).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 2), 2, "|")
    
        txtAMRLimit(2).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 3), 1, "|")
        txtAMRResult(2).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 3), 2, "|")
    
        txtAMRLimit(3).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 4), 1, "|")
        txtAMRResult(3).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 4), 2, "|")
    
        txtAMRLimit(4).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 5), 1, "|")
        txtAMRResult(4).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 5), 2, "|")
    
        txtAMRLimit(5).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 6), 1, "|")
        txtAMRResult(5).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 6), 2, "|")
    
        txtAMRLimit(6).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 7), 1, "|")
        txtAMRResult(6).Text = mGetP(GetText(spdTest, Row, colLFHIGH + 7), 2, "|")
    
        If GetText(spdTest, Row, colLFHIGH + 8) = "" Then
            chkAMR.Value = "0"
        Else
            chkAMR.Value = GetText(spdTest, Row, colLFHIGH + 8)
        End If
    End With
    
    'txtTestCd.SetFocus
End Sub

