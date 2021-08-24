VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "검사설정"
   ClientHeight    =   11385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   Icon            =   "frmTestSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11385
   ScaleWidth      =   19755
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraHidden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hidden"
      Height          =   6405
      Left            =   19140
      TabIndex        =   31
      Top             =   5130
      Visible         =   0   'False
      Width           =   12435
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   5625
         Left            =   5910
         TabIndex        =   45
         Top             =   510
         Width           =   5895
         Begin VB.TextBox txtRstApi 
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
            Height          =   330
            Left            =   1320
            TabIndex        =   55
            Top             =   4950
            Width           =   4365
         End
         Begin VB.TextBox txtURL 
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
            Height          =   330
            Left            =   1320
            TabIndex        =   54
            Top             =   4170
            Width           =   4365
         End
         Begin VB.TextBox txtOrdApi 
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
            Height          =   330
            Left            =   1320
            TabIndex        =   53
            Top             =   4560
            Width           =   4365
         End
         Begin VB.TextBox txtIN1 
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
            Left            =   1320
            TabIndex        =   52
            Top             =   1410
            Width           =   4335
         End
         Begin VB.TextBox txtIN2 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFC0&
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
            Left            =   1320
            TabIndex        =   51
            Top             =   1740
            Width           =   4335
         End
         Begin VB.TextBox txtFD1 
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
            Left            =   1320
            TabIndex        =   50
            Top             =   690
            Width           =   4335
         End
         Begin VB.TextBox txtFD2 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFC0&
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
            Left            =   1320
            TabIndex        =   49
            Top             =   1020
            Width           =   4335
         End
         Begin VB.TextBox txtResult 
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
            Height          =   330
            Left            =   1320
            TabIndex        =   48
            Top             =   3480
            Width           =   4365
         End
         Begin VB.TextBox txtOrder 
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
            Height          =   330
            Left            =   1320
            TabIndex        =   47
            Top             =   3090
            Width           =   4365
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3960
            Picture         =   "frmTestSet.frx":1272
            ScaleHeight     =   300
            ScaleWidth      =   300
            TabIndex        =   46
            Top             =   2730
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "결과 API"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   450
            TabIndex        =   67
            Top             =   5025
            Width           =   720
         End
         Begin VB.Shape Shape11 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   3
            Height          =   315
            Left            =   210
            Top             =   4950
            Width           =   1095
         End
         Begin VB.Label Label27 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "URL"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   450
            TabIndex        =   66
            Top             =   4245
            Width           =   810
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "오더 API"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   450
            TabIndex        =   65
            Top             =   4635
            Width           =   720
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "[URL 설정]"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   270
            TabIndex        =   64
            Top             =   3900
            Width           =   1020
         End
         Begin VB.Shape Shape10 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   3
            Height          =   315
            Left            =   210
            Top             =   4170
            Width           =   1095
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   3
            Height          =   315
            Left            =   210
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   3
            Height          =   315
            Left            =   210
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   3
            Height          =   315
            Left            =   210
            Top             =   3090
            Width           =   1095
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   3
            Height          =   615
            Left            =   180
            Top             =   1410
            Width           =   1095
         End
         Begin VB.Label Label9 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Inhalant"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   300
            TabIndex        =   63
            Top             =   1605
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   3
            Height          =   615
            Left            =   180
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label13 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "Food"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   300
            TabIndex        =   62
            Top             =   885
            Width           =   360
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "[Assay명 설정]"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   270
            TabIndex        =   61
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "[경로 설정]"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   270
            TabIndex        =   60
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "결과경로"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   450
            TabIndex        =   59
            Top             =   3555
            Width           =   720
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "오더경로"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   450
            TabIndex        =   58
            Top             =   3165
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "* 대소문자 정확히 입력하세요"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Left            =   1770
            TabIndex        =   57
            Top             =   390
            Width           =   2520
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "* 경로끝의 \는 제외"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Left            =   1830
            TabIndex        =   56
            Top             =   2790
            Width           =   1710
         End
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
         Left            =   3360
         TabIndex        =   35
         Top             =   330
         Width           =   1275
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
         Left            =   3360
         TabIndex        =   34
         Top             =   750
         Width           =   1275
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
         Left            =   3360
         TabIndex        =   33
         Top             =   1170
         Width           =   1275
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
         Left            =   3360
         TabIndex        =   32
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "High(M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   41
         Left            =   2550
         TabIndex        =   40
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Low(M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   40
         Left            =   2550
         TabIndex        =   39
         Top             =   420
         Width           =   675
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
         Left            =   1470
         TabIndex        =   38
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "High(F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   2550
         TabIndex        =   37
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Low(F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2550
         TabIndex        =   36
         Top             =   1260
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   240
      TabIndex        =   26
      Top             =   180
      Width           =   19425
      Begin VB.OptionButton optGubun 
         BackColor       =   &H00FFFFFF&
         Caption         =   "전체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Index           =   0
         Left            =   1860
         Style           =   1  '그래픽
         TabIndex        =   29
         Top             =   210
         Value           =   -1  'True
         Width           =   1650
      End
      Begin VB.OptionButton optGubun 
         BackColor       =   &H00C0FFFF&
         Caption         =   "FOOD"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   3555
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   165
         Width           =   1650
      End
      Begin VB.OptionButton optGubun 
         BackColor       =   &H00C0C0FF&
         Caption         =   "INHALANT"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   5220
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   165
         Width           =   1650
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "검사구분 선택 : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   30
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.Frame frameTestSet 
      BackColor       =   &H00FFFFFF&
      Height          =   10335
      Left            =   13740
      TabIndex        =   16
      Top             =   930
      Width           =   5925
      Begin VB.CheckBox chkType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용'"
         Height          =   390
         Left            =   1650
         TabIndex        =   43
         Top             =   4080
         Width           =   1305
      End
      Begin VB.ComboBox cboGubun 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmTestSet.frx":13BC
         Left            =   1650
         List            =   "frmTestSet.frx":13BE
         Style           =   2  '드롭다운 목록
         TabIndex        =   41
         Top             =   930
         Width           =   2145
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2790
         TabIndex        =   25
         Top             =   4650
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1650
         TabIndex        =   14
         Top             =   4650
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "취소"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3930
         TabIndex        =   15
         Top             =   4650
         Width           =   1095
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
         Left            =   3810
         TabIndex        =   12
         Top             =   3660
         Width           =   435
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
         Left            =   4260
         TabIndex        =   13
         Top             =   3660
         Width           =   435
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
         Left            =   2910
         TabIndex        =   3
         Top             =   1320
         Width           =   405
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
         Left            =   3330
         TabIndex        =   4
         Top             =   1320
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
         TabIndex        =   2
         Top             =   1350
         Width           =   1245
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
         Left            =   2940
         TabIndex        =   11
         Top             =   3690
         Width           =   825
      End
      Begin VB.TextBox txtAbbrNm 
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
         Left            =   1650
         TabIndex        =   9
         Top             =   3300
         Width           =   4100
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
         TabIndex        =   5
         Top             =   1740
         Width           =   2115
      End
      Begin VB.TextBox txtTestNm 
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
         Left            =   1650
         TabIndex        =   8
         Top             =   2910
         Width           =   4100
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
         TabIndex        =   7
         Top             =   2520
         Width           =   2115
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
         TabIndex        =   1
         Top             =   480
         Width           =   2115
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
         TabIndex        =   6
         Top             =   2130
         Width           =   2115
      End
      Begin VB.CheckBox chkResSpec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "소수점사용"
         Height          =   390
         Left            =   1650
         TabIndex        =   10
         Top             =   3660
         Width           =   1305
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "수기입력"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   600
         TabIndex        =   44
         Top             =   4185
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사구분"
         Height          =   180
         Left            =   600
         TabIndex        =   42
         Top             =   1005
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
         TabIndex        =   24
         Top             =   3765
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
         TabIndex        =   23
         Top             =   3405
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
         TabIndex        =   22
         Top             =   2985
         Width           =   540
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
         TabIndex        =   21
         Top             =   2625
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
         TabIndex        =   20
         Top             =   1860
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
         TabIndex        =   19
         Top             =   540
         Width           =   720
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
         TabIndex        =   18
         Top             =   2265
         Width           =   720
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
         TabIndex        =   17
         Top             =   1410
         Width           =   360
      End
   End
   Begin FPSpread.vaSpread spdTest 
      Height          =   10215
      Left            =   270
      TabIndex        =   0
      Top             =   1020
      Width           =   13425
      _Version        =   393216
      _ExtentX        =   23680
      _ExtentY        =   18018
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
      MaxCols         =   15
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestSet.frx":13C0
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
    Dim Test_Property As Scripting.Dictionary
    Dim objTest_Property As clsCommon

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
            .Add "GUBUN", cboGubun.Text
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
            .Add "GUBUN", cboGubun.Text
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
            .Add "EXAMTYPE", IIf(chkType.Value = "1", "수기", "")
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
        Call GetTestMaster(spdTest, cboGubun.ListIndex)
        

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

Private Sub Form_Load()

    cboGubun.Clear
    cboGubun.AddItem "FOOD"
    cboGubun.AddItem "INHALANT"
    cboGubun.ListIndex = 0
    
    Call GetTestMaster(spdTest)
    
End Sub


Private Sub optGubun_Click(Index As Integer)
    
    Call GetTestMaster(spdTest, Index)
    
End Sub

Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        Exit Sub
    End If

    With spdTest
        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
        cboGubun.Text = GetText(spdTest, Row, colLGUBUN)
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

        If GetText(spdTest, Row, colLTYPE) = "" Then
            chkType.Value = "0"
        Else
            chkType.Value = "1"
        End If
'        txtRefMLow.Text = GetText(spdTest, Row, colLMLOW)
'        txtRefMHigh.Text = GetText(spdTest, Row, colLMHIGH)
'        txtRefFLow.Text = GetText(spdTest, Row, colLFLOW)
'        txtRefFHigh.Text = GetText(spdTest, Row, colLFHIGH)
    End With
    
End Sub
