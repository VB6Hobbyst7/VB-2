VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmScreenSet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "화면설정"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   Icon            =   "frmScreenSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   9405
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtRowHeight 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   7230
      TabIndex        =   82
      Top             =   6900
      Width           =   1485
   End
   Begin HSCotrol.HSLabel HSLabel1 
      Height          =   345
      Left            =   150
      Top             =   150
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   609
      BackColor       =   16311496
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ▶ 화면설정"
      BevelOut        =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   4515
      Left            =   4920
      TabIndex        =   53
      Top             =   1020
      Width           =   4095
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   150
         TabIndex        =   81
         Top             =   4080
         Width           =   2235
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   2430
         TabIndex        =   80
         Top             =   4060
         Width           =   1515
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FFF56E&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   150
         TabIndex        =   79
         Top             =   3750
         Width           =   2235
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   2430
         TabIndex        =   78
         Top             =   3730
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   2430
         TabIndex        =   77
         Top             =   3400
         Width           =   1515
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FFF56E&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   150
         TabIndex        =   76
         Top             =   3420
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   150
         TabIndex        =   75
         Top             =   3090
         Width           =   2235
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   2430
         TabIndex        =   74
         Top             =   3070
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   2430
         TabIndex        =   71
         Top             =   2740
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   2430
         TabIndex        =   70
         Top             =   2410
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   2430
         TabIndex        =   69
         Top             =   2080
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2430
         TabIndex        =   68
         Top             =   1750
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2430
         TabIndex        =   67
         Top             =   1420
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2430
         TabIndex        =   66
         Top             =   1090
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2430
         TabIndex        =   65
         Top             =   760
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2430
         TabIndex        =   64
         Top             =   430
         Width           =   1515
      End
      Begin VB.TextBox txtColumnResult 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2430
         TabIndex        =   63
         Text            =   "10"
         Top             =   100
         Width           =   1515
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   150
         TabIndex        =   62
         Top             =   2760
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   150
         TabIndex        =   61
         Top             =   2430
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   60
         Top             =   2100
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FAFA96&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   150
         TabIndex        =   59
         Top             =   1770
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FAFA96&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   150
         TabIndex        =   58
         Top             =   1440
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FAFAB4&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   57
         Top             =   1110
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FAFAB4&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   56
         Top             =   780
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FAFAD2&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   55
         Top             =   450
         Width           =   2235
      End
      Begin VB.CheckBox chkColumnResult 
         BackColor       =   &H00FAFAD2&
         Caption         =   "선택"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   54
         Top             =   120
         Width           =   2235
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   4395
         Left            =   90
         Top             =   60
         Width           =   3915
      End
   End
   Begin VB.TextBox txtColWidth 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   7230
      TabIndex        =   43
      Top             =   6510
      Width           =   1485
   End
   Begin VB.Frame fraView 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   7155
      Left            =   270
      TabIndex        =   0
      Top             =   1020
      Width           =   4095
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         Left            =   2430
         TabIndex        =   42
         Top             =   6700
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAD2&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   150
         TabIndex        =   41
         Top             =   6720
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         Left            =   2430
         TabIndex        =   40
         Top             =   6370
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAD2&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   150
         TabIndex        =   39
         Top             =   6390
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         Left            =   2430
         TabIndex        =   38
         Top             =   6040
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAB4&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   150
         TabIndex        =   37
         Top             =   6060
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAD2&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   120
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAD2&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   35
         Top             =   450
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAB4&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   34
         Top             =   780
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAB4&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   33
         Top             =   1110
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFA96&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   150
         TabIndex        =   32
         Top             =   1440
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFA96&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   150
         TabIndex        =   31
         Top             =   1770
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   30
         Top             =   2100
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   150
         TabIndex        =   29
         Top             =   2430
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   150
         TabIndex        =   28
         Top             =   2760
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   150
         TabIndex        =   27
         Top             =   3090
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF56E&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   150
         TabIndex        =   26
         Top             =   3420
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF56E&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   150
         TabIndex        =   25
         Top             =   3750
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   150
         TabIndex        =   24
         Top             =   4080
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   150
         TabIndex        =   23
         Top             =   4410
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   150
         TabIndex        =   22
         Top             =   4740
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFA96&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   150
         TabIndex        =   21
         Top             =   5070
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2430
         TabIndex        =   20
         Text            =   "10"
         Top             =   100
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2430
         TabIndex        =   19
         Top             =   430
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2430
         TabIndex        =   18
         Top             =   760
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2430
         TabIndex        =   17
         Top             =   1090
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2430
         TabIndex        =   16
         Top             =   1420
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2430
         TabIndex        =   15
         Top             =   1750
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   2430
         TabIndex        =   14
         Text            =   "6"
         Top             =   2080
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   2430
         TabIndex        =   13
         Text            =   "10"
         Top             =   2410
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   2430
         TabIndex        =   12
         Top             =   2740
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   2430
         TabIndex        =   11
         Top             =   3070
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   2430
         TabIndex        =   10
         Top             =   3400
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   2430
         TabIndex        =   9
         Top             =   3730
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   2430
         TabIndex        =   8
         Top             =   4060
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   2430
         TabIndex        =   7
         Top             =   4390
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   2430
         TabIndex        =   6
         Top             =   4720
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   2430
         TabIndex        =   5
         Top             =   5050
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFA96&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   150
         TabIndex        =   4
         Top             =   5400
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAB4&
         Caption         =   "저장순번"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   150
         TabIndex        =   3
         Top             =   5730
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   2430
         TabIndex        =   2
         Top             =   5380
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         Left            =   2430
         TabIndex        =   1
         Top             =   5710
         Width           =   1515
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   7035
         Left            =   90
         Top             =   60
         Width           =   3915
      End
   End
   Begin HSCotrol.CButton cmdSave 
      Height          =   495
      Left            =   5940
      TabIndex        =   46
      Top             =   7650
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " 설정저장"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmScreenSet.frx":08A8
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.CButton cmdCancel 
      Height          =   495
      Left            =   7410
      TabIndex        =   47
      Top             =   7650
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " 닫    기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmScreenSet.frx":0A02
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H00BF8B59&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   4500
      TabIndex        =   48
      Top             =   7650
      Visible         =   0   'False
      Width           =   2235
      Begin VB.TextBox txtTop 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   52
         Text            =   "상단색"
         Top             =   150
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdTop 
         Caption         =   "Set"
         Height          =   285
         Left            =   1020
         TabIndex        =   51
         Top             =   180
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdBottom 
         Caption         =   "Set"
         Height          =   285
         Left            =   1020
         TabIndex        =   50
         Top             =   510
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtBottom 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   49
         Text            =   "하단색"
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1560
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사항목 높이"
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   3
      Left            =   5880
      TabIndex        =   83
      Top             =   6990
      Width           =   1140
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   405
      Left            =   120
      Top             =   120
      Width           =   9195
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "[ 메인 화면 ]"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   2
      Left            =   390
      TabIndex        =   73
      Top             =   750
      Width           =   1245
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "[ 상세결과 화면 ]"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   1
      Left            =   5040
      TabIndex        =   72
      Top             =   720
      Width           =   1605
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "※ 체크한 항목만 화면에 보임"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   0
      Left            =   5160
      TabIndex        =   45
      Top             =   5760
      Width           =   2640
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사항목 넓이"
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   7
      Left            =   5880
      TabIndex        =   44
      Top             =   6600
      Width           =   1140
   End
End
Attribute VB_Name = "frmScreenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBottom_Click()
    Dim LetColor
    Dim GetColor
    
    '선택하기전 색을 갖고 있는다.
    LetColor = txtBottom.BackColor
    
    CommonDialog1.ShowColor
    
    '선택한 색이 없다면
    If CommonDialog1.Color = 0 Then
        txtBottom.BackColor = LetColor
        GetColor = LetColor
    Else
        GetColor = CommonDialog1.Color
        txtBottom.BackColor = GetColor
    End If
    
    'Call WritePrivateProfileString("VIEW", "BOTTOMCOLOR", CStr(GetColor), App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "BOTTOMCOLOR", CStr(GetColor))

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i          As Integer
    Dim strSPDView As String
    Dim strSPDSize As String
    
    '-- 메인화면
    strSPDView = ""
    strSPDSize = ""
    
    For i = 0 To 20
        strSPDView = strSPDView & IIf(chkColumn(i).Value = "1", "1", "0")
        strSPDSize = strSPDSize & txtColumn(i).Text & "|"
    Next
    
    'Call WritePrivateProfileString("VIEW", "SPDVIEW", strSPDView, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("VIEW", "SPDSIZE", strSPDSize, App.PATH & "\INI\" & gMACH & ".ini")

    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDVIEW", strSPDView)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDSIZE", strSPDSize)

    '-- 상세결과
    strSPDView = ""
    strSPDSize = ""
    
    For i = 0 To 12
        strSPDView = strSPDView & IIf(chkColumnResult(i).Value = "1", "1", "0")
        strSPDSize = strSPDSize & txtColumnResult(i).Text & "|"
    Next
    
    'Call WritePrivateProfileString("VIEW", "SPDVIEW_R", strSPDView, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("VIEW", "SPDSIZE_R", strSPDSize, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("VIEW", "COLWIDTH", txtColWidth.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDVIEW_R", strSPDView)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "SPDSIZE_R", strSPDSize)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "COLWIDTH", txtColWidth.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "ROWHEIGHT", txtRowHeight.Text)
    
    MsgBox "컬럼정보가 변경되었습니다.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub cmdTop_Click()
    Dim LetColor
    Dim GetColor
    
    '선택하기전 색을 갖고 있는다.
    LetColor = txtTop.BackColor
    
    CommonDialog1.ShowColor
    
    '선택한 색이 없다면
    If CommonDialog1.Color = 0 Then
        txtTop.BackColor = LetColor
        GetColor = LetColor
    Else
        GetColor = CommonDialog1.Color
        txtTop.BackColor = GetColor
    End If
    
    'Call WritePrivateProfileString("VIEW", "TOPCOLOR", CStr(GetColor), App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub Form_Load()

    '-- 화면설정
    Call SetColumnName
    
    Call SetColumnName_Result
        
    txtColWidth.Text = gCOLWIDTH
    txtRowHeight.Text = gROWHEIGHT

    txtTop.BackColor = frmInterface.picHeader.BackColor
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub SetColumnName()
    Dim i       As Integer
    Dim varView As Variant
    Dim varSize As Variant
    
    chkColumn(0).Caption = "선택"
    chkColumn(1).Caption = "검사일자"
    chkColumn(2).Caption = "검사시간"
    chkColumn(3).Caption = "검사순번"
    chkColumn(4).Caption = "ER"
    chkColumn(5).Caption = "RT"
    chkColumn(6).Caption = "접수일자"
    chkColumn(7).Caption = "검체번호"
    chkColumn(8).Caption = "검체"
    chkColumn(9).Caption = "RackNo"
    chkColumn(10).Caption = "TubePos"
    chkColumn(11).Caption = "SeqNo"
    chkColumn(12).Caption = "이름"
    chkColumn(13).Caption = "Sex"
    chkColumn(14).Caption = "Age"
    chkColumn(15).Caption = "병록번호"
    chkColumn(16).Caption = "챠트번호"
    chkColumn(17).Caption = "의뢰과"
    chkColumn(18).Caption = "입/외구분"
    chkColumn(19).Caption = "오더갯수"
    chkColumn(20).Caption = "결과갯수"
    
    For i = 0 To 20
        chkColumn(i).Value = Mid(gCOLVIEW, i + 1, 1)
        txtColumn(i).BackColor = chkColumn(i).BackColor
        chkColumn(i).Font = "굴림"
        txtColumn(i).Font = "굴림"
    Next
    
    varSize = Split(gCOLSIZE, "|")
    
    For i = 0 To 20
        txtColumn(i).Alignment = 2
        txtColumn(i).Text = varSize(i)
        txtColumn(i).FontSize = 11
        If Mid(gCOLVIEW, i + 1, 1) = "1" Then
            txtColumn(i).FontBold = False
        Else
            txtColumn(i).FontBold = False
        End If
    Next

End Sub

Private Sub SetColumnName_Result()
    Dim i       As Integer
    Dim varView As Variant
    Dim varSize As Variant
    
    chkColumnResult(0).Caption = "선택"
    chkColumnResult(1).Caption = "번호"
    chkColumnResult(2).Caption = "처방코드"
    chkColumnResult(3).Caption = "검사코드"
    chkColumnResult(4).Caption = "SUB코드"
    chkColumnResult(5).Caption = "검사명(약어)"
    chkColumnResult(6).Caption = "장비채널"
    chkColumnResult(7).Caption = "장비결과"
    chkColumnResult(8).Caption = "검사결과"
    chkColumnResult(9).Caption = "FLAG"
    chkColumnResult(10).Caption = "판정"
    chkColumnResult(11).Caption = "참고치"
    chkColumnResult(12).Caption = "이전결과"
    
    For i = 0 To 12
        chkColumnResult(i).Value = Mid(gCOLVIEW_R, i + 1, 1)
        txtColumnResult(i).BackColor = chkColumnResult(i).BackColor
        chkColumnResult(i).Font = "굴림"
        txtColumnResult(i).Font = "굴림"
    Next
    
    varSize = Split(gCOLSIZE_R, "|")
    
    For i = 0 To 12
        txtColumnResult(i).Alignment = 2
        txtColumnResult(i).Text = varSize(i)
        txtColumnResult(i).FontSize = 11
        If Mid(gCOLVIEW_R, i + 1, 1) = "1" Then
            txtColumnResult(i).FontBold = False
        Else
            txtColumnResult(i).FontBold = False
        End If
    Next

End Sub
