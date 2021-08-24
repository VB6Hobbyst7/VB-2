VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmTestSet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "검사설정"
   ClientHeight    =   10065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17385
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   17385
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame frameTestSet 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   9885
      Left            =   10590
      TabIndex        =   35
      Top             =   60
      Width           =   6705
      Begin VB.TextBox txtEqpCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   240
         Width           =   585
      End
      Begin VB.Frame fraResultTrans 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4395
         Left            =   120
         TabIndex        =   36
         Top             =   5400
         Width           =   6375
         Begin VB.Frame fraN 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            ForeColor       =   &H80000008&
            Height          =   2085
            Left            =   60
            TabIndex        =   37
            Top             =   0
            Width           =   6255
            Begin VB.TextBox txtAMRNumLimit 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
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
               Left            =   360
               TabIndex        =   72
               Top             =   1740
               Width           =   2535
            End
            Begin VB.TextBox txtAMRNumResult 
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
               Index           =   4
               Left            =   3330
               TabIndex        =   71
               Top             =   1740
               Width           =   2865
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   70
               Text            =   "<>"
               Top             =   1740
               Width           =   345
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   41
               Text            =   ">="
               Top             =   1410
               Width           =   345
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   40
               Text            =   ">"
               Top             =   1080
               Width           =   345
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   39
               Text            =   "<="
               Top             =   750
               Width           =   345
            End
            Begin VB.TextBox TtxtCmp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   0
               Locked          =   -1  'True
               TabIndex        =   38
               Text            =   "<"
               Top             =   420
               Width           =   345
            End
            Begin VB.TextBox txtAMRNumResult 
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
               Index           =   3
               Left            =   3330
               TabIndex        =   26
               Top             =   1410
               Width           =   2865
            End
            Begin VB.TextBox txtAMRNumLimit 
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
               Left            =   360
               TabIndex        =   25
               Top             =   1410
               Width           =   2535
            End
            Begin VB.TextBox txtAMRNumResult 
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
               Index           =   2
               Left            =   3330
               TabIndex        =   24
               Top             =   1080
               Width           =   2865
            End
            Begin VB.TextBox txtAMRNumLimit 
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
               Left            =   360
               TabIndex        =   23
               Top             =   1080
               Width           =   2535
            End
            Begin VB.TextBox txtAMRNumResult 
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
               Index           =   1
               Left            =   3330
               TabIndex        =   22
               Top             =   750
               Width           =   2865
            End
            Begin VB.TextBox txtAMRNumLimit 
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
               Left            =   360
               TabIndex        =   21
               Top             =   750
               Width           =   2535
            End
            Begin VB.TextBox txtAMRNumResult 
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
               Left            =   3330
               TabIndex        =   20
               Top             =   420
               Width           =   2865
            End
            Begin VB.TextBox txtAMRNumLimit 
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
               Index           =   0
               Left            =   360
               TabIndex        =   19
               Top             =   420
               Width           =   2535
            End
            Begin HSCotrol.HSLabel HSLabel24 
               Height          =   315
               Left            =   0
               Top             =   90
               Width           =   6285
               _ExtentX        =   11086
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "수치형 결과변환"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   7
               Left            =   2910
               Top             =   420
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   8
               Left            =   2910
               Top             =   750
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   9
               Left            =   2910
               Top             =   1080
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   10
               Left            =   2910
               Top             =   1410
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   11
               Left            =   2910
               Top             =   1740
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
         End
         Begin VB.Frame fraC 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   60
            TabIndex        =   42
            Top             =   2100
            Width           =   6255
            Begin VB.TextBox txtAMRCharResult 
               Appearance      =   0  '평면
               BackColor       =   &H80000018&
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
               Left            =   3330
               TabIndex        =   74
               Top             =   1770
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharLimit 
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
               Left            =   0
               TabIndex        =   73
               Top             =   1770
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharResult 
               Appearance      =   0  '평면
               BackColor       =   &H80000018&
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
               Left            =   3330
               TabIndex        =   34
               Top             =   1440
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharLimit 
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
               Left            =   0
               TabIndex        =   33
               Top             =   1440
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharLimit 
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
               Left            =   0
               TabIndex        =   31
               Top             =   1110
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharResult 
               Appearance      =   0  '평면
               BackColor       =   &H80000018&
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
               Left            =   3330
               TabIndex        =   32
               Top             =   1110
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharLimit 
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
               Left            =   0
               TabIndex        =   29
               Top             =   780
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharResult 
               Appearance      =   0  '평면
               BackColor       =   &H80000018&
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
               Left            =   3330
               TabIndex        =   30
               Top             =   780
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharLimit 
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
               Index           =   0
               Left            =   0
               TabIndex        =   27
               Top             =   450
               Width           =   2895
            End
            Begin VB.TextBox txtAMRCharResult 
               Appearance      =   0  '평면
               BackColor       =   &H80000018&
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
               Left            =   3330
               TabIndex        =   28
               Top             =   450
               Width           =   2895
            End
            Begin HSCotrol.HSLabel HSLabel20 
               Height          =   315
               Left            =   0
               Top             =   120
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   556
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "문자형 결과변환"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   0
               Left            =   2910
               Top             =   450
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   1
               Left            =   2910
               Top             =   780
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   2
               Left            =   2910
               Top             =   1110
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   3
               Left            =   2910
               Top             =   1440
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
            Begin HSCotrol.HSLabel HSLabel1 
               Height          =   345
               Index           =   4
               Left            =   2910
               Top             =   1770
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               BackColor       =   8421504
               ForeColor       =   12648447
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "▶"
               BevelOut        =   0
               Alignment       =   2
            End
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         ForeColor       =   &H00FFFFFF&
         Height          =   4755
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   6375
         Begin VB.CheckBox chkCalTest 
            BackColor       =   &H00FFFFFF&
            Caption         =   "계산식 적용검사"
            Height          =   270
            Left            =   4650
            TabIndex        =   48
            Top             =   180
            Width           =   1665
         End
         Begin VB.Frame fraNC 
            BackColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   1110
            TabIndex        =   47
            Top             =   4230
            Width           =   4335
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "수치(판정)"
               Height          =   255
               Index           =   2
               Left            =   2910
               TabIndex        =   18
               Top             =   150
               Width           =   1185
            End
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "판정(수치)"
               Height          =   255
               Index           =   1
               Left            =   1500
               TabIndex        =   17
               Top             =   150
               Width           =   1185
            End
            Begin VB.OptionButton optINQuant 
               BackColor       =   &H00FFFFFF&
               Caption         =   "변환없음"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   16
               Top             =   150
               Value           =   -1  'True
               Width           =   1035
            End
         End
         Begin HSCotrol.HSLabel HSLabel10 
            Height          =   405
            Left            =   1110
            Top             =   3030
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   714
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "여성 참고치(F)"
            BevelOut        =   0
            Begin VB.TextBox txtRefFLow 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               TabIndex        =   9
               Top             =   30
               Width           =   1500
            End
            Begin VB.TextBox txtRefFHigh 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3420
               TabIndex        =   10
               Top             =   30
               Width           =   1500
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
               Left            =   3240
               TabIndex        =   46
               Top             =   135
               Width           =   135
            End
         End
         Begin HSCotrol.HSLabel HSLabel9 
            Height          =   405
            Left            =   1110
            Top             =   2640
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   714
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "남성 참고치(M)"
            BevelOut        =   0
            Begin VB.TextBox txtRefMLow 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               TabIndex        =   7
               Top             =   30
               Width           =   1500
            End
            Begin VB.TextBox txtRefMHigh 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3420
               TabIndex        =   8
               Top             =   30
               Width           =   1500
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
               Left            =   3240
               TabIndex        =   45
               Top             =   150
               Width           =   135
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   1110
            TabIndex        =   44
            Top             =   3810
            Width           =   4335
            Begin VB.OptionButton optResType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "수치결과"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   240
               TabIndex        =   13
               Top             =   150
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton optResType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "문자결과"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   1500
               TabIndex        =   14
               Top             =   150
               Width           =   1185
            End
            Begin VB.OptionButton optResType 
               BackColor       =   &H00FFFFFF&
               Caption         =   "혼합결과"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   2910
               TabIndex        =   15
               Top             =   150
               Visible         =   0   'False
               Width           =   1125
            End
         End
         Begin HSCotrol.HSLabel HSLabel2 
            Height          =   330
            Left            =   60
            Top             =   150
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777215
            ForeColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "순번"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin VB.TextBox txtSeq 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1110
            TabIndex        =   0
            Top             =   150
            Width           =   1035
         End
         Begin VB.TextBox txtResSpec 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3990
            TabIndex        =   12
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox txtAbbrNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4410
            TabIndex        =   2
            Top             =   520
            Width           =   1905
         End
         Begin VB.TextBox txtOChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1110
            TabIndex        =   3
            Top             =   890
            Width           =   1905
         End
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1110
            TabIndex        =   1
            Top             =   520
            Width           =   1905
         End
         Begin VB.TextBox txtTestCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1110
            TabIndex        =   5
            Top             =   1240
            Width           =   1905
         End
         Begin VB.TextBox txtRChannel 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4410
            TabIndex        =   4
            Top             =   890
            Width           =   1905
         End
         Begin VB.CheckBox chkResSpec 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00FFFFFF&
            Caption         =   "소수점 변환사용"
            Height          =   180
            Left            =   1110
            TabIndex        =   11
            Top             =   3540
            Width           =   1665
         End
         Begin VB.ListBox lstTestCode 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   4410
            TabIndex        =   6
            Top             =   1245
            Width           =   1905
         End
         Begin HSCotrol.HSLabel HSLabel3 
            Height          =   330
            Left            =   60
            Top             =   870
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "오더채널"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel4 
            Height          =   330
            Left            =   3360
            Top             =   870
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "결과채널"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel5 
            Height          =   330
            Left            =   60
            Top             =   510
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "검사명"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel6 
            Height          =   330
            Left            =   3360
            Top             =   510
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "검사약어"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel7 
            Height          =   330
            Left            =   60
            Top             =   1240
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "검사코드"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel8 
            Height          =   795
            Left            =   60
            Top             =   2640
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   1402
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "참고치"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel11 
            Height          =   345
            Left            =   60
            Top             =   3480
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "소수점"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel12 
            Height          =   345
            Left            =   2940
            Top             =   3480
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "자릿수"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel14 
            Height          =   345
            Left            =   60
            Top             =   3900
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "결과형태"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel23 
            Height          =   1170
            Left            =   3360
            Top             =   1245
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   2064
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "검사코드'S"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel lblNC 
            Height          =   345
            Left            =   60
            Top             =   4320
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "결과표기"
            BevelOut        =   0
            Alignment       =   2
         End
         Begin HSCotrol.CButton cmdSpecUP 
            Height          =   315
            Left            =   4650
            TabIndex        =   63
            Top             =   3480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   33023
            Caption         =   "▲"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin HSCotrol.CButton cmdSpecDown 
            Height          =   315
            Left            =   5040
            TabIndex        =   64
            Top             =   3480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   16744576
            Caption         =   "▼"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin HSCotrol.CButton cmdSeqUp 
            Height          =   315
            Left            =   2190
            TabIndex        =   65
            Top             =   150
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   33023
            Caption         =   "▲"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin HSCotrol.CButton cmdSeqDown 
            Height          =   315
            Left            =   2580
            TabIndex        =   66
            Top             =   150
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   16744576
            Caption         =   "▼"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin HSCotrol.CButton cmdAdd 
            Height          =   315
            Left            =   1140
            TabIndex        =   67
            Top             =   1620
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            BackColor       =   33023
            Caption         =   "Add"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin HSCotrol.CButton cmdRemove 
            Height          =   315
            Left            =   2100
            TabIndex        =   68
            Top             =   1620
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            BackColor       =   16744576
            Caption         =   "Remove"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   4210752
            HoverColor      =   -2147483630
         End
         Begin VB.Image Image1 
            Height          =   1260
            Left            =   5520
            Picture         =   "frmTestSet.frx":1272
            Top             =   3540
            Width           =   705
         End
      End
      Begin HSCotrol.CButton cmdConfirm 
         Height          =   375
         Index           =   8
         Left            =   1650
         TabIndex        =   59
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "신규(New)"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
      End
      Begin HSCotrol.CButton cmdExit 
         Height          =   375
         Left            =   5250
         TabIndex        =   60
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         BackColor       =   14737632
         Caption         =   "닫기(Close)"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
      End
      Begin HSCotrol.CButton cmdConfirm 
         Height          =   375
         Index           =   0
         Left            =   2850
         TabIndex        =   61
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         BackColor       =   12640511
         Caption         =   "저장(Save)"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   0
      End
      Begin HSCotrol.CButton cmdConfirm 
         Height          =   375
         Index           =   1
         Left            =   4050
         TabIndex        =   62
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         BackColor       =   16761024
         Caption         =   "삭제(Delete)"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   0
      End
      Begin VB.Frame fraCal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   4245
         Left            =   120
         TabIndex        =   49
         Top             =   5400
         Visible         =   0   'False
         Width           =   6345
         Begin VB.TextBox txtTestCal 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3000
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   58
            Top             =   1170
            Width           =   6240
         End
         Begin VB.ComboBox cboCalTest 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   4380
            TabIndex        =   57
            Top             =   780
            Width           =   1935
         End
         Begin VB.ComboBox cboCalTest 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   4380
            TabIndex        =   56
            Top             =   420
            Width           =   1935
         End
         Begin VB.ComboBox cboCalTest 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1230
            TabIndex        =   55
            Top             =   780
            Width           =   1935
         End
         Begin VB.ComboBox cboCalTest 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   1230
            TabIndex        =   54
            Top             =   420
            Width           =   1935
         End
         Begin VB.TextBox txtCal 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "ⓑ"
            Top             =   780
            Visible         =   0   'False
            Width           =   315
         End
         Begin HSCotrol.HSLabel HSLabel15 
            Height          =   345
            Left            =   60
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            BackColor       =   14737632
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "계산항목"
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel21 
            Height          =   315
            Left            =   60
            Top             =   90
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   556
            BackColor       =   15780518
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "계산식"
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel22 
            Height          =   345
            Left            =   60
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            BackColor       =   14737632
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "계산항목"
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel26 
            Height          =   345
            Left            =   3210
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            BackColor       =   14737632
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "계산항목"
            Alignment       =   2
         End
         Begin HSCotrol.HSLabel HSLabel27 
            Height          =   345
            Left            =   3210
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            BackColor       =   14737632
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "계산항목"
            Alignment       =   2
         End
         Begin VB.TextBox txtCal 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   4050
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "ⓒ"
            Top             =   420
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtCal 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   4050
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "ⓓ"
            Top             =   780
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtCal 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   780
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "ⓐ"
            Top             =   420
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin HSCotrol.HSLabel HSLabel13 
         Height          =   375
         Left            =   180
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         BackColor       =   16777215
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "장비코드"
         BevelOut        =   0
         Alignment       =   2
      End
   End
   Begin FPSpreadADO.fpSpread spdTest 
      CausesValidation=   0   'False
      Height          =   8895
      Left            =   30
      TabIndex        =   69
      Tag             =   "20001"
      Top             =   120
      Width           =   7500
      _Version        =   524288
      _ExtentX        =   13229
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      MaxCols         =   26
      MaxRows         =   489
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmTestSet.frx":2AE4
      VisibleCols     =   3
      VisibleRows     =   10
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
      CellNoteIndicatorColor=   16576
   End
End
Attribute VB_Name = "frmTestSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pSort As Integer

Private Sub cboCalTest_Click(Index As Integer)
    
    If txtTestCd.Text = "" Then
        MsgBox "계산식을 적용할 검사코드를 먼저 선택하세요", vbOKOnly + vbCritical, "계산식"
        Exit Sub
    End If
    
    txtTestCal.Text = txtTestCal.Text & "%" & cboCalTest(Index).Text & "%"
    
End Sub

Private Sub chkCalTest_Click()

    If chkCalTest.Value = "0" Then
        fraResultTrans.Visible = True
        fraCal.Visible = False
    Else
        fraResultTrans.Visible = False
        fraCal.Visible = True
    End If
    
End Sub

Private Sub cmdAdd_Click()
    Dim i As Integer
    
    With lstTestCode
        For i = 0 To .ListCount
            If txtTestCd.Text = .List(i) Then
                Exit Sub
            End If
        Next
        .AddItem txtTestCd.Text
        txtTestCd.Text = ""
    End With
    
End Sub

Private Sub cmdConfirm_Click(Index As Integer)
    Dim Test_Property       As Scripting.Dictionary
    Dim objTest_Property    As clsCommon
    Dim i                   As Integer
    Dim strTmp              As String
    Dim intINQuant          As Integer
    Dim intResUse           As Integer
    Dim strItemCodes        As String
    
    '결과표기
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
    
    '결과형태
    If optResType(0).Value = True Then
        intResUse = 0       '수치
    ElseIf optResType(1).Value = True Then
        intResUse = 1       '판정결과 (문자형)
    ElseIf optResType(2).Value = True Then
        intResUse = 2       '수치/판정결과 (문자형)
    End If
    
    '신규
    If Index = 8 Then
        Call frmClear
        txtTestNm.SetFocus
        
    '저장
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

        If lstTestCode.ListCount <= 0 Then
            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
            txtTestCd.SetFocus
            Exit Sub
        End If

        If Trim(txtTestNm.Text) = "" Then
            MsgBox "검사명을 입력하세요", vbCritical, Me.Caption
            txtTestNm.SetFocus
            Exit Sub
        End If

        'EQPMASTER 저장
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "SEQ", txtSeq.Text
            .Add "OCH", txtOChannel.Text
            .Add "RCH", txtRChannel.Text
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
            '소수점 사용여부
            .Add "RESUSE", IIf(chkResSpec.Value = "0", "0", "1")
            '변환소수점
            .Add "RES", txtResSpec.Text
            .Add "REFML", txtRefMLow.Text
            .Add "REFMH", txtRefMHigh.Text
            .Add "REFFL", txtRefFLow.Text
            .Add "REFFH", txtRefFHigh.Text
            '결과형태 : 0:정량,1:정성,2:정량/정성
            .Add "USERESULT", intResUse
            .Add "CALYN", IIf(chkCalTest.Value = "0", "0", "1")
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetEqpInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With

        'TESTMASTER 저장
        Set Test_Property = New Scripting.Dictionary
        
        strItemCodes = ""
        For i = 0 To lstTestCode.ListCount - 1
            strItemCodes = strItemCodes & lstTestCode.List(i) & "#"
        Next
        With Test_Property
            .Add "RCH", txtRChannel.Text
            .Add "SEQ", txtSeq.Text
            .Add "TESTCD", strItemCodes
            .Add "TESTCALCD", txtTestCd.Text
            .Add "CALCULATE", txtTestCal.Text
        End With

        Set objTest_Property = New clsCommon

        With objTest_Property
            .SetAdoCn AdoCn_Local
            If Not .LetTestInfo(Test_Property) Then
                '-- 저장 오류
                'Call GetTestList
            End If
        End With
        
        'AMRMASTER 저장
        Set Test_Property = New Scripting.Dictionary

        With Test_Property
            .Add "EQPCD", txtEqpCD.Text
            .Add "RCH", txtRChannel.Text
            .Add "AMRINRESULT", intINQuant
            .Add "AMRNUMLIMIT1", txtAMRNumLimit(0).Text
            .Add "AMRNUMLIMIT2", txtAMRNumLimit(1).Text
            .Add "AMRNUMLIMIT3", txtAMRNumLimit(2).Text
            .Add "AMRNUMLIMIT4", txtAMRNumLimit(3).Text
            .Add "AMRNUMLIMIT5", txtAMRNumLimit(4).Text
            
            .Add "AMRNUMRESULT1", txtAMRNumResult(0).Text
            .Add "AMRNUMRESULT2", txtAMRNumResult(1).Text
            .Add "AMRNUMRESULT3", txtAMRNumResult(2).Text
            .Add "AMRNUMRESULT4", txtAMRNumResult(3).Text
            .Add "AMRNUMRESULT5", txtAMRNumResult(4).Text
            
            .Add "AMRCHARLIMIT1", txtAMRCharLimit(0).Text
            .Add "AMRCHARLIMIT2", txtAMRCharLimit(1).Text
            .Add "AMRCHARLIMIT3", txtAMRCharLimit(2).Text
            .Add "AMRCHARLIMIT4", txtAMRCharLimit(3).Text
            .Add "AMRCHARLIMIT5", txtAMRCharLimit(4).Text
            
            .Add "AMRCHARRESULT1", txtAMRCharResult(0).Text
            .Add "AMRCHARRESULT2", txtAMRCharResult(1).Text
            .Add "AMRCHARRESULT3", txtAMRCharResult(2).Text
            .Add "AMRCHARRESULT4", txtAMRCharResult(3).Text
            .Add "AMRCHARRESULT5", txtAMRCharResult(4).Text

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

    '삭제
    ElseIf Index = 1 Then
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
            .Add "TESTNM", txtTestNm.Text
            .Add "ABBRNM", txtAbbrNm.Text
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
        
'    ElseIf Index = 2 Then
'        SQL = ""
'        SQL = SQL & "DELETE FROM EQPMASTER"
'
'        Call DBExec(AdoCn_Local, SQL)
'
'        With spdTest
'            For i = 1 To .MaxRows
'                SQL = ""
'                SQL = SQL & "INSERT INTO EQPMASTER " & vbCrLf
'                SQL = SQL & "(EQUIPCD, SEQNO, SENDCHANNEL, RSLTCHANNEL"
'                SQL = SQL & " , TESTCODE, TESTNAME, ABBRNAME, RESPRECUSE, RESPREC "
'                SQL = SQL & " , REFMLOW, REFMHIGH, REFFLOW, REFFHIGH,RESTYPE" & vbCrLf
'                '-- AMR
'                SQL = SQL & " , AMRLimit1, AMRResult1, AMRLimit2, AMRResult2, AMRLimit3, AMRResult3 " & vbCrLf
'                SQL = SQL & " , AMRLimit4, AMRResult4, AMRLimit5, AMRResult5, AMRLimit6, AMRResult6 " & vbCrLf
'                SQL = SQL & " , AMRLimit7, AMRResult7, AMRINResult)                                 " & vbCrLf
'                SQL = SQL & " VALUES (" & vbCrLf
'                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
'                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLOCHANNEL))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTNM))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLABBRNM))
'                SQL = SQL & "," & GetText(spdTest, i, colLRESSPECUSE)
'                SQL = SQL & "," & GetText(spdTest, i, colLRESSPEC)
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLMLOW))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLMHIGH))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLFLOW))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLFHIGH))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colRESTYPE))
'                '-- AMR
'                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 1))
'                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
'                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
'                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 2))
'                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
'                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
'                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 3))
'                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
'                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
'                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 4))
'                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
'                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
'                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 5))
'                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
'                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
'                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 6))
'                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
'                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
'                strTmp = Trim(GetText(spdTest, i, colRESTYPE + 7))
'                SQL = SQL & "," & STS(mGetP(strTmp, 1, "|"))
'                SQL = SQL & "," & STS(mGetP(strTmp, 2, "|"))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colRESTYPE + 8))
'                SQL = SQL & ")" & vbCrLf
'
'                Call DBExec(AdoCn_Local, SQL)
'
'            Next
'        End With
'
'        Call GetTestList
'        Call GetTestMaster(spdTest)
'
'    ElseIf Index = 3 Then
'
'        Call GetTestList
'        Call GetTestMaster(spdTest)
'
'    ElseIf Index = 4 Then
'        '문자현재코드적용
'        If Trim(txtEqpCD.Text) = "" Then
'            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
'            Exit Sub
'        End If
'
'        If Trim(txtRChannel.Text) = "" Then
'            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
'            txtRChannel.SetFocus
'            Exit Sub
'        End If
'
'        If Trim(txtTestCd.Text) = "" Then
'            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
'            txtTestCd.SetFocus
'            Exit Sub
'        End If
'
'        Set Test_Property = New Scripting.Dictionary
'
'        With Test_Property
'            .Add "EQPCD", txtEqpCD.Text
'            .Add "SEQ", txtSeq.Text
'            .Add "RCH", txtRChannel.Text
'            .Add "TESTCD", txtTestCd.Text
'
'            '-- 결과변환 : 문자형
'            .Add "AMRLIMIT8", txtAMRLimit(7).Text
'            .Add "AMRLIMIT9", txtAMRLimit(8).Text
'            .Add "AMRLIMIT10", txtAMRLimit(9).Text
'            .Add "AMRLIMIT11", txtAMRLimit(10).Text
'            .Add "AMRLIMIT12", txtAMRLimit(11).Text
'            .Add "AMRLIMIT13", txtAMRLimit(12).Text
'            .Add "AMRLIMIT14", txtAMRLimit(13).Text
'
'            .Add "AMRRESULT8", txtAMRResult(7).Text
'            .Add "AMRRESULT9", txtAMRResult(8).Text
'            .Add "AMRRESULT10", txtAMRResult(9).Text
'            .Add "AMRRESULT11", txtAMRResult(10).Text
'            .Add "AMRRESULT12", txtAMRResult(11).Text
'            .Add "AMRRESULT13", txtAMRResult(12).Text
'            .Add "AMRRESULT14", txtAMRResult(13).Text
'
'            '-- 결과변환 : 수치형
'            .Add "AMRLIMIT15", txtAMRLimit(14).Text
'            .Add "AMRLIMIT16", txtAMRLimit(15).Text
'            .Add "AMRLIMIT17", txtAMRLimit(16).Text
'            .Add "AMRLIMIT18", txtAMRLimit(17).Text
'            .Add "AMRLIMIT19", txtAMRLimit(18).Text
'
'            .Add "AMRRESULT15", txtAMRResult(14).Text
'            .Add "AMRRESULT16", txtAMRResult(15).Text
'            .Add "AMRRESULT17", txtAMRResult(16).Text
'            .Add "AMRRESULT18", txtAMRResult(17).Text
'            .Add "AMRRESULT19", txtAMRResult(18).Text
'
'
'        End With
'
'        Set objTest_Property = New clsCommon
'
'        With objTest_Property
'            .SetAdoCn AdoCn_Local
'            If Not .LetAMRInfo(Test_Property) Then
'                '-- 저장 오류
'                'Call GetTestList
'            End If
'        End With
'
'        Call GetTestList
'        Call GetTestMaster(spdTest)
'
'    ElseIf Index = 5 Then
'        '문자전체코드적용
'        SQL = ""
'        SQL = SQL & "DELETE FROM AMRMASTER"
'
'        Call DBExec(AdoCn_Local, SQL)
'
'        With spdTest
'            For i = 1 To .MaxRows
'                SQL = ""
'                SQL = SQL & "INSERT INTO AMRMASTER " & vbCrLf
'                SQL = SQL & "(EQUIPCD, SEQNO, RSLTCHANNEL, TESTCODE"
'                SQL = SQL & " , AMRLimit8, AMRLimit9, AMRLimit10, AMRLimit11, AMRLimit12, AMRLimit13, AMRLimit14 " & vbCrLf
'                SQL = SQL & " , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14 ) " & vbCrLf
'                SQL = SQL & " VALUES (" & vbCrLf
'                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
'                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
'                SQL = SQL & "," & STS(txtAMRLimit(7).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(8).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(9).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(10).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(11).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(12).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(13).Text)
'                SQL = SQL & "," & STS(txtAMRResult(7).Text)
'                SQL = SQL & "," & STS(txtAMRResult(8).Text)
'                SQL = SQL & "," & STS(txtAMRResult(9).Text)
'                SQL = SQL & "," & STS(txtAMRResult(10).Text)
'                SQL = SQL & "," & STS(txtAMRResult(11).Text)
'                SQL = SQL & "," & STS(txtAMRResult(12).Text)
'                SQL = SQL & "," & STS(txtAMRResult(13).Text)
'                SQL = SQL & ")" & vbCrLf
'
'                Call DBExec(AdoCn_Local, SQL)
'
'            Next
'        End With
'
'        Call GetTestList
'        Call GetTestMaster(spdTest)
'
'    ElseIf Index = 6 Then
'        '수치전체코드적용
'        SQL = ""
'        SQL = SQL & "DELETE FROM AMRMASTER"
'
'        Call DBExec(AdoCn_Local, SQL)
'
'        With spdTest
'            For i = 1 To .MaxRows
'                SQL = ""
'                SQL = SQL & "INSERT INTO AMRMASTER " & vbCrLf
'                SQL = SQL & "(EQUIPCD, SEQNO, RSLTCHANNEL, TESTCODE"
'                SQL = SQL & " , AMRLimit8, AMRLimit9, AMRLimit10, AMRLimit11, AMRLimit12, AMRLimit13, AMRLimit14 " & vbCrLf
'                SQL = SQL & " , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14 ) " & vbCrLf
'                SQL = SQL & " VALUES (" & vbCrLf
'                SQL = SQL & STS(GetText(spdTest, i, colLMACHCODE))
'                SQL = SQL & "," & GetText(spdTest, i, colLSEQNO)
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLRCHANNEL))
'                SQL = SQL & "," & STS(GetText(spdTest, i, colLTESTCD))
'                SQL = SQL & "," & STS(txtAMRLimit(7).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(8).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(9).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(10).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(11).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(12).Text)
'                SQL = SQL & "," & STS(txtAMRLimit(13).Text)
'                SQL = SQL & "," & STS(txtAMRResult(7).Text)
'                SQL = SQL & "," & STS(txtAMRResult(8).Text)
'                SQL = SQL & "," & STS(txtAMRResult(9).Text)
'                SQL = SQL & "," & STS(txtAMRResult(10).Text)
'                SQL = SQL & "," & STS(txtAMRResult(11).Text)
'                SQL = SQL & "," & STS(txtAMRResult(12).Text)
'                SQL = SQL & "," & STS(txtAMRResult(13).Text)
'                SQL = SQL & ")" & vbCrLf
'
'                Call DBExec(AdoCn_Local, SQL)
'
'            Next
'        End With
'
'        Call GetTestList
'        Call GetTestMaster(spdTest)
'
'    ElseIf Index = 7 Then
'        '수치현재코드적용
'        If Trim(txtEqpCD.Text) = "" Then
'            MsgBox "장비코드가 설정되지 않았습니다.", vbCritical, Me.Caption
'            Exit Sub
'        End If
'
'        If Trim(txtRChannel.Text) = "" Then
'            MsgBox "결과채널을 입력하세요", vbCritical, Me.Caption
'            txtRChannel.SetFocus
'            Exit Sub
'        End If
'
'        If Trim(txtTestCd.Text) = "" Then
'            MsgBox "검사코드를 입력하세요", vbCritical, Me.Caption
'            txtTestCd.SetFocus
'            Exit Sub
'        End If
'
'        Set Test_Property = New Scripting.Dictionary
'
'        With Test_Property
'            .Add "EQPCD", txtEqpCD.Text
'            .Add "SEQ", txtSeq.Text
'            .Add "RCH", txtRChannel.Text
'            .Add "TESTCD", txtTestCd.Text
'
'            '-- 결과변환 : 문자형
'            .Add "AMRLIMIT8", txtAMRLimit(7).Text
'            .Add "AMRLIMIT9", txtAMRLimit(8).Text
'            .Add "AMRLIMIT10", txtAMRLimit(9).Text
'            .Add "AMRLIMIT11", txtAMRLimit(10).Text
'            .Add "AMRLIMIT12", txtAMRLimit(11).Text
'            .Add "AMRLIMIT13", txtAMRLimit(12).Text
'            .Add "AMRLIMIT14", txtAMRLimit(13).Text
'
'            .Add "AMRRESULT8", txtAMRResult(7).Text
'            .Add "AMRRESULT9", txtAMRResult(8).Text
'            .Add "AMRRESULT10", txtAMRResult(9).Text
'            .Add "AMRRESULT11", txtAMRResult(10).Text
'            .Add "AMRRESULT12", txtAMRResult(11).Text
'            .Add "AMRRESULT13", txtAMRResult(12).Text
'            .Add "AMRRESULT14", txtAMRResult(13).Text
'
'            '-- 결과변환 : 수치형
'            .Add "AMRLIMIT15", txtAMRLimit(14).Text
'            .Add "AMRLIMIT16", txtAMRLimit(15).Text
'            .Add "AMRLIMIT17", txtAMRLimit(16).Text
'            .Add "AMRLIMIT18", txtAMRLimit(17).Text
'            .Add "AMRLIMIT19", txtAMRLimit(18).Text
'
'            .Add "AMRRESULT15", txtAMRResult(14).Text
'            .Add "AMRRESULT16", txtAMRResult(15).Text
'            .Add "AMRRESULT17", txtAMRResult(16).Text
'            .Add "AMRRESULT18", txtAMRResult(17).Text
'            .Add "AMRRESULT19", txtAMRResult(18).Text
'
'        End With
'
'        Set objTest_Property = New clsCommon
'
'        With objTest_Property
'            .SetAdoCn AdoCn_Local
'            If Not .LetAMRInfo(Test_Property) Then
'                '-- 저장 오류
'                'Call GetTestList
'            End If
'        End With
'
'        Call GetTestList
'        Call GetTestMaster(spdTest)
    End If
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

'Private Sub cmdNTypeChange_Click()
'    If fraNTypeChange.Visible = True Then
'        fraNTypeChange.Visible = False
'        cmdNTypeChange.Caption = "◀ 수치결과변환 보임"
'    Else
'        fraNTypeChange.Visible = True
'        cmdNTypeChange.Caption = "▶ 수치결과변환 숨김"
'        txtAMRLimit(14).SetFocus
'    End If
'
'End Sub

'Private Sub cmdNUnView_Click()
'
'    If fraNTypeChange.Visible = True Then
'        fraNTypeChange.Visible = False
'        cmdNTypeChange.Caption = "◀ 수치결과변환 보임"
'    Else
'        fraNTypeChange.Visible = True
'        cmdNTypeChange.Caption = "▶ 수치결과변환 숨김"
'    End If
'
'End Sub

Private Sub cmdSave_Click()

    
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    
    With lstTestCode
        If .ListCount = 0 Then
            Exit Sub
        End If
        For i = 0 To .ListCount
            If txtTestCd.Text = .List(i) Then
                .RemoveItem i
                txtTestCd.Text = ""
                Exit Sub
            End If
        Next
    End With

End Sub

Private Sub cmdSeqDown_Click()

    txtSeq.Text = txtSeq.Text - 1

End Sub

Private Sub cmdSeqUp_Click()
    
    txtSeq.Text = txtSeq.Text + 1
    
End Sub

Private Sub cmdSpecDown_Click()
    
    If IsNumeric(txtResSpec.Text) And txtResSpec.Text <> "" Then
        txtResSpec.Text = txtResSpec.Text - 1
    End If
    
End Sub

Private Sub cmdSpecUP_Click()
    
    If txtResSpec.Text = "" Then
        txtResSpec.Text = "0"
        chkResSpec.Value = "1"
    End If
    
    If IsNumeric(txtResSpec.Text) And txtResSpec.Text <> "" Then
        txtResSpec.Text = txtResSpec.Text + 1
    End If
    
End Sub


'Private Sub cmdTypeChange_Click()
'
'    If fraTypeChange.Visible = True Then
'        fraTypeChange.Visible = False
'        cmdTypeChange.Caption = "◀ 문자결과변환 보임"
'    Else
'        fraTypeChange.Visible = True
'        cmdTypeChange.Caption = "▶ 문자결과변환 숨김"
'        txtAMRLimit(7).SetFocus
'    End If
'
'End Sub


'Private Sub cmdUnView_Click()
'
'    If fraTypeChange.Visible = True Then
'        fraTypeChange.Visible = False
'        cmdTypeChange.Caption = "◀ 문자결과변환 보임"
'    Else
'        fraTypeChange.Visible = True
'        cmdTypeChange.Caption = "▶ 문자결과변환 숨김"
'    End If
'
'End Sub

'Private Sub cmdView_Click()
'
'    If fraResultTrans.Visible = False Then
'        cmdView.Caption = "검사결과 변환 ▲"
'        fraResultTrans.Visible = True
'    Else
'        cmdView.Caption = "검사결과 변환 ▼"
'        fraResultTrans.Visible = False
'    End If
'
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        If MsgBox("검사코드 설정화면을 닫으시겠습니까?", vbCritical + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    
    With spdTest
        Call SetText(spdTest, "검사약어", 0, 7):
        Call SetText(spdTest, "소수점 변환", 0, 8):
        Call SetText(spdTest, "변환 자릿수", 0, 9):
        Call SetText(spdTest, "남성 (하한치)", 0, 10):
        Call SetText(spdTest, "남성 (상한치)", 0, 11):
        Call SetText(spdTest, "여성 (하한치)", 0, 12):
        Call SetText(spdTest, "여성 (상한치)", 0, 13):
        Call SetText(spdTest, "결과형태", 0, 14):
        Call SetText(spdTest, "결과표기", 0, 25):
        Call SetText(spdTest, "계산항목", 0, 26):
        .ColWidth(1) = 0
        
        .ColWidth(2) = 4    '순번
        .ColWidth(3) = 15    '오더채널
        .ColWidth(4) = 15    '결과채널
        .ColWidth(5) = 0 '7    '검사코드
        .ColWidth(6) = 15    '검사명
        .ColWidth(7) = 0 '8    '검사약어
        
        .ColWidth(8) = 5    'V
        .ColWidth(9) = 5    'DEC
        
        .ColWidth(10) = 0 '6   '남(하)
        .ColWidth(11) = 0 '6   '남(상)
        .ColWidth(12) = 0 '6   '여(하)
        .ColWidth(13) = 0 '6   '여(상)
        
        .ColWidth(14) = 8   '결과형태
        .ColWidth(15) = 0   '수치변환1
        .ColWidth(16) = 0   '수치변환2
        .ColWidth(17) = 0   '수치변환3
        .ColWidth(18) = 0   '수치변환4
        .ColWidth(19) = 0   '수치변환5
        .ColWidth(20) = 0   '문자변환1
        .ColWidth(21) = 0   '문자변환2
        .ColWidth(22) = 0   '문자변환3
        .ColWidth(23) = 0   '문자변환4
        .ColWidth(24) = 0   '문자변환5
        .ColWidth(25) = 10
        .ColWidth(26) = 0   '계산항목
        .MaxRows = 0
    End With
    
    Call frmClear

    Call GetTestMaster(spdTest)
    
    txtEqpCD.Text = gHOSP.MACHCD
    
End Sub

Private Sub frmClear()
    Dim i As Integer
    
    txtEqpCD.Text = GetText(spdTest, 1, colLMACHCODE)
    txtSeq.Text = GetMaxSeqCode
    
    txtTestCd.Text = ""
    lstTestCode.Clear
    
    txtTestNm.Text = ""
    txtOChannel.Text = ""
    txtRChannel.Text = ""
    txtTestNm.Text = ""
    txtAbbrNm.Text = ""
        
    chkResSpec.Value = "0"
    txtResSpec.Text = ""
    txtRefMLow.Text = ""
    txtRefMHigh.Text = ""
    txtRefFLow.Text = ""
    txtRefFHigh.Text = ""
        
    optResType(0).Value = True
    optINQuant(0).Value = True
    
    For i = 0 To 4
        txtAMRNumLimit(i).Text = ""
        txtAMRNumResult(i).Text = ""
        txtAMRCharLimit(i).Text = ""
        txtAMRCharResult(i).Text = ""
    Next
    
    chkCalTest.Value = "0"
    fraCal.Visible = False
    txtTestCal.Text = ""
'    txtTestCalContents.Text = ""

End Sub


Private Sub Form_Resize()
    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    'spdTest.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 160
    spdTest.HEIGHT = Me.ScaleHeight - 160
    spdTest.WIDTH = Me.ScaleWidth - frameTestSet.WIDTH - 160
    frameTestSet.LEFT = Me.ScaleWidth - frameTestSet.WIDTH
    frameTestSet.HEIGHT = spdTest.HEIGHT

End Sub

Private Sub lstTestCode_Click()
    
    txtTestCd.Text = lstTestCode.Text
    
End Sub

'Private Sub optResType_Click(Index As Integer)
'
'    If Index = 0 Then
'        fraN.Enabled = True
'        fraC.Enabled = False
'        fraNC.Enabled = False
'
'        lblNC.Visible = False
'        fraNC.Visible = False
'    ElseIf Index = 1 Then
'        fraN.Enabled = False
'        fraC.Enabled = True
'        fraNC.Enabled = False
'
'        lblNC.Visible = False
'        fraNC.Visible = False
'    ElseIf Index = 2 Then
'        fraN.Enabled = True
'        fraC.Enabled = True
'        fraNC.Enabled = True
'
'        lblNC.Visible = True
'        fraNC.Visible = True
'    End If
'
'End Sub

Private Sub spdTest_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SetSpreadSort(spdTest, pSort)
        If pSort = 0 Then
            pSort = 1
        Else
            pSort = 1
        End If
        
        Exit Sub
    End If
    
    Call spdContentsView(Row)

'''    Dim strResUse   As String
'''    Dim varTestCode As Variant
'''    Dim intCnt      As Integer
'''
'''    Call frmClear
'''
'''    varTestCode = ""
'''
'''    If Row = 0 Then
'''        cmdNTypeChange.Enabled = False
'''        cmdTypeChange.Enabled = False
'''        Exit Sub
'''    End If
'''
'''    With spdTest
'''        varTestCode = GetTestCode(GetText(spdTest, Row, colLRCHANNEL))
'''        varTestCode = Split(varTestCode, "@")
'''        lstTestCode.Clear
'''        txtTestCd.Text = ""
'''        If UBound(varTestCode) > 0 Then
'''            For intCnt = 0 To UBound(varTestCode) - 1
'''                lstTestCode.AddItem varTestCode(intCnt)
'''            Next
'''            txtTestCd.Text = lstTestCode.List(0)
'''        End If
'''
'''        cmdNTypeChange.Enabled = True
'''        cmdTypeChange.Enabled = True
'''
''''        fraNTypeChange.Visible = False
''''        cmdNTypeChange.Caption = "◀ 수치결과변환 보임"
''''
''''        fraTypeChange.Visible = False
''''        cmdTypeChange.Caption = "◀ 문자결과변환 보임"
'''
'''        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
'''        txtSeq.Text = GetText(spdTest, Row, colLSEQNO)
'''        txtOChannel.Text = GetText(spdTest, Row, colLOCHANNEL)
'''        txtRChannel.Text = GetText(spdTest, Row, colLRCHANNEL)
'''        'txtTestCd.Text = GetText(spdTest, Row, colLTESTCD)
'''        txtTestNm.Text = GetText(spdTest, Row, colLTESTNM)
'''        txtAbbrNm.Text = GetText(spdTest, Row, colLABBRNM)
'''
'''        If GetText(spdTest, Row, colLRESSPECUSE) = "0" Then
'''            chkResSpec.Value = "0"
'''        Else
'''            chkResSpec.Value = "1"
'''        End If
'''        txtResSpec.Text = GetText(spdTest, Row, colLRESSPEC)
'''        txtRefMLow.Text = GetText(spdTest, Row, colLMLOW)
'''        txtRefMHigh.Text = GetText(spdTest, Row, colLMHIGH)
'''        txtRefFLow.Text = GetText(spdTest, Row, colLFLOW)
'''        txtRefFHigh.Text = GetText(spdTest, Row, colLFHIGH)
'''
'''        strResUse = GetText(spdTest, Row, colRESTYPE)
'''
''''        If strResUse = "" Or strResUse = "0" Then
''''            optResType(0).Value = True
''''        ElseIf strResUse = "1" Then
''''            optResType(1).Value = True
''''        ElseIf strResUse = "2" Then
''''            optResType(2).Value = True
''''        End If
'''        If strResUse = "" Or strResUse = "수치형" Then
'''            optResType(0).Value = True
'''        ElseIf strResUse = "문자형" Then
'''            optResType(1).Value = True
'''        ElseIf strResUse = "수치/문자형" Then
'''            optResType(2).Value = True
'''        End If
'''
'''        'AMR
'''        txtAMRLimit(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 1, "|")
'''        txtAMRResult(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 2, "|")
'''
'''        txtAMRLimit(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 1, "|")
'''        txtAMRResult(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 2, "|")
'''
'''        txtAMRLimit(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 1, "|")
'''        txtAMRResult(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 2, "|")
'''
'''        txtAMRLimit(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 1, "|")
'''        txtAMRResult(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 2, "|")
'''
'''        txtAMRLimit(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 1, "|")
'''        txtAMRResult(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 2, "|")
'''
'''        txtAMRLimit(5).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 1, "|")
'''        txtAMRResult(5).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 2, "|")
'''
'''        txtAMRLimit(6).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 1, "|")
'''        txtAMRResult(6).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 2, "|")
'''
'''        If GetText(spdTest, Row, colRESTYPE + 8) = "" Then
'''            optINQuant(0).Value = True
'''        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "변환없음" Then
'''            optINQuant(0).Value = True
'''        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "판정(수치)" Then
'''            optINQuant(1).Value = True
'''        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "수치(판정)" Then
'''            optINQuant(2).Value = True
'''        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "판정 수치" Then
'''            optINQuant(3).Value = True
'''        ElseIf GetText(spdTest, Row, colRESTYPE + 8) = "수치 판정" Then
'''            optINQuant(4).Value = True
'''        End If
'''
'''        'Call frmClear
'''        Call GetAMRMaster(txtSeq.Text, txtRChannel.Text, txtTestCd.Text)
'''
'''    End With
'''
'''    'txtTestCd.SetFocus
End Sub

Private Sub spdContentsView(ByVal Row As Integer)

    Dim strResUse   As String
    Dim varTestCode As Variant
    Dim intCnt      As Integer
    
    Call frmClear
    
    varTestCode = ""
    
    If Row = 0 Then
'        cmdNTypeChange.Enabled = False
'        cmdTypeChange.Enabled = False
        Exit Sub
    End If

    With spdTest
        txtEqpCD.Text = GetText(spdTest, Row, colLMACHCODE)
        txtSeq.Text = GetText(spdTest, Row, colLSEQNO)
        txtOChannel.Text = GetText(spdTest, Row, colLOCHANNEL)
        txtRChannel.Text = GetText(spdTest, Row, colLRCHANNEL)
        
        varTestCode = GetTestCode(GetText(spdTest, Row, colLRCHANNEL))
        varTestCode = Split(varTestCode, "@")
        lstTestCode.Clear
        txtTestCd.Text = ""
        If UBound(varTestCode) > 0 Then
            For intCnt = 0 To UBound(varTestCode) - 1
                lstTestCode.AddItem varTestCode(intCnt)
            Next
            txtTestCd.Text = lstTestCode.List(0)
        End If

'        cmdNTypeChange.Enabled = True
'        cmdTypeChange.Enabled = True
        
        'txtTestCd.Text = GetText(spdTest, Row, colLTESTCD)
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
        
'        If strResUse = "" Or strResUse = "0" Then
'            optResType(0).Value = True
'        ElseIf strResUse = "1" Then
'            optResType(1).Value = True
'        ElseIf strResUse = "2" Then
'            optResType(2).Value = True
'        End If
        If strResUse = "" Or strResUse = "수치형" Then
            optResType(0).Value = True
        ElseIf strResUse = "문자형" Then
            optResType(1).Value = True
        ElseIf strResUse = "수치/문자형" Then
            optResType(2).Value = True
        End If
        
        'AMR
        txtAMRNumLimit(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 1, "|")
        txtAMRNumResult(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 1), 2, "|")
        txtAMRNumLimit(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 1, "|")
        txtAMRNumResult(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 2), 2, "|")
        txtAMRNumLimit(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 1, "|")
        txtAMRNumResult(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 3), 2, "|")
        txtAMRNumLimit(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 1, "|")
        txtAMRNumResult(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 4), 2, "|")
        txtAMRNumLimit(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 1, "|")
        txtAMRNumResult(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 5), 2, "|")
        
        txtAMRCharLimit(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 1, "|")
        txtAMRCharResult(0).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 6), 2, "|")
        txtAMRCharLimit(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 1, "|")
        txtAMRCharResult(1).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 7), 2, "|")
        txtAMRCharLimit(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 8), 1, "|")
        txtAMRCharResult(2).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 8), 2, "|")
        txtAMRCharLimit(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 9), 1, "|")
        txtAMRCharResult(3).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 9), 2, "|")
        txtAMRCharLimit(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 10), 1, "|")
        txtAMRCharResult(4).Text = mGetP(GetText(spdTest, Row, colRESTYPE + 10), 2, "|")
    
        If GetText(spdTest, Row, colRESTYPE + 11) = "" Then
            optINQuant(0).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 11) = "변환없음" Then
            optINQuant(0).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 11) = "판정(수치)" Then
            optINQuant(1).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 11) = "수치(판정)" Then
            optINQuant(2).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 11) = "판정 수치" Then
            optINQuant(3).Value = True
        ElseIf GetText(spdTest, Row, colRESTYPE + 11) = "수치 판정" Then
            optINQuant(4).Value = True
        End If
        
        If GetText(spdTest, Row, colRESTYPE + 12) = "" Or GetText(spdTest, Row, colRESTYPE + 12) = "0" Then
            chkCalTest.Value = "0"
        Else
            chkCalTest.Value = "1"
        End If
        
        If chkCalTest.Value = "1" Then
            fraCal.Visible = True
            fraResultTrans.Visible = False
        Else
            fraCal.Visible = False
            fraResultTrans.Visible = True
        End If
        
        'Call frmClear
        Call GetAMRMaster(txtSeq.Text, txtRChannel.Text, txtTestCd.Text)
        
    End With
    
    'txtTestCd.SetFocus
    
End Sub

Private Sub spdTest_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)

    Call spdContentsView(NewRow)
    
'''    Dim strResUse   As String
'''    Dim varTestCode As Variant
'''    Dim intCnt      As Integer
'''
'''    Call frmClear
'''
'''    varTestCode = ""
'''
'''    If NewRow = 0 Then
'''        cmdNTypeChange.Enabled = False
'''        cmdTypeChange.Enabled = False
'''        Exit Sub
'''    End If
'''
'''    With spdTest
'''        varTestCode = GetTestCode(GetText(spdTest, NewRow, colLRCHANNEL))
'''        varTestCode = Split(varTestCode, "@")
'''        lstTestCode.Clear
'''        txtTestCd.Text = ""
'''        If UBound(varTestCode) > 0 Then
'''            For intCnt = 0 To UBound(varTestCode) - 1
'''                lstTestCode.AddItem varTestCode(intCnt)
'''            Next
'''            txtTestCd.Text = lstTestCode.List(0)
'''        End If
'''
'''        cmdNTypeChange.Enabled = True
'''        cmdTypeChange.Enabled = True
'''
''''        fraNTypeChange.Visible = False
''''        cmdNTypeChange.Caption = "◀ 수치결과변환 보임"
''''
''''        fraTypeChange.Visible = False
''''        cmdTypeChange.Caption = "◀ 문자결과변환 보임"
'''
'''        txtEqpCD.Text = GetText(spdTest, NewRow, colLMACHCODE)
'''        txtSeq.Text = GetText(spdTest, NewRow, colLSEQNO)
'''        txtOChannel.Text = GetText(spdTest, NewRow, colLOCHANNEL)
'''        txtRChannel.Text = GetText(spdTest, NewRow, colLRCHANNEL)
'''        'txtTestCd.Text = GetText(spdTest, NewRow, colLTESTCD)
'''        txtTestNm.Text = GetText(spdTest, NewRow, colLTESTNM)
'''        txtAbbrNm.Text = GetText(spdTest, NewRow, colLABBRNM)
'''
'''        If GetText(spdTest, NewRow, colLRESSPECUSE) = "0" Then
'''            chkResSpec.Value = "0"
'''        Else
'''            chkResSpec.Value = "1"
'''        End If
'''        txtResSpec.Text = GetText(spdTest, NewRow, colLRESSPEC)
'''        txtRefMLow.Text = GetText(spdTest, NewRow, colLMLOW)
'''        txtRefMHigh.Text = GetText(spdTest, NewRow, colLMHIGH)
'''        txtRefFLow.Text = GetText(spdTest, NewRow, colLFLOW)
'''        txtRefFHigh.Text = GetText(spdTest, NewRow, colLFHIGH)
'''
'''        strResUse = GetText(spdTest, NewRow, colRESTYPE)
'''
''''        If strResUse = "" Or strResUse = "0" Then
''''            optResType(0).Value = True
''''        ElseIf strResUse = "1" Then
''''            optResType(1).Value = True
''''        ElseIf strResUse = "2" Then
''''            optResType(2).Value = True
''''        End If
'''        If strResUse = "" Or strResUse = "수치형" Then
'''            optResType(0).Value = True
'''        ElseIf strResUse = "문자형" Then
'''            optResType(1).Value = True
'''        ElseIf strResUse = "수치/문자형" Then
'''            optResType(2).Value = True
'''        End If
'''
'''        'AMR
'''        txtAMRLimit(0).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 1), 1, "|")
'''        txtAMRResult(0).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 1), 2, "|")
'''
'''        txtAMRLimit(1).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 2), 1, "|")
'''        txtAMRResult(1).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 2), 2, "|")
'''
'''        txtAMRLimit(2).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 3), 1, "|")
'''        txtAMRResult(2).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 3), 2, "|")
'''
'''        txtAMRLimit(3).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 4), 1, "|")
'''        txtAMRResult(3).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 4), 2, "|")
'''
'''        txtAMRLimit(4).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 5), 1, "|")
'''        txtAMRResult(4).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 5), 2, "|")
'''
'''        txtAMRLimit(5).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 6), 1, "|")
'''        txtAMRResult(5).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 6), 2, "|")
'''
'''        txtAMRLimit(6).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 7), 1, "|")
'''        txtAMRResult(6).Text = mGetP(GetText(spdTest, NewRow, colRESTYPE + 7), 2, "|")
'''
'''        If GetText(spdTest, NewRow, colRESTYPE + 8) = "" Then
'''            optINQuant(0).Value = True
'''        ElseIf GetText(spdTest, NewRow, colRESTYPE + 8) = "변환없음" Then
'''            optINQuant(0).Value = True
'''        ElseIf GetText(spdTest, NewRow, colRESTYPE + 8) = "판정(수치)" Then
'''            optINQuant(1).Value = True
'''        ElseIf GetText(spdTest, NewRow, colRESTYPE + 8) = "수치(판정)" Then
'''            optINQuant(2).Value = True
'''        ElseIf GetText(spdTest, NewRow, colRESTYPE + 8) = "판정 수치" Then
'''            optINQuant(3).Value = True
'''        ElseIf GetText(spdTest, NewRow, colRESTYPE + 8) = "수치 판정" Then
'''            optINQuant(4).Value = True
'''        End If
'''
'''        'Call frmClear
'''        Call GetAMRMaster(txtSeq.Text, txtRChannel.Text, txtTestCd.Text)
'''
'''    End With
    
End Sub

Private Sub txtOChannel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtRChannel.Text = txtOChannel.Text
    End If
End Sub

Private Sub txtTestCal_KeyPress(KeyAscii As Integer)

    If txtTestCd.Text = "" Then
        MsgBox "계산식을 적용할 검사코드를 먼저 선택하세요", vbOKOnly + vbCritical, "계산식"
        Exit Sub
    End If
    
End Sub

Private Sub txtTestCd_Change()

    If txtTestCd.Text <> "" Then
        txtTestCal.Text = GetCalContents(txtRChannel.Text, txtTestCd.Text)
    End If
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If txtTestCd.Text <> "" Then
            Call cmdAdd_Click
        End If
    End If

End Sub

Private Sub txtTestNm_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtAbbrNm.Text = txtTestNm.Text
    End If

End Sub
