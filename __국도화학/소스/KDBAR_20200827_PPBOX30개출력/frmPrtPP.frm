VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPrtPP 
   BackColor       =   &H00FFFFFF&
   Caption         =   " PP Box 출력"
   ClientHeight    =   12270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12270
   ScaleWidth      =   22260
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   90
      TabIndex        =   58
      Top             =   60
      Width           =   21855
      Begin VB.TextBox txtTopPrtVal 
         Height          =   270
         Left            =   14880
         MultiLine       =   -1  'True
         TabIndex        =   96
         Top             =   210
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5610
         Style           =   1  '그래픽
         TabIndex        =   61
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6750
         Style           =   1  '그래픽
         TabIndex        =   60
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "작업완료포함조회"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   9210
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   1650
         TabIndex        =   62
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139460609
         CurrentDate     =   43884
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   3750
         TabIndex        =   63
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139460609
         CurrentDate     =   43884
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   9870
         Top             =   330
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":0000
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":059A
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":0B34
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":10CE
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":1960
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":1ABA
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":1C14
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":1D6E
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPP.frx":2648
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblstrPrtLabelName 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   17400
         TabIndex        =   80
         Top             =   270
         Width           =   2265
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3450
         TabIndex        =   66
         Top             =   420
         Width           =   195
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "▶ 생산일자 "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   330
         TabIndex        =   65
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label lblComStatus 
         BackStyle       =   0  '투명
         Caption         =   "Com1 연결실패"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11685
         TabIndex        =   64
         Top             =   420
         Width           =   3255
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   11340
         Picture         =   "frmPrtPP.frx":2F22
         Top             =   420
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   11325
      Left            =   90
      TabIndex        =   4
      Top             =   1050
      Width           =   21855
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   10785
         Left            =   8160
         TabIndex        =   5
         Top             =   210
         Width           =   13515
         Begin VB.PictureBox picSide 
            BackColor       =   &H00FFFFFF&
            Height          =   3915
            Left            =   7200
            ScaleHeight     =   3855
            ScaleWidth      =   5775
            TabIndex        =   70
            Top             =   6630
            Width           =   5835
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   34
               Left            =   6150
               Picture         =   "frmPrtPP.frx":34AC
               Stretch         =   -1  'True
               Top             =   2610
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   33
               Left            =   6270
               Picture         =   "frmPrtPP.frx":7479
               Stretch         =   -1  'True
               Top             =   2010
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   32
               Left            =   6090
               Picture         =   "frmPrtPP.frx":B446
               Stretch         =   -1  'True
               Top             =   1470
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   31
               Left            =   6090
               Picture         =   "frmPrtPP.frx":F413
               Stretch         =   -1  'True
               Top             =   900
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   30
               Left            =   6180
               Picture         =   "frmPrtPP.frx":133E0
               Stretch         =   -1  'True
               Top             =   390
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   29
               Left            =   2580
               Picture         =   "frmPrtPP.frx":173AD
               Stretch         =   -1  'True
               Top             =   6960
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   28
               Left            =   90
               Picture         =   "frmPrtPP.frx":1B37A
               Stretch         =   -1  'True
               Top             =   6990
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   27
               Left            =   2580
               Picture         =   "frmPrtPP.frx":1F347
               Stretch         =   -1  'True
               Top             =   6480
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   26
               Left            =   150
               Picture         =   "frmPrtPP.frx":23314
               Stretch         =   -1  'True
               Top             =   6570
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   25
               Left            =   2550
               Picture         =   "frmPrtPP.frx":272E1
               Stretch         =   -1  'True
               Top             =   6150
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   24
               Left            =   210
               Picture         =   "frmPrtPP.frx":2B2AE
               Stretch         =   -1  'True
               Top             =   6240
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   23
               Left            =   2550
               Picture         =   "frmPrtPP.frx":2F27B
               Stretch         =   -1  'True
               Top             =   5760
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   22
               Left            =   150
               Picture         =   "frmPrtPP.frx":33248
               Stretch         =   -1  'True
               Top             =   5790
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   21
               Left            =   2490
               Picture         =   "frmPrtPP.frx":37215
               Stretch         =   -1  'True
               Top             =   5280
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   20
               Left            =   210
               Picture         =   "frmPrtPP.frx":3B1E2
               Stretch         =   -1  'True
               Top             =   5340
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   19
               Left            =   2550
               Picture         =   "frmPrtPP.frx":3F1AF
               Stretch         =   -1  'True
               Top             =   4920
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   18
               Left            =   60
               Picture         =   "frmPrtPP.frx":4317C
               Stretch         =   -1  'True
               Top             =   4890
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   17
               Left            =   2520
               Picture         =   "frmPrtPP.frx":47149
               Stretch         =   -1  'True
               Top             =   4470
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   16
               Left            =   90
               Picture         =   "frmPrtPP.frx":4B116
               Stretch         =   -1  'True
               Top             =   4500
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   15
               Left            =   2520
               Picture         =   "frmPrtPP.frx":4F0E3
               Stretch         =   -1  'True
               Top             =   4020
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   14
               Left            =   90
               Picture         =   "frmPrtPP.frx":530B0
               Stretch         =   -1  'True
               Top             =   4080
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   13
               Left            =   2520
               Picture         =   "frmPrtPP.frx":5707D
               Stretch         =   -1  'True
               Top             =   3570
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   12
               Left            =   180
               Picture         =   "frmPrtPP.frx":5B04A
               Stretch         =   -1  'True
               Top             =   3600
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   11
               Left            =   2490
               Picture         =   "frmPrtPP.frx":5F017
               Stretch         =   -1  'True
               Top             =   3030
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   10
               Left            =   180
               Picture         =   "frmPrtPP.frx":62FE4
               Stretch         =   -1  'True
               Top             =   3030
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   9
               Left            =   2490
               Picture         =   "frmPrtPP.frx":66FB1
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   8
               Left            =   180
               Picture         =   "frmPrtPP.frx":6AF7E
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   7
               Left            =   2490
               Picture         =   "frmPrtPP.frx":6EF4B
               Stretch         =   -1  'True
               Top             =   1890
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   6
               Left            =   180
               Picture         =   "frmPrtPP.frx":72F18
               Stretch         =   -1  'True
               Top             =   1890
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   5
               Left            =   2490
               Picture         =   "frmPrtPP.frx":76EE5
               Stretch         =   -1  'True
               Top             =   1350
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   4
               Left            =   180
               Picture         =   "frmPrtPP.frx":7AEB2
               Stretch         =   -1  'True
               Top             =   1350
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   3
               Left            =   2490
               Picture         =   "frmPrtPP.frx":7EE7F
               Stretch         =   -1  'True
               Top             =   780
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   2
               Left            =   180
               Picture         =   "frmPrtPP.frx":82E4C
               Stretch         =   -1  'True
               Top             =   780
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   1
               Left            =   2520
               Picture         =   "frmPrtPP.frx":86E19
               Stretch         =   -1  'True
               Top             =   210
               Visible         =   0   'False
               Width           =   2205
            End
            Begin VB.Image imgPpBar 
               Height          =   465
               Index           =   0
               Left            =   210
               Picture         =   "frmPrtPP.frx":8ADE6
               Stretch         =   -1  'True
               Top             =   210
               Visible         =   0   'False
               Width           =   2205
            End
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   5175
            Left            =   7200
            ScaleHeight     =   5115
            ScaleWidth      =   5775
            TabIndex        =   7
            Top             =   1440
            Width           =   5835
            Begin VB.Image imgBar1 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtPP.frx":8EDB3
               Stretch         =   -1  'True
               Top             =   1950
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.Image imgQrBar 
               Height          =   750
               Left            =   660
               Picture         =   "frmPrtPP.frx":92D80
               Stretch         =   -1  'True
               Top             =   3000
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Image imgBar2 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtPP.frx":993B8
               Stretch         =   -1  'True
               Top             =   2430
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   0
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   1
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   2
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   3
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   4
               Left            =   0
               TabIndex        =   34
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   5
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   6
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   7
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   8
               Left            =   0
               TabIndex        =   30
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   9
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   10
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   11
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   12
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   13
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   14
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   15
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   1
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   2
               Left            =   0
               TabIndex        =   21
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   3
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   4
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   5
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   6
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   7
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   8
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   9
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   10
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   11
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   12
               Left            =   390
               TabIndex        =   11
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   13
               Left            =   690
               TabIndex        =   10
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   14
               Left            =   870
               TabIndex        =   9
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.Label lblContents 
               AutoSize        =   -1  'True
               Caption         =   "Name"
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
               Index           =   15
               Left            =   900
               TabIndex        =   8
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
         End
         Begin VB.CommandButton cmdVisible 
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6360
            TabIndex        =   89
            Top             =   6450
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtReelQTY 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8130
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   83
            Top             =   -90
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10410
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   82
            Top             =   -60
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11310
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   81
            Top             =   -60
            Width           =   1485
         End
         Begin VB.TextBox txtInBarcode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   77
            Top             =   5010
            Width           =   2520
         End
         Begin VB.TextBox txtMsg 
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
            Height          =   3015
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   76
            Top             =   6900
            Width           =   4245
         End
         Begin VB.CheckBox chkReelPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "측면Reel출력"
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
            Height          =   405
            Left            =   4560
            TabIndex        =   75
            Top             =   5910
            Value           =   1  '확인
            Width           =   1455
         End
         Begin FPSpread.vaSpread spdScan 
            Height          =   2985
            Left            =   4530
            TabIndex        =   74
            Top             =   6900
            Width           =   2625
            _Version        =   393216
            _ExtentX        =   4630
            _ExtentY        =   5265
            _StockProps     =   64
            ColsFrozen      =   8
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   15921919
            GridShowVert    =   0   'False
            MaxCols         =   19
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ScrollBars      =   2
            ShadowColor     =   16774120
            SpreadDesigner  =   "frmPrtPP.frx":9D385
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.TextBox txtScanCount 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   6630
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   73
            Top             =   5460
            Width           =   480
         End
         Begin VB.TextBox txtPrtPPBoxNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4830
            MaxLength       =   5
            TabIndex        =   72
            Top             =   -120
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox txtMaxTot 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1620
            MaxLength       =   5
            TabIndex        =   3
            Top             =   5910
            Width           =   1530
         End
         Begin VB.CheckBox chAutoPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "자동출력"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3390
            TabIndex        =   69
            Top             =   5910
            Value           =   1  '확인
            Width           =   1185
         End
         Begin VB.TextBox txtReelBarcode 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "2X2707R0202001P10110000"
            Top             =   5460
            Width           =   4980
         End
         Begin VB.TextBox txtPPBoxNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   1
            Top             =   5010
            Width           =   1530
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   735
            Left            =   240
            TabIndex        =   45
            Top             =   9900
            Width           =   6915
            Begin VB.CommandButton cmdErrClear 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   60
               Style           =   1  '그래픽
               TabIndex        =   93
               Top             =   150
               Width           =   765
            End
            Begin VB.CommandButton cmdBC 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "BC"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1350
               Style           =   1  '그래픽
               TabIndex        =   92
               Top             =   180
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.CommandButton cmdMakeBar 
               BackColor       =   &H00E0E0E0&
               Caption         =   "내부바코드출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   2370
               Style           =   1  '그래픽
               TabIndex        =   78
               Top             =   150
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00E0E0E0&
               Caption         =   "닫기"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   5580
               Style           =   1  '그래픽
               TabIndex        =   48
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdPrint 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   4470
               Style           =   1  '그래픽
               TabIndex        =   47
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdView 
               BackColor       =   &H00E0E0E0&
               Caption         =   "보기"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1800
               Style           =   1  '그래픽
               TabIndex        =   46
               Top             =   150
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.TextBox txtProdNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1770
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   44
            Top             =   360
            Width           =   1485
         End
         Begin VB.TextBox txtPackNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1770
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   43
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtProdOrderDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   42
            Top             =   360
            Width           =   1500
         End
         Begin VB.TextBox txtCompNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   41
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtSlittingNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8190
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   40
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtLotNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8190
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   6
            Top             =   360
            Width           =   1500
         End
         Begin FPSpread.vaSpread spdPrtReelDetail 
            Height          =   3435
            Left            =   240
            TabIndex        =   49
            Top             =   1440
            Width           =   6885
            _Version        =   393216
            _ExtentX        =   12144
            _ExtentY        =   6059
            _StockProps     =   64
            ColsFrozen      =   8
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   15921919
            GridShowVert    =   0   'False
            MaxCols         =   15
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            ScrollBarShowMax=   0   'False
            ShadowColor     =   16775150
            SpreadDesigner  =   "frmPrtPP.frx":9E4C6
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Frame fraTop 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   585
            Left            =   120
            TabIndex        =   84
            Top             =   6300
            Visible         =   0   'False
            Width           =   6075
            Begin VB.TextBox txtSidePrtVal 
               Height          =   285
               Left            =   4770
               MultiLine       =   -1  'True
               TabIndex        =   95
               Top             =   180
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.CommandButton cmdSidePrint 
               Appearance      =   0  '평면
               BackColor       =   &H00C0FFFF&
               Caption         =   "측면 출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2190
               Style           =   1  '그래픽
               TabIndex        =   94
               Top             =   150
               Width           =   2145
            End
            Begin VB.CommandButton cmdUnvisible 
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4380
               TabIndex        =   88
               Top             =   180
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.TextBox txtTopPrtNo 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1500
               MaxLength       =   5
               TabIndex        =   86
               Text            =   "1"
               Top             =   150
               Width           =   630
            End
            Begin VB.CommandButton cmdTopPrint 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "상단 출력"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2190
               Style           =   1  '그래픽
               TabIndex        =   85
               Top             =   0
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.Label lblComp 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  '단일 고정
               Caption         =   "출력수량"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   87
               Top             =   150
               Width           =   1350
            End
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "1 매"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   6540
            TabIndex        =   91
            Top             =   5970
            Width           =   435
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "생산 Reel 수량"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   11
            Left            =   6300
            TabIndex        =   90
            Top             =   -90
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "내부바코드"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   3210
            TabIndex        =   79
            Top             =   5010
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "출력 Box No"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   3270
            TabIndex        =   71
            Top             =   -90
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "바코드 스캔"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   68
            Top             =   5460
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "PP Box No"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   67
            Top             =   5010
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품명"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "고객사"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   3300
            TabIndex        =   55
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Slitting 작업번호"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   6360
            TabIndex        =   54
            Top             =   810
            Width           =   1800
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "생산일자"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   3300
            TabIndex        =   53
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "포장코드"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   52
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "박스당 릴수량"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   51
            Top             =   5910
            Width           =   1350
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "Lot No"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   6360
            TabIndex        =   50
            Top             =   360
            Width           =   1800
         End
      End
      Begin FPSpread.vaSpread spdPrtReel 
         Height          =   10695
         Left            =   210
         TabIndex        =   57
         Top             =   300
         Width           =   7905
         _Version        =   393216
         _ExtentX        =   13944
         _ExtentY        =   18865
         _StockProps     =   64
         ColsFrozen      =   8
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   25
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmPrtPP.frx":9F004
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.TextBox txtComm 
      Appearance      =   0  '평면
      Height          =   5955
      Left            =   20490
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1770
      Visible         =   0   'False
      Width           =   5475
   End
End
Attribute VB_Name = "frmPrtPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmPrtPPBox.frm
'   작성자  : 오세원
'   내  용  : PP Box 라벨출력
'   작성일  : 2020-02-29
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'


Private Sub cmdBC_Click()

    If spdScan.Visible = False Then
        spdScan.Visible = True
        spdScan.ZOrder 0
    Else
        spdScan.Visible = False
    End If
    
End Sub

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdPrtReel.MaxRows = 0
    spdPrtReelDetail.MaxRows = 0
    spdScan.MaxRows = 0
    
    'spdRegOrderDetail.MaxRows = 0
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now


    For i = 0 To 15
        'barReel.Visible = False
        'barPart.Visible = False
        lblTitle(i).Visible = False
    Next
    
    For i = 0 To 34 '11
        imgPpBar(i).Visible = False
    Next


    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
'    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    'txtReelQTY.Text = ""
    
    txtPPBoxNo.Text = ""
    txtPrtPPBoxNo.Text = ""
    txtReelBarcode.Text = ""
    txtMaxTot.Text = "0"
    txtScanCount.Text = "0"

    chkReelPrint.Value = "0"
    txtMsg.Text = ""
    txtInBarcode.Text = ""
    txtScanCount.Text = "0"

    lblstrPrtLabelName.Caption = ""
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub


' 바코드 리스트 가져옴
Private Sub GetOrderList(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String)
    Dim strLabelType    As String
    
    Set AdoRs = Get_OrderList(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType)
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReel
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 1)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 3)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_NO").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReel, AdoRs.Fields("PACK_CD").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReel, AdoRs.Fields("REEL_QTY").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReel, AdoRs.Fields("ROOL_INFO").Value & "", .MaxRows, 9)
                Call SetText(spdPrtReel, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 10)
                Call SetText(spdPrtReel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_MEMO").Value & "", .MaxRows, 12)
                Call SetText(spdPrtReel, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 13)
                Call SetText(spdPrtReel, AdoRs.Fields("CLOSE_YN").Value & "", .MaxRows, 14)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LABEL_CD").Value & "", .MaxRows, 15)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub

' Reel 작업 리스트 가져옴
Private Sub GetReelOrderList_PP(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String)

    Dim strLabelType    As String

    Set AdoRs = Get_OrderList_PP(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType)

    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReel
                .MaxRows = .MaxRows + 1

                Call SetText(spdPrtReel, "1", .MaxRows, 1)
                Call SetText(spdPrtReel, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReel, Format(AdoRs.Fields("PROD_ORDER_DT").Value & "", "####-##-##"), .MaxRows, 3)
'                Call SetText(spdPrtReel, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_PRT_NAME").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReel, AdoRs.Fields("PACK_CD").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReel, AdoRs.Fields("ORDER_MEMO").Value & "", .MaxRows, 8)
                'Call SetText(spdPrtReel, AdoRs.Fields("ROOL_INFO").Value & "", .MaxRows, 9)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 10)
                Call SetText(spdPrtReel, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("REEL_QTY").Value & "", .MaxRows, 12)
                'Call SetText(spdPrtReel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 11)
                Call SetText(spdPrtReel, AdoRs.Fields("COMP_VIEW").Value & "", .MaxRows, 13)
                Call SetText(spdPrtReel, AdoRs.Fields("CLOSE_YN").Value & "", .MaxRows, 14)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_LABEL_CD").Value & "", .MaxRows, 15)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdPrtReel, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                'Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdPrtReel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 18)
                Call SetText(spdPrtReel, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With

            AdoRs.MoveNext
        Loop
    
        AdoRs.Close

    End If


End Sub


Private Sub cmdErrClear_Click()
    
    txtMsg.Text = ""
    
End Sub


'내부 관리용 바코드 출력한다
Private Function GetMakeInBar(pBarcode As String) As String
    Dim strAFont    As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strHeader   As String
    
    GetMakeInBar = ""
    
    If pBarcode = "" Then
        Exit Function
    End If
    
    strAFont = "^AJN,50,30"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strHeader = strHeader & "^MD10"
    strHeader = strHeader & "^PRD"
    strOutput = ""
    
    If lblstrPrtLabelName.Caption = "P0003" Then
        strOutput = strOutput & "^FO500,50^CI26" & strAFont & "^FD국도화학 내부관리용코드^FS" & vbLf
        'strOutput = strOutput & "^FO500,100^CI26^BY3,3,100^BC^FD" & pBarcode
        strOutput = strOutput & "^FO400,100^CI26^BY4,3,100^BC^FD" & pBarcode
    Else
        strOutput = strOutput & "^FO100,100^CI26" & strAFont & "^FD국도화학 내부관리용코드^FS" & vbLf
        strOutput = strOutput & "^FO100,100^CI26^BY3,3,100^BC^FD" & pBarcode
    End If
    strOutput = strOutput & "^FS" & vbLf
    strOutput = strHeader & strOutput & "^XZ" & vbLf
    
    GetMakeInBar = strOutput

End Function

'내부 관리용 바코드 출력한다
Private Sub cmdMakeBar_Click()
    Dim strPrtData  As String
    
    strPrtData = GetMakeInBar(txtInBarcode.Text)
    
    If strPrtData <> "" Then
        comEqp.Output = strPrtData
        
        Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strPrtData, "A")
        
    End If

End Sub


'^XA : Opening Bracket으로 Format의 시작을 알린다.
'^FO : 인쇄 할 항목의 인쇄 할 위치(X축,Y축)를 정의한다.
'^A1 : 폰트설정
'      o => 회전 : N(0),R(90),I(180),B(270)
'      h => 높이 : 20
'      w => 넓이 : 25
'^BY : 바코드 사이즈 조정
'      w => module width        : 1 ~ 10    (Default : 2)
'      r => wide bar to narrow  : 2.0 ~ 3.0 (Default : 3)
'      h => bar code height     : 10
'^BQ : QR Code Bar Code

'^BC : Code 128 Bar Code (Subsets A, B, and C)  ==> Format ^BCo,h,f,g,e,m
'       o = orientation Accepted Values:
'           N = Normal
'           R = rotated 90 degrees (clockwise)
'           I = inverted 180 degrees
'           B = read from bottom up, 270 degrees
'           Default Value: current ^FW value
'       h = bar code height (in dots)
'           Accepted Values: 1 to 32000
'           Default Value: value set by ^BY
'       f = print interpretation line
'           Accepted Values: y (yes) Or N(NO)
'           Default Value: y
'           The interpretation line can be printed in any font by
'               placing the font command before the bar code command.
'       g = print interpretation line above code
'           Accepted Values: y (yes) Or N(NO)
'           Default Value: N
'       e = UCC check digit Accepted Values: Y (turns on) or N (turns off)
'           Mod 103 check digit is always there. It cannot be turned on or off. Mod 10 and 103 appear together with e turned on.
'           Default Value: N

'-----------------------------------------------------------------------------------------------------------------------
'  COMMAND SUMMARY
'-----------------------------------------------------------------------------------------------------------------------
' ^XA : start of label format
' ^XZ : end   of label format
' ^LH : set label home position
' ^FD : start of field data
' ^FS : end   of fiels data
' ^B3 : select Code39 bar code
' ^FO : set field position
' ^PQ : set print quantity  (^PQ10)
' ^B  : set barcode type
' ^BY : set barcode style   (code 128)
'-----------------------------------------------------------------------------------------------------------------------
' ^BY2,3,70                       ^BCN,70,Y,N,N
'
'  2 -> 가는선 font               ^BC -> code 128
'  3 -> 굵은선 font                70 -> barcode 높이
' 70 -> barcode 높이               Y -> barcode 하단에 문자열 출력
'                                  N -> barcode 상단에 문자열 출력
'                                  N -> check digit 표시유무
'-----------------------------------------------------------------------------------------------------------------------
' ^A  (FONT type)
'    font type       matrix     interchar gap     baseline
'                                 (in dots)       (in dots)
'        A             9X5            1               7
'        B            11X7            2               11
'       C,D           18X10           2               14
'        E            28X15           5               23
'        F            26X13           3               21
'        G            60X40           8               15
'        H            21X13           6                21
'        GS           24X24                       3XHEIGHT/4
'        0 DEFAULT    15X12                       3XHEIGHT/4
'
' ^A0,N,26,22 : D TYPE 높이 26, 폭 22dot
'
' ^CI  (Change International Font/Encoding)
' 26 = Multibyte Asian Encodings with ASCII Transparency a And c

        'asc("~") : R_7E
        '? Hex(126)
            
        'asc("℃") : R_A1C9
         '?Hex(126)
            
                                'strData = Replace(strData, "~", "_7E")
                                'strData = Replace(strData, "℃", "_A1C9")
            


Private Sub cmdPrint_Click()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    '-- 재출력용
    Dim strPrintVal()   As Variant
    Dim strPPTopLabel   As String
    
    Dim intY            As Integer
    
    Dim strInBarcode    As String
    Dim strInBarData    As String
    
    Dim strOutput3   As String
    
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal
    intCnt = 0
    
    blnPrint = False
    strBarcode = ""
    strOutput = ""
    strPPTopLabel = ""
    intCnt = 0
    intPrt = 0
    i = 0
    
    If spdScan.MaxRows = 0 Then
        MsgBox "바코드를 먼저 스캔하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    strPlusXPos = 630
    'strAFont = "^A0N,35,25"
    strAFont = "^A0N,60,50"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'    strHeader = strHeader & "^MD8"
'    strHeader = strHeader & "^PRD"
    
    strFooter = "^XZ" & vbLf
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0011", "P0018", "P0019", "P0020"
            strHeader = strHeader & "~SD12"
            strHeader = strHeader & "^MD12"
            strHeader = strHeader & "^PR0"
                        
            strAFont = "^A0N,40,30"
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        '이미 계산됨
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^BY2,2,80^BC"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,2,80"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            If strType = "발행시간" Then
                                strPrtData(i) = strPrtData(i) & "^A0N,30,20"
                                strData = Format(MONTH(Now), "00") & "/" & Format(Day(Now), "00") & " " & Format(Now, "hh:mm:ss") 'Format(Now, "mm/dd hh:mm:ss")
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            strPrtData(i) = strPrtData(i) & "^FH"
                            
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
            
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            '-- 측면출력여부
            If chkReelPrint.Value = "1" Then
                '-- O 출력한다.
                strPPTopLabel = strOutput
            Else
                '-- X 출력안한다.
                strPPTopLabel = ""
                strOutput = strHeader & strOutput & strFooter
                comEqp.Output = strOutput
                
                Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
                
            End If
            
            blnPrint = True
            txtComm.Text = strOutput
            strOutput = ""
                
        Case "P0001", "P0002"
            'strHeader = strHeader & "~SD12"
            strHeader = strHeader & "^MD25"
            '2020.07.24 수정 ^PR0 에서 ^PRC 로 수정 (노란색에서 잘나옴)
            strHeader = strHeader & "^PRA"
            
'^PR p, s
': Slew SPEED
'A or 2 50.8mm /Sec
'B or 3 76.2mm /Sec
'C or 4 101.6mm /Sec
'5 127mm /Sec
'D or 6 152.4mm /Sec
'E or 8 203.2mm /Sec
': Print SPEED
'A or 2 50.8mm /Sec
'B or 3 76.2mm /Sec
'C or 4 101.6mm /Sec
'5 127mm /Sec
'D or 6 152.4mm /Sec
'E or 8 203.2mm /Sec
                        
            strAFont = "^A0N,40,30"
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strLotSub = "P" & strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                        '이미 계산됨
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^BY2,2,80^BC"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BY" & "2,2,80"
                            strPrtData(i) = strPrtData(i) & "^BQ"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            strLot = mGetP(strData, 1, "(")
                            strLot = strLot & "(" & strLotSub & ")"
                            
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            If strType <> "Provider" Then
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^CI26"
                                If strType = "발행시간" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,30,20"
                                    strData = Format(MONTH(Now), "00") & "/" & Format(Day(Now), "00") & " " & Format(Now, "hh:mm:ss") 'Format(Now, "mm/dd hh:mm:ss")
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont
                                End If
                                strPrtData(i) = strPrtData(i) & "^FH"
                                
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            End If
                        End If
                    End If
                Next
            End With
            
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            '-- 측면출력여부
            If chkReelPrint.Value = "1" Then
                '-- O 출력한다.
                strPPTopLabel = strOutput
            Else
                '-- X 출력안한다.
                strPPTopLabel = ""
                strOutput = strHeader & strOutput & strFooter
                comEqp.Output = strOutput
                
                Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
                
            End If
            
            '트래킹용
            'ReDim Preserve strTrackBC(intCnt)
            'strTrackBC(intCnt) = strBarcode
            
            '재출력용
            'ReDim Preserve strPrintVal(intCnt)
            'strPrintVal(intCnt) = strOutput
            
            'intCnt = intCnt + 1
            
            blnPrint = True
            txtComm.Text = strOutput
            strOutput = ""
        
            '국도화학 내부관리용코드
        Case "P0003"
            strAFont = "^A0N,45,35"
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        strLotSub = Format(intPrt, "0000")
                        strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "보관온도" Then
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        strPrtData(i) = strPrtData(i) & "^A0N,20,15"
                        strPrtData(i) = strPrtData(i) & "^FH"
                        strPrtData(i) = strPrtData(i) & "^FD" & "(" & strData & ")"
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    Else
                        ReDim Preserve strPrtData(i) As String
                        strPrtData(i) = ""
                        strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                        strPrtData(i) = strPrtData(i) & "^CI26"
                        If strType = "Name" Then
                            strPrtData(i) = strPrtData(i) & "^A0N,75,55"
                        ElseIf strType = "Material code" Or strType = "생산일자" Then
                            strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                        Else
                            strPrtData(i) = strPrtData(i) & strAFont
                        End If
                        strPrtData(i) = strPrtData(i) & "^FH"
                        If strNamePrt = "Y" Then
                            strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                        Else
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                        End If
                        strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                        i = i + 1
                    End If
                Next
            End With
        
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            If chkReelPrint.Value = "1" Then
                strPPTopLabel = strOutput
            Else
                strPPTopLabel = ""
                strOutput = strHeader & strOutput & strFooter
                
                strOutput = Replace(strOutput, "~", "_7E")
                strOutput = Replace(strOutput, "℃", "'C")
                
                comEqp.Output = strOutput
            
                Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
            End If
            
            'strOutput = strHeader & strOutput & "^XZ" & vbLf
            'comEqp.Output = strOutput
            
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
'            '재출력용
'            ReDim Preserve strPrintVal(intCnt)
'            strPrintVal(intCnt) = strOutput
'
'            intCnt = intCnt + 1
            
            blnPrint = True
            txtComm.Text = strOutput
            'strOutput = ""
        
        Case "P0004", "P0005", "P0008", "P0009"
            strAFont = "^A0N,45,35"
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                    
                    '^FO50,900^CI26^BY2,3,130^BC^FD0201002240E1B4AAN1K3B1SE108270^FS
                    
                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select
                    
                    If strType = "바코드" Then
                        'strLotSub = strSlt & Format(intPrt, "00")
                        'strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        
                        strData = Mid(strData, 1, 24) & Format(100 + CCur(txtPPBoxNo.Text), "000") & Mid(strData, 28, 3)
                        
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,6"
                            strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    Else
                        If strType = "LotNo" Then
                            'strLot = mGetP(strData, 1, "(")
                            'strLot = strLot & "(P" & strLotSub & ")"
                            strLot = strData
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        ElseIf strType = "보관온도" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^A0N,30,20"
                            strPrtData(i) = strPrtData(i) & "^FH"
                            strPrtData(i) = strPrtData(i) & "^FD" & "(" & strData & ")"
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        ElseIf strType = "바코드값" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            strPrtData(i) = strPrtData(i) & "^FD" & strBarcode
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            
                            If strType = "Name" Then
                                strPrtData(i) = strPrtData(i) & "^A0N,75,55"
                            ElseIf strType = "Material code" Or strType = "생산일자" Then
                                strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                            Else
                                strPrtData(i) = strPrtData(i) & strAFont
                            End If
                            
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With
        
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            If chkReelPrint.Value = "1" Then
                strPPTopLabel = strOutput
            Else
                strPPTopLabel = ""
                strOutput = strHeader & strOutput & strFooter
                
                strOutput = Replace(strOutput, "~", "_7E")
                strOutput = Replace(strOutput, "℃", "'C")
                
                comEqp.Output = strOutput
                
                Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
                
            End If
            
            'strOutput = strHeader & strOutput & "^XZ" & vbLf
            'comEqp.Output = strOutput
            
'            '트래킹용
'            ReDim Preserve strTrackBC(intCnt)
'            strTrackBC(intCnt) = strBarcode
'
            '재출력용
            ReDim Preserve strPrintVal(intCnt)
            strPrintVal(intCnt) = strOutput

'            intCnt = intCnt + 1
            
            blnPrint = True
            txtComm.Text = strOutput
'            strOutput = ""
            
        Case "P0006", "P0007", "P0010"
            strAFont = "^A0N,45,35"
            With spdPrtReelDetail
                For intRow = 1 To .MaxRows
                    '좌측 출력
                    strType = GetText(spdPrtReelDetail, intRow, 1)
                    strData = GetText(spdPrtReelDetail, intRow, 3)
                    strXPos = GetText(spdPrtReelDetail, intRow, 4)
                    strYPos = GetText(spdPrtReelDetail, intRow, 5)
                    strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
                    strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
                    strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
                    strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                    strNamePrt = GetText(spdPrtReelDetail, intRow, 10)

                    Select Case strRot
                        Case "0":   strRot = "N"
                        Case "90":  strRot = "R"
                        Case "180": strRot = "I"
                        Case "270": strRot = "B"
                    End Select

                    If strType = "바코드" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                        If Mid(strBarType, 1, 1) = "1" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                        strBarcode = strData
                    ElseIf strType = "PartsID" Then
                        strLotSub = strSlt & Format(intPrt, "00")
                        If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                            strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                            strPrtData(i) = strPrtData(i) & "^FD" & strData
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    Else
                        If strType = "LotNo" Then
                            strLot = strData

                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strLot
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            'strLotNo = strData
                        ElseIf strType = "Material code" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                            strMatCd = strData
                        ElseIf strType = "Name" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^A0N,65,55" 'strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        ElseIf strType = "보관온도" Then
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & "^A0N,30,20"
                            strPrtData(i) = strPrtData(i) & "^FH"
                            strPrtData(i) = strPrtData(i) & "^FD" & "(" & strData & ")"
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        Else
                            ReDim Preserve strPrtData(i) As String
                            strPrtData(i) = ""
                            strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                            strPrtData(i) = strPrtData(i) & "^CI26"
                            strPrtData(i) = strPrtData(i) & strAFont
                            strPrtData(i) = strPrtData(i) & "^FH"
                            If strNamePrt = "Y" Then
                                strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                            Else
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                            End If
                            strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                            i = i + 1
                        End If
                    End If
                Next
            End With

            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            strPPTopLabel = ""
            strOutput = strHeader & strOutput & strFooter
            
            
            strOutput = Replace(strOutput, "~", "_7E")
            strOutput = Replace(strOutput, "℃", "'C")
        
            comEqp.Output = strOutput
            Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
            txtTopPrtVal.Text = strOutput
            
            intCnt = 0
            '재출력용
            ReDim Preserve strPrintVal(intCnt)
            strPrintVal(intCnt) = strOutput

            blnPrint = True
            txtComm.Text = strOutput
'            strOutput = ""

    End Select
   
    '-- PP Box라벨(바코드) 출력  ==> 연속용지일 경우만 가능
    If chkReelPrint.Value = "1" Then
        Erase strPrtData
        i = 0
        
        strPlusXPos = 680
        '-- 측면 바코드 만들기
        With spdScan
            '내부 바코드 출력(상단에 2개 포함하여 출력)
            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
                Dim intMaxNo    As Integer
                Dim strDate     As String
                
                i = 0
                Erase strPrtData
                
                strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
                
                'MAX NO 조회만
                Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "P")
                If AdoRs.RecordCount = 0 Then
                    'INSERT
                    intMaxNo = 1
                Else
                    'UPDATE
                    intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
                    intMaxNo = intMaxNo + 1
                End If
                AdoRs.Close
            
                'PP 박스 M+200302(제조일자) + 100(박스번호) (100번부터 시작)
                txtInBarcode.Text = "M" & Mid(strDate, 3, 6) & Format(100 + intMaxNo, "000")
                
                strXPos = 50: strYPos = 70
                i = 0
                ReDim Preserve strPrtData(i) As String
                strPrtData(i) = ""
                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                strPrtData(i) = strPrtData(i) & "^CI26"
                'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                strPrtData(i) = strPrtData(i) & "^BY2,3,100^BC"
                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                i = i + 1
                ReDim Preserve strPrtData(i) As String
                strPrtData(i) = ""
                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
                strPrtData(i) = strPrtData(i) & "^CI26"
                'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                strPrtData(i) = strPrtData(i) & "^BY2,3,100^BC"
                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                
                i = i + 1
                
                '내부바코드를 사용한다.
                strBarcode = txtInBarcode.Text
                
                For intY = 1 To .MaxRows
                    .Row = intY
                    .Col = 1
                    
                    If intY = 1 Then
                        strXPos = 50: strYPos = 250
                    Else
                        If intY Mod 2 = 0 Then
                            strXPos = strPlusXPos:  strYPos = strYPos
                        Else
                            strXPos = 50:           strYPos = strYPos + 160
                        End If
                    End If
                    
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BY2,3,100^BC"
                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Next
            
                strOutput = ""
                For J = 0 To UBound(strPrtData)
                    strOutput = strOutput & strPrtData(J)
                Next
                txtSidePrtVal.Text = strHeader & strOutput & strFooter
                strOutput = ""
            
            Else

                '-- 연속용지 출력 시(TP203C)
                '-- PP BOX 옆면
                i = 0
                For intY = 1 To .MaxRows
                    .Row = intY
                    .Col = 1
                    If intY = 1 Then
                        strXPos = 50: strYPos = 1000 '650
                    ElseIf intY = 2 Then
                        strXPos = 400: strYPos = 1000 '650
                    ElseIf intY = 3 Then
                        strXPos = 50: strYPos = 1120 '800
                    ElseIf intY = 4 Then
                        strXPos = 400: strYPos = 1120 '800
                    ElseIf intY = 5 Then
                        strXPos = 50: strYPos = 1240 '950
                    ElseIf intY = 6 Then
                        strXPos = 400: strYPos = 1240 '950
                    ElseIf intY = 7 Then
                        strXPos = 50: strYPos = 1360 '1100
                    ElseIf intY = 8 Then
                        strXPos = 400: strYPos = 1360 '1100
                    ElseIf intY = 9 Then
                        strXPos = 50: strYPos = 1480 '1250
                    ElseIf intY = 10 Then
                        strXPos = 400: strYPos = 1480 '1250
                    End If
                    
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BY1,3,80^BCN"
                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    'Debug.Print strPrtData(i)
                    i = i + 1
                Next
            End If
        End With

        If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
            '측면출력 데이터 저장
'            Dim i           As Integer
'            Dim strOutput   As String
'            Dim strData(0)  As Variant
            'If txtSidePrtVal.Text <> "" Then
                'PROD_ORDER_DT  strDate
                'PROD_CD        lblstrPrtLabelName.Caption
                'PROD_REEL_BAR
                'PROD_PP_BAR
                'PROD_ICE_BAR
                'GUBUN          PPBOX
                'PRINT_DATA     txtSidePrtVal.Text
                'OUT_YN         N
                
                'spdScan
                
            '    strOutput = txtSidePrtVal.Text
            'End If
'            If strOutput <> "" Then
'                '출력수량
'                For i = 1 To txtTopPrtNo.Text
'                    comEqp.Output = strOutput
'
'                    If txtTopPrtVal.Text <> "" Then
'                        strData(0) = txtTopPrtVal.Text & ETX & strOutput
'                    Else
'                        strData(0) = strOutput
'                    End If
'
'                    Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
'
'                    Call SetPackTrack("", txtInBarcode.Text, strData)
'                Next
'            End If
        Else
            strOutput = ""
            For J = 0 To UBound(strPrtData)
                strOutput = strOutput & strPrtData(J)
            Next
            
            'PP라벨 상단라벨을 위에 측면을 아래에 놓고 한번에 출력
            If strPPTopLabel <> "" Then
                strOutput = strHeader & strPPTopLabel & strOutput & strFooter
            Else
                strOutput = strHeader & strOutput & strFooter
            End If
            
            '트래킹용
            ReDim Preserve strTrackBC(intCnt)
            strTrackBC(intCnt) = strBarcode
            
            '재출력용
            ReDim Preserve strPrintVal(intCnt)
            strPrintVal(intCnt) = strOutput
            
            If strOutput <> "" Then
                If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0011" _
                 Or lblstrPrtLabelName.Caption = "P0018" Or lblstrPrtLabelName.Caption = "P0019" Or lblstrPrtLabelName.Caption = "P0020" Then
                    strOutput3 = strOutput
                Else
                    comEqp.Output = strOutput
                    Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
                End If
            End If
            
            blnPrint = True
            txtComm.Text = txtComm.Text & vbCrLf & strOutput
            
            strOutput = ""
            strPPTopLabel = ""
        End If
    End If
    
    If blnPrint = True Then
        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
        'MAX NO
        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "P")
        If AdoRs.RecordCount = 0 Then
            'INSERT
            intMaxNo = 1
            If Set_MAX_NO("IN", "P", intMaxNo) Then
            End If
        Else
            'UPDATE
            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
            intMaxNo = intMaxNo + 1
            If Set_MAX_NO("UP", "P", intMaxNo) Then
            End If
        End If
        AdoRs.Close
    
        'PP 박스 M+200302(제조일자) + 100(박스번호) (100번부터 시작)
        txtInBarcode.Text = "M" & Mid(strDate, 3, 6) & Format(100 + intMaxNo, "000")
        strInBarcode = ""
        
        '-- PP BOX 상단에 바코드가 없는 경우나  PP BOX 상단에 바코드가 있는 경우라도 바코드내에 릴에 대한 정보가 없다면 ==> 측면 바코드를 내부바코드로 사용한다.
        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Or lblstrPrtLabelName.Caption = "P0011" _
             Or lblstrPrtLabelName.Caption = "P0018" Or lblstrPrtLabelName.Caption = "P0019" Or lblstrPrtLabelName.Caption = "P0020" Then
            strInBarcode = txtInBarcode.Text
            
            '상단 라벨에 바코드 없음, 측면바코드 없음
            If lblstrPrtLabelName.Caption = "P0003" Then
                '트래킹용
                ReDim Preserve strTrackBC(intCnt)
                strTrackBC(intCnt) = strInBarcode
                
                '재출력용
                ReDim Preserve strPrintVal(intCnt)
                strPrintVal(intCnt) = strOutput
            End If
            
            '내부 관리용 바코드 출력한다
            If strInBarcode <> "" Then
                strInBarData = GetMakeInBar(strInBarcode)
                If strInBarData <> "" Then
                    '출력
                    If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0011" _
                        Or lblstrPrtLabelName.Caption = "P0018" Or lblstrPrtLabelName.Caption = "P0019" Or lblstrPrtLabelName.Caption = "P0020" Then
                        
                        strInBarData = Replace(strInBarData, "^XZ", "")
                        strOutput3 = Replace(strOutput3, "^XA", "")
                        strInBarData = strInBarData & strOutput3
                    End If
                    comEqp.Output = strInBarData
                    
                    Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strInBarData, "A")
                    
                    '재출력용
                    '-2020.07.27 수정
                    'strPrintVal(intCnt) = strPrintVal(intCnt) & ETX & strInBarData
                    strPrintVal(intCnt) = strInBarData
                    
                End If
            End If
        ElseIf lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Or lblstrPrtLabelName.Caption = "P0010" Then
            strInBarcode = txtInBarcode.Text
            strBarcode = ""
            
            strPrintVal(0) = strPrintVal(0) & ETX & txtSidePrtVal.Text
        End If
        
        '트래킹 저장
        Call SetPackTrack(strBarcode, strInBarcode, strPrintVal)
        
        txtPPBoxNo.Text = intMaxNo + 1
        txtScanCount.Text = "0"
        spdScan.MaxRows = 0
        
        For intCnt = 0 To 34 '9
            imgPpBar(intCnt).Visible = False
        Next
        
    End If
    
End Sub

Private Sub SetPackTrack(ByVal pPPBarcode As String, ByVal pPPINBarcode As String, ByVal pPPPrtVal As Variant)
    Dim intCnt      As Integer
    Dim intMaxNo    As Integer
    
    With spdScan
        For intCnt = 1 To .MaxRows
            .Row = intCnt
            .Col = 1
            gPackTrack.ORDERDT = Format(txtProdOrderDt.Text, "yyyymmdd")     'Key
            gPackTrack.PRODCD = txtProdCd.Text                               'Key
            gPackTrack.REELBAR = .Text
            gPackTrack.PPBAR = pPPBarcode
            gPackTrack.ICEBAR = ""
            gPackTrack.PPBARIN = pPPINBarcode
            gPackTrack.ICEBARIN = ""
            gPackTrack.LOTNO = txtLotNo.Text
            gPackTrack.REELPRTID = ""
            gPackTrack.REELPRTDT = ""
            gPackTrack.PPPRTID = gKUKDO.USERID
            gPackTrack.PPPRTDT = ""
            gPackTrack.ICEPRTID = ""
            gPackTrack.ICEPRTDT = ""
            '재출력용
            gPackTrack.REELVAL = ""
            '제일 처음 릴바코드에 저장한다.
            If intCnt = 1 Then
                gPackTrack.PPVAL = pPPPrtVal(0)
            Else
                gPackTrack.PPVAL = ""
            End If
            gPackTrack.ICEVAL = ""
            
            '트래킹 저장
            '-- PP Box 는 Insert 없음
            If Set_Pack_Track("UP", "P") Then
            End If

        Next
    End With
      
End Sub

Private Sub cmdSearch_Click()
    Dim strFromDt    As String
    Dim strToDt      As String
    Dim strYN        As String
    
    strFromDt = Format(dtpFromDate, "yyyymmdd")
    strToDt = Format(dtpToDate, "yyyymmdd")
    
    Call cmdClear_Click
    
    Call GetReelOrderList_PP(strFromDt, strToDt, "", "", "P")

End Sub


Private Sub cmdSidePrint_Click()
    
    Dim i           As Integer
    Dim strOutput   As String
    Dim strData(0)  As Variant
    
    'Call setSidePrint(False)
    
    strOutput = txtSidePrtVal.Text
    If strOutput <> "" Then
        '출력수량
        For i = 1 To txtTopPrtNo.Text
            
            comEqp.Output = strOutput
            
            If txtTopPrtVal.Text <> "" Then
                strData(0) = txtTopPrtVal.Text & ETX & strOutput
            Else
                strData(0) = strOutput
            End If
            
            'Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
            
            'Call SetPackTrack("", txtInBarcode.Text, strData)
                    
        Next
        'txtSidePrtVal.Text = ""
    End If

End Sub


'Private Function setSidePrint(ByVal pUpFlag As Boolean) As String
'    Dim i           As Integer
'    Dim strOutput   As String
'    Dim strData(0)  As Variant
'
'    setSidePrint = ""
'    strOutput = txtSidePrtVal.Text
'    If strOutput <> "" Then
'        '출력수량
'        For i = 1 To txtTopPrtNo.Text
'
'            If pUpFlag = False Then
'                comEqp.Output = strOutput
'            End If
'
'            If txtTopPrtVal.Text <> "" Then
'                strData(0) = txtTopPrtVal.Text & ETX & strOutput
'            Else
'                strData(0) = strOutput
'            End If
'
'            Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
'
'            Call SetPackTrack("", txtInBarcode.Text, strData)
'
'            setSidePrint = strData(0)
'
'        Next
'        txtSidePrtVal.Text = ""
'    End If
'End Function


Private Sub cmdTopPrint_Click()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strFooter   As String
    
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strType     As String
    Dim strNamePrt  As String
    Dim strXPos     As String
    Dim strXPos2    As String
    Dim strXPos_B   As String
    Dim strXPos_N   As String
    Dim strYPos     As String
    Dim strData     As String
    Dim strBarType  As String
    Dim strFont     As String
    Dim strRot      As String
    Dim strSlt      As String
    Dim strLot      As String
    Dim strLotSub   As String
    Dim strPlusXPos As String
    Dim strAFont    As String
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim blnPrint    As Boolean
    Dim strPrtData()    As String
    '- TP500B 상부바코드 두개찍는데 사용
    Dim strMatCd    As String
    Dim strLotNo    As String
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    
    Dim strPPTopLabel       As String
    
    Erase strPrtData
    Erase strTrackBC
    blnPrint = False
    strBarcode = ""
    strOutput = ""
    strPPTopLabel = ""
    intCnt = 0
    i = 0
    
    
    strPlusXPos = 630
    strAFont = "^A0N,60,50"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strHeader = strHeader & "^MD9"
    
    strFooter = "^XZ" & vbLf

'Case "P0006", "P0007"
    strAFont = "^A0N,45,35"
    With spdPrtReelDetail
        For intRow = 1 To .MaxRows
            '좌측 출력
            strType = GetText(spdPrtReelDetail, intRow, 1)
            strData = GetText(spdPrtReelDetail, intRow, 3)
            strXPos = GetText(spdPrtReelDetail, intRow, 4)
            strYPos = GetText(spdPrtReelDetail, intRow, 5)
            strBarType = GetText(spdPrtReelDetail, intRow, 6)   '바코드타입
            strFont = GetText(spdPrtReelDetail, intRow, 7)      '폰트
            strRot = GetText(spdPrtReelDetail, intRow, 8)       '회전
            strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
            strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
            
            Select Case strRot
                Case "0":   strRot = "N"
                Case "90":  strRot = "R"
                Case "180": strRot = "I"
                Case "270": strRot = "B"
            End Select
            
            If strType = "바코드" Then
                strLotSub = strSlt & Format(intPrt, "00")
                strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                If Mid(strBarType, 1, 1) = "1" Then
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                    strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Else
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                End If
                strBarcode = strData
            ElseIf strType = "PartsID" Then
                strLotSub = strSlt & Format(intPrt, "00")
                If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
                    strPrtData(i) = strPrtData(i) & "^BY3,3,100^BC"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                Else
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                End If
            Else
                If strType = "LotNo" Then
                    strLot = strData
                    
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & strAFont
                    strPrtData(i) = strPrtData(i) & "^FH"
                    If strNamePrt = "Y" Then
                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                    Else
                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                    End If
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                    'strLotNo = strData
                ElseIf strType = "Material code" Then
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & strAFont
                    strPrtData(i) = strPrtData(i) & "^FH"
                    If strNamePrt = "Y" Then
                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                    Else
                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                    End If
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                    strMatCd = strData
                ElseIf strType = "Name" Then
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^A0N,65,55" 'strAFont
                    strPrtData(i) = strPrtData(i) & "^FH"
                    If strNamePrt = "Y" Then
                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                    Else
                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                    End If
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                    strMatCd = strData
                ElseIf strType = "수량" Then
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & "^A0N,65,55" 'strAFont
                    strPrtData(i) = strPrtData(i) & "^FH"
                    If strNamePrt = "Y" Then
                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                    Else
                        strPrtData(i) = strPrtData(i) & "^FD" & txtMaxTot.Text
                    End If
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                    strMatCd = strData
                Else
                    ReDim Preserve strPrtData(i) As String
                    strPrtData(i) = ""
                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                    strPrtData(i) = strPrtData(i) & "^CI26"
                    strPrtData(i) = strPrtData(i) & strAFont
                    strPrtData(i) = strPrtData(i) & "^FH"
                    If strNamePrt = "Y" Then
                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                    Else
                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                    End If
                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                    i = i + 1
                End If
            End If
        Next
    End With

    strOutput = ""
    For J = 0 To UBound(strPrtData)
        strOutput = strOutput & strPrtData(J)
    Next
    
    strOutput = strHeader & strOutput & "^XZ" & vbLf
    '출력수량
    For i = 1 To txtTopPrtNo.Text
        comEqp.Output = strOutput
        
        Call SetPrtData("PPBOX" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strOutput, "A")
        
    Next
    
    blnPrint = True
    txtComm.Text = strOutput
    strOutput = ""
   
    fraTop.Visible = False

'    '-- PP Box라벨(바코드) 출력
'    If chkReelPrint.Value = "1" Then
'        Erase strPrtData
'        i = 0
'
'        With spdScan
'            '내부 바코드 출력(상단에 2개 포함하여 출력)
'            If lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Then
'                Dim intMaxNo    As Integer
'                Dim strDate     As String
'
'                strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
'
'                '-- PP BOX 상단에 바코드가 없는 경우 측면 바코드를 내부바코드로 사용한다.
'                'strBarcode = txtInBarcode.Text
'                'txtInBarcode.Text = "M" & Mid(strDate, 3, 6) & Format(100 + CCur(txtPPBoxNo.Text), "000")
'                'strBarcode = txtInBarcode.Text
'
'                strXPos = 50: strYPos = 100
'                i = 0
'                ReDim Preserve strPrtData(i) As String
'                strPrtData(i) = ""
'                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                strPrtData(i) = strPrtData(i) & "^CI26"
'                'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
'                strPrtData(i) = strPrtData(i) & "^BY2,3,130^BC"
'                strPrtData(i) = strPrtData(i) & "^FD" & strMatCd
'                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
'                i = i + 1
'                ReDim Preserve strPrtData(i) As String
'                strPrtData(i) = ""
'                strPrtData(i) = strPrtData(i) & "^FO" & strPlusXPos & "," & strYPos
'                strPrtData(i) = strPrtData(i) & "^CI26"
'                'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
'                strPrtData(i) = strPrtData(i) & "^BY2,3,130^BC"
'                strPrtData(i) = strPrtData(i) & "^FD" & txtInBarcode.Text
'                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
'                i = i + 1
'
'                For intCnt = 1 To .MaxRows
'                    .Row = intCnt
'                    .Col = 1
'
'                    If intCnt = 1 Then
'                        strXPos = 50: strYPos = 900
'                    Else
'                        If intCnt Mod 2 = 0 Then
'                            strXPos = strPlusXPos:  strYPos = strYPos
'                        Else
'                            strXPos = 50:          strYPos = strYPos + 200
'                        End If
'                    End If
'
'                    ReDim Preserve strPrtData(i) As String
'                    strPrtData(i) = ""
'                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                    strPrtData(i) = strPrtData(i) & "^CI26"
'                    'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
'                    strPrtData(i) = strPrtData(i) & "^BY2,3,130^BC"
'
'                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
'                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
'                    i = i + 1
'                Next
'            Else
'                'If lblstrPrtLabelName.Caption = "P0006" Then
'                For intCnt = 1 To .MaxRows
'                    .Row = intCnt
'                    .Col = 1
'
'                    If intCnt = 1 Then
'                        strXPos = 50: strYPos = 900
'                    Else
'                        If intCnt Mod 2 = 0 Then
'                            strXPos = strPlusXPos:  strYPos = strYPos
'                        Else
'                            strXPos = 50:          strYPos = strYPos + 200
'                        End If
'                    End If
'
'                    ReDim Preserve strPrtData(i) As String
'                    strPrtData(i) = ""
'                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                    strPrtData(i) = strPrtData(i) & "^CI26"
'                    'strPrtData(i) = strPrtData(i) & "^BC" & "N,80,Y,N,N"
'                    strPrtData(i) = strPrtData(i) & "^BY2,3,130^BC"
'                    strPrtData(i) = strPrtData(i) & "^FD" & .Text
'                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
'                    i = i + 1
'                Next
'            End If
'        End With
'
'        For J = 0 To UBound(strPrtData)
'            strOutput = strOutput & strPrtData(J)
'        Next
'
'        '측면을 먼저 출력
'        'strOutput = strHeader & strOutput & "^XZ" & vbLf
'        'comEqp.Output = strOutput
'
'        'PP라벨을 나중에 출력
'        If strPPTopLabel <> "" Then
'            strOutput = strHeader & strPPTopLabel & strOutput & strFooter
'        Else
'            strOutput = strHeader & strOutput & strFooter
'        End If
'
'        If UBound(strPrtData) > 0 Then
'            comEqp.Output = strOutput
'        End If
'        strOutput = ""
'        strPPTopLabel = ""
'
'        ReDim Preserve strTrackBC(intCnt)
'        strTrackBC(intCnt) = strBarcode
'        intCnt = intCnt + 1
'
'        blnPrint = True
'        txtComm.Text = txtComm.Text & vbCrLf & strOutput
'        strOutput = ""
'
'    End If
'
'    If blnPrint = True Then
'
'        strDate = Format(txtProdOrderDt.Text, "yyyymmdd")
'
'        'MAX NO
'        Set AdoRs = Get_MAX_NO(strDate, lblstrPrtLabelName.Caption, "P")
'        If AdoRs.RecordCount = 0 Then
'            'INSERT
'            intMaxNo = 1
'            If Set_MAX_NO("IN", "P", intMaxNo) Then
'            End If
'        Else
'            'UPDATE
'            intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
'            intMaxNo = intMaxNo + 1
'            If Set_MAX_NO("UP", "P", intMaxNo) Then
'            End If
'        End If
'        AdoRs.Close
'
'        'PP 박스 M+200302(제조일자) + 100(박스번호) (100번부터 시작)
'        txtInBarcode.Text = "M" & Mid(strDate, 3, 6) & Format(100 + intMaxNo, "000")
'
'        '-- PP BOX 상단에 바코드가 없는 경우 측면 바코드를 내부바코드로 사용한다.
'        If lblstrPrtLabelName.Caption = "P0001" Or lblstrPrtLabelName.Caption = "P0002" Or lblstrPrtLabelName.Caption = "P0003" Or lblstrPrtLabelName.Caption = "P0006" Or lblstrPrtLabelName.Caption = "P0007" Then
'            strBarcode = txtInBarcode.Text
'            Call cmdMakeBar_Click
'        End If
'
'        Call SetPackTrack(strBarcode)
'
'        txtPPBoxNo.Text = intMaxNo + 1
'        txtScanCount.Text = "0"
'        spdScan.MaxRows = 0
'
'        For intCnt = 0 To 9
'            imgPpBar(intCnt).Visible = False
'        Next
'
'    End If
    
End Sub

Private Sub cmdUnvisible_Click()
    
    fraTop.Visible = False

End Sub

Private Sub cmdView_Click()
    
    If txtComm.Visible = False Then
        txtComm.Visible = True
    Else
        txtComm.Visible = False
    End If
    
End Sub

Private Sub cmdVisible_Click()
    
    fraTop.Visible = True
    
End Sub

Private Sub Form_Load()

    
    Unload frmPrtReel
    Unload frmPrtICE
    Unload frmPrtReprint
    
    Call CtlInitializing
    
    '-- 통신열기
    Call OpenCommunication
    
End Sub

Private Sub OpenCommunication()

On Error GoTo ErrHandle
    
'    frmPrtReel.comEqp.PortOpen = False
'    frmPrtICE.comEqp.PortOpen = False

    comEqp.CommPort = gComm.COMPORT
    comEqp.RTSEnable = gComm.RTSEnable
    comEqp.DTREnable = gComm.DTREnable
    comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

    If comEqp.PortOpen Then
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결성공"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
    End If

    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "포트 연결실패"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            
            Resume Next
        Else
            
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "위    치 : " & "Public Sub OpenCommunication()" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If


End Sub

'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim i           As Integer
    
    With spdPrtReel
        Call SetText(spdPrtReel, "선택", 0, 1):              .ColWidth(1) = 0
        Call SetText(spdPrtReel, "Lot No", 0, 2):            .ColWidth(2) = 12
        Call SetText(spdPrtReel, "제조일자", 0, 3):          .ColWidth(3) = 10
        Call SetText(spdPrtReel, "공정No", 0, 4):            .ColWidth(4) = 0
        Call SetText(spdPrtReel, "제품코드", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdPrtReel, "제품명", 0, 6):            .ColWidth(6) = 12
        Call SetText(spdPrtReel, "포장코드", 0, 7):          .ColWidth(7) = 0
        Call SetText(spdPrtReel, "메모", 0, 8):              .ColWidth(8) = 0
        Call SetText(spdPrtReel, "작업내용설명", 0, 9):      .ColWidth(9) = 0 'Roll정보
        Call SetText(spdPrtReel, "길이", 0, 10):             .ColWidth(10) = 4
        Call SetText(spdPrtReel, "SLT No", 0, 11):           .ColWidth(11) = 0
        Call SetText(spdPrtReel, "수량", 0, 12):             .ColWidth(12) = 0
        Call SetText(spdPrtReel, "고객사", 0, 13):           .ColWidth(13) = 10
        Call SetText(spdPrtReel, "작업완료여부", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdPrtReel, "라벨코드", 0, 15):         .ColWidth(15) = 10
        Call SetText(spdPrtReel, "입력자", 0, 16):           .ColWidth(16) = 0
        Call SetText(spdPrtReel, "입력일시", 0, 17):         .ColWidth(17) = 0
        Call SetText(spdPrtReel, "수정자", 0, 18):           .ColWidth(18) = 0
        Call SetText(spdPrtReel, "수정일시", 0, 19):         .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    With spdPrtReelDetail
        Call SetText(spdPrtReelDetail, "항목", 0, 1):          .ColWidth(1) = 18
        Call SetText(spdPrtReelDetail, "순서", 0, 2):          .ColWidth(2) = 4
        Call SetText(spdPrtReelDetail, "내용", 0, 3):          .ColWidth(3) = 32
        Call SetText(spdPrtReelDetail, "X", 0, 4):          .ColWidth(4) = 0
        Call SetText(spdPrtReelDetail, "Y", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdPrtReelDetail, "", 0, 6):            .ColWidth(6) = 0
        Call SetText(spdPrtReelDetail, "", 0, 7):          .ColWidth(7) = 0
        Call SetText(spdPrtReelDetail, "", 0, 8):          .ColWidth(8) = 0
        Call SetText(spdPrtReelDetail, "", 0, 9):          .ColWidth(9) = 0
        Call SetText(spdPrtReelDetail, "", 0, 10):      .ColWidth(10) = 0
        Call SetText(spdPrtReelDetail, "", 0, 11):           .ColWidth(11) = 0
        Call SetText(spdPrtReelDetail, "", 0, 12):             .ColWidth(12) = 0
        Call SetText(spdPrtReelDetail, "", 0, 13):           .ColWidth(13) = 0
        Call SetText(spdPrtReelDetail, "", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdPrtReelDetail, "", 0, 15):         .ColWidth(15) = 0
        Call SetText(spdPrtReelDetail, "", 0, 16):           .ColWidth(16) = 0
        Call SetText(spdPrtReelDetail, "", 0, 17):         .ColWidth(17) = 0
        Call SetText(spdPrtReelDetail, "", 0, 18):           .ColWidth(18) = 0
        Call SetText(spdPrtReelDetail, "", 0, 19):         .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    With spdScan
        Call SetText(spdScan, "바코드", 0, 1):          .ColWidth(1) = 18
        .MaxRows = 0
        .MaxCols = 1
        .RowHeight(-1) = 12
    End With
    
'    With spdRegOrderDetail
'        Call SetText(spdRegOrderDetail, "제조일자", 0, 1):        .ColWidth(1) = 0
'        Call SetText(spdRegOrderDetail, "순번", 0, 2):            .ColWidth(2) = 0
'        Call SetText(spdRegOrderDetail, "제품코드", 0, 3):        .ColWidth(3) = 0
'        Call SetText(spdRegOrderDetail, "SLT No", 0, 4):          .ColWidth(4) = 0
'        Call SetText(spdRegOrderDetail, "일련번호", 0, 5):        .ColWidth(5) = 8
'        Call SetText(spdRegOrderDetail, "SLT내용", 0, 6):         .ColWidth(6) = 28
'        Call SetText(spdRegOrderDetail, "시작번호", 0, 7):        .ColWidth(7) = 10
'        Call SetText(spdRegOrderDetail, "끝번호", 0, 8):          .ColWidth(8) = 10
'        Call SetText(spdRegOrderDetail, "", 0, 9):                .ColWidth(9) = 0
'        Call SetText(spdRegOrderDetail, "No", 0, 10):             .ColWidth(10) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 11):               .ColWidth(11) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 12):               .ColWidth(12) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 13):               .ColWidth(13) = 0
'        Call SetText(spdRegOrderDetail, "", 0, 14):               .ColWidth(14) = 0
'        Call SetText(spdRegOrderDetail, "사용여부", 0, 15):       .ColWidth(15) = 0
'        Call SetText(spdRegOrderDetail, "입력자", 0, 16):         .ColWidth(16) = 0
'        Call SetText(spdRegOrderDetail, "입력일시", 0, 17):       .ColWidth(17) = 0
'        Call SetText(spdRegOrderDetail, "수정자", 0, 18):         .ColWidth(18) = 0
'        Call SetText(spdRegOrderDetail, "수정일시", 0, 19):       .ColWidth(19) = 0
'
'        .MaxRows = 0
'    End With
    
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtLotNo.Text = ""
    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
    'txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    
    txtPPBoxNo.Text = ""
    txtPrtPPBoxNo.Text = ""
    txtReelBarcode.Text = ""
    txtMaxTot.Text = "0"
    txtScanCount.Text = "0"

    chkReelPrint.Value = "0"
    txtMsg.Text = ""
    
    'txtReelQTY.Text = ""
    
    gSORT = 0

End Sub



Private Sub lblComp_DblClick(Index As Integer)
    Dim i As Integer
    Dim strNum As String
    
    strNum = txtReelBarcode.Text
    For i = 1 To 10
        txtReelBarcode.Text = strNum
        Call txtReelBarcode_KeyPress(vbKeyReturn)
        Select Case txtProdNm
            Case "TP203C(KCF)"
                strNum = Mid(strNum, 1, 15) & Format(CCur(Mid(strNum, 16, 2)) + 1, "00") & Mid(strNum, 18)
            Case "TP203C(KAF)"
                strNum = Mid(strNum, 1, 15) & Format(CCur(Mid(strNum, 16, 2)) + 1, "00") & Mid(strNum, 18)
            Case "KAF-TP408A"
                '32401002000001KK370101020000
                strNum = Mid(strNum, 1, 20) & Format(CCur(Mid(strNum, 21, 2)) + 1, "00") & Mid(strNum, 23)
            Case "KAF-TP400E", "KAF-TP500B"
                '0201002240E1B4AAN1K371SE101270
                strNum = Mid(strNum, 1, 24) & Format(CCur(Mid(strNum, 25, 3)) + 1, "000") & Mid(strNum, 28)
            Case "KAF-TP500B", "KAF-TP500B"
                strNum = Mid(strNum, 1, 12) & Format(CCur(Mid(strNum, 13, 3)) + 1, "000") & Mid(strNum, 16)
        
        End Select
    Next
    txtReelBarcode.Text = strNum
    
End Sub

Private Sub spdPrtReel_Click(ByVal Col As Long, ByVal Row As Long)
    Dim pAdoRS1      As ADODB.Recordset
    Dim pAdoRS2      As ADODB.Recordset
    Dim i               As Integer
    Dim J               As Integer
    Dim strPrtSide      As String
    Dim strBarCd        As String
    Dim strDate         As String
    Dim strQty          As String
    Dim strProdLabelCd  As String
    Dim strProdCd       As String
    Dim strCompCd       As String
    Dim strTmp          As String
    Dim strGu           As String
    Dim strContents     As String
    Dim strLeft         As String
    Dim strLotNo        As String
    Dim strLotNoFull    As String
    '자재코드
    Dim strBarData      As String
    Dim strProdName     As String
    Dim strProdLen      As String
    Dim strProdMatCd    As String
    Dim strExMonth      As String
    Dim strProdSize     As String
    Dim strChimeiCd     As String
    Dim strVendorCd     As String
    Dim strProdLineFA   As String
    Dim strProdSlitFA   As String
    Dim strContYN       As String
    Dim strPcnNO        As String
    Dim strMaxTot       As String
    Dim strExDate       As String
    Dim strProdTemp     As String
    Dim strPrtLabelName As String
    Dim strProdCalLen   As String
    Dim intMaxNo        As Integer
    Dim strPrtLbllNm    As String
    Dim strTmp1         As String
    
    If Row = 0 Then
        If Col = 1 Then
            If GetText(spdPrtReel, 1, 1) = "1" Then
                For i = 1 To spdPrtReel.DataRowCnt
                    Call SetText(spdPrtReel, "0", i, 1)
                Next
            Else
                For i = 1 To spdPrtReel.DataRowCnt
                    Call SetText(spdPrtReel, "1", i, 1)
                Next
            End If
        Else
            '-- 정렬 추가
            Call SetSpreadSort(spdPrtReel, 0)
        End If
        Exit Sub
    End If
    
    For i = 0 To 15
        'barReel.Visible = False
        'barPart.Visible = False
        lblTitle(i).Visible = False
    Next
    
    For i = 0 To 34 '11
        imgPpBar(i).Visible = False
    Next
    
    txtReelBarcode.Text = ""
    txtInBarcode.Text = ""
    txtScanCount.Text = "0"
    spdScan.MaxRows = 0
    
    imgBar1.Visible = False
    imgBar2.Visible = False
    imgQrBar.Visible = False
    
    strPrtLbllNm = GetText(spdPrtReel, Row, 6) & "|" & GetText(spdPrtReel, Row, 10)
    
    strPrtLabelName = GetText(spdPrtReel, Row, 5)
    lblstrPrtLabelName.Caption = strPrtLabelName
    
    strDate = GetText(spdPrtReel, Row, 3)
    txtProdOrderDt.Text = strDate
    strProdCd = GetText(spdPrtReel, Row, 5)
    txtProdCd.Text = strProdCd
    txtProdNm.Text = GetText(spdPrtReel, Row, 6)
    strProdLen = GetText(spdPrtReel, Row, 10)
    strProdLen = strProdLen * 100 '미터를 cm으로 변환
    txtProdLen.Text = strProdLen
'    txtProdPosNo.Text = GetText(spdPrtReel, Row, 4)
    txtPackNm.Text = GetText(spdPrtReel, Row, 7)
    txtReelQTY.Text = GetText(spdPrtReel, Row, 12)
    txtSlittingNo.Text = GetText(spdPrtReel, Row, 11)
    txtCompNm.Text = GetText(spdPrtReel, Row, 13)
    strLotNo = GetText(spdPrtReel, Row, 2)
    txtLotNo.Text = strLotNo
    strProdLabelCd = GetText(spdPrtReel, Row, 15)
    strCompCd = GetText(spdPrtReel, Row, 18)
    strQty = txtReelQTY.Text
    
    gPackTrack.PRODCD = strProdCd   '5자리
    gPackTrack.LOTNO = strLotNo
    gPackTrack.ORDERDT = strDate    '8자리
   
    txtReelBarcode.Text = ""
    txtReelBarcode.Enabled = True
    'txtReelBarcode.BackColor = vbWhite
    txtTopPrtNo.Text = "1"
    txtSidePrtVal.Text = ""
    txtTopPrtVal.Text = ""
    
    fraTop.Visible = False
    
    'MAX NO 찾기
    Set AdoRs = Get_MAX_NO(gPackTrack.ORDERDT, gPackTrack.PRODCD, "P")
    If AdoRs.RecordCount = 0 Then
        intMaxNo = 1
    Else
        intMaxNo = AdoRs.Fields("MAX_NO").Value & ""
        intMaxNo = intMaxNo + 1
    End If
    AdoRs.Close
    
    txtPPBoxNo.Text = intMaxNo

    With spdPrtReelDetail
        .MaxRows = 0
    End With

    Set AdoRs = Get_LabelDetail(strProdLabelCd, "P")
            
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        '-- 제품정보 찾아오기
        SQL = ""
        SQL = SQL & "SELECT                                 " & vbCrLf
        SQL = SQL & "  M.PROD_NAME                          " & vbCrLf
        SQL = SQL & ", M.PROD_LENGTH                        " & vbCrLf
        SQL = SQL & ", M.PROD_MATERIAL_CD                   " & vbCrLf
        SQL = SQL & ", M.EXPIR_MONTH                        " & vbCrLf
        SQL = SQL & ", M.PROD_STOR_TEMP                     " & vbCrLf
        SQL = SQL & ", M.PROD_SIZE                          " & vbCrLf
        SQL = SQL & ", M.PROD_CHIMEI_PN                     " & vbCrLf
        SQL = SQL & ", M.VENDER_CD                          " & vbCrLf
        SQL = SQL & ", M.PROD_LINE_FA                       " & vbCrLf
        SQL = SQL & ", M.PROD_SLIT_FA                       " & vbCrLf
        SQL = SQL & ", M.PROD_CONTROL_YN                    " & vbCrLf
        SQL = SQL & ", M.PROD_PCN_NO                        " & vbCrLf
        SQL = SQL & ", L.PROD_MAX_TOT                       " & vbCrLf
        SQL = SQL & ", L.LABEL_PRT_SIDE                     " & vbCrLf
        SQL = SQL & "  FROM LBL_M_PROD M                    " & vbCrLf
        SQL = SQL & "     , LBL_LABEL_MASTER L              " & vbCrLf
        SQL = SQL & " WHERE M.PROD_CD  = '" & strProdCd & "'  " & vbCrLf
        SQL = SQL & "   AND M.COMP_CD  = '" & strCompCd & "'  " & vbCrLf
        SQL = SQL & "   AND M.USED_YN  = 'Y'                " & vbCrLf
        SQL = SQL & "   AND M.PROD_CD = L.PROD_CD           " & vbCrLf
        SQL = SQL & "   AND M.COMP_CD = L.COMP_CD           " & vbCrLf
        SQL = SQL & "   AND L.PROD_LABEL_TYPE = 'P'         " & vbCrLf
   
        Set pAdoRS2 = New ADODB.Recordset
        Call GetRecordset(AdoCn, SQL, pAdoRS2, "")
        If Not pAdoRS2 Is Nothing Then
            If Not pAdoRS2.EOF Then
                'strBarData = strBarData & Trim(pAdoRS2("PROD_MATERIAL_CD") & "")
                strBarData = pAdoRS2("PROD_MATERIAL_CD") & ""
                strProdName = pAdoRS2("PROD_NAME") & ""
                strProdLen = pAdoRS2("PROD_LENGTH") & ""
                strProdMatCd = pAdoRS2("PROD_MATERIAL_CD") & ""
                strExMonth = pAdoRS2("EXPIR_MONTH") & ""
                strExDate = DateAdd("m", strExMonth, strDate)
                strExDate = DateAdd("d", -1, strExDate)
                strProdTemp = pAdoRS2("PROD_STOR_TEMP") & ""
                strProdSize = pAdoRS2("PROD_SIZE") & ""
                strChimeiCd = pAdoRS2("PROD_CHIMEI_PN") & ""
                strVendorCd = pAdoRS2("VENDER_CD") & ""
                strProdLineFA = pAdoRS2("PROD_LINE_FA") & ""
                strProdSlitFA = pAdoRS2("PROD_SLIT_FA") & ""
                strContYN = pAdoRS2("PROD_CONTROL_YN") & ""
                strPcnNO = pAdoRS2("PROD_PCN_NO") & ""
                strMaxTot = pAdoRS2("PROD_MAX_TOT") & ""
                txtMaxTot.Text = strMaxTot
                strPrtSide = pAdoRS2("LABEL_PRT_SIDE") & ""
                If strPrtSide = "Y" Then
                    picSide.Visible = True
                    chkReelPrint.Value = "1"
                Else
                    picSide.Visible = False
                    chkReelPrint.Value = "0"
                End If
            End If
        End If
        pAdoRS2.Close
        Set pAdoRS2 = Nothing
        '-- 제품정보 찾아오기
        
        'chkReelPrint.Value = "1"
        
        Do Until AdoRs.EOF
            With spdPrtReelDetail
                .MaxRows = .MaxRows + 1
                strGu = AdoRs.Fields("LABEL_ITEM_GU").Value & ""
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_NAME").Value & "", .MaxRows, 1)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_SEQ").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_NAME").Value & "", .MaxRows, 3)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "", .MaxRows, 4)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "", .MaxRows, 5)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_GU").Value & "", .MaxRows, 6)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_FONT").Value & "", .MaxRows, 7)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_ROT").Value & "", .MaxRows, 8)
                Call SetText(spdPrtReelDetail, txtSlittingNo.Text, .MaxRows, 9)
                Call SetText(spdPrtReelDetail, AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "", .MaxRows, 10)
    
                '============================== 제품별로 분기 ==============================
                Select Case strPrtLabelName
                    Case "P0011", "P0018", "P0019", "P0020"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "P")
                            strBarData = ""
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드":       ' strBarData = strBarData
                                        Case "유효기간_년":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strExDate))
                                        Case "유효기간_월":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strExDate))
                                        Case "유효기간_일":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strExDate))
                                        Case "N/A":             strBarData = strBarData & "0000"
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "Slitting순번":    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
                                        Case "제품길이":
                                                                strProdCalLen = (1 * strProdLen)
                                                                'x=0의 갯수 : chr(x+64)
                                                                J = 0
                                                                For i = Len(strProdCalLen) To 1 Step -1
                                                                    If Mid(strProdCalLen, i, 1) = "0" Then
                                                                        J = J + 1
                                                                    End If
                                                                Next
                                                                strProdCalLen = Mid(strProdCalLen, 1, Len(strProdCalLen) - J)
                                                                strProdCalLen = Format(strProdCalLen, "0000")
                                                                strBarData = strBarData & strProdCalLen & Chr(J + 64)
                                    End Select
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Provider" Then
                            Call SetText(spdPrtReelDetail, gKUKDO.COMPNM, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Product Date" Then
                            If Len(strDate) = 10 Then
                                strTmp1 = Mid(strDate, 1, 4) & Mid(strDate, 6, 2) & Mid(strDate, 9, 2)
                                strTmp1 = Mid(strTmp1, 1, 4) & strGu & Mid(strTmp1, 5, 2) & strGu & Mid(strTmp1, 7, 2)
                            End If
                            Call SetText(spdPrtReelDetail, strTmp1, .MaxRows, 3)
                            strContents = strTmp1
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Chimei P/N" Then
                            Call SetText(spdPrtReelDetail, strChimeiCd, .MaxRows, 3)
                            strContents = strChimeiCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Product" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Total Quantity/Length" Then
                            'strContents = strQty & "Reels/" & CCur(strProdLen * strQty) * 100 & "cm"
                            strContents = "1Reel/" & strProdLen & "M"
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "발행시간" Then
                            strContents = Format(Now, "yyyy/mm/dd hh:mm:ss")
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Expiration Date" Then
                            Call SetText(spdPrtReelDetail, Year(strExDate) & strGu & Format(MONTH(strExDate), "00") & strGu & Format(Day(strExDate), "00"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        End If
                        
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgBar1.Visible = True
                            imgBar1.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar1.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar1.WIDTH = 4365
                        Else
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                            '항목명 출력여부
                            If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                            Else
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                            End If
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                        End If
                        strContents = ""
                        
                    '============== TP203C(ACF) ======================================================================
                    Case "P0001", "P0002"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "P")
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드":       ' strBarData = strBarData
                                        Case "유효기간_년":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strExDate))
                                        Case "유효기간_월":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strExDate))
                                        Case "유효기간_일":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strExDate))
                                        Case "N/A":             strBarData = strBarData & "0000"
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "Slitting순번":    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
                                        Case "제품길이":
                                                                strProdCalLen = (1 * strProdLen)
                                                                'x=0의 갯수 : chr(x+64)
                                                                J = 0
                                                                For i = Len(strProdCalLen) To 1 Step -1
                                                                    If Mid(strProdCalLen, i, 1) = "0" Then
                                                                        J = J + 1
                                                                    End If
                                                                Next
                                                                strProdCalLen = Mid(strProdCalLen, 1, Len(strProdCalLen) - J)
                                                                strProdCalLen = Format(strProdCalLen, "0000")
                                                                strBarData = strBarData & strProdCalLen & Chr(J + 64)
                                    End Select
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Provider" Then
                            Call SetText(spdPrtReelDetail, gKUKDO.COMPNM, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Chimei P/N" Then
                            Call SetText(spdPrtReelDetail, strChimeiCd, .MaxRows, 3)
                            strContents = strChimeiCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Total Quantity/Length" Then
                            'strContents = strQty & "Reels/" & CCur(strProdLen * strQty) * 100 & "cm"
                            strContents = "1Reel/" & CCur(strProdLen) * 100 & "cm"
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "발행시간" Then
                            strContents = Format(Now, "yyyy/mm/dd hh:mm:ss")
                            Call SetText(spdPrtReelDetail, strContents, .MaxRows, 3)
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Expiration Date" Then
                            Call SetText(spdPrtReelDetail, Year(strExDate) & strGu & Format(MONTH(strExDate), "00") & strGu & Format(Day(strExDate), "00"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        End If
                        
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgBar1.Visible = True
                            imgBar1.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar1.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar1.WIDTH = 4365
                        Else
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                            '항목명 출력여부
                            If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                            Else
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                            End If
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                        End If
                        strContents = ""
                                        

                    '============== CF-TP408A ======================================================================
                    Case "P0003"
                        chkReelPrint.Value = "0"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "P")
                        
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드"
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "ACF고객사":       strBarData = strBarData & "K"
                                        Case "제품순번":        strBarData = strBarData & "0001"
                                        Case "제품길이":        strBarData = strBarData & Format(strProdLen, "0000")
                                        Case "고객사내용":      strBarData = strBarData & "00"
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material code" Then
                            Call SetText(spdPrtReelDetail, strProdMatCd, .MaxRows, 3)
                            strContents = strProdMatCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "수량" Then
                            Call SetText(spdPrtReelDetail, strMaxTot, .MaxRows, 3)
                            strContents = strMaxTot
'                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
'                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
'                            strContents = strBarData
'                            gPackTrack.REELBAR = strContents
'                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "보관온도" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "생산일자" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "유효기간" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
'                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
'                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
'                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgBar1.Visible = True
                            imgBar1.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar1.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar1.WIDTH = 4365
                        Else
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                            '항목명 출력여부
                            If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                            Else
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                            End If
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                        End If
                            
                        strContents = ""
                        
                    '============== CF-TP400E ======================================================================
                    Case "P0004", "P0005", "P0008", "P0009"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "P")
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드"
                                        Case "Vendor코드":      strBarData = strBarData & strVendorCd
                                        Case "제조라인공장":    strBarData = strBarData & strProdLineFA
                                        Case "Sliting공장":     strBarData = strBarData & strProdSlitFA
                                        Case "관리선이탈여부":  strBarData = strBarData & strContYN
                                        Case "PCN차수":         strBarData = strBarData & strPcnNO
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        
                                        'Case "생산LOT":         strBarData = strBarData & Format(txtSlittingNo.Text, "0") & txtCompNm.Text & 100 + CCur(txtPPBoxNo.Text)
                                                                                                                                            '100부터 시작
                                        Case "생산LOT":         strBarData = strBarData & Format(txtSlittingNo.Text, "0") & txtCompNm.Text & 99 + CCur(txtPPBoxNo.Text)
                                        Case "REEL단위":        strBarData = strBarData & "R" & Mid(strProdLen, 1, 2)
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material code" Then
                            Call SetText(spdPrtReelDetail, strProdMatCd, .MaxRows, 3)
                            strContents = strProdMatCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "수량" Then
                            Call SetText(spdPrtReelDetail, strMaxTot, .MaxRows, 3)
                            strContents = strMaxTot
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "보관온도" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "생산일자" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "유효기간" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgQrBar.Visible = True
                            imgQrBar.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgQrBar.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgQrBar.WIDTH = 1000
                        Else
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                            '항목명 출력여부
                            If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                            Else
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                            End If
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                        End If

                        strContents = ""
                                                        
                    '============== CF-TP400E ======================================================================
                    Case "P0006", "P0007", "P0010"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            'strBarData = ""
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "P")
                        
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "ACF":         strBarData = "C"
                                        Case "생산LOT":     strBarData = strBarData & strLotNo
                                        Case "_":           strBarData = strBarData & "_"
                                        Case "P/N":         strBarData = strBarData & "101"
                                        Case "유효기간":    strBarData = strBarData & Format(strExDate, "yyyymmdd")
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== 바코드 정보 찾아오기 ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드값" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "생산일자" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "수량" Then
                            Call SetText(spdPrtReelDetail, strMaxTot, .MaxRows, 3)
                            strContents = strMaxTot
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material code" Then
                            Call SetText(spdPrtReelDetail, strProdMatCd, .MaxRows, 3)
                            strContents = strProdMatCd
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "보관온도" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "유효기간" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            imgBar1.Visible = True
                            imgBar1.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            imgBar1.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            imgBar1.WIDTH = 4365
                        Else
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                            '항목명 출력여부
                            If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                            Else
                                lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                            End If
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                            lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                        End If
                            
                        strContents = ""
            
                End Select
                
                AdoRs.MoveNext
                
'                MsgBox strPrtLbllNm & " 제품은 상단 PP Box 만 먼저 출력하세요", vbOKOnly + vbInformation, Me.Caption
                
            End With
            
        Loop
        
    
    End If
    
    AdoRs.Close

    If strPrtLabelName = "P0006" Or strPrtLabelName = "P0007" Or strPrtLabelName = "P0010" Then
        '상단만 출력하고 용지를 바꿔서 옆면 출력
'        txtReelBarcode.Text = ""
'        txtReelBarcode.Enabled = False
'        txtReelBarcode.BackColor = &HE0E0E0
'
'        cmdTopPrint.Visible = True
    
        fraTop.Visible = True
        txtReelBarcode.SetFocus
    Else
        txtReelBarcode.SetFocus
    End If
    
    

'    If strPrtLabelName = "P0006" Or strPrtLabelName = "P0007" Then
'        For i = 1 To 12
'            If i = 1 Then
'                txtReelBarcode.Text = strProdMatCd
'            Else
'                txtReelBarcode.Text = strLotNo
'
'            End If
'            Call txtReelBarcode_KeyPress(vbKeyReturn)
'
'        Next
'    End If
'
'    txtReelBarcode.SetFocus

'    Call GetReelOrderList(strFromDt, strToDt, "", "", "R")

End Sub


Private Sub txtReelBarcode_KeyPress(KeyAscii As Integer)
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim strDate     As String
    Dim strProdLen  As String
    Dim strProdCalLen   As String
    Dim strContents As String
    Dim intMaxNo    As Integer
    Dim strBarData  As String
    
    If KeyAscii = vbKeyReturn Then
        If txtReelBarcode.Text <> "" Then
            
            If Len(gPackTrack.ORDERDT) = 10 Then
                strDate = Format(gPackTrack.ORDERDT, "yyyymmdd")
            Else
                strDate = gPackTrack.ORDERDT
            End If
            '트래킹 테이블에 존재하는지 체크
            Set AdoRs = Get_Pack_Track(strDate, gPackTrack.PRODCD, txtReelBarcode.Text, "", "")
        
            If AdoRs.RecordCount = 0 Then
                'MsgBox txtReelBarcode.Text & "는 발행된 Reel 바코드가 아닙니다.", vbOKOnly + vbInformation, Me.Caption
                txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 발행된 Reel 바코드가 아닙니다." & vbCrLf
                txtReelBarcode.SelStart = 0
                txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                Exit Sub
            Else
                If AdoRs.Fields("PROD_PP_BAR").Value & "" <> "" Then
                    txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "는 트래킹된 Reel 바코드입니다." & vbCrLf
                    txtReelBarcode.SelStart = 0
                    txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                    Exit Sub
                End If
            End If
            
            With spdScan
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 1
                    If .Text = txtReelBarcode.Text Then
                        'MsgBox txtReelBarcode.Text & "와 같은 바코드가 있습니다.", vbOKOnly + vbInformation, Me.Caption
                        txtMsg.Text = txtMsg.Text & txtReelBarcode.Text & "와 같은 바코드가 있습니다." & vbCrLf
                        txtReelBarcode.SelStart = 0
                        txtReelBarcode.SelLength = Len(txtReelBarcode.Text)
                        Exit Sub
                    End If
                Next
                
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdScan, txtReelBarcode.Text, .MaxRows, 1)
                
                'PP Box당 갖고있는 Reel 수량
                'txtMaxTot.Text = AdoRs.RecordCount
            
            End With
                
'            imgPpBar(i - 1).Visible = True
'            txtScanCount.Text = txtScanCount.Text + 1
'            If txtMaxTot.Text = txtScanCount.Text Then
'                If chAutoPrint.Value = "1" Then
'                    Call cmdPrint_Click
'                End If
'            End If
'            txtReelBarcode.Text = ""
        
            imgPpBar(i - 1).Visible = True
            txtScanCount.Text = txtScanCount.Text + 1
            txtReelBarcode.Text = ""
        
        
            For i = 1 To spdPrtReelDetail.MaxRows
                Select Case lblstrPrtLabelName.Caption
                    Case "P0011", "P0018", "P0019", "P0020"
                        If GetText(spdPrtReelDetail, i, 1) = "Total Quantity/Length" Then
                            'cm 단위로 하였기에  M단위로 바꿈
                            strProdLen = txtProdLen.Text / 100
                            strContents = txtScanCount.Text & "Reel/" & CCur(strProdLen * txtScanCount.Text) & "M"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strContents = GetText(spdPrtReelDetail, i, 3)
                            '2X2708800000309010001B
                            'strContents = Mid
                            
                            strProdCalLen = (txtScanCount.Text * txtProdLen.Text)
                            'x=0의 갯수 : chr(x+64)
                            J = 0
                            For k = Len(strProdCalLen) To 1 Step -1
                                If Mid(strProdCalLen, k, 1) = "0" Then
                                    J = J + 1
                                End If
                            Next
                            strProdCalLen = Mid(strProdCalLen, 1, Len(strProdCalLen) - J)
                            strProdCalLen = Format(strProdCalLen, "0000")
                            strBarData = strBarData & strProdCalLen & Chr(J + 64)
                            strContents = Mid(strContents, 1, 17) & strBarData
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                    Case "P0001", "P0002"
                        If GetText(spdPrtReelDetail, i, 1) = "Total Quantity/Length" Then
                            strProdLen = txtProdLen.Text
                            strContents = txtScanCount.Text & "Reel/" & CCur(strProdLen * txtScanCount.Text) & "cm"
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strContents = GetText(spdPrtReelDetail, i, 3)
                            '2X2708800000309010001B
                            'strContents = Mid
                            
                            strProdCalLen = (txtScanCount.Text * txtProdLen.Text)
                            'x=0의 갯수 : chr(x+64)
                            J = 0
                            For k = Len(strProdCalLen) To 1 Step -1
                                If Mid(strProdCalLen, k, 1) = "0" Then
                                    J = J + 1
                                End If
                            Next
                            strProdCalLen = Mid(strProdCalLen, 1, Len(strProdCalLen) - J)
                            strProdCalLen = Format(strProdCalLen, "0000")
                            strBarData = strBarData & strProdCalLen & Chr(J + 64)
                            strContents = Mid(strContents, 1, 17) & strBarData
                            Call SetText(spdPrtReelDetail, strContents, i, 3)
                        End If
                    Case "P0004", "P0008", "P0009"
                        If GetText(spdPrtReelDetail, i, 1) = "바코드" Then
                            strBarData = GetText(spdPrtReelDetail, i, 3)
                            strContents = (CCur(txtProdLen.Text) / 100) * CCur(txtScanCount.Text)
                            strContents = Format(strContents, "0000")
                            'strContents = Mid(strContents, 1, 2)
                            'A=65,B=66....R=82
                            If CCur(Mid(strContents, 1, 2)) < 10 Then
                                strContents = Mid(strContents, 2, 3)
                            Else
                                strContents = Chr(CCur(Mid(strContents, 1, 2)) + 55) & Mid(strContents, 3, 2)
                            End If
                            strBarData = Mid(strBarData, 1, 27) & strContents
                            Call SetText(spdPrtReelDetail, strBarData, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "바코드값" Then
                            strBarData = GetText(spdPrtReelDetail, i, 3)
                            strContents = (CCur(txtProdLen.Text) / 100) * CCur(txtScanCount.Text)
                            strContents = Format(strContents, "0000")
                            'strContents = Mid(strContents, 1, 2)
                            'A=65,B=66....R=82
                            If CCur(Mid(strContents, 1, 2)) < 10 Then
                                strContents = Mid(strContents, 2, 3)
                            Else
                                strContents = Chr(CCur(Mid(strContents, 1, 2)) + 55) & Mid(strContents, 3, 2)
                            End If
                            strBarData = Mid(strBarData, 1, 27) & strContents
                            Call SetText(spdPrtReelDetail, strBarData, i, 3)
                        ElseIf GetText(spdPrtReelDetail, i, 1) = "수량" Then
                            Call SetText(spdPrtReelDetail, txtScanCount.Text, i, 3)
                        End If
                    Case Else
                        If GetText(spdPrtReelDetail, i, 1) = "수량" Then
                            Call SetText(spdPrtReelDetail, txtScanCount.Text, i, 3)
                        End If

                End Select
            Next
            
            DoEvents
            
            Call Sleep(200)
            
            If txtMaxTot.Text = txtScanCount.Text Then
                If chAutoPrint.Value = "1" Then
                    Call cmdPrint_Click
                End If
            End If
            
        End If
    End If
    
End Sub
