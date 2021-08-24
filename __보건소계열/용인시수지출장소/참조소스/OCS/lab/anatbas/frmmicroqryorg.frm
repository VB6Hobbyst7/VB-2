VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Begin VB.Form frmMicroQryOrg 
   Caption         =   "세균List 조건별 조회확인"
   ClientHeight    =   4590
   ClientLeft      =   2190
   ClientTop       =   2415
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7245
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   3795
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   6795
      _Version        =   131072
      _ExtentX        =   11986
      _ExtentY        =   6694
      _StockProps     =   100
      BorderStyle     =   1
      TabsPerRow      =   1
      TabCount        =   5
      AlignTextH      =   2
      Orientation     =   1
      TabShape        =   3
      AlignPictureV   =   1
      OffsetFromClientTop=   -1  'True
      Mode            =   1
      BookCornerType  =   1
      BookShowCornerGuard=   -1  'True
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      TextOrientation =   4
      TabCaption      =   "frmMicroQryOrg.frx":0000
      Begin VB.ListBox lstGrpCode 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   1680
         Left            =   -17219
         TabIndex        =   76
         Top             =   -17219
         Width           =   1095
      End
      Begin Threed.SSCommand cmdQry4 
         Height          =   375
         Left            =   -18599
         TabIndex        =   75
         Top             =   -15914
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "조회확인"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   0
         Left            =   -15779
         TabIndex        =   66
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "+"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin VB.TextBox txtQryCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -17939
         TabIndex        =   35
         Top             =   -16274
         Width           =   1335
      End
      Begin Threed.SSCommand cmdQry0 
         Height          =   1095
         Left            =   4440
         TabIndex        =   3
         Top             =   1260
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1931
         _StockProps     =   78
         Caption         =   "조회확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmMicroQryOrg.frx":0583
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Left            =   540
         TabIndex        =   2
         Top             =   480
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "조회를 원하는 세균명의 시작 alphabet 입력 하세요"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Alignment       =   1
      End
      Begin VB.TextBox txtQryName 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   1260
         Width           =   3255
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   1140
         TabIndex        =   4
         Top             =   1620
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   15
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelOuter      =   1
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   432
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "A"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   1
            Left            =   300
            TabIndex        =   6
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "B"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   2
            Left            =   540
            TabIndex        =   7
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "C"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   3
            Left            =   780
            TabIndex        =   8
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "D"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   4
            Left            =   1020
            TabIndex        =   9
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "E"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   5
            Left            =   1260
            TabIndex        =   10
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "F"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   6
            Left            =   1500
            TabIndex        =   11
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "G"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   7
            Left            =   1740
            TabIndex        =   12
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "H"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   8
            Left            =   1980
            TabIndex        =   13
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "I"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   9
            Left            =   2220
            TabIndex        =   14
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "J"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   10
            Left            =   2460
            TabIndex        =   15
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "K"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   11
            Left            =   2700
            TabIndex        =   16
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "L"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   12
            Left            =   2940
            TabIndex        =   17
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "M"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   13
            Left            =   60
            TabIndex        =   18
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "N"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   14
            Left            =   300
            TabIndex        =   19
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "O"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   15
            Left            =   540
            TabIndex        =   20
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "P"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   16
            Left            =   780
            TabIndex        =   21
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "Q"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   17
            Left            =   1020
            TabIndex        =   22
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "R"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   18
            Left            =   1260
            TabIndex        =   23
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "S"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   19
            Left            =   1500
            TabIndex        =   24
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "T"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   20
            Left            =   1740
            TabIndex        =   25
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "U"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   21
            Left            =   1980
            TabIndex        =   26
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "V"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   22
            Left            =   2220
            TabIndex        =   27
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "W"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   23
            Left            =   2460
            TabIndex        =   28
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "X"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   24
            Left            =   2700
            TabIndex        =   29
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "Y"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   300
            Index           =   25
            Left            =   2940
            TabIndex        =   30
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   538
            _StockProps     =   78
            Caption         =   "Z"
            BevelWidth      =   1
            Outline         =   0   'False
         End
      End
      Begin Threed.SSCommand cmdQry1 
         Height          =   1095
         Left            =   -19679
         TabIndex        =   34
         Top             =   -17054
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1931
         _StockProps     =   78
         Caption         =   "조회확인"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmMicroQryOrg.frx":0E5D
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   -18899
         TabIndex        =   36
         Top             =   -17054
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   15
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelOuter      =   1
         Enabled         =   0   'False
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   39
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "a"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   1
            Left            =   300
            TabIndex        =   40
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "b"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   2
            Left            =   540
            TabIndex        =   41
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "c"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   3
            Left            =   780
            TabIndex        =   42
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "d"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   4
            Left            =   1020
            TabIndex        =   43
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "e"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   5
            Left            =   1260
            TabIndex        =   44
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "f"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   6
            Left            =   1500
            TabIndex        =   45
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "g"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   7
            Left            =   1740
            TabIndex        =   46
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "h"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   8
            Left            =   1980
            TabIndex        =   47
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "i"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   9
            Left            =   2220
            TabIndex        =   48
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "j"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   10
            Left            =   2460
            TabIndex        =   49
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "k"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   11
            Left            =   2700
            TabIndex        =   50
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "l"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   12
            Left            =   2940
            TabIndex        =   51
            Top             =   60
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "m"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   13
            Left            =   60
            TabIndex        =   52
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "n"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   14
            Left            =   300
            TabIndex        =   53
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "o"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   15
            Left            =   540
            TabIndex        =   54
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "p"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   16
            Left            =   780
            TabIndex        =   55
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "q"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   17
            Left            =   1020
            TabIndex        =   56
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "r"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   18
            Left            =   1260
            TabIndex        =   57
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "s"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   19
            Left            =   1500
            TabIndex        =   58
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "t"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   20
            Left            =   1740
            TabIndex        =   59
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "u"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   21
            Left            =   1980
            TabIndex        =   60
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "v"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   22
            Left            =   2220
            TabIndex        =   61
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "w"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   23
            Left            =   2460
            TabIndex        =   62
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "x"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   24
            Left            =   2700
            TabIndex        =   63
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "y"
            BevelWidth      =   1
            Outline         =   0   'False
         End
         Begin Threed.SSCommand cmdSearch1 
            Height          =   300
            Index           =   25
            Left            =   2940
            TabIndex        =   64
            Top             =   360
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "z"
            BevelWidth      =   1
            Outline         =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   375
         Left            =   -19679
         TabIndex        =   38
         Top             =   -15674
         Width           =   4515
         _Version        =   65536
         _ExtentX        =   7964
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "조회를 원하는 세균코드의 시작 alphabet 입력 하세요"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Alignment       =   1
         Enabled         =   0   'False
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   375
         Left            =   -19739
         TabIndex        =   65
         Top             =   -15794
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "조회를 원하는 Gram Stain Result Word 를 선택하세요!"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Alignment       =   1
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   1
         Left            =   -16334
         TabIndex        =   67
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "-"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   2
         Left            =   -16889
         TabIndex        =   68
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "a"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   3
         Left            =   -17444
         TabIndex        =   69
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "b"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   4
         Left            =   -17999
         TabIndex        =   70
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "f"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   5
         Left            =   -18554
         TabIndex        =   71
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "m"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   6
         Left            =   -19109
         TabIndex        =   72
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "o"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdQryGram 
         Height          =   375
         Index           =   7
         Left            =   -19664
         TabIndex        =   73
         Top             =   -16514
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "w"
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "약제그룹코드"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -17219
         TabIndex        =   74
         Top             =   -15434
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "세균코드"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -16439
         TabIndex        =   37
         Top             =   -16214
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "세균명"
         Height          =   195
         Left            =   540
         TabIndex        =   31
         Top             =   1320
         Width           =   555
      End
   End
   Begin Threed.SSPanel panelPercent 
      Height          =   375
      Left            =   180
      TabIndex        =   32
      Top             =   60
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   4
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   60
         TabIndex        =   33
         Top             =   60
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMicroQryOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbGrpCode_Click()

    
End Sub

Private Sub cmdQry0_Click()
        
    ProgressBar1.Value = 0
    
    GoSub Get_OrgList_Data
    Exit Sub
    
'/----------------------------------------------------------
Get_OrgList_Data:
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    If Trim(txtQryName.Text) <> "" Then
        strSql = strSql & " WHERE  ORG_NAME LIKE '" & txtQryName.Text & "%'"
    End If
    strSql = strSql & " ORDER  BY ORG_name"
    
    frmMicroOrg.ssOrgList.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    frmMicroOrg.ssOrgList.MaxRows = adoSet.RecordCount
    ProgressBar1.Min = 0
    ProgressBar1.Max = adoSet.RecordCount
    
    Do Until adoSet.EOF
        frmMicroOrg.ssOrgList.Row = frmMicroOrg.ssOrgList.DataRowCnt + 1
        frmMicroOrg.ssOrgList.Col = 1: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Code").Value & ""
        frmMicroOrg.ssOrgList.Col = 2: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Name").Value & ""
        frmMicroOrg.ssOrgList.Col = 3: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Gram").Value & ""
        frmMicroOrg.ssOrgList.Col = 4: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_AntiGr").Value & ""
        frmMicroOrg.ssOrgList.Col = 5: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Status").Value & ""
        adoSet.MoveNext
        DoEvents
        ProgressBar1.Value = ProgressBar1.Value + 1
        panelPercent = ProgressBar1.Value & "/" & ProgressBar1.Max & " "
    Loop
    Call adoSetClose(adoSet)
    Unload Me
    Return
    
    

End Sub

Private Sub cmdQry1_Click()
    ProgressBar1.Value = 0
    
    GoSub Get_OrgList_Data
    Exit Sub
    
'/----------------------------------------------------------
Get_OrgList_Data:
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    If Trim(txtQryCode.Text) <> "" Then
        strSql = strSql & " WHERE  ORG_Code LIKE '" & txtQryCode.Text & "%'"
    End If
    strSql = strSql & " ORDER  BY ORG_Code"
    
    frmMicroOrg.ssOrgList.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    frmMicroOrg.ssOrgList.MaxRows = adoSet.RecordCount
    ProgressBar1.Min = 0
    ProgressBar1.Max = adoSet.RecordCount
    
    Do Until adoSet.EOF
        frmMicroOrg.ssOrgList.Row = frmMicroOrg.ssOrgList.DataRowCnt + 1
        frmMicroOrg.ssOrgList.Col = 1: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Code").Value & ""
        frmMicroOrg.ssOrgList.Col = 2: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Name").Value & ""
        frmMicroOrg.ssOrgList.Col = 3: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Gram").Value & ""
        frmMicroOrg.ssOrgList.Col = 4: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_AntiGr").Value & ""
        frmMicroOrg.ssOrgList.Col = 5: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Status").Value & ""
        adoSet.MoveNext
        DoEvents
        ProgressBar1.Value = ProgressBar1.Value + 1
        panelPercent = ProgressBar1.Value & "/" & ProgressBar1.Max & " "
    Loop
    Call adoSetClose(adoSet)
    Unload Me
    Return

End Sub

Private Sub cmdQry4_Click()
        
    ProgressBar1.Value = 0
    If lstGrpCode.ListIndex = -1 Then Exit Sub
    
    GoSub Get_OrgList_Data
    Exit Sub
    
'/----------------------------------------------------------
Get_OrgList_Data:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    strSql = strSql & " WHERE  Org_AntiGr  =  '" & Trim(lstGrpCode.List(lstGrpCode.ListIndex)) & "'"
    
    frmMicroOrg.ssOrgList.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    frmMicroOrg.ssOrgList.MaxRows = adoSet.RecordCount
    ProgressBar1.Min = 0
    ProgressBar1.Max = adoSet.RecordCount
    
    Do Until adoSet.EOF
        frmMicroOrg.ssOrgList.Row = frmMicroOrg.ssOrgList.DataRowCnt + 1
        frmMicroOrg.ssOrgList.Col = 1: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Code").Value & ""
        frmMicroOrg.ssOrgList.Col = 2: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Name").Value & ""
        frmMicroOrg.ssOrgList.Col = 3: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Gram").Value & ""
        frmMicroOrg.ssOrgList.Col = 4: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_AntiGr").Value & ""
        frmMicroOrg.ssOrgList.Col = 5: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Status").Value & ""
        adoSet.MoveNext
        DoEvents
        ProgressBar1.Value = ProgressBar1.Value + 1
        panelPercent = ProgressBar1.Value & "/" & ProgressBar1.Max & " "
    Loop
    Call adoSetClose(adoSet)
    Unload Me
    Return

End Sub

Private Sub cmdQryGram_Click(Index As Integer)
    
    ProgressBar1.Value = 0
    
    GoSub Get_OrgList_Data
    Exit Sub
    
'/----------------------------------------------------------
Get_OrgList_Data:
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    strSql = strSql & " WHERE  ORG_Gram = '" & Trim(cmdQryGram(Index).Caption) & "'"
    strSql = strSql & " ORDER  BY ORG_Gram, ORG_Code"
    
    frmMicroOrg.ssOrgList.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    frmMicroOrg.ssOrgList.MaxRows = adoSet.RecordCount
    ProgressBar1.Min = 0
    ProgressBar1.Max = adoSet.RecordCount
    
    Do Until adoSet.EOF
        frmMicroOrg.ssOrgList.Row = frmMicroOrg.ssOrgList.DataRowCnt + 1
        frmMicroOrg.ssOrgList.Col = 1: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Code").Value & ""
        frmMicroOrg.ssOrgList.Col = 2: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Name").Value & ""
        frmMicroOrg.ssOrgList.Col = 3: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Gram").Value & ""
        frmMicroOrg.ssOrgList.Col = 4: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_AntiGr").Value & ""
        frmMicroOrg.ssOrgList.Col = 5: frmMicroOrg.ssOrgList.Text = adoSet.Fields("Org_Status").Value & ""
        adoSet.MoveNext
        DoEvents
        ProgressBar1.Value = ProgressBar1.Value + 1
        panelPercent = ProgressBar1.Value & "/" & ProgressBar1.Max & " "
    Loop
    Call adoSetClose(adoSet)
    Unload Me
    Return
    
End Sub

Private Sub CmdSearch_Click(Index As Integer)
    
    DoEvents
    
    txtQryName.Text = CmdSearch(Index).Caption
    Call cmdQry0_Click
    

End Sub

Private Sub cmdSearch1_Click(Index As Integer)
    DoEvents
    
    txtQryCode.Text = cmdSearch1(Index).Caption
    Call cmdQry1_Click

End Sub

Private Sub Form_Load()
    
    Me.Height = 5280
    Me.Width = 7365
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    
    GoSub ANTI_GroupCode_Get
    Exit Sub
    
ANTI_GroupCode_Get:
    strSql = ""
    strSql = strSql & " SELECT org_AntiGr"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    strSql = strSql & " GROUP  BY Org_AntiGr"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Me.lstGrpCode.Clear
    Do Until adoSet.EOF
        lstGrpCode.AddItem Trim(adoSet.Fields("org_antiGr").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

