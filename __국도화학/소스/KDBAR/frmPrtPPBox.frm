VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "barcod32.ocx"
Begin VB.Form frmPrtPPBox 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PP Box ????"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20505
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   20505
   WindowState     =   2  '?ִ?ȭ
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   90
      TabIndex        =   65
      Top             =   60
      Width           =   19425
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "??ȸ"
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5610
         Style           =   1  '?׷???
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ȭ??????"
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6750
         Style           =   1  '?׷???
         TabIndex        =   67
         ToolTipText     =   "????ȭ???? ???? ?????ϴ?"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "?۾??Ϸ???????ȸ"
         BeginProperty Font 
            Name            =   "???? ????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         TabIndex        =   66
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
         TabIndex        =   69
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   70057985
         CurrentDate     =   43884
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   3750
         TabIndex        =   70
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "???? ????"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   70057985
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
               Picture         =   "frmPrtPPBox.frx":0000
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":059A
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":0B34
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":10CE
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":1960
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":1ABA
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":1C14
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":1D6E
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtPPBox.frx":2648
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  '??? ????
         BackStyle       =   0  '????
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "???? ????"
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
         TabIndex        =   73
         Top             =   420
         Width           =   195
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '????
         Caption         =   "?? ???????? "
         BeginProperty Font 
            Name            =   "???? ????"
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
         TabIndex        =   72
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label lblComStatus 
         BackStyle       =   0  '????
         Caption         =   "Com1 ???Ἲ??"
         BeginProperty Font 
            Name            =   "???? ????"
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
         TabIndex        =   71
         Top             =   420
         Width           =   3255
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   11340
         Picture         =   "frmPrtPPBox.frx":2F22
         Top             =   420
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   8205
      Left            =   90
      TabIndex        =   1
      Top             =   1050
      Width           =   19395
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   7725
         Left            =   8160
         TabIndex        =   2
         Top             =   300
         Width           =   10995
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '????
            Height          =   1185
            Left            =   8250
            TabIndex        =   47
            Top             =   6270
            Width           =   2475
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00E0E0E0&
               Caption         =   "?ݱ?"
               BeginProperty Font 
                  Name            =   "???? ????"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1260
               Style           =   1  '?׷???
               TabIndex        =   51
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdPrint 
               Appearance      =   0  '????
               BackColor       =   &H00E0E0E0&
               Caption         =   "????"
               BeginProperty Font 
                  Name            =   "???? ????"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   150
               Style           =   1  '?׷???
               TabIndex        =   50
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdAllPrint 
               Appearance      =   0  '????
               BackColor       =   &H00E0E0E0&
               Caption         =   "?ϰ?????"
               BeginProperty Font 
                  Name            =   "???? ????"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   150
               Style           =   1  '?׷???
               TabIndex        =   49
               Top             =   660
               Width           =   1095
            End
            Begin VB.CommandButton cmdView 
               BackColor       =   &H00E0E0E0&
               Caption         =   "????"
               BeginProperty Font 
                  Name            =   "???? ????"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1260
               Style           =   1  '?׷???
               TabIndex        =   48
               Top             =   660
               Width           =   1095
            End
         End
         Begin VB.TextBox txtProdNm 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   46
            Top             =   360
            Width           =   1485
         End
         Begin VB.TextBox txtPackNm 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   45
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtProdOrderDt 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   44
            Top             =   360
            Width           =   1500
         End
         Begin VB.TextBox txtCompNm 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   43
            Top             =   810
            Width           =   1500
         End
         Begin VB.TextBox txtProdPosNo 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7410
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   42
            Top             =   810
            Width           =   1635
         End
         Begin VB.TextBox txtSlittingNo 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9510
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   41
            Top             =   810
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox txtReelQTY 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   40
            Top             =   5640
            Width           =   1500
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   4185
            Left            =   5640
            ScaleHeight     =   4125
            ScaleWidth      =   5055
            TabIndex        =   5
            Top             =   1440
            Width           =   5115
            Begin BarcodLib.Barcod barReel 
               Height          =   555
               Left            =   300
               TabIndex        =   6
               Top             =   300
               Visible         =   0   'False
               Width           =   4365
               _Version        =   65543
               _ExtentX        =   7699
               _ExtentY        =   979
               _StockProps     =   75
               Caption         =   "2X2707R0202001P10110000"
               BarWidth        =   0
               Direction       =   0
               Style           =   7
               UPCNotches      =   3
               Alignment       =   2
               Extension       =   ""
            End
            Begin BarcodLib.Barcod barPart 
               Height          =   555
               Left            =   270
               TabIndex        =   7
               Top             =   1230
               Visible         =   0   'False
               Width           =   4365
               _Version        =   65543
               _ExtentX        =   7699
               _ExtentY        =   979
               _StockProps     =   75
               Caption         =   "2X2707R0202001P10110000"
               BarWidth        =   0
               Direction       =   0
               Style           =   7
               UPCNotches      =   3
               Alignment       =   2
               Extension       =   ""
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
                  Name            =   "????"
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
         Begin VB.TextBox txtLotNo 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7410
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   4
            Top             =   360
            Width           =   1635
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00D0E0E0&
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9510
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   1000
         End
         Begin FPSpread.vaSpread spdPrtReelDetail 
            Height          =   4125
            Left            =   240
            TabIndex        =   52
            Top             =   1440
            Width           =   5325
            _Version        =   393216
            _ExtentX        =   9393
            _ExtentY        =   7276
            _StockProps     =   64
            ColsFrozen      =   8
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "???? ????"
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
            SpreadDesigner  =   "frmPrtPPBox.frx":34AC
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin FPSpread.vaSpread spdRegOrderDetail 
            Height          =   1455
            Left            =   240
            TabIndex        =   53
            Top             =   6060
            Width           =   7455
            _Version        =   393216
            _ExtentX        =   13150
            _ExtentY        =   2566
            _StockProps     =   64
            ColsFrozen      =   8
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "???? ????"
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
            MaxCols         =   19
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ScrollBars      =   2
            ShadowColor     =   16774120
            SpreadDesigner  =   "frmPrtPPBox.frx":40EE
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "??ǰ??"
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   63
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "??????"
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   62
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "Slitting No"
            BeginProperty Font 
               Name            =   "???? ????"
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
            Left            =   8460
            TabIndex        =   61
            Top             =   810
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "???? No"
            BeginProperty Font 
               Name            =   "???? ????"
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
            Left            =   6360
            TabIndex        =   60
            Top             =   810
            Width           =   1000
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "????????"
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   59
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "?????ڵ?"
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   58
            Top             =   810
            Width           =   1500
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "Reel ????"
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   57
            Top             =   5640
            Width           =   1500
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '????
            Caption         =   "2 ??"
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   4290
            TabIndex        =   56
            Top             =   5700
            Width           =   435
         End
         Begin VB.Label lblstrPrtLabelName 
            BackStyle       =   0  '????
            BeginProperty Font 
               Name            =   "???? ????"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   8370
            TabIndex        =   55
            Top             =   6000
            Width           =   2265
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '??? ????
            Appearance      =   0  '????
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '???? ????
            Caption         =   "Lot No"
            BeginProperty Font 
               Name            =   "???? ????"
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
            TabIndex        =   54
            Top             =   360
            Width           =   1000
         End
      End
      Begin FPSpread.vaSpread spdPrtReel 
         Height          =   7635
         Left            =   210
         TabIndex        =   64
         Top             =   390
         Width           =   7905
         _Version        =   393216
         _ExtentX        =   13944
         _ExtentY        =   13467
         _StockProps     =   64
         ColsFrozen      =   8
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "???? ????"
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
         SpreadDesigner  =   "frmPrtPPBox.frx":532A
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.TextBox txtComm 
      Appearance      =   0  '????
      Height          =   5955
      Left            =   19620
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1770
      Visible         =   0   'False
      Width           =   5475
   End
End
Attribute VB_Name = "frmPrtPPBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   ???ϸ?  : frmPrtPPBox.frm
'   ?ۼ???  : ??????
'   ??  ??  : PP Box ????????
'   ?ۼ???  : 2020-02-29
'   ??  ??  : 1.0.0
'   ??  ??  : ????ȭ??
'-----------------------------------------------------------------------------'

Private Sub cmdAllPrint_Click()
    Dim i As Integer
    
    If MsgBox("?ϰ??????? ?Ͻðڽ??ϱ??", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbYes Then
        For i = 1 To spdPrtReel.MaxRows
            If GetText(spdPrtReel, i, 1) = "1" Then
                Call spdPrtReel_Click(2, i)
                DoEvents
                Call cmdPrint_Click
                DoEvents
                Call SetText(spdPrtReel, "", i, 1)
            End If
        Next
    End If
    
End Sub

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdPrtReel.MaxRows = 0
    spdPrtReelDetail.MaxRows = 0
    spdRegOrderDetail.MaxRows = 0
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    txtReelQTY.Text = ""
    
    lblstrPrtLabelName.Caption = ""
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub


' ???ڵ? ????Ʈ ??????
Private Sub GetOrderList(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String)
    Dim strLabelType    As String
    
    Set AdoRs = Get_OrderList(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType)
    
    If AdoRs Is Nothing Then
        '???ϵ? ???? ????
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

' Reel ?۾? ????Ʈ ??????
Private Sub GetReelOrderList_PP(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String)

    Dim strLabelType    As String

    Set AdoRs = Get_OrderList_PP(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType)

    If AdoRs Is Nothing Then
        '???ϵ? ???? ????
    Else
        Do Until AdoRs.EOF
            With spdPrtReel
                .MaxRows = .MaxRows + 1

                Call SetText(spdPrtReel, "1", .MaxRows, 1)
                Call SetText(spdPrtReel, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReel, Format(AdoRs.Fields("PROD_ORDER_DT").Value & "", "####-##-##"), .MaxRows, 3)
                Call SetText(spdPrtReel, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 4)
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

    End If

    AdoRs.Close

End Sub


'^XA : Opening Bracket???? Format?? ?????? ?˸???.
'^FO : ?μ? ?? ?׸??? ?μ? ?? ??ġ(X??,Y??)?? ?????Ѵ?.
'^A1 : ??Ʈ????
'      o => ȸ?? : N(0),R(90),I(180),B(270)
'      h => ???? : 20
'      w => ???? : 25
'^BY : ???ڵ? ?????? ????
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
'  2 -> ???¼? font               ^BC -> code 128
'  3 -> ?????? font                70 -> barcode ????
' 70 -> barcode ????               Y -> barcode ?ϴܿ? ???ڿ? ????
'                                  N -> barcode ???ܿ? ???ڿ? ????
'                                  N -> check digit ǥ??????
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
' ^A0,N,26,22 : D TYPE ???? 26, ?? 22dot
'
' ^CI  (Change International Font/Encoding)
' 26 = Multibyte Asian Encodings with ASCII Transparency a And c

        'asc("~") : R_7E
        '? Hex(126)
            
        'asc("??") : R_A1C9
         '?Hex(126)
            
                                'strData = Replace(strData, "~", "_7E")
                                'strData = Replace(strData, "??", "_A1C9")
            


Private Sub cmdPrint_Click()
    Dim intPRow     As Integer
    Dim intPrt      As Integer
    Dim strPrtFNo   As String
    Dim strPrtTNo   As String
    Dim intPrtFNo   As Integer
    Dim intPrtTNo   As Integer
    Dim intRow      As Integer
    Dim strHeader   As String
    Dim strOutput   As String
    Dim strBarcode  As String
    Dim strPrtData()   As String
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
    
    blnPrint = False
    strBarcode = ""
    
    If txtProdNm.Text = "" Then
        MsgBox "?????? Reel?????? ?????ϼ???", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    Select Case lblstrPrtLabelName.Caption
        Case "TP203C(ACF)|100", "TP203C(ACF)|200"
            i = 0
               
            strPlusXPos = 630
            strAFont = "^A0N,35,30"
            
            strHeader = "^XA" & vbLf
            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
            strHeader = strHeader & "^PON^FS" & vbLf
            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
            strHeader = strHeader & "^MD9"
            
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
                    With spdPrtReelDetail
                        For intRow = 1 To .MaxRows
                            '???? ????
                            strType = GetText(spdPrtReelDetail, intRow, 1)
                            strData = GetText(spdPrtReelDetail, intRow, 3)
                            strXPos = GetText(spdPrtReelDetail, intRow, 4)
                            strYPos = GetText(spdPrtReelDetail, intRow, 5)
                            strBarType = GetText(spdPrtReelDetail, intRow, 6)   '???ڵ?Ÿ??
                            strFont = GetText(spdPrtReelDetail, intRow, 7)      '??Ʈ
                            strRot = GetText(spdPrtReelDetail, intRow, 8)       'ȸ??
                            strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                            strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                            
                            Select Case strRot
                                Case "0":   strRot = "N"
                                Case "90":  strRot = "R"
                                Case "180": strRot = "I"
                                Case "270": strRot = "B"
                            End Select
                            
                            If strType = "???ڵ?" Then
                                strLotSub = "P" & strSlt & Format(intPrt, "00")
                                strData = Mid(strData, 1, 14) & strLotSub & Mid(strData, 19, 5)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                    strPrtData(i) = strPrtData(i) & "^BQ"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                    comEqp.Output = strOutput
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
'                    If intPrt = 1 Then Exit Sub
                Next
            Next
            
        Case "CF-TP408A|200"
            i = 0
            strPlusXPos = 630
            strAFont = "^A0N,35,25"
            
            strHeader = "^XA" & vbLf
            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
            strHeader = strHeader & "^PON^FS" & vbLf
            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
            strHeader = strHeader & "^MD9"
            
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
                    With spdPrtReelDetail
                        For intRow = 1 To .MaxRows
                            '???? ????
                            strType = GetText(spdPrtReelDetail, intRow, 1)
                            strData = GetText(spdPrtReelDetail, intRow, 3)
                            strXPos = GetText(spdPrtReelDetail, intRow, 4)
                            strYPos = GetText(spdPrtReelDetail, intRow, 5)
                            strBarType = GetText(spdPrtReelDetail, intRow, 6)   '???ڵ?Ÿ??
                            strFont = GetText(spdPrtReelDetail, intRow, 7)      '??Ʈ
                            strRot = GetText(spdPrtReelDetail, intRow, 8)       'ȸ??
                            strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                            strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                            
                            Select Case strRot
                                Case "0":   strRot = "N"
                                Case "90":  strRot = "R"
                                Case "180": strRot = "I"
                                Case "270": strRot = "B"
                            End Select
                            
                            If strType = "???ڵ?" Then
                                strLotSub = Format(intPrt, "0000")
                                strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
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
                            
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                        Next
                    End With
                
                    strOutput = ""
                    For J = 0 To UBound(strPrtData)
                        strOutput = strOutput & strPrtData(J)
                    Next
                    
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    comEqp.Output = strOutput
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
'                    If intPrt = 1 Then Exit Sub
                Next
            Next
            
        Case "CF-TP400E|270", "CF-TP500B|270"
            i = 0
            strPlusXPos = 630
            strAFont = "^A0N,35,25"
            
            strHeader = "^XA" & vbLf
            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
            strHeader = strHeader & "^PON^FS" & vbLf
            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
            strHeader = strHeader & "^MD9"
            
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
                    With spdPrtReelDetail
                        For intRow = 1 To .MaxRows
                            '???? ????
                            strType = GetText(spdPrtReelDetail, intRow, 1)
                            strData = GetText(spdPrtReelDetail, intRow, 3)
                            strXPos = GetText(spdPrtReelDetail, intRow, 4)
                            strYPos = GetText(spdPrtReelDetail, intRow, 5)
                            strBarType = GetText(spdPrtReelDetail, intRow, 6)   '???ڵ?Ÿ??
                            strFont = GetText(spdPrtReelDetail, intRow, 7)      '??Ʈ
                            strRot = GetText(spdPrtReelDetail, intRow, 8)       'ȸ??
                            strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                            strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                            
                            Select Case strRot
                                Case "0":   strRot = "N"
                                Case "90":  strRot = "R"
                                Case "180": strRot = "I"
                                Case "270": strRot = "B"
                            End Select
                            
                            If strType = "???ڵ?" Then
                                strLotSub = strSlt & Format(intPrt, "00")
                                strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                    comEqp.Output = strOutput
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
'                    If intPrt = 1 Then Exit Sub
                Next
            Next
            
        Case "CF-TP500B|350", "CF-TP500E|350"
            i = 0
            strPlusXPos = 630
            strAFont = "^A0N,35,25"
            
            strHeader = "^XA" & vbLf
            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
            strHeader = strHeader & "^PON^FS" & vbLf
            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
            strHeader = strHeader & "^MD9"
            
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
                    With spdPrtReelDetail
                        For intRow = 1 To .MaxRows
                            '???? ????
                            strType = GetText(spdPrtReelDetail, intRow, 1)
                            strData = GetText(spdPrtReelDetail, intRow, 3)
                            strXPos = GetText(spdPrtReelDetail, intRow, 4)
                            strYPos = GetText(spdPrtReelDetail, intRow, 5)
                            strBarType = GetText(spdPrtReelDetail, intRow, 6)   '???ڵ?Ÿ??
                            strFont = GetText(spdPrtReelDetail, intRow, 7)      '??Ʈ
                            strRot = GetText(spdPrtReelDetail, intRow, 8)       'ȸ??
                            strSlt = GetText(spdPrtReelDetail, intRow, 9)       'Slitting No
                            strNamePrt = GetText(spdPrtReelDetail, intRow, 10)
                            
                            Select Case strRot
                                Case "0":   strRot = "N"
                                Case "90":  strRot = "R"
                                Case "180": strRot = "I"
                                Case "270": strRot = "B"
                            End Select
                            
                            If strType = "???ڵ?" Then
                                strLotSub = strSlt & Format(intPrt, "00")
                                strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 19, 5)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                    comEqp.Output = strOutput
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
'                    If intPrt = 1 Then Exit For
                Next
            Next
            
    End Select
   
    If blnPrint = True Then
        Call SetPackTrack(strBarcode)
    End If
    
End Sub

Private Sub SetPackTrack(ByVal pBarcode As String)
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intItemNo   As Integer
    
    gPackTrack.ORDERDT = Format(txtProdOrderDt.Text, "yyyymmdd")     'Key
    gPackTrack.PRODCD = txtProdCd.Text                               'Key
    gPackTrack.REELBAR = pBarcode
    gPackTrack.PPBAR = ""
    gPackTrack.ICEBAR = ""
    gPackTrack.PPBARIN = ""
    gPackTrack.ICEBARIN = ""
    gPackTrack.LOTNO = txtLotNo.Text
    gPackTrack.REELPRTID = gKUKDO.USERID
    gPackTrack.REELPRTDT = ""
    gPackTrack.PPPRTID = ""
    gPackTrack.PPPRTDT = ""
    gPackTrack.ICEPRTID = ""
    gPackTrack.ICEPRTDT = ""
    
    'Ʈ??ŷ ????
    '-- Insert / Update ã?ƿ???
    Set AdoRs = Get_Pack_Track(gPackTrack.ORDERDT, gPackTrack.PRODCD, gPackTrack.REELBAR)

    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Pack_Track("IN") Then
        End If
    Else
        'UPDATE
        If Set_Pack_Track("UP") Then
        End If
    End If

    gOrder.ORDDATE = Format(txtProdOrderDt.Text, "yyyymmdd")      'Key
    gOrder.PRODPOSNO = txtProdPosNo.Text                          'Key
    gOrder.PRODCD = txtProdCd.Text                                'Key
    gOrder.SLITINGNO = txtSlittingNo.Text                         'Key


    '?ͷ±??? UPDATE : LBL_PROD_ORDER..CLOSE_YN
    If Set_Order_CloseYN("UP") Then
        Call cmdSearch_Click
    End If


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


Private Sub cmdView_Click()
    
    If txtComm.Visible = False Then
        txtComm.Visible = True
    Else
        txtComm.Visible = False
    End If
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    '-- ???ſ???
    Call OpenCommunication
    
End Sub

Private Sub OpenCommunication()

On Error GoTo ErrHandle

    comEqp.CommPort = gComm.COMPORT
    comEqp.RTSEnable = gComm.RTSEnable
    comEqp.DTREnable = gComm.DTREnable
    comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

    If comEqp.PortOpen Then
        lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ???Ἲ??"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ????????"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
    End If

    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("??Ʈ ??ȣ?? ?߸??Ǿ????ϴ?." & vbNewLine & vbNewLine & "   ???? ?????Ͻðڽ??ϱ??", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & "??Ʈ ????????"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            
            Resume Next
        Else
            
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "??    ġ : " & "Public Sub OpenCommunication()" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "??????ȣ : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "???????? : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If


End Sub

'-- ??Ʈ???ʱ?ȭ
Private Sub CtlInitializing()
    Dim i           As Integer
    
    With spdPrtReel
        Call SetText(spdPrtReel, "????", 0, 1):              .ColWidth(1) = 4
        Call SetText(spdPrtReel, "Lot No", 0, 2):            .ColWidth(2) = 12
        Call SetText(spdPrtReel, "????????", 0, 3):          .ColWidth(3) = 10
        Call SetText(spdPrtReel, "????No", 0, 4):            .ColWidth(4) = 0
        Call SetText(spdPrtReel, "??ǰ?ڵ?", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdPrtReel, "??ǰ??", 0, 6):            .ColWidth(6) = 12
        Call SetText(spdPrtReel, "?????ڵ?", 0, 7):          .ColWidth(7) = 0
        Call SetText(spdPrtReel, "?޸?", 0, 8):              .ColWidth(8) = 0
        Call SetText(spdPrtReel, "?۾????뼳??", 0, 9):      .ColWidth(9) = 0 'Roll????
        Call SetText(spdPrtReel, "????", 0, 10):             .ColWidth(10) = 4
        Call SetText(spdPrtReel, "SLT No", 0, 11):           .ColWidth(11) = 6
        Call SetText(spdPrtReel, "????", 0, 12):             .ColWidth(12) = 4
        Call SetText(spdPrtReel, "??????", 0, 13):           .ColWidth(13) = 10
        Call SetText(spdPrtReel, "?۾??ϷῩ??", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdPrtReel, "?????ڵ?", 0, 15):         .ColWidth(15) = 10
        Call SetText(spdPrtReel, "?Է???", 0, 16):           .ColWidth(16) = 0
        Call SetText(spdPrtReel, "?Է??Ͻ?", 0, 17):         .ColWidth(17) = 0
        Call SetText(spdPrtReel, "??????", 0, 18):           .ColWidth(18) = 0
        Call SetText(spdPrtReel, "?????Ͻ?", 0, 19):         .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    With spdPrtReelDetail
        Call SetText(spdPrtReelDetail, "?׸?", 0, 1):          .ColWidth(1) = 15
        Call SetText(spdPrtReelDetail, "????", 0, 2):          .ColWidth(2) = 5
        Call SetText(spdPrtReelDetail, "????", 0, 3):          .ColWidth(3) = 21
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
    
    With spdRegOrderDetail
        Call SetText(spdRegOrderDetail, "????????", 0, 1):        .ColWidth(1) = 0
        Call SetText(spdRegOrderDetail, "????", 0, 2):            .ColWidth(2) = 0
        Call SetText(spdRegOrderDetail, "??ǰ?ڵ?", 0, 3):        .ColWidth(3) = 0
        Call SetText(spdRegOrderDetail, "SLT No", 0, 4):          .ColWidth(4) = 0
        Call SetText(spdRegOrderDetail, "?Ϸù?ȣ", 0, 5):        .ColWidth(5) = 8
        Call SetText(spdRegOrderDetail, "SLT????", 0, 6):         .ColWidth(6) = 28
        Call SetText(spdRegOrderDetail, "???۹?ȣ", 0, 7):        .ColWidth(7) = 10
        Call SetText(spdRegOrderDetail, "????ȣ", 0, 8):          .ColWidth(8) = 10
        Call SetText(spdRegOrderDetail, "", 0, 9):                .ColWidth(9) = 0
        Call SetText(spdRegOrderDetail, "No", 0, 10):             .ColWidth(10) = 0
        Call SetText(spdRegOrderDetail, "", 0, 11):               .ColWidth(11) = 0
        Call SetText(spdRegOrderDetail, "", 0, 12):               .ColWidth(12) = 0
        Call SetText(spdRegOrderDetail, "", 0, 13):               .ColWidth(13) = 0
        Call SetText(spdRegOrderDetail, "", 0, 14):               .ColWidth(14) = 0
        Call SetText(spdRegOrderDetail, "???뿩??", 0, 15):       .ColWidth(15) = 0
        Call SetText(spdRegOrderDetail, "?Է???", 0, 16):         .ColWidth(16) = 0
        Call SetText(spdRegOrderDetail, "?Է??Ͻ?", 0, 17):       .ColWidth(17) = 0
        Call SetText(spdRegOrderDetail, "??????", 0, 18):         .ColWidth(18) = 0
        Call SetText(spdRegOrderDetail, "?????Ͻ?", 0, 19):       .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtLotNo.Text = ""
    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    
    txtReelQTY.Text = ""
    
    gSORT = 0

End Sub

' ?۾????ü? ????Ʈ ?????? 'strDate, cboProdPosNo.Text, cboProdCd.Text, cboSlittingNo.Text
Private Sub GetOrderDetail(ByVal pDate As String, ByVal pProPosNo As String, ByVal pProCd As String, ByVal pSltNo As String)
    
    Set AdoRs = Get_OrderDetail(pDate, pProPosNo, pProCd, pSltNo)
            
    If AdoRs Is Nothing Then
        '???ϵ? ???? ????
    Else
        spdRegOrderDetail.MaxRows = 0
        Do Until AdoRs.EOF
            With spdRegOrderDetail
                .MaxRows = .MaxRows + 1
                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 1)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 2)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 3)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 4)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SEQ_NO").Value & "", .MaxRows, 5)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_INFO").Value & "", .MaxRows, 6)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_F").Value & "", .MaxRows, 7)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_T").Value & "", .MaxRows, 8)
  
            End With
            AdoRs.MoveNext
        Loop
        AdoRs.Close
    End If

End Sub



Private Sub spdPrtReel_Click(ByVal Col As Long, ByVal Row As Long)
    Dim pAdoRS1      As ADODB.Recordset
    Dim pAdoRS2      As ADODB.Recordset
    Dim i               As Integer
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
    '?????ڵ?
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
    Dim strExDate       As String
    Dim strProdTemp     As String
    Dim strPrtLabelName As String
    
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
            '-- ???? ?߰?
            Call SetSpreadSort(spdPrtReel, 0)
        End If
        Exit Sub
    End If
    
    For i = 0 To 15
        barReel.Visible = False
        barPart.Visible = False
        lblTitle(i).Visible = False
    Next
    
    strPrtLabelName = GetText(spdPrtReel, Row, 6) & "|" & GetText(spdPrtReel, Row, 10)
    lblstrPrtLabelName.Caption = strPrtLabelName
    
    strDate = GetText(spdPrtReel, Row, 3)
    txtProdOrderDt.Text = strDate
    strProdCd = GetText(spdPrtReel, Row, 5)
    txtProdCd.Text = strProdCd
    txtProdNm.Text = GetText(spdPrtReel, Row, 6)
    strProdLen = GetText(spdPrtReel, Row, 10)
    strProdLen = strProdLen * 100 '???͸? cm???? ??ȯ
    txtProdPosNo.Text = GetText(spdPrtReel, Row, 4)
    txtPackNm.Text = GetText(spdPrtReel, Row, 7)
    txtReelQTY.Text = GetText(spdPrtReel, Row, 12)
    txtSlittingNo.Text = GetText(spdPrtReel, Row, 11)
    txtCompNm.Text = GetText(spdPrtReel, Row, 13)
    strLotNo = GetText(spdPrtReel, Row, 2)
    txtLotNo.Text = strLotNo
    strProdLabelCd = GetText(spdPrtReel, Row, 15)
    strCompCd = GetText(spdPrtReel, Row, 18)
    strQty = txtReelQTY.Text
    
    gPackTrack.PRODCD = strProdCd   '5?ڸ?
    gPackTrack.LOTNO = strLotNo
    gPackTrack.ORDERDT = strDate    '8?ڸ?
              
    Call GetOrderDetail(Format(strDate, "yyyymmdd"), txtProdPosNo.Text, strProdCd, txtSlittingNo.Text)

    With spdPrtReelDetail
        .MaxRows = 0
    End With

    Set AdoRs = Get_LabelDetail(strProdLabelCd, "P")
            
    If AdoRs Is Nothing Then
        '???ϵ? ???? ????
    Else
        '-- ??ǰ???? ã?ƿ???
        SQL = ""
        SQL = SQL & "SELECT             " & vbCrLf
        SQL = SQL & "  PROD_NAME        " & vbCrLf
        SQL = SQL & ", PROD_LENGTH      " & vbCrLf
        SQL = SQL & ", PROD_MATERIAL_CD " & vbCrLf
        SQL = SQL & ", EXPIR_MONTH      " & vbCrLf
        SQL = SQL & ", PROD_STOR_TEMP   " & vbCrLf
        SQL = SQL & ", PROD_SIZE        " & vbCrLf
        SQL = SQL & ", PROD_CHIMEI_PN   " & vbCrLf
        SQL = SQL & ", VENDER_CD        " & vbCrLf
        SQL = SQL & ", PROD_LINE_FA     " & vbCrLf
        SQL = SQL & ", PROD_SLIT_FA     " & vbCrLf
        SQL = SQL & ", PROD_CONTROL_YN  " & vbCrLf
        SQL = SQL & ", PROD_PCN_NO                          " & vbCrLf
        SQL = SQL & "  FROM LBL_M_PROD                      " & vbCrLf
        SQL = SQL & " WHERE PROD_CD  = '" & strProdCd & "'  " & vbCrLf
        SQL = SQL & "   AND COMP_CD  = '" & strCompCd & "'  " & vbCrLf
        SQL = SQL & "   AND USED_YN  = 'Y'                  " & vbCrLf
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
            End If
        End If
        pAdoRS2.Close
        Set pAdoRS2 = Nothing
        '-- ??ǰ???? ã?ƿ???
        
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
    
                '============================== ??ǰ???? ?б? ==============================
                Select Case strPrtLabelName
                    '============== TP203C(ACF) ======================================================================
                    Case "TP203C(ACF)|100", "TP203C(ACF)|200"
                        '-- ???ڵ? ???? ???? ã?ƿ???
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            'strBarData = ""
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                            If pAdoRS1 Is Nothing Then
                                '???ϵ? ???? ????
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "?????ڵ?":       ' strBarData = strBarData
                                        Case "??ȿ?Ⱓ_??":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strExDate))
                                        Case "??ȿ?Ⱓ_??":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strExDate))
                                        Case "??ȿ?Ⱓ_??":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strExDate))
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "Slitting????":    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
                                        Case "Product No":      strBarData = strBarData & "P" & Format(txtSlittingNo.Text, "0") & "01"
                                        Case "??ǰ????":        strBarData = strBarData & strProdLen
                                    End Select
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
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
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Storage Temperature" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Production Date" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Expiration Date" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material Code" Then
                            Call SetText(spdPrtReelDetail, strChimeiCd, .MaxRows, 3)
                            strContents = strChimeiCd
                        End If
                        
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            barReel.Alignment = bcACenter
                            barReel.Style = msSCode128B
                            barReel.Visible = True
                            barReel.Caption = strContents
                            barReel.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            barReel.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
                            barReel.WIDTH = 4365
                        Else
                            If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            Else
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            End If
                        End If
                        strContents = ""
                                        

                    '============== CF-TP408A ======================================================================
                    Case "CF-TP408A|200"

                        '-- ???ڵ? ???? ???? ã?ƿ???
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            'strBarData = ""
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                        
                            If pAdoRS1 Is Nothing Then
                                '???ϵ? ???? ????
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "?????ڵ?"
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "ACF??????":       strBarData = strBarData & "K"
                                        Case "??ǰ????":        strBarData = strBarData & "0001"
                                        Case "??ǰ????":        strBarData = strBarData & Format(strProdLen, "0000")
                                        Case "?????系??":      strBarData = strBarData & "00"
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
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
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "?????µ?" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "????????" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "??ȿ?Ⱓ" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "00" & ")"
                            strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            barReel.Alignment = bcACenter
                            barReel.Style = msSCode128B
                            barReel.Visible = True
                            barReel.Caption = strContents
                            barReel.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            barReel.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
                            barReel.WIDTH = 4365
                        Else
                            If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            Else
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            End If
                        End If
                            
                        strContents = ""
                        
                    '============== CF-TP400E ======================================================================
                    Case "CF-TP400E|270", "CF-TP500B|270"
                        '-- ???ڵ? ???? ???? ã?ƿ???
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            'strBarData = ""
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                        
                            If pAdoRS1 Is Nothing Then
                                '???ϵ? ???? ????
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "?????ڵ?"
                                        Case "Vendor?ڵ?":      strBarData = strBarData & strVendorCd
                                        Case "???????ΰ???":    strBarData = strBarData & strProdLineFA
                                        Case "Sliting????":     strBarData = strBarData & strProdSlitFA
                                        Case "????????Ż????":  strBarData = strBarData & strContYN
                                        Case "PCN????":         strBarData = strBarData & strPcnNO
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "??????_??":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "????LOT":         strBarData = strBarData & Format(txtSlittingNo.Text, "0") & txtCompNm.Text & Format(txtSlittingNo.Text, "0") & "01"
                                        Case "REEL????":        strBarData = strBarData & strProdLen
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
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
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "?????µ?" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "????????" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "??ȿ?Ⱓ" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "00" & ")"
                            strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            barReel.Alignment = bcACenter
                            barReel.Style = msSCode128B
                            barReel.Visible = True
                            barReel.Caption = strContents
                            barReel.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            barReel.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
                            barReel.WIDTH = 4365
                        Else
                            If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            Else
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            End If
                        End If
                            
                        strContents = ""
                                                        
                    '============== CF-TP400E ======================================================================
                    Case "CF-TP500B|350", "CF-TP500E|350"
                        '-- ???ڵ? ???? ???? ã?ƿ???
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            'strBarData = ""
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                        
                            If pAdoRS1 Is Nothing Then
                                '???ϵ? ???? ????
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "ACF":         strBarData = "C"
                                        Case "????LOT":     strBarData = strBarData & strLotNo
                                        Case "_":           strBarData = strBarData & "_"
                                        Case "P/N":         strBarData = strBarData & "101"
                                        Case "??ȿ?Ⱓ":    strBarData = strBarData & Format(strExDate, "yyyymmdd")
                                    End Select
                        
                                    pAdoRS1.MoveNext
                                Loop
                                pAdoRS1.Close
                            End If
                            '============================== ???ڵ? ???? ã?ƿ??? ==============================
        
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                        
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            gPackTrack.REELBAR = strContents
                            strBarData = ""
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "?????µ?" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "????????" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "??ȿ?Ⱓ" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            'strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo & strLotNoFull, .MaxRows, 3)
                            strContents = strLotNo
                        End If
                        strLeft = 0
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ?" Then
                            barReel.Alignment = bcACenter
                            barReel.Style = msSCode128B
                            barReel.Visible = True
                            barReel.Caption = strContents
                            barReel.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            barReel.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
                            barReel.WIDTH = 4365
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
                            barPart.Alignment = bcACenter
                            barPart.Style = msSCode128B
                            barPart.Visible = True
                            barPart.Caption = strContents
                            barPart.LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                            barPart.TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 3
                            barPart.WIDTH = 3000
                        Else
                            If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "???ڵ尪" Then
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            Else
                                '?׸??? ???¿???
                                If AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y" Then
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = AdoRs.Fields("LABEL_ITEM_NAME").Value & " : " & strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                Else
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Visible = True
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").Caption = strContents
                                    strLeft = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") + Len(AdoRs.Fields("LABEL_ITEM_NAME").Value & "")
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").LEFT = CCur(AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "") * 3
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").TOP = CCur(AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "") * 6
                                    lblTitle(AdoRs.Fields("LABEL_ITEM_SEQ").Value & "").BackColor = vbWhite
                                End If
                            End If
                        End If
                            
                        strContents = ""
                                                        
                                                        
                End Select
                
                AdoRs.MoveNext
            End With
        Loop
    End If
    
    AdoRs.Close

'    Call GetReelOrderList(strFromDt, strToDt, "", "", "R")

End Sub



