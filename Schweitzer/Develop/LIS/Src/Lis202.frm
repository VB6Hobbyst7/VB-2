VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm202AccDataEntry 
   BackColor       =   &H00E0E0E0&
   Caption         =   "접수번호별 결과등록"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15165
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1022
   Icon            =   "Lis202.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15165
   Tag             =   "20200"
   WindowState     =   2  '최대화
   Begin VB.Frame frmSMS 
      BackColor       =   &H00F8E4D8&
      Caption         =   "SMS전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   5160
      TabIndex        =   59
      Top             =   1530
      Width           =   4515
      Begin VB.TextBox txtExDtNo 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   80
         Tag             =   "opt"
         Top             =   2190
         Width           =   2325
      End
      Begin VB.TextBox txtExDtNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   79
         Tag             =   "opt"
         Top             =   1800
         Width           =   1005
      End
      Begin VB.TextBox txtExDtId 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3030
         MaxLength       =   15
         TabIndex        =   78
         Tag             =   "opt"
         Top             =   1800
         Width           =   1305
      End
      Begin VB.CommandButton cmdTrans 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전송"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         Style           =   1  '그래픽
         TabIndex        =   71
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
      End
      Begin VB.CommandButton cmdCancle 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3030
         Style           =   1  '그래픽
         TabIndex        =   70
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
      End
      Begin VB.TextBox txtTransId 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   69
         Tag             =   "opt"
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtTransNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   68
         Tag             =   "opt"
         Top             =   300
         Width           =   1875
      End
      Begin VB.TextBox txtTransNo 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   67
         Tag             =   "opt"
         Top             =   630
         Width           =   3195
      End
      Begin VB.TextBox txtDtId 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3030
         MaxLength       =   15
         TabIndex        =   66
         Tag             =   "opt"
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txtDtNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   65
         Tag             =   "opt"
         Top             =   1020
         Width           =   1005
      End
      Begin VB.TextBox txtDetpCd 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   64
         Tag             =   "opt"
         Top             =   2580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDeptNm 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   63
         Tag             =   "opt"
         Top             =   2580
         Width           =   3195
      End
      Begin VB.TextBox txtDtNo 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   62
         Tag             =   "opt"
         Top             =   1410
         Width           =   2325
      End
      Begin VB.TextBox txtTransDt 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   61
         Tag             =   "opt"
         Top             =   4170
         Width           =   3195
      End
      Begin VB.TextBox txtTestCd 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5100
         MaxLength       =   15
         TabIndex        =   60
         Tag             =   "opt"
         Top             =   1350
         Width           =   1305
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   7
         Left            =   180
         TabIndex        =   72
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "전송자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   1905
         Index           =   8
         Left            =   180
         TabIndex        =   73
         Top             =   1020
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   3360
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
         Caption         =   "수신자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   74
         Top             =   2970
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
         Caption         =   "메시지"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   75
         Top             =   4200
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
         Caption         =   "전송일시"
         Appearance      =   0
      End
      Begin RichTextLib.RichTextBox rtfMessage 
         Height          =   1170
         Left            =   1140
         TabIndex        =   76
         Top             =   2970
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2064
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis202.frx":038A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   11
         Left            =   180
         TabIndex        =   77
         Top             =   630
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   765
         Index           =   13
         Left            =   1110
         TabIndex        =   81
         Top             =   1020
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1349
         BackColor       =   14737632
         ForeColor       =   0
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
         Height          =   765
         Index           =   14
         Left            =   1110
         TabIndex        =   82
         Top             =   1800
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1349
         BackColor       =   14737632
         ForeColor       =   0
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
         Caption         =   "주치의"
         Appearance      =   0
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   10035
      TabIndex        =   56
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "종합판독"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5880
      Style           =   1  '그래픽
      TabIndex        =   55
      Tag             =   "135"
      Top             =   8610
      Width           =   1080
   End
   Begin VB.CommandButton cmdSMS 
      BackColor       =   &H008080FF&
      Caption         =   "SMS"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9390
      Style           =   1  '그래픽
      TabIndex        =   54
      Tag             =   "135"
      Top             =   8610
      Width           =   1080
   End
   Begin VB.CommandButton cmdOrderView 
      BackColor       =   &H00F4F0F2&
      Caption         =   "처방별조회(&C)"
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
      Left            =   5100
      Style           =   1  '그래픽
      TabIndex        =   50
      Top             =   90
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "취소"
      Height          =   330
      Left            =   13740
      Style           =   1  '그래픽
      TabIndex        =   49
      Top             =   6270
      Width           =   690
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00ACCDD0&
      Caption         =   "적용"
      Enabled         =   0   'False
      Height          =   330
      Left            =   13035
      Style           =   1  '그래픽
      TabIndex        =   48
      Top             =   6270
      Width           =   690
   End
   Begin VB.TextBox txtBatchRst 
      Appearance      =   0  '평면
      BackColor       =   &H00F1F5F4&
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
      Left            =   11235
      MaxLength       =   15
      TabIndex        =   46
      Tag             =   "opt"
      Top             =   6270
      Width           =   1785
   End
   Begin MedControls1.LisLabel lblDisease 
      Height          =   270
      Left            =   8040
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   345
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   476
      BackColor       =   16777215
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
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H00DBE6E6&
      Caption         =   "특  수"
      Height          =   285
      Left            =   12795
      Style           =   1  '그래픽
      TabIndex        =   23
      Top             =   40
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdMicro 
      BackColor       =   &H00DBE6E6&
      Caption         =   "미생물"
      Height          =   285
      Left            =   13635
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   40
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdRmk 
      BackColor       =   &H008080FF&
      Caption         =   "처방비고"
      Height          =   285
      Left            =   11955
      Style           =   1  '그래픽
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   40
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Comment by Accession No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   75
      TabIndex        =   13
      Tag             =   "20003"
      Top             =   6600
      Width           =   6885
      Begin VB.CommandButton cmdRemarkTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         Picture         =   "Lis202.frx":0427
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   1560
         Width           =   315
      End
      Begin VB.CommandButton cmdCommentTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         Picture         =   "Lis202.frx":0959
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   945
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   990
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   1746
         _Version        =   393217
         BackColor       =   15857140
         ScrollBars      =   2
         TextRTF         =   $"Lis202.frx":0E8B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   360
         Left            =   90
         TabIndex        =   16
         Top             =   1530
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   635
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"Lis202.frx":10BD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCapRemark 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   1260
         Width           =   1545
      End
   End
   Begin VB.PictureBox picRst 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   75
      ScaleHeight     =   4605
      ScaleWidth      =   14340
      TabIndex        =   25
      Top             =   1500
      Width           =   14400
      Begin MSComctlLib.ProgressBar prgRst 
         Height          =   240
         Left            =   -15
         TabIndex        =   26
         ToolTipText     =   "자료를 가져오고 있읍니다."
         Top             =   4380
         Visible         =   0   'False
         Width           =   14070
         _ExtentX        =   24818
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread ssRst 
         CausesValidation=   0   'False
         Height          =   4620
         Left            =   -15
         TabIndex        =   27
         Tag             =   "20001"
         Top             =   -15
         Width           =   14355
         _Version        =   196608
         _ExtentX        =   25321
         _ExtentY        =   8149
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
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
         GrayAreaBackColor=   15857140
         MaxCols         =   19
         MaxRows         =   10
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "Lis202.frx":12F1
         VisibleCols     =   10
         VisibleRows     =   10
         TextTip         =   2
      End
      Begin VB.Label lblSpreadLoading 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "잠시 기다려 주세요. 결과 데이터를 로딩하고 있읍니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3330
         TabIndex        =   28
         Top             =   2520
         Width           =   6675
      End
   End
   Begin VB.CheckBox chkSpcNo 
      BackColor       =   &H00DBE6E6&
      Caption         =   "검체번호로 찾기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006B72A9&
      Height          =   285
      Left            =   5010
      TabIndex        =   20
      Top             =   660
      Width           =   1695
   End
   Begin VB.ComboBox cboRelTest 
      BackColor       =   &H00FFF9F7&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   630
      Width           =   6450
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "확인"
      CausesValidation=   0   'False
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
      TabIndex        =   5
      Tag             =   "135"
      Top             =   8595
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
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
      TabIndex        =   7
      Tag             =   "128"
      Top             =   8595
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
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
      TabIndex        =   6
      Tag             =   "124"
      Top             =   8595
      Width           =   1320
   End
   Begin VB.Frame fraText 
      BackColor       =   &H00DBE6E6&
      Caption         =   " Text Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   7215
      TabIndex        =   12
      Tag             =   "20002"
      Top             =   6600
      Width           =   7260
      Begin VB.CommandButton cmdTextTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Picture         =   "Lis202.frx":1B20
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   1005
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   1050
         Left            =   75
         TabIndex        =   4
         Top             =   270
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   1852
         _Version        =   393217
         BackColor       =   15663102
         Enabled         =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis202.frx":2052
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfFlagText 
         Height          =   525
         Left            =   600
         TabIndex        =   57
         Top             =   1350
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   926
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis202.frx":22C5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   525
         Index           =   12
         Left            =   90
         TabIndex        =   58
         Top             =   1350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
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
         Caption         =   "Flag"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ListView lvwPatient 
      Height          =   555
      Left            =   75
      TabIndex        =   2
      Tag             =   "20113"
      Top             =   945
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   979
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   15857140
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   1
      Left            =   6705
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   45
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
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
      Caption         =   "연 락 처"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   2
      Left            =   6705
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   345
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
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
      Caption         =   "상 병 명"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   3
      Left            =   6705
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   630
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
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
      Caption         =   "관련검사 결과"
      Appearance      =   0
   End
   Begin VB.Frame fraCul 
      BackColor       =   &H00DBE6E6&
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
      Height          =   555
      Left            =   7215
      TabIndex        =   29
      Top             =   8595
      Width           =   2175
      Begin VB.CheckBox chkCul 
         BackColor       =   &H00DBE6E6&
         Caption         =   "부분"
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   90
         Width           =   660
      End
      Begin VB.CommandButton cmdCul 
         BackColor       =   &H00F4F0F2&
         Caption         =   "누적결과조회"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         Style           =   1  '그래픽
         TabIndex        =   30
         Tag             =   "135"
         Top             =   15
         Width           =   1320
      End
   End
   Begin MedControls1.LisLabel lblTelno 
      Height          =   270
      Left            =   8040
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   45
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   476
      BackColor       =   16777215
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
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.Frame fraKey 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Index           =   0
      Left            =   75
      TabIndex        =   18
      Top             =   -75
      Width           =   4935
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
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
         Height          =   364
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   2715
         Begin MSMask.MaskEdBox mskBarNo 
            Height          =   315
            Left            =   975
            TabIndex        =   52
            Top             =   0
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "  &&-#########"
            PromptChar      =   "_"
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   4
            Left            =   0
            TabIndex        =   53
            Top             =   0
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
      End
      Begin VB.Frame fraAccNo 
         BackColor       =   &H00DBE6E6&
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
         Height          =   364
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2715
         Begin MSMask.MaskEdBox mskAccNo 
            Height          =   315
            Left            =   975
            TabIndex        =   0
            Top             =   0
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "&&-######-#####"
            PromptChar      =   "_"
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   6
            Left            =   0
            TabIndex        =   32
            Top             =   0
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
            Caption         =   "접수번호"
            Appearance      =   0
         End
      End
      Begin VB.CommandButton cmdNextData 
         BackColor       =   &H00EAE7E3&
         Caption         =   "(&N) >>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3915
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   330
         Width           =   825
      End
      Begin VB.CommandButton cmdNextData 
         BackColor       =   &H00EAE7E3&
         Caption         =   "<< (&P)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3000
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   330
         Width           =   870
      End
   End
   Begin VB.Frame fraKey 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   1
      Left            =   75
      TabIndex        =   21
      Top             =   -75
      Width           =   4935
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
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
         Height          =   364
         Left            =   105
         TabIndex        =   38
         Top             =   345
         Width           =   2715
         Begin MSMask.MaskEdBox mskSpcNo 
            Height          =   315
            Left            =   975
            TabIndex        =   39
            Top             =   0
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "############"
            PromptChar      =   "_"
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   0
            Left            =   50
            TabIndex        =   40
            Top             =   0
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
      End
      Begin VB.Label lblLabNo 
         BackStyle       =   0  '투명
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2865
         TabIndex        =   22
         Top             =   390
         Width           =   2070
      End
   End
   Begin VB.Frame fraMesg 
      BackColor       =   &H00DBE6E6&
      Height          =   2655
      Left            =   10380
      TabIndex        =   42
      Top             =   555
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "확인"
         Height          =   420
         Left            =   2940
         Style           =   1  '그래픽
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1095
      End
      Begin VB.TextBox txtMesg 
         Height          =   1785
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   43
         Top             =   390
         Width           =   4050
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   2
         Left            =   15
         TabIndex        =   45
         Top             =   90
         Width           =   4050
         _ExtentX        =   7144
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
         Caption         =   "처방 비고사항 조회"
         Appearance      =   0
         LeftGab         =   200
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   330
      Index           =   5
      Left            =   10290
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6270
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   582
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
      Caption         =   "배치결과"
      Appearance      =   0
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDF0F5&
      BackStyle       =   0  '투명
      Caption         =   "오류가 발생했다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00313D46&
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   8685
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Index           =   1
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   8670
      Width           =   5745
   End
End
Attribute VB_Name = "frm202AccDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Public WithEvents clsTemplete   As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
Private WithEvents objCuM       As frmTmpCumulative
Attribute objCuM.VB_VarHelpID = -1
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
'Private insForm     As Form
Private objLab032   As clsComcode032
Private objLab301   As clsWSList
Private objPtInfo   As clsPatientInfo

Private gblnNewObj      As Boolean
Private blnFirst        As Boolean
Private blnDayCount     As Boolean
Private gstrPtAddInfo   As String
Private gblnModify      As Boolean
Private gstrModifyData  As String

Private LastWorkArea    As String
Private LastAccDt       As String
Private LastAccSeq      As String
Private MsgFg           As Boolean
Private DataFg          As Boolean
Private ClearFg         As Boolean
Private LeaveCellFg As Boolean

Private strCombo        As String

Private gintTemplete    As Integer
Private blnRstChange    As Boolean
Private blnExpect       As Boolean

Private AdoCn_SQL       As ADODB.Connection
Private AdoRs_SQL       As ADODB.Recordset

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset
Dim strRcvDt  As String
Dim strTmpAge As String
Dim strTmpSex As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdApply_Click()
    Dim strBatchRst As String
    Dim i           As Long
    
    If Trim(txtBatchRst.Text) = "" Then
        txtBatchRst.SetFocus
        Exit Sub
    End If
    
    '** 중요사항
    ' * 컬럼 중 기존에 사용치 않고 있는 14는 배치결과 등록 시 사용한다.
    strBatchRst = Trim(txtBatchRst.Text)
    i = 1
    With ssRst
        If .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        For i = 1 To .DataRowCnt
            .Row = i
            
            .Col = 1
            If .BackColor = vbGrayText Then
                GoTo Skip
            End If
            
            .Col = 2:
            'If Trim(.Value) = "" Then
            If Len(Trim(.Value & "")) = 0 Then  '결과값없는것만 처리한다. 2014-08-13 PSK
                .Value = strBatchRst
                .ForeColor = DCM_LightRed
                
                '-- Batch Result
                .Col = .MaxCols: .Value = strBatchRst
                
                '-- Batch Result Check Flag
                .Col = 14: .Value = 1
            End If
            
            'Call ssRst_LeaveCell(2, i, -1, -1, False) 2014-08-14 PSK 맊음
            
Skip:

        Next
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Dim i       As Long
    
    i = 1
    With ssRst
        If .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        For i = 1 To .DataRowCnt
            .Row = i
            
            '-- Batch Result
            .Col = 14:
            If Trim(.Value) = 1 Then
                .Value = ""
                
                '-- Batch Result
                .Col = 2: .Value = ""
                
                '-- Batch Result Check Flag
                .Col = 14: .Value = ""
                
                '-- 판정 Clear
                .Col = 6: .Value = ""
                .Col = 7: .Value = ""
                .Col = 8: .Value = ""
                
            End If
        Next
    End With
    
End Sub

Private Sub cmdCancle_Click()
    frmSMS.Visible = False
End Sub

Private Sub cmdOrderView_Click()
' 2008.12.17. 양성현 작업중입니다.
' 2009.01.09 양성현 환자ID 파라메터 추가
' 2009.04.13 양성현 추가
    Dim i As Integer
    Dim pFrmName As String
'    Dim cxxx  As S2LIS_ReviewLib.clsLISResultReview
    pFrmName = "frm401ResultView"
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    medMain.lblSubMenu.Caption = "처방결과조회" 'medGetP(Button.Tag, 1, "(")
    
    
'   gPatientId = lblPtId.Caption
'  s2lis_reviewlib.PtId = lblPtId.Caption
    
'    gUsingInWardMenu = True
    frmLisReview.ButtonKey = "LIS155A" 'Button.Key
    frmLisReview.PtId = lvwPatient.ListItems(1)  'lblPtId.Caption
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

        Exit Sub

PermissionDenied:
   
'    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"

End Sub

Private Sub cmdSMS_Click()
    Dim SSQL As String
    
    Set AdoCn_ORACLE = New ADODB.Connection
    
    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
'        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("Persist Security Info") = True
        
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        
'        Screen.MousePointer = vbHourglass
        .Open
    End With
    
    frmSMS.Visible = True
    txtTransId.Text = Trim(ObjSysInfo.EmpId)
    txtTransNm.Text = GetEmpNm(Trim(ObjSysInfo.EmpId))
    txtTransNo.Text = mskAccNo.Text
    txtDtNo.Text = ""
    txtTransDt.Text = Format(Now, "YYYY-MM-DD HH:MM:DD")
    
    rtfMessage.Text = rtfMessage.Text & vbCRLF & "Critical value 즉시처치요함" & vbCr ' & rtfComment.Text
    If txtDtId.Text <> "" Then
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT TELNO,EMPNO FROM S2COM098"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
                       
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"

'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = (select orddoct from s2lab201 where workarea = '" & objPtInfo.WorkArea & "' and accdt =  '" & objPtInfo.AccDt & "' and accseq = '" & objPtInfo.AccSeq & "')"

        SSQL = ""
        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, EMPNM AS EMPNM from gainsamt"
        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = '" & txtDtId.Text & "' "

        Set AdoRs_ORACLE = New ADODB.Recordset
    
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
    
        If AdoRs_ORACLE.RecordCount > 0 Then
            txtDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
            txtDtNm.Text = AdoRs_ORACLE.Fields("EMPNM") & ""
        End If
'
'        Set AdoCn_ORACLE = Nothing
    End If
    
    If txtExDtId.Text <> "" Then
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT TELNO,EMPNO FROM S2COM098"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
                       
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"

'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = (select orddoct from s2lab201 where workarea = '" & objPtInfo.WorkArea & "' and accdt =  '" & objPtInfo.AccDt & "' and accseq = '" & objPtInfo.AccSeq & "')"

        SSQL = ""
        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, EMPNM AS EMPNM from gainsamt"
        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = '" & txtExDtId.Text & "' "

        Set AdoRs_ORACLE = New ADODB.Recordset
    
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
    
        If AdoRs_ORACLE.RecordCount > 0 Then
            txtExDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
            txtExDtNm.Text = AdoRs_ORACLE.Fields("EMPNM") & ""
        End If
        
        Set AdoCn_ORACLE = Nothing
    End If
End Sub

Private Sub cmdSpecial_Click()
    frmRealTestShow.DrFrame1.ZOrder 0
    frmRealTestShow.LisLabel7(0).Caption = "특수검사 관련검사 리스트"
    frmRealTestShow.Show
    Call frmRealTestShow.SpecialTest(lvwPatient.ListItems.Item(1).Text, lvwPatient.ListItems.Item(1).SubItems(1), cboRelTest, "1")
End Sub

Private Sub cmdMicro_Click()
    frmRealTestShow.DrFrame2.ZOrder 0
    frmRealTestShow.LisLabel7(0).Caption = "미생물 관련검사 리스트"
    frmRealTestShow.Show
    Call frmRealTestShow.SpecialTest(lvwPatient.ListItems.Item(1).Text, lvwPatient.ListItems.Item(1).SubItems(1), cboRelTest, "2")
End Sub

Public Sub chkSpcNo_Click()

    On Error GoTo Err_Trap:

    If chkSpcNo.Value = 1 Then
        fraKey(0).Visible = False
        fraKey(1).Visible = True
        mskAccNo.Enabled = False
        mskSpcNo.Enabled = True
        mskSpcNo.SetFocus
    Else
        fraKey(0).Visible = True
        fraKey(1).Visible = False
        mskSpcNo.Enabled = False
        mskAccNo.Enabled = True
        mskAccNo.SetFocus
    End If
    Exit Sub

Err_Trap:

End Sub

Private Sub cmdClear_Click()
    Call ClearData
End Sub

Private Sub cmdExit_Click()
    
    Dim intYesNo As VbMsgBoxResult

    If gblnModify = True Then
        If DataFetch <> gstrModifyData Then
            intYesNo = MsgBox("자료가 수정되었습니다." & vbNewLine & "수정된 자료를 저장하시겠습니까?", vbYesNo, "결과등록")
            If intYesNo = vbYes Then Call cmdSave_Click    '데이타 저장
        End If
        gblnModify = False: gstrModifyData = ""
    End If
    Set objCuM = Nothing
    Set clsTemplete = Nothing
    Set objLab301 = Nothing
    Set objPtInfo = Nothing
    Unload Me
    Set frm202AccDataEntry = Nothing
    
End Sub

Private Sub cmdNextData_Click(Index As Integer)

    Dim NextSeq As String
    Dim strMask As String
    Dim intYesNo As VbMsgBoxResult

    If gblnModify = True Then
        If DataFetch <> gstrModifyData Then
            intYesNo = MsgBox("자료가 수정되었읍니다." & vbNewLine & "수정된 자료를 저장하시겠슴니까?", vbYesNo, "결과등록")
            If intYesNo = vbYes Then Call cmdSave_Click    '데이타 저장
        End If
        gblnModify = False: gstrModifyData = ""
    End If

    If LastWorkArea <> "" And LastAccDt <> "" And LastAccSeq <> "" Then
        NextSeq = objPtInfo.GetAccSeq(LastWorkArea, LastAccDt, LastAccSeq, Index + 1)
        If NextSeq <> "" Then
            strMask = String(Len(LastWorkArea), "&") & "-"
            strMask = strMask & String(Len(Mid(LastAccDt, 3)), "#") & "-"
            strMask = strMask & String(Len(NextSeq), "#")
            'mskAccNo.Text = LastWorkArea & "-" & Mid(LastAccDt, 3) & String(6 - Len(Mid(LastAccDt, 3)), "_") & "-" & NextSeq & String(4 - Len(NextSeq), "_")
            mskAccNo.Mask = strMask
            mskAccNo.Text = LastWorkArea & "-" & Mid(LastAccDt, 3) & "-" & NextSeq
            Call Data_Load
        Else
            MsgBox "해당 데이타가 없습니다."
        End If
    End If

End Sub

Private Sub cmdRemarkTemplete_Click()
    
'    Dim SqlStmt As String
'
'    Set objCodeList = Nothing
'    Set objCodeList = New clsPopUpList
'
'    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    
    Dim RS          As Recordset
    Dim SqlStmt     As String
    Dim strWorkArea As String
    
    Set objCodeList = Nothing
    Set objCodeList = New clsPopUpList
    
    strWorkArea = Mid(mskAccNo.ClipText, 1, 2)
    
    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark) & " and " & DBW("field1=", strWorkArea)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    If RS.EOF Then
        SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    End If
    Set RS = Nothing
    
    With objCodeList
'        Set .MyOraSE = OraSE
'        Set .MyDb = dbconn
        .Connection = DBConn
        .FormCaption = "Remark"
        .ColumnHeaderText = "Code;Remark"
'        .HideColumnHeaders = True
        .ColumnHeaderWidth = "840.189;5309.858"
        .FormHeight = 3105
        .FormWidth = 6605
        .HideSearchTool = True
        .SelectByClick = True
        .LoadPopUp SqlStmt
        
'        .ListCaption = "Remark"
'        .ListColHeader = "Code" & vbTab & "Remark"
'        .Top = Me.cmdRemarkTemplete.Top + 5600
'        .Left = Me.cmdRemarkTemplete.Left + 200
'        .Width = 6250
'        .Height = 3000
'        .Tag = "Remark"
'        .CaptionOn = True
'        .MultiSel = False
'        .PopupList SqlStmt, 2
'        .ListAdd vbTab & "< 없 음 > ", 2, 1
    End With

End Sub
Private Function DiffSaveCheck() As Boolean
    '===================================================================
    'DIFF COUNT CHECK
    '마스터에 DIFF COUNT 코드에 등록된 코드의 합이 100이 아니면 안된다.
    'S2LAB032 에 CDINDEX=LC3_WBCDiffCode 이며 검사코드는 CDVAL1임
    '해당 CDVAL1의 모든 값의 합이 100이 아니면 안됩니다.
    '===================================================================
    Dim objDic As New clsDictionary
    Dim SSQL   As String
    Dim RS     As Recordset
    Dim ii     As Long
    
    Dim sValue As String
    Dim tValue As String
    Dim blnCheck As Boolean
    
    objDic.Clear
    objDic.FieldInialize "testcd", "rstcd"
    SSQL = objPtInfo.DiffCountSQL
    tValue = "0"
    
    blnCheck = False
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            objDic.AddNew RS.Fields("cdval1").Value & "", ""
            RS.MoveNext
        Loop
        For ii = 1 To ssRst.MaxRows
            With objPtInfo.Result.Item(ii)
                If objDic.Exists(.TestCd) Then
                    If .SpcCd = P_DiffSpcCd Then
                        blnCheck = True
                        objDic.KeyChange .TestCd
                        ssRst.Row = ii
                        ssRst.Col = objPtInfo.SSCol("RESULT")
                        objDic.Fields("rstcd") = ssRst.Value
                    End If
                End If
            End With
        Next
        objDic.MoveFirst
        Do Until objDic.EOF
            tValue = CDbl(tValue) + Val(objDic.Fields("rstcd"))
            objDic.MoveNext
        Loop
        
        If blnCheck = True And CDbl(tValue) <> 100 Then
            MsgBox "Diff Count 결과입력오류입니다." & _
                   "입력 총합계는 " & tValue & " 입니다.", vbCritical + vbOKOnly, "결과등록 오류"
            Set RS = Nothing
            Set objDic = Nothing
            Exit Function
        End If
    End If
    
    DiffSaveCheck = True
    Set RS = Nothing
    Set objDic = Nothing


End Function

Private Function ABORhCheck() As Boolean
    Dim ii As Integer
    
    For ii = 1 To ssRst.MaxRows
        With objPtInfo.Result.Item(ii)
            ssRst.Row = ii
            ssRst.Col = 19
            If .TestCd = P_ABOTestCD Or .TestCd = "LB2000" Then
                If .LastRst <> "" Then
                    If UCase(ssRst.Value) <> .LastRst Then
                        MsgBox "혈액형이 이전결과와 일치하지 않습니다.", vbOKOnly + vbCritical, "혈액형등록"
                        Exit Function
                    End If
                End If
            ElseIf .TestCd = P_RHTestCD Or .TestCd = "LB2021" Then
                If .LastRst <> "" Then
                    If UCase(ssRst.Value) <> .LastRst Then
                        MsgBox "RH가 이전결과와 일치하지 않습니다.", vbOKOnly + vbCritical, "RH등록"
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
    ABORhCheck = True
End Function

Private Sub cmdSave_Click()
    
    Dim ii As Long
    Dim blnDBSuccess As Boolean
    Dim strYesNo     As String
    
    ssRst.Row = 1: ssRst.Col = 2
    ssRst.Action = ActionActiveCell
    
    Call SpDispRtfText
            
    'WBC DIFF COUNT 결과체크
    If P_DiffFg Then
        If DiffSaveCheck = False Then
            strYesNo = MsgBox("결과등록을 하시겠습니까?.", vbInformation + vbYesNo, "결과등록")
            If strYesNo = vbNo Then Exit Sub
        End If
    End If
    
    '혈액형 결과체크
    If P_ABOCHK Then
        If ABORhCheck = False Then
            strYesNo = MsgBox("결과등록을 하시겠습니까?.", vbInformation + vbYesNo, "결과등록")
            If strYesNo = vbNo Then Exit Sub
        End If
    End If

    With objPtInfo
        .FootNote = rtfComment.Text
        .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
    End With
    '/*
    For ii = 1 To ssRst.MaxRows
        With objPtInfo.Result.Item(ii)
            ssRst.Row = ii
            ssRst.Col = objPtInfo.SSCol("RESULT")
'            If ssRst.Value = CS_EqpError Then
            If UCase(ssRst.Value) = UCase(CS_EqpError) Then
                ssRst.Action = ActionActiveCell
                Exit Sub
            End If
            'If .TxtType = "2" Then
            If .TxtType = "2" And .RstDiv = "R" Then
                If .TextRst = "" Or ssRst.Value = "" Then
                    '검사는 일반결과와 텍스트 결과를 같이 입력요. 결과보류 처리. (Required 항목일 경우만.. by KMK)
                    ssRst.Col = objPtInfo.SSCol("EC")
                    ssRst.Value = 1
                End If
            End If
            ssRst.Col = objPtInfo.SSCol("RESULT")
        End With
    Next ii
   '
    blnDBSuccess = objPtInfo.DataEntry 'objPtInfo                  '결과등록을 수행한다.
    If blnDBSuccess = False Then
        ClearData
        MsgBox objPtInfo.ErrNo & _
                " - " & objPtInfo.ErrText, vbCritical + vbOKOnly, "결과등록 ERROR"
        Exit Sub
    Else
        ClearData
    End If
   '
    If LastWorkArea <> "" And LastAccDt <> "" And LastAccSeq <> "" Then
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'            mskAccNo.Text = LastWorkArea & "-" & Mid(LastAccDt, 3) & String(6 - Len(Mid(LastAccDt, 3)), "_") & "-" & LastAccSeq & String(4 - Len(LastAccSeq), "_")
'            mskAccNo.SelStart = 10
'            mskAccNo.SelLength = 4
            mskAccNo.Text = LastWorkArea & "-" & Mid(LastAccDt, 3) & String(6 - Len(Mid(LastAccDt, 3)), "_") & "-" & LastAccSeq & String(5 - Len(LastAccSeq), "_")
            mskAccNo.SelStart = 10
            mskAccNo.SelLength = 5
    End If

    ssRst.MaxRows = 0
    lvwPatient.ListItems.Clear
    rtfText.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""
   '
   
    If P_RealPrinter = True Then
'    프린트 응급실 수술실 보내기
'        DoEvents
        Call PrintEROP24
'        DoEvents
    End If
End Sub


Private Sub PrintEROP24()
    Dim RS          As Recordset
    Dim objReport   As clsBatchReport
    Dim objSQL      As clsLISSqlReport
    Dim objDisease  As S2LIS_ReportLib.clsDisease
    Dim picESign    As Object
    Dim strSQL      As String
    Dim strEmpId    As String
    Dim strAge      As String
    Dim strWardID   As String
    
    Set objSQL = New clsLISSqlReport
    Set objReport = New clsBatchReport
    Set objDisease = New S2LIS_ReportLib.clsDisease
    
    ' 오라클기준으로 변경
    strSQL = " SELECT a.ptid,a.workarea,a.accdt,a.accseq,a.stscd,a.vfydt,a.vfytm, " & _
             "        d." & F_PTNM & " as ptnm, " & F_DOB2("d") & " as dob, d." & F_SEX & " as sex, " & _
             "        c.deptcd, c.wardid, c.majdoct " & _
             "   FROM " & T_HIS001 & " d, " & T_LAB101 & " c, " & T_LAB102 & " b, " & T_LAB201 & " a " & _
             "  WHERE " & DBW("a.workarea", LastWorkArea, 2) & _
             "    AND " & DBW("a.accdt", LastAccDt, 2) & _
             "    AND " & DBW("a.accseq", LastAccSeq, 2) & _
             "    AND " & DBW("a.stscd", StsCd_LIS_FinRst, 2) & _
             "    AND a.reqtotcnt=a.reqinputcnt " & _
             "    AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
             "    AND b.ptid=c.ptid AND b.orddt=c.orddt AND b.ordno=c.ordno " & _
             "    AND c.deptcd in ('ER','24') AND b.ptid = d." & F_PTID
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.RecordCount > 0 Then '
        With objReport
            .PtId = RS.Fields("ptid").Value & ""
            .ptnm = RS.Fields("ptnm").Value & ""
            .PtSex = RS.Fields("sex").Value & ""
            strAge = RS.Fields("dob").Value & ""
            If Len(strAge) = 6 Then strAge = strAge & "01"
            .PtAge = DateDiff("yyyy", Format(strAge, CS_DateMask), GetSystemDate)
                
            .FromDt = RS.Fields("vfydt").Value & ""
            .ToDt = RS.Fields("vfydt").Value & ""
            
            If Trim(RS.Fields("bussdiv").Value & "") = "1" Then
                .Dept = RS.Fields("deptcd").Value & ""
            Else
                .Dept = RS.Fields("wardid").Value & ""
            End If
            
            .VfyDt = RS.Fields("vfydt").Value & ""
'            If objLisComCode.DeptCd.Exists(Rs.Fields("deptcd").Value & "") Then
'                objLisComCode.DeptCd.KeyChange (Rs.Fields("deptcd").Value & "")
                .DeptNm = GetDeptNm(RS.Fields("deptcd").Value & "") 'objLisComCode.DeptCd.Fields("deptnm")
'            End If
            strWardID = Trim(RS.Fields("wardid").Value & "")
            .WardId = strWardID
'            If objLisComCode.WardId.Exists(strWardID) = True Then
'                objLisComCode.WardId.KeyChange (strWardID)
                objReport.Ward = GetWardNm(strWardID) 'objLisComCode.WardId.Fields("wardnm")
'            End If
            .Doct = GetEmpNm(RS.Fields("majdoct").Value & "")
            .VfyNM = ObjMyUser.EmpLngNm
            .MdfDt = ""
            objDisease.PtId = RS.Fields("ptid").Value & ""
            .ICD = objDisease.Disease
            Call .ReportForOneERPatient(RS.Fields("ptid").Value & "", RS.Fields("vfydt").Value & "", _
                                   RS.Fields("workarea").Value & "", RS.Fields("accdt").Value & "", _
                                   RS.Fields("accseq").Value & "", picESign, RS.Fields("vfydt").Value & "", _
                                   RS.Fields("vfydt").Value & "")
        End With
    End If
    
    If P_PrinterChkFg = True Then
        strSQL = " update " & T_LAB302 & _
                 "    set " & _
                            DBW("rptfg", "Y", 3) & _
                            DBW("rptdt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
                            DBW("rptid", ObjSysInfo.EmpId, 2) & _
                 "  WHERE " & DBW("a.workarea", LastWorkArea, 2) & _
                 "    AND " & DBW("a.accdt", LastAccDt, 2) & _
                 "    AND " & DBW("a.accseq", LastAccSeq, 2)
        
        DBConn.Execute strSQL
    End If
    
NoData:
    Set RS = Nothing
    Set objSQL = Nothing
    Set objReport = Nothing
    Set objDisease = Nothing
End Sub

'Private Function GetEmpNm(ByVal vEmpID As String) As String
'    Dim objData As New clsBasisData
'
'    GetEmpNm = objData.GetEmpNm(vEmpID)
'    Set objData = Nothing
'End Function

'Private Function GetDeptNm(ByVal vDeptCd As String) As String
'    Dim objData As New clsBasisData
'
'    GetDeptNm = objData.GetDeptNm(vDeptCd)
'    Set objData = Nothing
'End Function

'Private Function GetWardNm(ByVal vWardId As String) As String
'    Dim objData As New clsBasisData
'
'    GetWardNm = objData.GetWardNm(vWardId)
'    Set objData = Nothing
'End Function

Private Sub cmdTextTemplete_Click()
    If rtfText.Enabled = False Then Exit Sub
    Call CallTemplete(2, 0)
End Sub

Private Sub cmdCommentTemplete_Click()
    If ssRst.MaxRows < 1 Then Exit Sub
    Call CallTemplete(3, 0)
End Sub

'추가 : 온승호 SMS전송 모듈
'2010.01.18

Private Sub cmdTrans_Click()
    Dim ServerName   As String
    Dim DatabaseName As String
    Dim UserName     As String
    Dim Password     As String
    Dim strTransCd   As String
    Dim strDoctCd    As String
    Dim strTransDt   As String
    Dim strTransStatus As String
    Dim strTansEtc   As String
    Dim strMessage   As String
    Dim strTransNo   As String
    Dim strDoctNo    As String
    Dim strSQL       As String
    Dim strDeptNm    As String
    Dim strTranNm    As String
    Dim strSMSIP     As String
    Dim strBackNo    As String
    Dim strTmpTestCd As String
    Dim strMaDtId  As String
    Dim strMaTransNo As String
    
    Set AdoCn_ORACLE = New ADODB.Connection
    
    On Error Resume Next    '2013-09-11 PSK
    
    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
        .Properties("Persist Security Info") = True
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        .Open
    End With
           
    Set AdoRs_ORACLE = New ADODB.Recordset
        
    strSQL = ""
    strSQL = "SELECT * FROM S2lab032  "
    strSQL = strSQL + " WHERE cdindex = 'C232'"
    strSQL = strSQL + "   AND cdval1 = 'SVR1'  "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open strSQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            strSMSIP = AdoRs_ORACLE.Fields("FIELD4") & ""
        Else
            strSMSIP = "172.16.200.37"
        End If
        .Close
    End With
    
    Set AdoCn_SQL = New ADODB.Connection

    ServerName = strSMSIP
    DatabaseName = "medicalCRM_jesus"
    UserName = "jesus"
    Password = "jesus"
   
    With AdoCn_SQL
        .ConnectionTimeout = 10
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = ServerName
        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("User ID").Value = UserName
        .Properties("Password").Value = Password
        Screen.MousePointer = vbHourglass
        .Open
    End With
    Screen.MousePointer = vbDefault
    
'    If txtDtNo.Text = "" Then
'        MsgBox "수신번호를 입력하세요.", vbCritical + vbOKOnly, "수신번호등록 Message"
'        txtDtNo.SetFocus
'        Exit Sub
'    End If
    
    strTransCd = ObjSysInfo.EmpId
    strTransNo = txtTransNo.Text
    strDoctCd = txtDtId.Text
    strMaDtId = txtExDtId.Text
    strMaTransNo = txtExDtNo.Text
    strTransDt = Format(Now, "YYYY-MM-DD HH:MM:SS")
    strDoctNo = txtDtNo.Text
    strTransStatus = "1"
    strTansEtc = "LIS"
    strDeptNm = txtDeptNm.Text
    strTranNm = txtTransNm.Text
    strMessage = rtfMessage.Text & vbCRLF & "- " & strTranNm
    strBackNo = "063-230-8753"
    strTmpTestCd = txtTestCd.Text
    
    If Len(strMessage) > 80 Then
        MsgBox "메시지의 크기를 줄여주세요.", vbCritical + vbOKOnly, "메시지내용수정 Message"
        rtfMessage.SetFocus
        Exit Sub
    End If
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
    strSQL = strSQL & " values('" & strTransCd & "' ,"
    strSQL = strSQL & "        '" & strDoctNo & "' ,"
    strSQL = strSQL & "        '" & strBackNo & "' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '" & strTransDt & "' ,"
    strSQL = strSQL & "        '" & strTransStatus & "' ,"
    strSQL = strSQL & "        '" & strTansEtc & "')"
    
    AdoCn_SQL.Execute strSQL
    
    ' 검사코드 추가
    ' 2019-05-03 SMS 조회 검사코드로 조회 용
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT, TESTCD)"
    strSQL = strSQL & " values('" & strTransDt & "' ,"
    strSQL = strSQL & "        '" & strTransCd & "' ,"
    strSQL = strSQL & "        '" & strDoctNo & "' ,"
    strSQL = strSQL & "        '" & Trim(txtDtId.Text) & "' ,"
    strSQL = strSQL & "        '" & Trim(txtDtNm.Text) & "' ,"
    strSQL = strSQL & "        '" & strDeptNm & "' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '정상' ,"
    strSQL = strSQL & "        '" & strTransNo & "',"
    strSQL = strSQL & "        '" & strRcvDt & "',"
    strSQL = strSQL & "        '" & strTmpTestCd & "')"
    
    AdoCn_ORACLE.Execute strSQL
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
    strSQL = strSQL & " (select '" & strDoctCd & "' ,"
    strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
    strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
    strSQL = strSQL & "        '7' ,"
    strSQL = strSQL & "        SYSDATE ,"
    strSQL = strSQL & "        '[CVR(이상결과보고)]' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strDoctCd & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
    
    AdoCn_ORACLE.Execute strSQL
    
    If Trim(txtDtId.Text) <> Trim(txtExDtId.Text) Then
        strSQL = ""
        strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
        strSQL = strSQL & " values('" & strTransCd & "' ,"
        strSQL = strSQL & "        '" & txtExDtNo.Text & "' ,"
        strSQL = strSQL & "        '" & strBackNo & "' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '" & strTransDt & "' ,"
        strSQL = strSQL & "        '" & strTransStatus & "' ,"
        strSQL = strSQL & "        '" & strTansEtc & "')"
        
        AdoCn_SQL.Execute strSQL
        
        ' 검사코드 추가
        ' 2019-05-03 SMS 조회 검사코드로 조회 용
        
        strSQL = ""
        strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT, TESTCD)"
        strSQL = strSQL & " values('" & strTransDt & "' ,"
        strSQL = strSQL & "        '" & strTransCd & "' ,"
        strSQL = strSQL & "        '" & txtExDtNo.Text & "' ,"
        strSQL = strSQL & "        '" & Trim(txtExDtId.Text) & "' ,"
        strSQL = strSQL & "        '" & Trim(txtExDtNm.Text) & "' ,"
        strSQL = strSQL & "        '" & strDeptNm & "' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '정상' ,"
        strSQL = strSQL & "        '" & strTransNo & "',"
        strSQL = strSQL & "        '" & strRcvDt & "',"
        strSQL = strSQL & "        '" & strTmpTestCd & "')"
        
        AdoCn_ORACLE.Execute strSQL
        
        strSQL = ""
        strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
        strSQL = strSQL & " (select '" & strMaDtId & "' ,"
        strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
        strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
        strSQL = strSQL & "        '7' ,"
        strSQL = strSQL & "        SYSDATE ,"
        strSQL = strSQL & "        '[CVR(이상결과보고)]' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strMaDtId & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
        
        AdoCn_ORACLE.Execute strSQL
    End If
    
    strRcvDt = ""
    
    frmSMS.Visible = False
    Set AdoCn_SQL = Nothing
    Set AdoCn_ORACLE = Nothing
    
End Sub

Private Sub Command1_Click()
    Set objMyCmt = New clsLabComments
    With objMyCmt
        Set .SysInfo = ObjSysInfo
'                Set .MyDb = DBConn
        .DoctId = ObjMyUser.EmpId
        .DoctNm = ObjMyUser.EmpLngNm
        .PatId = lvwPatient.ListItems(1)
        .BedInDt = Text1.Text
        .ShowForm
    End With
End Sub

Private Sub Form_Activate()
   '
    medMain.lblSubMenu.Caption = Me.Caption
    If blnFirst = False Then
        Call LoadLvwHead
        blnFirst = True
        ClearData
    End If
    
    '누적결과및 관련검사(미생물/특수조회여부)
    If P_RealTestMicSpecial = True Then fraCul.Visible = True
End Sub

Private Sub Form_Load()
    
    Me.Show
    Call ClearData
    blnFirst = False
    gblnModify = False
'    chkSpcNo.Value = 0
'    Call chkSpcNo_Click
    
    prgRst.Align = vbAlignBottom
    prgRst.Visible = False
    ssRst.RowHeight(-1) = 13.6
    'ssRst.RowHeight(-1) = 15.6
    '
    cboRelTest.Clear
    cboRelTest.AddItem "관련 검사의 최근결과"
    cboRelTest.ListIndex = 0

    LastWorkArea = ""
    LastAccDt = ""
    LastAccSeq = ""
    frmSMS.Visible = False

End Sub

Private Sub clsTemplete_CopyTemplete()
   '
    If ssRst.MaxRows < 1 Then Exit Sub
    With objPtInfo
        Select Case gintTemplete
            Case 1:
                If clsTemplete.rtfText.Text <> "" Then
                    rtfRemark.Text = clsTemplete.rtfText.Text
                    .RmkCd = frm230TempSearch.lblCode.Caption
                    .RmkNm = rtfRemark.Text
                Else
                    rtfRemark.Text = ""
                    .RmkCd = ""
                    .RmkNm = ""
                End If
            Case 2:
                rtfText.Text = clsTemplete.rtfText.Text
                .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
                rtfText.SetFocus
            Case 3:
                rtfComment.Text = clsTemplete.rtfText.Text
                .FootNote = rtfComment.Text
                rtfComment.SetFocus
        End Select
    End With
    Set clsTemplete = Nothing
End Sub

Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    
    Dim strTitle As String
    Dim strWorkArea As String
    
    strWorkArea = Mid(mskAccNo.ClipText, 1, 2)
   
    Set clsTemplete = New frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
    With clsTemplete
        .qField1 = strWorkArea
        .Show
        If pintMode = 0 Then
            .lblName.Caption = "Edit " & strTitle
        Else
            .lblName.Caption = "Modify " & strTitle
        End If
        .Caption = strTitle & " " & "Templete Editor"
        .lblInfo.Caption = pintMode & "$" & pintPrg
        Select Case pintPrg
            Case 1:
                .lblCode.Caption = objPtInfo.RmkCd
                .rtfText = rtfRemark.Text
            Case 2:
                .rtfText = rtfText.Text
            Case 3:
                .rtfText = rtfComment.Text
        End Select
    End With
    gintTemplete = pintPrg
    
End Sub

Private Sub LoadLvwHead()
    
    Dim colHead As ColumnHeader
    Dim intMode As Integer
   
    '국가별 설정 모드
    intMode = 1         'Korea
    'intMode = 2         'English
    If intMode = 1 Then
        medInitLvwHead lvwPatient, "환자ID,환자성명,성/나이,생년월일,병상,주치의,검체,접수일자,비고(외부QC)", _
                                   "-100,150,-400,0,100,100,250,0"
'        medInitLvwHead lvwPatient, "환자ID,환자성명,성/나이,생년월일,병상,주치의,검체,접수일자", _
'                                   "-100,150,-400,0,100,100,0"
    Else
        medInitLvwHead lvwPatient, "Patient ID,Patient Name,Sex/Age,Location,Physician", _
                                    "0,0,0,0,0,0"
    End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCuM = Nothing
    Set clsTemplete = Nothing
    Set objLab301 = Nothing
    Set objPtInfo = Nothing
    Call ICSPatientMark
End Sub

Private Sub mskAccNo_KeyPress(KeyAscii As Integer)
    
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyReturn Then lvwPatient.SetFocus

End Sub


Private Sub mskAccNo_LostFocus()
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdNextData(0).Name Then Exit Sub
    If ActiveControl.Name = cmdNextData(1).Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If ActiveControl.Name = chkSpcNo.Name Then Exit Sub

    If Trim(mskAccNo.ClipText) = "" Then
        'Cancel = True
        lblErr.Caption = ""
        Exit Sub
    End If
    
    Call Data_Load

End Sub
   '
Public Sub Data_Load()
    
    Dim strBk As String
    Dim strSQL      As String
    Dim M2LAB302    As New ADODB.Recordset
    Dim M2LABFLAG   As New ADODB.Recordset
    
    strBk = mskAccNo.Text
    
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
    
    Call PtResultLoad(Trim(mskAccNo.FormattedText))
    
    If fraMesg.Visible Then fraMesg.Visible = False
    If cmdRmk.Visible Then cmdRmk.Visible = False
    
    If objPtInfo.TestCount > 0 Then
        ClearFg = False
        EditData
        lblErr.Caption = ""
        lvwPatient.SetFocus
        SendKeys "{TAB}"
        LastWorkArea = objPtInfo.WorkArea
        LastAccDt = objPtInfo.AccDt
        LastAccSeq = objPtInfo.AccSeq
        
        Dim MyResult As New clsLISResultReview
        Dim RS       As Recordset
        Dim SSQL     As String
        Dim ii       As Integer
        
        Call MyResult.GetRelTest(cboRelTest, mskAccNo.FormattedText)
        
        '------------------------추가 사항----------------------------
        '관련검사
        strCombo = ""
        For ii = 0 To cboRelTest.ListCount - 1
            strCombo = strCombo & cboRelTest.List(ii) & COL_DIV
        Next
        If strCombo <> "" Then strCombo = Mid(strCombo, 1, Len(strCombo) - 1)
        
        Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(1).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
        
        '처방리마크 조회(있는지 여부만 조회)
        SSQL = MyResult.GetOrderRemark(LastWorkArea, LastAccDt, LastAccSeq)
        Set RS = New Recordset
        RS.Open SSQL, DBConn
        
        If Not RS.EOF Then cmdRmk.Visible = True
        
        cboRelTest.ListIndex = 0
        Set RS = Nothing
        Set MyResult = Nothing
        fraAccNo.Enabled = False
        '------------------------추가 사항----------------------------
        gstrModifyData = DataFetch()
   '
        ssRst.Row = 1
        ssRst.Col = objPtInfo.SSCol("RESULT")
        ssRst.Action = ActionActiveCell
        DoEvents
        On Error Resume Next
        ssRst.SetFocus
        mskBarNo.Text = ""
        cmdApply.Enabled = True
        cmdOrderView.Visible = True
    Else
        'ClearData
        mskAccNo.Text = strBk
        ssRst.Visible = True
        MsgBox "해당 접수번호는 입력할 검사가 없습니다.", vbCritical + vbOKOnly, "결과등록 Message"
        'Cancel = True
        ClearData
        'mskAccNo.Mask = "&&-######-####"
        'mskAccNo.Text = "__-______-____"
        DoEvents
        If chkSpcNo.Value = 0 Then
            mskAccNo.SetFocus
        Else
            mskSpcNo.SetFocus
        End If
        
        cmdApply.Enabled = False
        
        cmdOrderView.Visible = False
        Exit Sub
    End If
    
    '-- 추가 By M.G.Choi 2007.04.06
'       등록된 결과가 있으면 Data Load 시 판정 실시
    Dim i       As Integer

    With ssRst
        For i = 1 To .DataRowCnt
            .Col = 2: .Row = i
            If IsNumeric(.Text) Then
                Call objPtInfo.NumValCheck(i)
            Else
                Call ssRst_LeaveCell(2, i, 2, i, False)
            End If
        Next
        .Col = 2: .Row = 1
        If IsNumeric(.Text) Then
            Call objPtInfo.NumValCheck(1)
        Else
            Call ssRst_LeaveCell(2, 1, 2, 1, False)
        End If
    End With
'    '---------------------------------------------------

    '==> 2014-07-15 PSK
    '======================================================================================
    ' B004101 : Micro. WBC (UA) 결과값  '1-4개' 이며
    ' B004118 : RBC morphology          항목이 존재할경우 FootNote 결과 뿌려주자...
    '======================================================================================
    Debug.Print mskAccNo.FormattedText
    Debug.Print medGetP(mskAccNo.FormattedText, 1, "-")
    Debug.Print medGetP(mskAccNo.FormattedText, 2, "-")
    Debug.Print medGetP(mskAccNo.FormattedText, 3, "-")
    
    strSQL = "         SELECT a.testcd, b.testnm, c.r_cnt, a.WORKAREA, a.ACCDT, a.ACCSEQ " & vbCRLF
    strSQL = strSQL & "FROM   S2LAB302 a, S2LAB001 b, " & vbCRLF
    strSQL = strSQL & "       ( SELECT workarea, accdt, accseq, nvl(COUNT(*),0) AS r_cnt FROM S2LAB302 where testcd = 'B004118' GROUP BY  workarea, accdt, accseq ) c " & vbCRLF
    strSQL = strSQL & "Where  b.TestCd = a.TestCd " & vbCRLF
    strSQL = strSQL & "AND    a.workarea = '" & medGetP(mskAccNo.FormattedText, 1, "-") & "' " & vbCRLF
    strSQL = strSQL & "AND    a.accdt = '" & "20" & medGetP(mskAccNo.FormattedText, 2, "-") & "' " & vbCRLF
    strSQL = strSQL & "AND    a.accseq = '" & medGetP(mskAccNo.FormattedText, 3, "-") & "' " & vbCRLF
    strSQL = strSQL & "AND    c.workarea = a.workarea " & vbCRLF
    strSQL = strSQL & "AND    c.accdt = a.accdt " & vbCRLF
    strSQL = strSQL & "AND    c.accseq = a.accseq " & vbCRLF
    strSQL = strSQL & "AND    a.TESTCD = 'B004102' " & vbCRLF
    strSQL = strSQL & "AND    LTRIM(a.rstcd) In ('1-4개','1개미만') "
    Debug.Print strSQL
    M2LAB302.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    If Not M2LAB302.EOF() And Not M2LAB302.BOF() Then
       M2LAB302.MoveFirst
       If Len(Trim(M2LAB302!TestCd & "")) > 0 And M2LAB302!r_cnt > 0 Then
          If Len(Trim(rtfComment.Text & "")) = 0 Then
             rtfComment.Text = "inadequate RBC number :" & vbCRLF & "RBC 1~4/H.P.F이하로 수가 너무 적은 경우" & vbCRLF & "결과의 정확도에 문제가 있으므로 RBC 형태 관찰은 보고하지 않습니다." & vbCRLF
          End If
       End If
    End If
    M2LAB302.Close: Set M2LAB302 = Nothing
    'rtfComment.TextRTF = "inadequate RBC number :" & vbCRLF & "RBC 1~4/H.P.F이하로 수가 너무 적은 경우" & vbCRLF & "결과의 정확도에 문제가 있으므로 RBC 형태 관찰은 보고하지 않습니다." & vbCRLF
    '======================================================================================
   
   strSQL = ""
   strSQL = " SELECT rsttxt FROM s2labflag WHERE workarea = '" & medGetP(mskAccNo.FormattedText, 1, "-") & "' AND accdt = '" & "20" & medGetP(mskAccNo.FormattedText, 2, "-") & "'  AND accseq = '" & medGetP(mskAccNo.FormattedText, 3, "-") & "' "
   
   M2LABFLAG.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
   
    If Not M2LABFLAG.EOF() And Not M2LABFLAG.BOF() Then
       M2LABFLAG.MoveFirst
       If Len(Trim(M2LABFLAG!RSTTXT & "")) > 0 Then
           rtfFlagText.Text = M2LABFLAG!RSTTXT & ""
       End If
    Else
       rtfFlagText.Text = ""
    End If
    M2LABFLAG.Close: Set M2LABFLAG = Nothing
    
    ssRst.Row = 1: ssRst.Col = 2
    ssRst.Action = ActionActiveCell
    
    Call SpDispRtfText
End Sub


Private Sub mskSpcNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lvwPatient.SetFocus
End Sub

Private Sub mskSpcNo_LostFocus()

    Dim strMsk As String
    Dim sSpcYy As String, sSpcNo As String
    Dim sWorkArea As String, sAccDt As String, sAccSeq As String

    If DataFg Then Exit Sub
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If ActiveControl.Name = chkSpcNo.Name Then Exit Sub

    If Len(mskSpcNo.ClipText) > 0 And Len(mskSpcNo.ClipText) < 12 Then
        mskSpcNo.SetFocus
    End If

    If Trim(mskSpcNo.ClipText) = "" Then Exit Sub
    
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
    
    sSpcYy = Mid(mskSpcNo.ClipText, 1, P_SpcYyLength)
    sSpcNo = Format(Mid(mskSpcNo.ClipText, P_SpcYyLength + 1, P_SpcNoLength), "#0")
    
    lblLabNo.Caption = objPtInfo.GetLabNoBySpcNo(sSpcYy, sSpcNo)
    
    sWorkArea = medGetP(lblLabNo.Caption, 1, "-")
    sAccDt = medGetP(lblLabNo.Caption, 2, "-")
    If Len(sAccDt) = 8 Then sAccDt = Mid(sAccDt, 3)
    
    sAccSeq = medGetP(lblLabNo.Caption, 3, "-")
    lblLabNo.Caption = sWorkArea & "-" & sAccDt & "-" & sAccSeq
    
    
    strMsk = String(Len(sWorkArea), "&") & "-"
    strMsk = strMsk & String(Len(sAccDt), "#") & "-"
    strMsk = strMsk & String(Len(sAccSeq), "#")
    mskAccNo.Mask = strMsk
    
    mskAccNo.Text = lblLabNo.Caption
    
    DoEvents

    DataFg = True
    
    Call Data_Load

End Sub


Private Sub objCodeList_SendCode(ByVal SelString As String)


    Dim strTmp As String
    '
    If Not IsNull(SelString) And SelString <> "" Then
        Select Case objCodeList.Tag
            Case "Remark":
                objPtInfo.RmkCd = medGetP(SelString, 1, vbTab)
                If Trim(objPtInfo.RmkCd) <> "" Then
                    objPtInfo.RmkNm = medGetP(SelString, 2, vbTab)
                Else
                    objPtInfo.RmkNm = ""
                End If
                rtfRemark.Text = objPtInfo.RmkNm
        End Select
    End If
    Set objCodeList = Nothing
   '
End Sub

'Private Sub objCodeList_ListClick(ByVal SelList As String)
'    objPtInfo.RmkCd = medGetP(SelList, 1, vbTab)
'    objPtInfo.RmkNm = medGetP(SelList, 2, vbTab)
'    rtfRemark.Text = objPtInfo.RmkNm
'End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    objPtInfo.RmkCd = objCodeList.SelectedItems(0)
    objPtInfo.RmkNm = objCodeList.SelectedItems(1)
    rtfRemark.Text = objPtInfo.RmkNm
End Sub

Private Sub rtfComment_Change()

    If ClearFg Then Exit Sub
    gblnModify = True

End Sub

Private Sub rtfRemark_Change()

    If ClearFg Then Exit Sub
    gblnModify = True

End Sub

Private Sub rtfText_Change()

    If ClearFg Then Exit Sub
    gblnModify = True
    '   objPtInfo.Result.Item(ssRst.Row).TextRst = rtfText.Text

End Sub

Private Sub rtfText_DblClick()
    Dim Domain As String
    Dim lngURLCnt   As Integer
    Dim strURL(10)  As String
    Dim strFlag As String
    Dim lngFCnt As Integer
    
    lngURLCnt = 0
    strURL(1) = "": strURL(2) = "": strURL(3) = "": strURL(4) = "": strURL(5) = ""
    strURL(6) = "": strURL(7) = "": strURL(8) = "": strURL(9) = "": strURL(10) = ""
    Debug.Print Trim(rtfText.Text)
    Domain = Trim(rtfText.Text)
    If InStr(Domain, vbCr) > 0 Or InStr(Domain, vbLf) > 0 Or InStr(Domain, vbCRLF) > 0 Then
        strFlag = "E"
        For lngFCnt = 1 To Len(Domain)
            Select Case Mid(Domain, lngFCnt, 1)
             Case vbCr, vbLf, vbCRLF
                  If Mid(Domain, lngFCnt + 1, 4) = "http" Then
                     lngURLCnt = lngURLCnt + 1
                     strFlag = "S"
                  Else
                     If lngURLCnt > 0 Then strFlag = "E"
                  End If
             Case Else
                If lngURLCnt > 0 And strFlag = "S" Then
                   strURL(lngURLCnt) = strURL(lngURLCnt) & Mid(Domain, lngFCnt, 1)
                End If
            End Select
        Next
        
        
        Debug.Print strURL(1)
        Debug.Print strURL(2)
        
        
        If Len(strURL(1)) > 0 And Mid(strURL(1), 1, 4) = "http" Then
           ShellExecute 1, vbNullString, strURL(1), vbNullString, vbNullString, 1
           Sleep 2000
        End If
        
        If Len(strURL(2)) > 0 And Mid(strURL(2), 1, 4) = "http" Then
           ShellExecute 2, vbNullString, strURL(2), vbNullString, vbNullString, 1
           Sleep 2000
        End If
        
        If Len(strURL(3)) > 0 And Mid(strURL(3), 1, 4) = "http" Then
           ShellExecute 0, vbNullString, strURL(3), vbNullString, vbNullString, 1
           Sleep 2000
        End If
        
        If Len(strURL(4)) > 0 And Mid(strURL(4), 1, 4) = "http" Then
           ShellExecute 0, vbNullString, strURL(4), vbNullString, vbNullString, 1
           Sleep 2000
        End If
        
        If Len(strURL(5)) > 0 And Mid(strURL(5), 1, 4) = "http" Then
           ShellExecute 0, vbNullString, strURL(5), vbNullString, vbNullString, 1
           Sleep 2000
        End If
    Else
        Domain = Trim(rtfText.Text)
        If InStr(Domain, "http") > 0 Then
            ShellExecute 0, vbNullString, Domain, vbNullString, vbNullString, 1
        End If
    End If
End Sub

Private Sub rtfText_LostFocus()
   '
    objPtInfo.Result.Item(ssRst.Row).TextRst = rtfText.Text
   '
End Sub

Private Function DataFetch() As String
    Dim ii As Integer
    
    DataFetch = ""
    With ssRst
        .Col = objPtInfo.SSCol("RESULT")
        .COL2 = objPtInfo.SSCol("EC")
        .Row = 1: .Row2 = .MaxRows
        DataFetch = .Clip & "$"
    End With
    With objPtInfo
        DataFetch = DataFetch & rtfComment.Text & "$" & rtfRemark.Text & "$"
        For ii = 1 To ssRst.MaxRows
            DataFetch = DataFetch & .Result.Item(ii).TextRst
        Next ii
    End With
    
End Function

Public Sub ClearData()
    ClearFg = True
    gblnModify = False
    gstrModifyData = ""
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'    mskAccNo.Mask = "&&-######-####"
'    mskAccNo.Text = "__-______-____"
    mskAccNo.Mask = "&&-######-#####"
    mskAccNo.Text = "__-______-_____"
    
    mskSpcNo.Text = "____________"
    lblLabNo.Caption = ""
    lblErr.Caption = ""
    lblDisease.Caption = ""
    lblTelno.Caption = ""
    If blnFirst = True Then
        fraAccNo.Enabled = True
        'mskAccNo.SetFocus
    End If
   '
    fraAccNo.Enabled = True
    ssRst.MaxRows = 0
    ssRst.Enabled = False
    chkSpcNo.Enabled = True
    mskAccNo.BackColor = vbWhite
    mskSpcNo.BackColor = vbWhite
    cmdSave.Enabled = False
    CmdTemplete False
   '
    cboRelTest.Clear
    lvwPatient.ListItems.Clear
    lvwPatient.BackColor = DCM_LightGray
    rtfComment.BackColor = DCM_LightGray
    rtfText.BackColor = DCM_LightGray
   '
    fraComment.Enabled = False
    lblCapRemark.Enabled = False
    MsgFg = False
    DataFg = False
    LeaveCellFg = False
   '
    rtfComment.Text = ""
    rtfText.Text = ""
    rtfRemark.Text = ""
    
    chkSpcNo.Value = 0
    Call chkSpcNo_Click
    cmdRmk.Visible = False
    fraMesg.Visible = False
    blnExpect = False
    
    cmdApply.Enabled = False
    txtBatchRst.Text = ""
    
End Sub

Private Sub EditData()
   '
    ssRst.Enabled = True
    '
    mskAccNo.BackColor = DCM_LightGray
    mskSpcNo.BackColor = DCM_LightGray
    chkSpcNo.Enabled = False
    cmdSave.Enabled = True
    '
    fraComment.Enabled = True
    lblCapRemark.Enabled = True
    '
    lvwPatient.BackColor = vbWhite
    rtfComment.BackColor = &HF1F5F4     'vbWhite
   '
End Sub

Private Function FormatUnder(ByRef strval As String, _
                             ByVal strSign As String) As String
    
    Dim intLen As Integer
    Dim ii As Integer
   
    If strSign = "+" Then
        FormatUnder = FormatUnder & CStr(Val(strval) + 1)
        strval = Val(strval) + 1
    Else
        FormatUnder = FormatUnder & CStr(Val(strval) - 1)
        strval = Val(strval) - 1
    End If
   '
    intLen = 4 - Len(strval)
    For ii = 1 To intLen
        FormatUnder = "_" & FormatUnder
    Next

    If Val(strval) < 1 Then
        FormatUnder = "___1"
    End If
    
End Function

Private Sub PtResultLoad(ByVal strAccNo As String)

    Dim strSpcYY  As String
    Dim strSpcNo  As String
    Dim valPtInfo As Variant
    Dim SSQL      As String

    lvwPatient.ListItems.Clear
    ssRst.Visible = False
    DoEvents
    
    MouseRunning
    
    Set objPtInfo.PrgBar = prgRst
    objPtInfo.PrgBarInit
    
    strTmpAge = ""
    
    With objPtInfo
        .PtType = RESULT_BY_ACCESSION             '/* 결과등록 유형, 반드시 셋팅 해야 됨./
        .AccNo = strAccNo       '/* 접수번호, 반드시 셋팅 해야 됨./
        
        .LoadTable , ObjMyUser.EmpId
        
        If .TestCount > 0 Then
            CmdTemplete True
            If lvwPatient.Enabled = False Then
                lvwPatient.Enabled = True
            End If
            medDataLoadLvw lvwPatient, vbNewLine, vbTab, .GetStringPtInfo
            
            valPtInfo = Split(.GetStringPtInfo, vbTab)
            
            txtDeptNm.Text = valPtInfo(4)
'            txtDtNm.Text = valPtInfo(5)
            txtDtId.Text = objPtInfo.OrdDoct
            txtExDtId.Text = objPtInfo.MajDoct
            
            strRcvDt = valPtInfo(7)
            Text1.Text = Format(.BedInDt, "####-##-##")
            rtfMessage.Text = "환자명 : " & valPtInfo(1) & "(" & valPtInfo(0) & ")" & vbCRLF
            strTmpAge = Trim(medGetP(valPtInfo(2), 2, "/"))
            strTmpSex = Trim(medGetP(valPtInfo(2), 1, "/"))
            
            Set AdoCn_ORACLE = New ADODB.Connection
            
            With AdoCn_ORACLE
                .ConnectionTimeout = 25
        '        .Provider = "OraOLEDB.Oracle.1"
                .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
                .Properties("Data Source").Value = "PMC"
        '        .Properties("Initial Catalog").Value = DatabaseName
                .Properties("Persist Security Info") = True
                
                .Properties("User ID").Value = "oral1"
                .Properties("Password").Value = "oral1"
                
        '        Screen.MousePointer = vbHourglass
                .Open
            End With
                       
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.PtId = lvwPatient.ListItems(1).Text
            lblDisease.Caption = objDisease.Disease
            lblDisease.ToolTipText = objDisease.Disease
            Set objDisease = Nothing
                
            rtfRemark.Text = .RmkNm
            rtfComment.Text = .FootNote
            
            If objPtInfo.Result.Item(1).TxtType <> "0" Then
                rtfText.Text = objPtInfo.Result.Item(1).TextRst
                rtfText.Enabled = True
                rtfText.BackColor = &HEEFFFE    'vbWhite
                cmdTextTemplete.Enabled = True
            Else
                rtfText.Enabled = False
                rtfText.BackColor = DCM_LightGray
                cmdTextTemplete.Enabled = False
            End If
            .GetResultSpread ssRst, RESULT_BY_ACCESSION
            '감염관리
            Call ICSPatientMark(lvwPatient.ListItems(1).Text, enICSNum.LIS_ALL)
            '병동/진료과 연락처(환자ID,CONTROL)
            
            Call GetPtTelInfo(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq, lblTelno, strSpcYY, strSpcNo)
            mskBarNo.Text = "  " & strSpcYY & "-" & Format(strSpcNo, "000000000")
        Else
            Call ICSPatientMark
        End If
    End With
    
    Dim ii As Integer
    
    With ssRst
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 5: .ForeColor = DCM_LightRed: .FontBold = True
        Next
    End With
    
    
    MouseDefault
    
    objPtInfo.PrgBarClear
    DoEvents
    
    Set AdoCn_ORACLE = Nothing
   '
End Sub

Private Sub ssRst_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If ClearFg Then Exit Sub
    gblnModify = True

End Sub

Private Sub ssRst_Click(ByVal Col As Long, ByVal Row As Long)
   '
    Dim i As Long
    Dim strTestCd As String
    Dim strSpcCd As String
    Dim strCalType As String
    Dim strTmpVal As String
    
    Dim dblTotVolume As Double
    Dim dblSerumCrea As Double
    Dim dblUrineCrea As Double
    
    Dim dblCal1     As Double
    Dim dblCal2     As Double
    Dim dblCal3     As Double
    Dim dblCal4     As Double

    Dim strTmp       As String
    
    If ClearFg Then Exit Sub
     '## 보류표시 Clear
    If Row = 0 And Col = 4 Then
        With ssRst
            .Col = 4
            blnExpect = IIf(blnExpect, False, True)
            For i = 1 To .MaxRows
                .Row = i
                If .CellType = CellTypeCheckBox Then
                    .Value = IIf(blnExpect, 0, 1)
                End If
'                If .CellType = CellTypeCheckBox Then .Value = 0
'                gblnModify = True
            Next
            ssRst.Row = 1: ssRst.Col = 2
            ssRst.Action = ActionActiveCell
            
            Call SpDispRtfText
        End With
    End If
    
    If Row <= 0 Then Exit Sub
    If Col = 10 Then '검사장비
    End If
       
    Call SpDispRtfText
        
    If Col = 1 Then
        '부분누적결과
        If Row = 0 Then Exit Sub
        If Not P_RealTestMicSpecial Then Exit Sub
        ssRst.Row = Row:        ssRst.Col = Col
        If objPtInfo.Result.Item(Row).RstDiv = "*" Then
            If ssRst.ForeColor = vbWhite Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = vbWhite
            End If
        Else
            If ssRst.ForeColor = DCM_MidBlue Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = DCM_MidBlue
            End If
        End If
        chkCul.Value = 0
        For i = 1 To ssRst.DataRowCnt
            ssRst.Row = i: ssRst.Col = 1
            If ssRst.ForeColor = DCM_LightRed Then
                chkCul.Value = 1
            End If
        Next
    
    ElseIf Col = 3 Then
        '2002-01-03추가 :계산공식 적용
        ssRst.Row = Row: ssRst.Col = 3
        If P_ApplyCalculation Then
            strTestCd = objPtInfo.Result.Item(Row).TestCd
            strSpcCd = objPtInfo.Result.Item(Row).SpcCd
            strCalType = objPtInfo.GetCalType(strTestCd, strSpcCd)
            
            If strCalType <> "" Then
                Select Case strCalType
                    Case "1", "2", "3"
                        '## 1: Creatinine, MTP, Ca, UA, BUN (24H Urine)
                        '## 2: Na, K, Cl, Amylase (24H Urine)
                        '## 3: Amylase (2H Urine)
                        '## Total Volume
                        strTmpVal = InputBox("Total Volume", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCRLF
                            End If
                        End If
                        
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "4"    '## CCR (24H Urine)
                        '## 1.Total Volume
                        strTmpVal = InputBox("Total Volume", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCRLF
                            End If
                        End If
                        
                        '## 2.Urine Creatinine
                        strTmpVal = InputBox("Urine Creatinine", "계산", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblUrineCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Urine Creatinine: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 3.Serum Creatinine
                        strTmpVal = InputBox("Serum Creatinine", "계산", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Serum Creatinine: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 4.키,몸무게 Factor
                        Dim dblHuman As Double
                        
                        strTmpVal = InputBox("체표면적", "계산", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblHuman = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "체표면적: " & strTmpVal & vbCRLF
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea, dblHuman)
                    Case "5"    '## LDL-Cholesterol (Serum)
                        '## 1.Cholesterol
                        strTmpVal = InputBox("Cholesterol", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Cholesterol: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 2.HDL-Cholesterol
                        strTmpVal = InputBox("HDL-Cholesterol", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblUrineCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "HDL-Cholesterol: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 3.TG
                        Dim dblTG As Double
                        
                        strTmpVal = InputBox("TG", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTG = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "TG: " & strTmpVal & vbCRLF
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea, dblTG)
                    Case "6"
                        '## 1.MPV
                        strTmpVal = InputBox("MPV", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                        End If
                        
                        '## 2.PLT
                        strTmpVal = InputBox("PLT", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "7"    '## ACCR 계산공식
                        '## 5.1.12: 이상대(2005-06-03)
                        '   - ACCR 계산공식 추가
                        '## 1.Amylase(Serum)
                        strTmpVal = InputBox("Amylase(Serum)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal1 = Val(strTmpVal)
                        End If
                        
                        '## 2.Creatinine(Serum)
                        strTmpVal = InputBox("Creatinine(Serum)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal2 = Val(strTmpVal)
                        End If
                        
                        '## 3.Amylase(24Urine)
                        strTmpVal = InputBox("Amylase(24Urine)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal3 = Val(strTmpVal)
                        End If
                        
                        '## 4.Creatinine(24Urine)
                        strTmpVal = InputBox("Creatinine(24Urine)", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal4 = Val(strTmpVal)
                        End If
                        
                        '## 5.Total Volumn
                        strTmpVal = InputBox("Total Volumn", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblCal1, dblCal2, dblCal3, dblCal4)
                    Case "8"
                        '## 2007.10.09 계산공식 추가 : Result = 검사결과값 / 특정 Creatnine 결과값
                        strTmpVal = InputBox("Creatnine", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Creatinine: " & strTmpVal & vbCRLF
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "10"
                        '## 1: Creatinine, MTP, Ca, UA, BUN (24H Urine)
                        '## 2: Na, K, Cl, Amylase (24H Urine)
                        '## 3: Amylase (2H Urine)
                        '## Total Volume
                        strTmpVal = InputBox("Total Volume", "계산", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCRLF
                            End If
                        End If
                        
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                End Select
            End If

            ssRst.Row = Row: ssRst.Col = 3
            ssRst.CellType = CellTypeStaticText
            ssRst.Text = "√"
            ssRst.ForeColor = DCM_Blue
            
            ssRst.Row = Row: ssRst.Col = 2
            ssRst.Action = ActionActiveCell
            ssRst.SetFocus
            DoEvents
            SendKeys "{Enter}"
        End If
    End If
End Sub

Private Sub ssRst_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim varTmp As Variant
    Dim strTest As String
    Dim strResult As String
    Dim strTestCd As String
    Dim strSQL          As String
    Dim RS              As New Recordset
    
    txtTestCd.Text = ""
    
    With ssRst
        .GetText 1, Row, varTmp: strTest = Trim(varTmp)
        .GetText 2, Row, varTmp: strResult = Trim(varTmp)
    End With
    rtfMessage.Text = rtfMessage.Text & strTest & " : " & strResult & vbCRLF
    
    strTestCd = objPtInfo.Result.Item(ssRst.ActiveRow).TestCd
    
    ' 2019-05-03 검사코드 추가
    txtTestCd.Text = strTestCd
    
    Select Case strTestCd
        Case "B2021", "LB2021"
            If UCase(strResult) = "NEGATIVE" Then
                Call cmdSMS_Click
            End If
        Case "B2061"
            If UCase(strResult) = "POSITIVE" Then
                Call cmdSMS_Click
            End If
        Case "ABOC22"
            If UCase(strResult) = "NEGATIVE" Then
                strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '00051'"
                RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
                If RS.RecordCount > 0 Then
                    RS.MoveFirst
                    rtfComment.Text = RS.Fields("text1") & ""
                End If
                RS.Close
            End If
        Case "B2602"
            If Val(strTmpAge) < 20 Or InStr(UCase(strTmpAge), "D") > 0 Then
                If strTmpSex = "M" Then
                    strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = 'ALP20M'"
                Else
                    strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = 'ALP20F'"
                End If
                RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
                If RS.RecordCount > 0 Then
                    RS.MoveFirst
                    rtfComment.Text = RS.Fields("text1") & ""
                End If
                RS.Close
            End If
        Case "CZ396"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0030'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CZ394"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0031'"
            RS.Open strSQL, DBConn
            
            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CZ394D"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0031'"
            RS.Open strSQL, DBConn
            
            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X274"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3066'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
'        Case "B2061"
'            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0050'"
'            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
'
'            If RS.RecordCount > 0 Then
'                RS.MoveFirst
'                rtfComment.Text = RS.Fields("text1") & ""
'            End If
'            RS.Close
        Case "B568"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3080'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B0260"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0029-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B109113"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1016'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4612"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3029-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C35901"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3063'"
            RS.Open strSQL, DBConn
            
            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435A"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435B"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435C"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435D"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435E"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435F"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435G"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C404"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "S641"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3064'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3260"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3065'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3530"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3068'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3630"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3069'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' 외부검사 코메트 처리 2016.03.28
        Case "27LB"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8116'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "27LC"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8191'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "27LN"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8051'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B145"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8025'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B1712"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8173'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B2700"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8106'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C208"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8196'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3241"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8004'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B3380"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8230'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3460"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8207'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3470"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8002'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3580"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8017'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3600"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8210'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3823"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8101'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3931"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8192'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C432"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8182'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450339"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8050'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4503390"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8012'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450401"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8206'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C45041"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8171'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450410"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8111'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450412"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8131'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450413"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8183'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450415"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8109'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450416"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8250'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452361"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8172'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452363"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8200'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452364"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8153'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452399"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8015'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4523991"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8047'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468242"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8108'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468244"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8112'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468245"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8114'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468246"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8151'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' 외부검사 필수항목
        Case "C468249"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8251'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468246"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8151'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468250"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8184'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468251"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8013'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468253"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8215'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468344"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8113'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468349"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8252'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468350"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8185'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468351"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8014'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468353"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8216'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C472261"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8234'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C472361"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8235'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C474281"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8103'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E300"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8100'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E426"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8119'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X146"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8170'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "S095"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8020'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X700"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8048'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X701"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8048'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X728"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8030'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X730"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8029'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "Z133"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8031'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "Z134"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8032'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "Z982"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8162'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23202"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23201"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' 2016-07-11 추가
        Case "C4912"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6020'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4913"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6020'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468342"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6012-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468343"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6017-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C549"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6010-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C46901"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6004'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724TB"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C46501"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6009'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468241"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6014'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B568"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3080'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C474381"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6011-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "H976"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6019-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4690596"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6003'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724TB"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23202A"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23201A"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' 2018.04.25 분자유전부 COMMENT 추가
        Case "S876"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1500'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "S606"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1501'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "RV1201G"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1510'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "H977"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1502'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "PNBPCR5G"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1525'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "AFBR5957"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1540'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CREPCR5"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1541'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C5154"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1503'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CY7512", "B4064", "S729TB", "X560TB", "B4021AC", "B4021D", "B4052A"

            rtfComment.Text = "결핵협회 결과입니다."
' 2018-09-13 추가
        Case "B2640", "C3942T"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3070'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' 코드변경 "C3842", "27EGER", "C4712ER
' 2020-01-28 변경
        Case "C3842", "27EGER", "C4712ER"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3070-2'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X274"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3066'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C2285"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3400"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062-2'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C2532"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062-3'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CZ394HIM", "CZ394ERM"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3059-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
    End Select
            
End Sub

Private Sub ssRst_EditChange(ByVal Col As Long, ByVal Row As Long)
    ssRst.Row = Row
    ssRst.Col = objPtInfo.SSCol("MAXCOL")
    ssRst.Value = ""
    If ClearFg Then Exit Sub
    gblnModify = True
End Sub

Private Sub ssRst_GotFocus()
    If MsgFg Then Exit Sub
    If LeaveCellFg Then Exit Sub

    With ssRst
        If .MaxRows = 0 Then Exit Sub
        .Row = 1
        .Col = objPtInfo.SSCol("RESULT")
        .Action = ActionActiveCell
        .EditEnterAction = EditEnterActionDown
    End With
    fraAccNo.Enabled = False
    
End Sub

Private Sub ssRst_KeyUp(KeyCode As Integer, Shift As Integer)
   '
    If KeyCode = 38 Or KeyCode = 40 Then
        SpDispRtfText
    ElseIf KeyCode = vbKeyF2 Then
        Call ssRst_RightClick(1, ssRst.ActiveCol, ssRst.ActiveRow, 100, 100)
    End If
  '
End Sub

'Private Sub ssRst_LostFocus()
'    Dim strTmp          As String
'    Dim strTmp1         As String
'    Dim strUTmp         As String
'    Dim strRstVal       As String
'
'    Dim strResultVal    As String
'    Dim strResultChk    As String
'    Dim strTestCd       As String
'
'    If ssRst.ActiveRow < 1 Then Exit Sub
'
'    ssRst.Row = ssRst.ActiveRow
'    ssRst.Col = objPtInfo.SSCol("RESULT")
'    strTestCd = objPtInfo.Result.Item(ssRst.ActiveRow).TestCd
'
'    strTmp = UCase(ssRst.Value)
'    strUTmp = ssRst.Value
'
'    ssRst.Col = objPtInfo.SSCol("MAXCOL"): strTmp1 = ssRst.Value
'    strRstVal = Trim(medGetP(objPtInfo.GetRstCdValString(strTestCd, strTmp1), 1, COL_DIV))
'
'    If strTmp = strRstVal Or strUTmp = strRstVal Then
'        blnRstChange = True
'        Exit Sub
'    End If
'
'    strResultVal = objPtInfo.GetRstCdValString(strTestCd, strTmp)
'    strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
'    strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
'
'    If strTmp <> strResultVal Then
'    '결과코드값이 있다.
'        ssRst.Col = objPtInfo.SSCol("RESULT"): ssRst.Value = strResultVal
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"): ssRst.Value = strTmp
'        If strResultChk <> "" Then
'            objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = ""
'            objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = ""
'            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        End If
'
'        Select Case strResultChk
'            Case "*"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = "N"
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
''                    objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = "N"
''                    ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
''                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "L"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "H"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'        End Select
'        blnRstChange = True
'    Else
'    '결과코드값이 없다
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"): ssRst.Value = strTmp
'        ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = ""
'    End If
'
'End Sub

Private Sub ssRst_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim strWorkArea     As String
Dim strAccDt        As String
Dim strAccSeq       As String
Dim strSQL          As String

Dim M2LAB302        As New ADODB.Recordset

    If ClickType <> 1 Then Exit Sub

    If MsgFg Then Exit Sub
    
    If Row <= 0 Then Exit Sub
    
    objPtInfo.SsTop = picRst.Top + 220
    objPtInfo.SsLeft = picRst.Left - 740
    ssRst.Row = Row
    ssRst.Col = Col
    ssRst.Action = ActionActiveCell
    objPtInfo.MfyFg = False
    MsgFg = True
    Call objPtInfo.PopUp(, Col)
    MsgFg = False
    
    '2014-08-13 PSK 결과변경시 결과값확인하여 FootNote 작성한다.
    Call objPtInfo.ResultCheck(Row)
    Select Case objPtInfo.Result.Item(Row).TestCd
     Case "B004102"
          strWorkArea = objPtInfo.Result.Item(Row).WorkArea
          strAccDt = objPtInfo.Result.Item(Row).AccDt
          strAccSeq = objPtInfo.Result.Item(Row).AccSeq
          
          strSQL = "SELECT * FROM S2LAB302 where testcd = 'B004118' AND workarea = '" & strWorkArea & "' " & vbCRLF
          strSQL = strSQL & "AND accdt = '" & strAccDt & "' AND accseq = " & strAccSeq & " "
          M2LAB302.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
          If Not M2LAB302.EOF() And Not M2LAB302.BOF() Then
             ssRst.Row = Row: ssRst.Col = Col
             Select Case Trim(ssRst.Text & "")
              Case "1-4개", "1개미만"
                   If Len(Trim(rtfComment.Text & "")) = 0 Then
                     rtfComment.Text = "inadequate RBC number :" & vbCRLF & "RBC 1~4/H.P.F이하로 수가 너무 적은 경우" & vbCRLF & "결과의 정확도에 문제가 있으므로 RBC 형태 관찰은 보고하지 않습니다." & vbCRLF
                  End If
             End Select
          End If
          M2LAB302.Close: Set M2LAB302 = Nothing
    End Select
End Sub

Private Sub ssRst_Advance(ByVal AdvanceNext As Boolean)
    Dim strCodeValue    As String
    Dim strRstType      As String
    Dim strErr          As String
    Dim strTestCd       As String
    Dim strResultVal    As String
    Dim strResultChk    As String
    Dim lngMaxCol       As String
    Dim lngResultCol    As String
    Dim strWorkArea     As String
    
    Dim Col             As Long
    Dim Row             As Long

    Row = ssRst.ActiveRow
    If Row < 0 Then Exit Sub
    lngResultCol = objPtInfo.SSCol("RESULT")
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    
    On Error GoTo ErrLevaeCell:

    Col = ssRst.ActiveCol
    If Col = lngResultCol Then
        Call objPtInfo.ResultCheck
        strRstType = objPtInfo.Result.Item(Row).RstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(Row).AvalVal & "자리)"
                Else
                    strErr = "유효숫자 입력 오류. (정수형만 입력)"
                End If
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
                Call objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
                strErr = "결과 입력 오류!"
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
            End If
        ElseIf strRstType = "R" Then
            If objPtInfo.IsRateCd = False Then
                strErr = "비율결과 입력 오류!"
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
                strErr = "FREE결과 입력 오류! (10자리이내)"
                GoTo ErrLevaeCell
            Else
                Call objPtInfo.NumValCheck
                lblErr.Caption = ""
            End If
        End If
    End If
    
    strTestCd = objPtInfo.Result.Item(Row).TestCd
    strWorkArea = objPtInfo.Result.Item(Row).WorkArea
    
    If Col = lngResultCol Then
        If strWorkArea = "05" Then
            ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = Trim(ssRst.Value)
        Else
            ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.Value))
        End If
        If strCodeValue = "" Then
            If strWorkArea = "05" Then
                ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = Trim(ssRst.Value)
            Else
                ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
            End If
        End If
'        ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue <> "" Then
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
        
            If strResultVal <> ssRst.Value Then
                ssRst.Row = Row: ssRst.Col = lngResultCol:  ssRst.Value = strResultVal
                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
            End If
            
        Else
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
            
            If strResultVal <> strCodeValue Then
                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                If strRstType = "F" Then
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                        ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = lngResultCol:   ssRst.Value = ""
                        ssRst.Col = lngMaxCol:      ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If
    
    LeaveCellFg = False

    Exit Sub
   '
ErrLevaeCell:
    lblErr.Caption = strErr
    ssRst.Value = ""
    
    MsgFg = True
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    MsgFg = False
    
    LeaveCellFg = True
    
    With ssRst
        .Row = Row
        .Col = objPtInfo.SSCol("RESULT")
        .Value = ""
        .Action = ActionActiveCell
        On Error Resume Next
        .SetFocus
    End With
    objPtInfo.ResultCheck
End Sub

Private Sub ssRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim strCodeValue    As String       '입력값
    Dim strRstType      As String       '결과타입
    Dim strErr          As String       '에러메세지
    Dim strTestCd       As String       '결과등록 검사코드
    Dim strResultVal    As String       '결과값
    Dim strResultChk    As String       '결과코드입력값 체크
    Dim lngResultCol    As Long         '결과입력 Col
    Dim lngMaxCol       As Long         '결과저장 Col
    Dim strWorkArea     As String
    
    strResultVal = "": strResultChk = ""
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    lngResultCol = objPtInfo.SSCol("RESULT")
    
    If Row < 1 Then Exit Sub
    If MsgFg Then Exit Sub

    On Error GoTo ErrLevaeCell
    If NewRow > 0 Then Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(NewRow).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
    
    If Row = ssRst.MaxRows Then
        'Advance 이벤트에서 포커스가 스프레드에서 다른컨트롤로 넘어갈시
        'LeaveCell이벤트의 뼁뼁이를 방지하기 위해서 exit sub를 씀
        '허나, ESR이 아닌 다른 아이템에 대해서는 항목이 하나일때 EXIT SUb를 빼면
        '참고치 체크가 안된다.
        blnRstChange = False
        If lngResultCol <> Col Then blnRstChange = True
        If blnRstChange = True Then Exit Sub
'        If lngResultCol = Col Then Call ssRst_LostFocus
'
'        If Me.ActiveControl.Name = ssRst.Name Then Exit Sub
        If blnRstChange = True Then Exit Sub
    Else
'        If NewRow > 0 Then Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(NewRow).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
'        If lngResultCol <> Col Then blnRstChange = True
'        If lngResultCol = Col Then Call ssRst_LostFocus
'        If blnRstChange = True Then Exit Sub
    End If
    
    lblErr.Caption = ""
    If Col = lngResultCol Then
'        Call objPtInfo.ResultCheck --온승호
        Call objPtInfo.ResultCheck(Row)
        strRstType = objPtInfo.Result.Item(Row).RstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(Row).AvalVal & "자리)"
                Else
                    strErr = "유효숫자 입력 오류. (정수형만 입력)"
                End If
                GoTo ErrLevaeCell
            Else
                objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            'If objPtInfo.IsAlphaCd = False Then --온승호
            If objPtInfo.IsAlphaCd(Row) = False Then
               strErr = "결과 입력 오류!"
               GoTo ErrLevaeCell
            End If
        ElseIf strRstType = "R" Then
            If objPtInfo.IsRateCd = False Then
               strErr = "비율결과 입력 오류!"
               GoTo ErrLevaeCell
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
               strErr = "FREE결과 입력 오류! (10자리이내)"
               GoTo ErrLevaeCell
            End If
            objPtInfo.NumValCheck
        End If
        ssRst.EditEnterAction = EditEnterActionDown
    End If
   '
    Call SpDispRtfText(NewRow)
    
    strTestCd = objPtInfo.Result.Item(Row).TestCd
    strWorkArea = objPtInfo.Result.Item(Row).WorkArea
    
    ' 2020-01-28
    ' 검사코드 E439 일시 예외조건 추가
    If Col = lngResultCol Then
        If strWorkArea = "05" Then
            ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = Trim(ssRst.Value)
        Else
            If strTestCd = "E439" Then
                ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = Trim(ssRst.Value)
            Else
                ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.Value))
            End If
        End If
        If strCodeValue = "" Then
            If strWorkArea = "05" Then
                ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = Trim(ssRst.Value)
            Else
                ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
            End If
        End If
'        ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue <> "" Then
            '저장 Col에 값이 있을경우(popup이용)
'            ssRst.Col = lngMaxCol:          ssRst.Value = strCodeValue
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '결과값
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '결과체크값
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '결과값
            
            ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
            ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'            If strResultChk <> "" Then
'                objPtInfo.Result.Item(Row).DPDiv = ""
'                objPtInfo.Result.Item(Row).HLDiv = ""
'            End If
            Select Case strResultChk
                Case "*"
                        objPtInfo.Result.Item(Row).HLDiv = "N"
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                                                                
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                Case "L"
                        objPtInfo.Result.Item(Row).HLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                Case "H"
                        objPtInfo.Result.Item(Row).HLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
            End Select
'            If strResultVal <> ssRst.Value Then
'                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
'                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
'                Select Case strResultChk
'                    Case "*"
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "L"
'                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "H"
'                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                End Select
'            Else
'                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
'            End If
        Else
            '저장Col에 값이 없을경우(직접입력)
            ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '결과값
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '결과체크값
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '결과값
            If strResultVal <> strCodeValue Then
                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = strResultChk
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = strResultChk
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = strResultChk
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                If strRstType = "F" Then
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                        ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = lngResultCol:   ssRst.Value = ""
                        ssRst.Col = lngMaxCol:      ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If
            
    ssRst.Row = Row
    ssRst.Col = 2
    If Trim(ssRst.Value) = "" Then
        ssRst.Col = 14: ssRst.Value = ""
    End If
       
    LeaveCellFg = False
    
    Exit Sub
   '
ErrLevaeCell:
    lblErr.Caption = strErr
    ssRst.EditEnterAction = EditEnterActionDown
    
    DoEvents
    
    With ssRst
        .Row = Row
        .Col = lngResultCol
        .Value = ""
        .Action = ActionActiveCell
    End With
    objPtInfo.ResultCheck
    
    MsgFg = True
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    MsgFg = False
    
    LeaveCellFg = True
        
    Cancel = True
    
    On Error Resume Next
    ssRst.SetFocus
End Sub

Private Sub ssRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
   '
    If Row < 1 Then Exit Sub
    objPtInfo.SpToolTip Row, Col, MultiLine, ShowTip, TipText, TipWidth
    ssRst.TextTip = TextTipFloatingFocusOnly
   '
End Sub

Private Sub SpDispRtfText(Optional Row As Long = 0)
   '
    If Row < 0 Then Exit Sub
    If Row = 0 Then
       ssRst.Row = ssRst.ActiveRow
    Else
       ssRst.Row = Row
    End If
    ssRst.Col = objPtInfo.SSCol("TXT")
    With objPtInfo.Result.Item(ssRst.Row)
        If ssRst.CellType = CellTypePicture Or ssRst.Text = "T" Then
            If .TxtType <> "0" Then
                rtfText.Text = .TextRst
                rtfText.Enabled = True
                cmdTextTemplete.Enabled = True
                rtfText.BackColor = &HEEFFFE    'vbWhite
            Else
                rtfText.Text = ""
                rtfText.Enabled = False
                cmdTextTemplete.Enabled = False
                rtfText.BackColor = DCM_LightGray
            End If
        Else
            rtfText.Text = ""
            rtfText.Enabled = False
            cmdTextTemplete.Enabled = False
            rtfText.BackColor = DCM_LightGray
        End If
    End With
   '
End Sub

Private Sub CmdTemplete(ByVal blnVisible As Boolean)
   '
    cmdTextTemplete.Enabled = blnVisible
    cmdRemarkTemplete.Enabled = blnVisible
    cmdCommentTemplete.Enabled = blnVisible
   '
End Sub

Private Function AccTrim(ByVal strval As String) As String
    
    Dim aryTmp() As String
    Dim ii As Integer
    
    aryTmp = Split(strval, "-")
    For ii = 0 To 2
        aryTmp(ii) = Trim(aryTmp(ii))
    Next
    AccTrim = Join(aryTmp, "-")
    
End Function

Private Sub cmdCul_Click()
    Dim objTestCd   As New clsDictionary
    Dim sPtid       As String
    Dim ii          As Integer
    
    Me.MousePointer = vbHourglass
    
    Set objCuM = New frmTmpCumulative

    objTestCd.Clear
    objTestCd.FieldInialize "testcd", "spccd"

    objTestCd.Sort = False
    
    For ii = 1 To ssRst.MaxRows
        ssRst.Row = ii
        ssRst.Col = 1
        With objPtInfo.Result.Item(ii)
            If chkCul.Value = 0 Then
                If objTestCd.Exists("testcd") = False Then
                    objTestCd.AddNew .TestCd, .SpcCd
                End If
            Else
                If ssRst.ForeColor = DCM_LightRed Then
                    If objTestCd.Exists("testcd") = False Then
                        objTestCd.AddNew .TestCd, .SpcCd
                    End If
                End If
            End If
            sPtid = objPtInfo.PtId
        End With
    Next ii
    objTestCd.Sort = True

    With objCuM
        .Top = Me.Top + 2000
        .Left = Me.Left + 200
        .MousePointer = vbDefault
        .Caption = "환자ID: " & sPtid & " 누적결과"
        Call .DisplayItem(objTestCd, sPtid)
        DoEvents
        
        .Show vbModal
        .WindowState = 0
        DoEvents
    End With

    Set objTestCd = Nothing
    
    Me.MousePointer = vbDefault
End Sub


Private Sub cmdRmk_Click()
    Dim objSQL   As clsLISResultReview
    Dim RS       As Recordset
    Dim aryTmp() As String
    Dim strTmp   As String
    Dim SSQL     As String
    Dim ii       As Integer
    
    txtMesg.Text = ""
    Set objSQL = New clsLISResultReview
    SSQL = objSQL.GetOrderRemark(LastWorkArea, LastAccDt, LastAccSeq)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        strTmp = "처방일자 : " & Format(RS.Fields("orddt").Value & "", "####-##-##") & vbCRLF
        strTmp = strTmp & "처방번호 : " & RS.Fields("ordno").Value & "" & vbCRLF
        strTmp = strTmp & "처방비고  " & vbCRLF
        txtMesg.Text = strTmp
        strTmp = ""
        aryTmp = Split(RS.Fields("mesg").Value & "", vbCRLF)
        For ii = LBound(aryTmp) To UBound(aryTmp)
            strTmp = strTmp & " " & aryTmp(ii) & vbCRLF
        Next
        txtMesg.Text = txtMesg.Text & strTmp
        fraMesg.Visible = True
        fraMesg.ZOrder 0
    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub

Private Sub cmdOk_Click()
    fraMesg.Visible = False
    ssRst.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Comment내에 "Total Volume:" 문자열 조회
'   반환 : 존재(True), 비존재(False)
'-----------------------------------------------------------------------------'
Private Function CheckComment() As Boolean
    Dim strTemp As String
    
    strTemp = rtfComment.Text
    If InStr(strTemp, "Total Volume:") > 0 Then
        CheckComment = True
    End If
End Function

Private Sub txtBatchRst_GotFocus()
    With txtBatchRst
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtBatchRst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdApply.SetFocus
    End If
End Sub

