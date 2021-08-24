VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C491A66B-3FD4-425B-A0A5-1773B78C83B0}#1.0#0"; "f_bsctrl.ocx"
Begin VB.MDIForm medMain 
   BackColor       =   &H00DEDBDD&
   Caption         =   "SCHWEITZER - LIS 1.0"
   ClientHeight    =   10650
   ClientLeft      =   1140
   ClientTop       =   2145
   ClientWidth     =   13260
   Icon            =   "medMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   NegotiateToolbars=   0   'False
   Picture         =   "medMain.frx":0FEA
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '최대화
   Begin VB.PictureBox picMain 
      Align           =   1  '위 맞춤
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  '단색
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   13200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   13260
      Begin VB.Frame Frame1 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   795
         Left            =   13890
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   1290
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '없음
            Caption         =   "Frame4"
            Height          =   435
            Index           =   0
            Left            =   795
            TabIndex        =   13
            Top             =   330
            Width           =   375
            Begin VB.Label Label3 
               BackStyle       =   0  '투명
               Caption         =   "R"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   165
               TabIndex        =   14
               Top             =   150
               Width           =   270
            End
            Begin VB.Shape Shape3 
               FillColor       =   &H000000FF&
               FillStyle       =   0  '단색
               Height          =   360
               Left            =   60
               Shape           =   3  '원형
               Top             =   60
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '없음
            Caption         =   "Frame4"
            Height          =   435
            Index           =   1
            Left            =   405
            TabIndex        =   11
            Top             =   330
            Width           =   375
            Begin VB.Label Label3 
               Alignment       =   2  '가운데 맞춤
               BackStyle       =   0  '투명
               Caption         =   "I"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   12
               Top             =   150
               Width           =   270
            End
            Begin VB.Shape Shape2 
               FillColor       =   &H000000FF&
               FillStyle       =   0  '단색
               Height          =   345
               Left            =   60
               Shape           =   3  '원형
               Top             =   60
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '없음
            Caption         =   "Frame4"
            Height          =   435
            Index           =   2
            Left            =   15
            TabIndex        =   9
            Top             =   330
            Width           =   375
            Begin VB.Label Label3 
               BackStyle       =   0  '투명
               Caption         =   "D"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   150
               TabIndex        =   10
               Top             =   150
               Width           =   270
            End
            Begin VB.Shape Shape1 
               FillColor       =   &H000000FF&
               FillStyle       =   0  '단색
               Height          =   345
               Left            =   45
               Shape           =   3  '원형
               Top             =   60
               Width           =   315
            End
         End
         Begin VB.Label lblPtid 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Caption         =   "ID:123456789"
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
            Left            =   15
            TabIndex        =   15
            Top             =   135
            Width           =   1260
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  '투명하지 않음
            Height          =   810
            Left            =   0
            Top             =   0
            Width           =   1290
         End
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   390
         Left            =   13890
         TabIndex        =   4
         Top             =   495
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   688
         BackColor       =   -2147483643
         ForeColor       =   5658923
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "위치"
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MSComctlLib.Toolbar tbrSubTool 
         Height          =   525
         Left            =   4185
         TabIndex        =   5
         Top             =   -15
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   926
         ButtonWidth     =   609
         ButtonHeight    =   926
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         Begin F_BSCTRLLib.xBSCtrl xBSCtrl1 
            Height          =   285
            Left            =   270
            TabIndex        =   16
            Top             =   3060
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   503
            _StockProps     =   0
         End
      End
      Begin MSComctlLib.TabStrip tabSubMenu 
         Height          =   360
         Left            =   30
         TabIndex        =   6
         Top             =   630
         Width           =   13050
         _ExtentX        =   23019
         _ExtentY        =   635
         Style           =   2
         Placement       =   1
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   8
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "채취/접수"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "결과등록"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "미생물/기타검사"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "조회/출력"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "QC"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Manager"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "통계"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "종합검증/판독"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   405
         Left            =   13890
         Picture         =   "medMain.frx":2854A
         Stretch         =   -1  'True
         Top             =   75
         Width           =   1290
      End
      Begin VB.Label lblSubMenu 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Laboratory Information System"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.Shape shpSubMenu 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   60
         Top             =   90
         Width           =   4065
      End
      Begin VB.Shape Shape10 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   525
         Left            =   45
         Top             =   75
         Width           =   4095
      End
   End
   Begin VB.PictureBox picComTool 
      Align           =   4  '오른쪽 맞춤
      Height          =   9285
      Left            =   12660
      ScaleHeight     =   9225
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   1065
      Width           =   600
      Begin MSComctlLib.Toolbar tbrComTool 
         Height          =   4560
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   8043
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "imlComTool"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_PTINFO"
               Object.ToolTipText     =   "미입력결과"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_EXIT"
               Object.ToolTipText     =   "종료"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_SCHEDULE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_READ"
               Object.ToolTipText     =   "공지사항 읽기"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_WRITE"
               Object.ToolTipText     =   "공지사항 쓰기"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_MAIL"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_CALCUL"
               Object.ToolTipText     =   "계산기"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_HELP"
               Object.ToolTipText     =   "도움말"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_TELNO"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_SCRLOCK"
               Object.ToolTipText     =   "화면 잠금"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C_DOWNLOAD"
               Object.ToolTipText     =   "새 버전 받기"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "C_DOWNLOAD1"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  '아래 맞춤
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   10350
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5997
            MinWidth        =   5997
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17110
            MinWidth        =   17110
            Text            =   "Message will be showed here."
            TextSave        =   "Message will be showed here."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4145
            MinWidth        =   4145
            Text            =   "POMIS"
            TextSave        =   "POMIS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlComTool 
      Left            =   10035
      Top             =   4950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2AB0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2B3E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2BCC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2C5A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2CE7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2D758
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2E034
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2E910
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2F1EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":2FAC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":303A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":30AA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog diaComDialog 
      Left            =   9360
      Top             =   5025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save as "
      Filter          =   "Excel worksheet (*.xls)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico|Text (*.txt)|*.txt"
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   0
      Left            =   405
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3137C
            Key             =   "LIS201"
            Object.Tag             =   "처방등록(처방)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":31D76
            Key             =   "LIS214"
            Object.Tag             =   "병동채혈(병동)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":32770
            Key             =   "LIS204"
            Object.Tag             =   "간호사채혈(간호사)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3316A
            Key             =   "LIS205"
            Object.Tag             =   "일반접수(접수)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":34F9C
            Key             =   "LIS206"
            Object.Tag             =   "외래접수(외래)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":35996
            Key             =   "LIS207"
            Object.Tag             =   "외부검사(외부)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":36390
            Key             =   "LIS208"
            Object.Tag             =   "바코드재발행(재발행)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":36D8A
            Key             =   "LIS209"
            Object.Tag             =   "접수대기자(대기)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":37784
            Key             =   "LIS210"
            Object.Tag             =   "접수취소(취소)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3817E
            Key             =   "LIS217"
            Object.Tag             =   "외래부서 일괄채혈(진료과)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":38B78
            Key             =   "LIS212"
            Object.Tag             =   "일괄재발행(일괄재)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":38E92
            Key             =   "LIS222"
            Object.Tag             =   "현장검사(현장검사)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3A2EC
            Key             =   "LIS223"
            Object.Tag             =   "특수부재발행(특수부)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   1
      Left            =   1350
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3A73E
            Key             =   "LIS301"
            Object.Tag             =   "업무나열서작성(WS작성)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3B138
            Key             =   "LIS302"
            Object.Tag             =   "접수번호별결과등록(LabNo)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3BB32
            Key             =   "LIS303"
            Object.Tag             =   "장비별결과등록(장비별)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3C52C
            Key             =   "LIS304"
            Object.Tag             =   "WorkSheet별결과등록(WS별)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3D57E
            Key             =   "LIS305"
            Object.Tag             =   "아이템별결과등록(Item별)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3DF78
            Key             =   "LIS309"
            Object.Tag             =   "항산성결과등록(항산성)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3E972
            Key             =   "LIS306"
            Object.Tag             =   "결과수정(수정)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3F36C
            Key             =   "LIS307"
            Object.Tag             =   "WBC Diff Count(Diff)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3FD66
            Key             =   "LIS308"
            Object.Tag             =   "결과 미입력리스트(미입력)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":40760
            Key             =   "LIS310"
            Object.Tag             =   "이미지관련결과등록(이미지)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4115A
            Key             =   "LIS311"
            Object.Tag             =   "WorkSheet 일괄결과등록(WS일괄)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":41B54
            Key             =   "LIS312"
            Object.Tag             =   "장비별 일괄결과등록(장비Bat)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4254E
            Key             =   "LIS313"
            Object.Tag             =   "판독소견결과등록(판독소견)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   2
      Left            =   2775
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":42868
            Key             =   "LIS401"
            Object.Tag             =   "미생물 업무나열서(WS작성)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":43262
            Key             =   "LIS402"
            Object.Tag             =   "Nogrowth(No.Gro)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":43C5C
            Key             =   "LIS411"
            Object.Tag             =   "바코드재발행(재발행)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":44656
            Key             =   "LIS410"
            Object.Tag             =   "Growth(Growth)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":45050
            Key             =   "LIS403"
            Object.Tag             =   "G.S결과등록(G.S등록)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":45A4A
            Key             =   "LIS404"
            Object.Tag             =   "G.S결과수정(G.S수정)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":46444
            Key             =   "LIS405"
            Object.Tag             =   "감수성결과등록(Cul 등록)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":46E3E
            Key             =   "LIS406"
            Object.Tag             =   "감수성결과수정(Cul 수정)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":47838
            Key             =   "LIS407"
            Object.Tag             =   "미생물 QC(Q . C)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":48232
            Key             =   "LIS408"
            Object.Tag             =   "특수검사 WorkSheet작성(S.WS)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":48C2C
            Key             =   "LIS409"
            Object.Tag             =   "특수검사 결과등록(S.결과)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":49626
            Key             =   "LIS412"
            Object.Tag             =   "항생제 내성율(내성율)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4AA80
            Key             =   "LIS413"
            Object.Tag             =   "환불수가내역(환불)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   3
      Left            =   4125
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4AD9A
            Key             =   "LIS501"
            Object.Tag             =   "처방별결과조회(처방별)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4B794
            Key             =   "LIS501N"
            Object.Tag             =   "전체결과조회(통합)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4C18E
            Key             =   "LIS502"
            Object.Tag             =   "Cumulative Result(누적)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4CB88
            Key             =   "LIS503"
            Object.Tag             =   "Preselected Item Review(항목)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4D582
            Key             =   "LIS504"
            Object.Tag             =   "전체결과조회(보고일)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4DF7C
            Key             =   "LIS505"
            Object.Tag             =   "결과보고대기(대기)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4E976
            Key             =   "LIS506"
            Object.Tag             =   "Report(Report)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4F370
            Key             =   "LIS507"
            Object.Tag             =   "접수조회(접수)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4FD6A
            Key             =   "LIS508"
            Object.Tag             =   "퇴원환자 결과조회(퇴원)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":50764
            Key             =   "LIS509"
            Object.Tag             =   "과거결과조회(과거)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5115E
            Key             =   "LIS510"
            Object.Tag             =   "환자별 결과조회(환자별)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":51B58
            Key             =   "LIS512"
            Object.Tag             =   "이미지관련 결과조회(이미지)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":52552
            Key             =   "LIS514"
            Object.Tag             =   "최근결과(최근결과)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   5
      Left            =   6060
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   6
      Left            =   6900
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":539AC
            Key             =   "LIS801"
            Object.Tag             =   "검사항목 통계(일별)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":543A6
            Key             =   "LIS802"
            Object.Tag             =   "TurnAroundTime(TAT)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":54DA0
            Key             =   "LIS803"
            Object.Tag             =   "균별 항생제 리스트(균종별)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5579A
            Key             =   "LIS804"
            Object.Tag             =   "이상결과리스트( 이상)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":56194
            Key             =   "LIS805"
            Object.Tag             =   "AnalysisList(소견)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":56B8E
            Key             =   "LIS806"
            Object.Tag             =   "감수성 추이( 추이)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":57588
            Key             =   "LIS807"
            Object.Tag             =   "WorkLoad(W . L)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":57F82
            Key             =   "LIS808"
            Object.Tag             =   "그룹별 검사항목 통계(그룹별)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5897C
            Key             =   "LIS809"
            Object.Tag             =   "병동별 Blood Culture 월 건수(B . C)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":59F2E
            Key             =   "LIS810"
            Object.Tag             =   "미생물 통계(미생물)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5A928
            Key             =   "LIS811"
            Object.Tag             =   "Case Study(CASE)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5B322
            Key             =   "LIS812"
            Object.Tag             =   "EMMA LIST(EMMA)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5D154
            Key             =   "LIS813"
            Object.Tag             =   "검사항목별 통계(항목별)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5DB4E
            Key             =   "LIS814"
            Object.Tag             =   "이미지관련통계(이미지)"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5E548
            Key             =   "LIS815"
            Object.Tag             =   "근무별 업무량 통계(업무량)"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5EF42
            Key             =   "LIS816"
            Object.Tag             =   "검사항목별 TAT(TAT)"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7B59E
            Key             =   "LIS817"
            Object.Tag             =   "월별 TAT 달성율(달성율)"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7B9F0
            Key             =   "LIS818"
            Object.Tag             =   "진료/진검결과 상이리스트(진료결과)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   10
      Left            =   5085
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7BB4A
            Key             =   "LIS6011"
            Object.Tag             =   "QC 컨트롤마스터"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7C726
            Key             =   "LIS601"
            Object.Tag             =   "QC 마스터"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7D302
            Key             =   "LIS610"
            Object.Tag             =   "QC 자동처방"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7DEDE
            Key             =   "LIS610N"
            Object.Tag             =   "QC 자동처방"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7EAB8
            Key             =   "LIS609"
            Object.Tag             =   "QC 처방등록"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7F694
            Key             =   "LIS611"
            Object.Tag             =   "내부정도관리 결과등록"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":80270
            Key             =   "LIS613"
            Object.Tag             =   "외부정도관리 결과등록"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":80E4C
            Key             =   "LIS614"
            Object.Tag             =   "미생물 QC 마스터"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":81A28
            Key             =   "LIS615"
            Object.Tag             =   "미생물 QC 결과등록"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8260C
            Key             =   "LIS616"
            Object.Tag             =   "혈액은행 QC 결과등록"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":831F0
            Key             =   "LIS602"
            Object.Tag             =   "QC 결과조회"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":83DCC
            Key             =   "LIS602N"
            Object.Tag             =   "QC 결과조회"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":849A6
            Key             =   "LIS630"
            Object.Tag             =   "통계"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":85580
            Key             =   "LIS605"
            Object.Tag             =   "온도관리"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8615C
            Key             =   "HIS601"
            Object.Tag             =   "장비이력"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":875B8
            Key             =   "LIS620"
            Object.Tag             =   "T-Test"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   7
      Left            =   7980
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   8
      Left            =   9165
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":88A12
            Key             =   "LIS901"
            Object.Tag             =   "Bypass & POCT(POCT)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8940C
            Key             =   "LIS902"
            Object.Tag             =   "추가처방(추가처방)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":89E06
            Key             =   "LIS903"
            Object.Tag             =   "미채혈 사유관리(아침채혈)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8A800
            Key             =   "LIS221"
            Object.Tag             =   "병동간호사 통합채혈(통합채혈)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8B1FA
            Key             =   "LIS220"
            Object.Tag             =   "산부인과 검체 채혈(산부인과)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8BBF4
            Key             =   "LIS906"
            Object.Tag             =   "병동수납처리(수납처리)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8C5EE
            Key             =   "LIS907"
            Object.Tag             =   "미실시검사내역(미실시)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8E420
            Key             =   "LIS908"
            Object.Tag             =   "아침채혈Schedule작성(Schedule)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8EE1A
            Key             =   "LIS909"
            Object.Tag             =   "연락처 작성 및 조회(연락처)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8F814
            Key             =   "LIS910"
            Object.Tag             =   "검사예약(검사예약)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   4
      Left            =   5085
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9020E
            Key             =   "QC01"
            Object.Tag             =   "QC 컨트롤마스터"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":91668
            Key             =   "QC02"
            Object.Tag             =   "QC 마스터"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":92AC2
            Key             =   "QC03"
            Object.Tag             =   "QC 일정"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":93F1C
            Key             =   "QC04"
            Object.Tag             =   "QC 자동처방"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":94AF6
            Key             =   "QC05"
            Object.Tag             =   "QC 처방"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":95F50
            Key             =   "QC06"
            Object.Tag             =   "내부정도관리"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":973AA
            Key             =   "QC07"
            Object.Tag             =   "QC 결과조회"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":98804
            Key             =   "QC08"
            Object.Tag             =   "QC 계산"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":99C5E
            Key             =   "QC09"
            Object.Tag             =   "T Test"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9B0B8
            Key             =   "QC10"
            Object.Tag             =   "Calibration"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9C512
            Key             =   "QC11"
            Object.Tag             =   "냉장고관리"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9D96C
            Key             =   "QC12"
            Object.Tag             =   "장비이력"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":9EDC6
            Key             =   "QC13"
            Object.Tag             =   "QC 결과조회(전체)"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuLogon 
         Caption         =   "다른 이름으로 로그온(&L)"
      End
      Begin VB.Menu mnuPasswd 
         Caption         =   "비밀번호 변경"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "&Version"
      End
      Begin VB.Menu mnuPrinterBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "프린터 설정(&P)"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료(&X)"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "추가기능(&O)"
      Begin VB.Menu mnuMenuSetting 
         Caption         =   "메뉴설정"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "&E-Mail 읽기"
      End
      Begin VB.Menu mnuDate 
         Caption         =   "날짜/시간 설정(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCalend 
         Caption         =   "달력(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCalcul 
         Caption         =   "계산기(&C)"
      End
      Begin VB.Menu mnuFrmSet 
         Caption         =   "화면 설정관리"
      End
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScrLock 
         Caption         =   "Screen &Lock"
      End
      Begin VB.Menu mnudiv7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "새 프로그램 받기"
      End
      Begin VB.Menu mnuTMP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegEdit 
         Caption         =   "Registry 등록"
      End
      Begin VB.Menu mnuFormMaster 
         Caption         =   "폼관리"
      End
      Begin VB.Menu mnuGroupMaster 
         Caption         =   "그룹관리"
      End
      Begin VB.Menu mnuUserMaster 
         Caption         =   "사용자관리"
      End
      Begin VB.Menu mnuEmpMaster 
         Caption         =   "직원정보관리"
      End
      Begin VB.Menu mnuDoctMaster 
         Caption         =   "의사정보관리"
      End
      Begin VB.Menu mnuBarMaster 
         Caption         =   "바코드출력양식설정"
      End
   End
   Begin VB.Menu mnuWins 
      Caption         =   "창(&W)"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "병원별 특수기능"
      Visible         =   0   'False
      Begin VB.Menu menu 
         Caption         =   "메뉴1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴9"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu menu 
         Caption         =   "메뉴10"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFbar 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuSet 
         Caption         =   "병원별 메뉴설정"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(&H)"
      Begin VB.Menu mnuWrite 
         Caption         =   "공지사항 쓰기(&W)"
      End
      Begin VB.Menu mnuInform 
         Caption         =   "공지사항 보기(&R)"
      End
      Begin VB.Menu mnuTopics 
         Caption         =   "도움말 목차(&C)"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "도움말 색인(&I)"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Schweitzer-LIS"
      End
   End
End
Attribute VB_Name = "medMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objS2DSM As clsS2DSM
Attribute objS2DSM.VB_VarHelpID = -1
Private objMyNote As New clsS2DCU
    
#Const UseLabCommentSystem = True

'Private frmThis As Form
Private LoadS2Code As Boolean
'Private MailConfirm As Boolean
Private blnDownload As Boolean

Private blnFormShow As Boolean

Private Sub MDIForm_Activate()
    tbrComTool.Height = picComTool.Height
    imgLogo.Left = picMain.Width - imgLogo.Width - 70
    lblLocation.Left = picMain.Width - lblLocation.Width - 70
End Sub

' 프로그램 기동시 Check 사항 : Splash 창 띄우기, 중복실행Check, DB연결
' - Coding by 김미경
Private Sub MDIForm_Initialize()

    Dim strTmp As String
    
    '컬럼명과 테이블명 활성화를 위한...
    strTmp = F_PTID
    strTmp = T_LAB001
    strTmp = P_HOSPITALNAME
    
    If InstallDir = "" Then
        Call SaveSetting("Schweitzer2000", "InstallDir", "InstallDir", App.Path & "\..\..\")
    End If
    
    Call GetRegInfo     'Registry 정보 읽어오기
    
    '// Splash 화면 로드...
    Set objS2DSM = New clsS2DSM
    If ObjSysInfo.RunSplash = "1" Then
        With objS2DSM
            .ProductName = App.ProductName
            .Version = App.Major & "." & App.Minor & "." & App.Revision
            .Copyright = App.LegalCopyright
            .LoadSplash
            .SetSplashMsg (.ProductName & " 프로그램을 기동 중입니다.")
        End With
    End If
    
    DoEvents
    
    '// 프로그램 중복실행 체크
    If App.PrevInstance = True Then
        If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
        MsgBox App.ProductName & " 이 이미 실행중입니다. " & vbCRLF & _
              "<Ctrl><Alt><Delete> Key를 눌러 확인 후 다시 실행하십시오.", _
              vbOKOnly + vbExclamation, "Schweitzer-" & App.FileDescription
        End
    End If
    
    Call GetDatabase        ' DB연결 및 Server Configuration 설정
    Call CheckVersion       ' 최신버전 Download
    Call LoadBuildingInfo   ' 건물정보 로드
    
    #If Not UseLabCommentSystem Then
        Call tabSubMenu.Tabs.Remove(8)
    #End If
    
    Set MainFrm = Me
End Sub

Private Sub GetRegInfo()

    '// Registry 정보 Update
    Set ObjSysInfo = New clsS2DSO
    With ObjSysInfo
        .ProjectId = App.FileDescription
        Call .SetAppName(App.LegalTrademarks & " " & App.FileDescription)   ' Registry등록시 Key값인 Application Name
        Call .CheckAppPath(App.Path & "\")      ' 현재 Application Path로 Update
        Call .ReadRegistryInfo                  ' Registry에 등록된 정보를 읽어온다
    End With
End Sub

' Registry에 DB정보가 등록되지 않았으면 Configuration창을 띄운다.
' DB연결은 3회까지 재시도 한후 정상적으로 연결되지 않으면 프로그램을 종료한다.
' - Coding by 김미경
Private Sub GetDatabase()

    With ObjSysInfo
    
        objS2DSM.SetSplashMsg ("DB를 연결중입니다.")
        
        If .ServerRegistered Then Call DBConnect

        If Not IsDBOpen Then
            .ButtonCheck = "SetDb"
            .LoadDatabaseConfig                     ' DB연결정보 등록 창 로드
            
            DCM_DbType = .DbTYPE
            If .RegCanceled Then
                If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
                Call AppExitRtn(True)               ' 취소했을 경우 Application 종료
            End If
            
            Call DBConnect
            
            If Not IsDBOpen Then
                If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
                MsgBox "Database 연결에 문제가 있습니다. 전산실 혹은 임상병리과로 연락바랍니다.(☎" & ObjSysInfo.HelpLine & ")", vbCritical + vbOKOnly, "Database 연결오류"
                ClearAllObject
                
                End
            End If
        End If
    End With
End Sub

' 서버에 등록된 최신버전과 현 Application의 버전을 비교하여 Upgrade 프로그램을 실행시킨다.
Private Sub CheckVersion(Optional ByVal blnChk As Boolean = True)
    Dim RS              As Recordset
    Dim SSQL            As String
    Dim strFileServer   As String
    Dim strCurVersion   As String
    Dim strNewVersion   As String
    Dim strGetNewExePath As String
    
    If blnChk Then objS2DSM.SetSplashMsg ("버전을 체크하고 있습니다.")
    
    If Dir(INIPath) = "" Then
        If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
        MsgBox INIPath & " 파일이 존재하지 않습니다." & vbCRLF & _
                        " 프로그램이 종료됩니다.", vbExclamation
        Call AppExitRtn(True)
    End If
    
    SSQL = " SELECT text1 as svrpath, field3 as version " & _
           " FROM  " & T_LAB032 & _
           " WHERE " & DBW("cdindex = ", LC3_FileServer) & _
           " AND   field1 = '1'"
    
    Set RS = New Recordset
    On Error GoTo Errors
    RS.Open SSQL, DBConn
'    If RS.DBerror Then GoTo Errors
    
    If RS.EOF Then
        Set RS = Nothing
        Exit Sub
    Else
        strFileServer = Trim(RS.Fields("SvrPath").Value & "")
        strNewVersion = Trim(RS.Fields("Version").Value & "")
    End If
    
    blnDownload = True
    '최근다운로드 날짜얻기
    strCurVersion = medGetINI("Version", "LastDate", INIPath)
    '다운로드 EXE실행 경로
'    strGetNewExePath = INIPath & "\..\GetNewVersion.EXE "
    strGetNewExePath = InstallDir & "GetNewVersion.exe"
    '다운로드 경로 설정
    Call medSetINI("DownLoad", "Path", strFileServer, INIPath)
'    Call SetInitINI("DownLoad", "Path", strFileServer)
    
    '버전비교
    If strNewVersion > strCurVersion Then
        If Dir(strGetNewExePath) <> "" Then
            Call Shell(strGetNewExePath, vbNormalFocus)
        Else
            If ObjSysInfo.RunSplash = "1" Then objS2DSM.UnloadSplash
            MsgBox "버전관리 프로그램이 설치되지 않았습니다. 전산실 혹은 임상병리과에 문의바랍니다.(☎" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation + vbOKOnly ', "파일누락"
            blnDownload = False
            If ObjSysInfo.RunSplash = "1" Then objS2DSM.LoadSplash
        End If
    Else
        If blnChk = False Then
            MsgBox "최신버전이 설치되어 있습니다.", vbInformation + vbOKOnly ', "버전정보"
            Exit Sub
        End If
    End If
    
Errors:
    Set RS = Nothing
End Sub

'건물정보 설정 창
'* Coding by 김미경
Private Sub LoadBuildingInfo()

    Dim strBldList As String
    
    With ObjSysInfo
        If .UseBuildingInfo = "1" Then      '건물정보를 사용하는 경우
'            Set objS2DSM.MyDb = dbconn
            objS2DSM.SetSplashMsg ("건물정보를 로드하고 있습니다.")
            If .BuildingNo = 0 Or .BuildingCd = "" Then
                strBldList = objS2DSM.GetBuildingList(LC3_Buildings)
                .ButtonCheck = "Onlyreg"
                .BuildingList = strBldList
                .LoadBuildingInfo
            End If
        Else
            .BuildingCd = "10"
            .BuildingNm = "본원"
            .BuildingNo = 1
        End If
    End With
End Sub

Private Sub MDIForm_Load()
    Dim ShowAtStartup As Integer
    Dim rst
    Dim strSQL        As String
    Dim RS            As Recordset
    
On Error Resume Next
    
    objS2DSM.SetSplashMsg ("메인화면을 로드하고 있습니다.")
    
    Me.Caption = App.LegalTrademarks & " - " & App.FileDescription & " " & _
                 App.Major & "." & App.Minor & "." & App.Revision & " (" & ObjSysInfo.DatabaseNm & ":" & ObjSysInfo.DBLoginId & ")"
    
'    If ObjSysInfo.UseBuildingInfo = "1" Then
        lblLocation.Visible = True
        lblLocation.Caption = ObjSysInfo.BuildingNm         '위치
'    Else
'        lblLocation.Visible = False
'    End If
'    App.HelpFile = App.Path & LoadResString(9)          'Help File 지정
        
    Me.Show
    DoEvents
    
    '
    '// Logon 화면 Display (S2DSM.dll)
    With objS2DSM
        '// Splash창은 Unload시킨다.
        If ObjSysInfo.RunSplash = "1" Then Call .UnloadSplash
        
        '// 로그인 화면 로드
'        Set objS2DSM = New clsS2DSM
        .CancelIsEnd = True
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
'        Set .MyDb = dbconn

'커맨드 라인에 사용자id, 사용자 pw가 넘어 온 경우에는 로긴화면을 표시하지 않고
'자체적으로 로긴 처리를 한다.

'################################################################
'2012-10-22 예수병원 보안 문제로 LIS자체 로그인을 못하고 막음
'OCS자체 로긴후 LIS 호출 시 사용자ID, 사용자PW를 확인 후 로긴 함
'################################################################

'' 컴파일소스 컴파일시 주석풀것 2013-11-30 PSK
'''==================================================================================================
'    If CmdLine = "" Then
'        Call MsgBox("모세에서 로긴 후 메뉴에서 임상병리를 사용하세요.!", vbExclamation, App.Title)
'        Call AppExitRtn(True)
'    Else
'        Call .LoadLogOn
'        If Not .SuccessLogIn Then Call AppExitRtn(True)     '로그온에 실패&취소 했을 경우 종료'
'                 strSQL = "SELECT * FROM CCCAPCKT                                          "
'        strSQL = strSQL + " WHERE EMPNO = '" & Trim(ObjSysInfo.EmpId) & "'         "
'        strSQL = strSQL + "   AND EXEID = 'SLIS'                                            "
'        strSQL = strSQL + "   AND TO_CHAR(SYSDATE, 'YYYYMMDD') BETWEEN STARTDTM AND ENDDTM "
'
'        Set RS = New Recordset
'        RS.Open strSQL, DBConn
'
'        If RS.EOF Then
'            rst = xBSCtrl1.SetBlockCapture(hwnd, 1)
'        Else
'            rst = xBSCtrl1.SetBlockCapture(hwnd, 0)
'        End If '
'
'        Set RS = Nothing
'    End If
''==================================================================================================
'''' 디버그 소스 컴파일시 주석처리 2013-11-30 PSK
'==================================================================================================
    If CmdLine = "" Then
        Call .LoadLogOn
        If Not .SuccessLogIn Then Call AppExitRtn(True)     '로그온에 실패&취소 했을 경우 종료
    Else
        .LoginId = Trim(medGetP(CmdLine, 1, ";"))
        .LoginPwd = Trim(medGetP(CmdLine, 2, ";"))
        Call .ProcessLogOn

        If Not .SuccessLogIn Then Call .LoadLogOn
        If Not .SuccessLogIn Then Call AppExitRtn(True)
    End If
'''==================================================================================================

        '// 코드 Dictionary 로드
'        Call LoadMasterData
    End With
    
    '// Status Bar : 병원명, 사용자, 회사명 Display
    With stsBar
        .Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjSysInfo.EmpNm
        .Panels(2).Text = "프로그램이 정상적으로 시작되었습니다."
        .Panels(3).Text = App.CompanyName
    End With
    
    tabSubMenu.Tabs(1).Selected = True

    
'    병원실정에 맞추어 작성한 프로그램 목록 로드
'    Call UseMenuSetting
    
    If Not ObjMyUser.IsDeveloper Then
        mnuFrmSet.Visible = False
    End If
    
    If ObjSysInfo.EmpId <> "9999" Then
        mnuFrmSet.Visible = False
        mnuMenuSetting.Visible = False
    End If
    
    DoEvents
End Sub

'==========================================================
'병원별 메뉴를 셋팅한다.
'병원마다 특수하게(어쩔수 없이) 사용하는
'메뉴에 대하여 별도로 관리 하기 위해서 추가 하였음
'S2Menu.dll
'이후병원부터 적용하자.

'Private Sub menu_Click(Index As Integer)
'    frmLisMenu.Show
'    frmLisMenu.ZOrder 0
'    Call frmLisMenu.ShowThisForm(CStr(Index))
'End Sub
'
'Private Sub menuSet_Click()
'    frmLisMenu.Show
'    frmLisMenu.ZOrder 0
'    Call frmLisMenu.ShowThisForm
'End Sub
'
'
'Private Sub UseMenuSetting()
'    Dim objMenu As New clsMenuSet
'
'    Set objMenu.SetForm = medMain
'    Call objMenu.MenuSetting
'
'    Set objMenu = Nothing
'End Sub
'==========================================================
'Private Sub LoadMasterData()
'
'    Dim objPrgBar As New clsprogress
'
'    '// 코드 Dictionary 로드
'    If LoadS2Code = False Then
'        Set objLisComCode = New clsHosComCode
''        ObjLISComCode.setDbConn dbconn
'        objLisComCode.SetForm medMain
'        objLisComCode.ProjectCd = objS2DSM.ProjectId         '프로젝트코드 : APS, BBS, LIS
'        If objLisComCode.LoadLISEntity = True Then
'            LoadS2Code = True
'        End If
'
'        Set objPrgBar.StatusBar = medMain.stsBar
'        objPrgBar.Max = 100
'        objPrgBar.Value = 80
'        DoEvents
'        objLisComCode.LoadBarcodeInfo
'        objPrgBar.Value = 90
'        DoEvents
'        objLisComCode.LoadLisItem
'        objPrgBar.Value = 100
'        DoEvents
'    End If
'
'    Set objPrgBar = Nothing
'
'End Sub

Private Sub ShowInformAtStart()
        
    '// 시작 시 공지사항 윈도우를 표시할 것인지를 확인한다.
    If ObjSysInfo.ShowAtStartup <> "0" Then
        'Set objMyNote = New clsS2DCU
        With objMyNote
'            Set .MyDb = dbconn
            .ProjectId = ObjSysInfo.ProjectId
            .TradeMark = App.LegalTrademarks
            .FormShow (f_TodayNote)
        End With
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   Dim Resp As VbMsgBoxResult
   Resp = AppExitRtn
   If Resp = vbNo Then Cancel = 1
End Sub

Private Sub mnuBarMaster_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If

'###########################################
'Con_Hos 가 프로젝트로 참조될때 오류나서 잠깐 막았음.
    Dim objBar As New clsBarcode

    With objBar
        Set .TableInfo = clsTables
        Set .FieldInfo = clsFields

        .SetBarConfig

    End With

    Set objBar = Nothing
End Sub

'[메뉴] - About 화면 로드...
Private Sub mnuAbout_Click()
    With ObjSysInfo
        .ProjectId = App.FileDescription
        .Version = App.Major & "." & App.Minor & "." & App.Revision
        .Copyright = App.LegalCopyright
        .LoadAbout
    End With
End Sub

'[메뉴] - 임시
Private Sub mnuBottom_Click()
    tabSubMenu.Top = 650
    tabSubMenu.Placement = tabPlacementBottom
    lblSubMenu.Top = 150
    shpSubMenu.Top = 50
    tbrSubTool.Top = 0
    If tabSubMenu.Width = 4020 Then
        tbrSubTool.Top = 90
    Else
        tbrSubTool.Top = 0
    End If
End Sub

'[메뉴] - 계산기
Private Sub mnuCalcul_Click()
    If Dir(GetSysDir & "CALC.EXE") = "" Then
        MsgBox "계산기 프로그램이 설치되지 않았습니다. " & vbCRLF & _
               "전산실 혹은 임상병리과로 연락 바랍니다. (☎" & ObjSysInfo.HelpLine & ")", vbCritical + vbOKOnly, "파일누락"
    Else
        Call Shell(GetSysDir & "CALC.EXE", vbNormalFocus)
    End If
End Sub

Private Sub mnuDoctMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(6)
End Sub

'[메뉴] - 최신버전 받기
Private Sub mnuDownload_Click()
    If MsgBox("새 버전을 받으시겠습니까?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    Call CheckVersion(False)
    
'    If Dir(InstallDir & "GetNewVersion.EXE ") <> "" Then      'GetNewVersion 실행
'        Call Shell(InstallDir & "GetNewVersion.EXE " & App.FileDescription, vbNormalFocus)
'    Else
'        MsgBox "버전관리 프로그램이 설치되지 않았습니다. 전산실 혹은 임상병리과에 문의바랍니다.(☎" & ObjSysInfo.HelpLine & ")", _
'                vbExclamation + vbOKOnly, "파일누락"
'        blnDownload = False
'    End If
End Sub

'[메뉴] - 프로그램 종료
Private Sub mnuExit_Click()
    Call AppExitRtn
End Sub

'[메뉴] - 화면 관리
Private Sub mnuFrmSet_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "사용권한이 없습니다.", vbExclamation
        Exit Sub
    End If
    frmSystem_manager.Show vbModal
End Sub

'[메뉴] - 도움말 색인
Private Sub mnuIndex_Click()
    
   With diaComDialog
      .HelpFile = App.HelpFile
      .HelpCommand = &H101&    'cdlHelpIndex
      .ShowHelp
   End With
   
End Sub

Private Sub mnuPrint_Click()
On Error GoTo PrintErr
    diaComDialog.ShowPrinter
PrintErr:
    Exit Sub
End Sub

'[메뉴] - Registry 정보 수정 : 개발자만 사용
Private Sub mnuRegEdit_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    ObjSysInfo.TradeMark = App.LegalTrademarks
    ObjSysInfo.LoadRegEdit
    ObjSysInfo.ReadRegistryInfo
End Sub

'[메뉴] - Screen Lock
Private Sub mnuScrLock_Click()

'    medScrLock.Show 1   'Screen Lock
    Call ObjSysInfo.ReadRegistryInfo
    With objS2DSM
        .CancelIsEnd = False
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
        .lockfg = True
        .OldLoginId = ObjMyUser.LoginId
        .OldLogInPass = ObjMyUser.LogInPass
'        Set .MyDb = dbconn
        Call .LoadLogOn
        Set ObjMyUser = .MyUser
    End With
    
End Sub

'[메뉴] - 도움말 목차
Private Sub mnuTopics_Click()
    
   With diaComDialog
      .HelpFile = App.HelpFile
      .HelpCommand = &HB Or &H5&  'HelpCNT Or cdlHelpSetContents
      .ShowHelp
   End With
   
End Sub

'[메뉴] - 공지사항 읽기
Private Sub mnuInform_Click()
    With objMyNote
'        Set .MyDb = dbconn
        .ProjectId = ObjSysInfo.ProjectId
        .TradeMark = App.LegalTrademarks
        .CanDelete = ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor
        .FormShow (f_ReadNote)
    End With
End Sub

Private Sub mnuVersion_Click()
    
    MsgBox "제품명 : " & App.LegalTrademarks & " " & App.FileDescription & vbNewLine & "버전 : " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation + vbOKOnly, "버전정보"

End Sub

'[메뉴] - 공지사항 쓰기
Private Sub mnuWrite_Click()
    With objMyNote
'        Set .MyDb = dbconn
        .EmpId = ObjMyUser.EmpId
        .ProjectId = ObjSysInfo.ProjectId
        .FormShow (f_WriteNote)
    End With
End Sub

'[메뉴] - Log On 화면
Private Sub mnuLogon_Click()
'    Set objS2DSM = New clsS2DSM
    Call ObjSysInfo.ReadRegistryInfo
    Call MyUnloadForms(medMain.Name)
    With objS2DSM
        .CancelIsEnd = False
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
        .lockfg = False
'        Set .MyDb = dbconn
        Call .LoadLogOn
    End With
End Sub

'[메뉴] - 비밀번호 변경 화면
Private Sub mnuPasswd_Click()
    Call UseS2DSM(5)
End Sub

'[메뉴] - 윈도우즈 리스트
Private Sub mnuWins_Click()
    
End Sub

'[메뉴] - 폼관리
Private Sub mnuFormMaster_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(1)
    
    
End Sub

'[메뉴] - 직원정보 등록 : 개발자,매니저만 사용
Private Sub mnuEmpMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(2)
End Sub

'[메뉴] - 그룹 관리
Private Sub mnuGroupMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(3)
End Sub

'[메뉴] - 사용자관리
Private Sub mnuUserMaster_Click()
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor) Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(4)
End Sub

'툴바를 클릭했을 경우 해당 폼을 띄운다.
Sub ShowForm(ByVal frmThis As Form, ByVal strFrmNm As String)

    Dim i As Integer
    
    If ObjMyUser(strFrmNm) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(strFrmNm).CanRead Then GoTo PermissionDenied

    Screen.MousePointer = vbHourglass
    If frmThis.MDIChild = True Then
        
        frmThis.Show
        frmThis.ZOrder 0
        
    Else
        frmThis.Show , Me
    End If
    lblSubMenu.Caption = frmThis.Caption
    Screen.MousePointer = vbDefault

    blnFormShow = True
    Exit Sub


PermissionDenied:
    Unload frmThis
    Set frmThis = Nothing
    
    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
'
End Sub

'[Event] - Logon 성공 !
Private Sub objS2DSM_LogonSuccess()
    
    Set ObjMyUser = objS2DSM.MyUser
    
    If ObjSysInfo.LogonId <> ObjMyUser.LoginId Then
        
        'Locking의 경우 최근 사용자와 현재 로긴한 사용자가 틀릴경우...
        If objS2DSM.lockfg Then
            Call MyUnloadForms(Me.Name)
        End If
        
        ObjSysInfo.LogonId = ObjMyUser.LoginId
        ObjSysInfo.EmpId = ObjMyUser.EmpId
        ObjSysInfo.EmpNm = ObjMyUser.EmpLngNm
        stsBar.Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjSysInfo.EmpNm
        
        Call ShowInformAtStart  '공지사항
        
    End If
    
End Sub

'[Event] - Logon 화면을 그냥 종료했을 경우...
Private Sub objS2DSM_QuitLogon()
    
    Dim Resp As VbMsgBoxResult
    
    If objS2DSM.CancelIsEnd Then Resp = AppExitRtn(True)
    
End Sub

'S2DSM Class를 사용하는 루틴
Private Sub UseS2DSM(ByVal intCase As Integer)
    
    If objS2DSM Is Nothing Then Set objS2DSM = New clsS2DSM
    
    With objS2DSM
        
        .ProjectId = App.FileDescription
'        Set .MyDb = dbconn
        Call .FormShow(intCase)
        
    End With

End Sub

'WardMenu Class를 사용하는 루틴
Private Sub UseS2WardMenu(ByVal intCase As Integer)
    
'    If objS2WardMenu Is Nothing Then Set objS2WardMenu = New clsShowForm1
'
'    With objS2WardMenu
'
'        Set .MyDb = DbConn
'        ' WhichForm : 1-처방및결과조회, 2-전체결과조회, 3-과거데이타 조회, 4-결과보고대기환자 리스트,
'        '             5-병동일괄채혈, 6-간호사채혈, 7-바코드재발행(local), 8-누적결과조회, 9-채혈리스트,
'        '             10-바코드재발행(일괄), 11-과거누적결과, 12-항목별 통계, 13-수혈상황조회
'        Call .ShowForm(intCase)
'
'    End With

End Sub

Private Sub tabSubMenu_Click()
    'objS2DMM.ShowButtons
    Dim intIDX As Integer
    
    intIDX = tabSubMenu.SelectedItem.Index
    lblSubMenu.Caption = medGetP(tabSubMenu.Tabs(intIDX).Caption, 1, "(")
    
    If intIDX = 6 Then
        'Manager Menu 권한
        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
            MsgBox "이 메뉴를 사용하실 권한이 없습니다.. 전산실 혹은 임상병리과에 문의하십시요.(☎" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation, "Security Check"
            Exit Sub
        End If
        frmLisMaster.Show: frmLisMaster.ZOrder 0
    End If
    
    If intIDX = 7 Then
        'Statistic Menu 권한
        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
            MsgBox "이 메뉴를 사용하실 권한이 없습니다.. 전산실 혹은 임상병리과에 문의하십시요.(☎" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation, "Security Check"
            Exit Sub
        End If
    End If
  
    #If UseLabCommentSystem Then
        If intIDX = 8 Then
            'Manager Menu 권한
            If Not (P_UseLabCommentSystem And (ObjMyUser.IsSupervisor Or ObjMyUser.IsDeveloper)) Then
                MsgBox "이 메뉴를 사용하실 권한이 없습니다.. 전산실 혹은 임상병리과에 문의하십시요.(☎" & ObjSysInfo.HelpLine & ")", _
                    vbExclamation, "Security Check"
                Exit Sub
            End If
            Set objMyCmt = New clsLabComments
            With objMyCmt
                Set .SysInfo = ObjSysInfo
'                Set .MyDb = DBConn
                .DoctId = ObjMyUser.EmpId
                .DoctNm = ObjMyUser.EmpLngNm
                .ShowForm
            End With
        End If
    #End If
    
    Call TabClickMenuSetting
    Exit Sub
'
'
'
'    Dim Count As Integer, i As Integer
'    Dim intIDX As Integer
'    Dim strTag As String
'    Dim btnX As Button
'
'
'    ' Job Group 선택....Sub Toolbar의 내용이 바뀐다.
'    intIDX = tabSubMenu.SELECTedItem.Index
'    lblSubMenu.Caption = medGetP(tabSubMenu.Tabs(intIDX).Caption, 1, "(")
'
'    If intIDX = 6 Then
'        'Manager Menu 권한
'        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
'            MsgBox "이 메뉴를 사용하실 권한이 없습니다.. 전산실 혹은 임상병리과에 문의하십시요.(☎" & ObjSysInfo.HelpLine & ")", _
'                    vbExclamation, "Security Check"
'            Exit Sub
'        End If
'        frmLisMaster.Show: frmLisMaster.ZOrder 0
'    End If
'
'    If intIDX = 7 Then
'        'Statistic Menu 권한
'        If Not (ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor) Then
'            MsgBox "이 메뉴를 사용하실 권한이 없습니다.. 전산실 혹은 임상병리과에 문의하십시요.(☎" & ObjSysInfo.HelpLine & ")", _
'                    vbExclamation, "Security Check"
'            Exit Sub
'        End If
'    End If
'
'    #If UseLabCommentSystem Then
'        If intIDX = 8 Then
'            'Manager Menu 권한
'            If Not (P_UseLabCommentSystem AND (ObjMyUser.IsSupervisor Or ObjMyUser.IsDeveloper)) Then
'                MsgBox "이 메뉴를 사용하실 권한이 없습니다.. 전산실 혹은 임상병리과에 문의하십시요.(☎" & ObjSysInfo.HelpLine & ")", _
'                    vbExclamation, "Security Check"
'                Exit Sub
'            End If
'            Set objMyCmt = New clsLabComments
'            With objMyCmt
'                Set .SysInfo = ObjSysInfo
'                .DoctId = ObjMyUser.EmpId
'                .DoctNm = ObjMyUser.EmpLngNm
'                .ShowForm
'            End With
'        End If
'    #End If
'
'    ' 올라있던 버튼을 삭제
'    For i = tbrSubTool.Buttons.Count To 1 Step -1
'        Call tbrSubTool.Buttons.Remove(i)
'    Next i
'
'
'    If imlSubList(intIDX - 1).ListImages.Count = 0 Then Exit Sub
'    tbrSubTool.ImageList = imlSubList(intIDX - 1)
'
'    Count = imlSubList(intIDX - 1).ListImages.Count
'
'    ' 버튼을 다시 그린다.
'    For i = 1 To Count   ' Step -1
'        strTag = imlSubList(intIDX - 1).ListImages(i).Tag
'        If Tag <> "-" Then
'            If intIDX = 7 Or intIDX = 4 Or intIDX = 1 Or intIDX = 2 Or intIDX = 3 Or intIDX = 9 Then
'                Call tbrSubTool.Buttons.Add(i, imlSubList(intIDX - 1).ListImages(i).Key, medGetP(medGetP(strTag, 2, "("), 1, ")"), , i)
'            Else
'                Call tbrSubTool.Buttons.Add(i, imlSubList(intIDX - 1).ListImages(i).Key, , , i)
'            End If
'            tbrSubTool.Buttons(i).ToolTipText = strTag
'            tbrSubTool.Buttons(i).Tag = strTag
'        Else
'            Call tbrSubTool.Buttons.Add(i, , , tbrSeparator, i)
'        End If
'    Next i
'컴팔할때 꼭 풀것~~~~~~~~~~~~~!!
'    Call SetInvisibleButton(intIDX)

End Sub

Private Sub TabClickMenuSetting()
   
    Dim i       As Integer
    Dim intIDX  As Integer
    Dim strTag  As String
    
    Dim objFrm      As clsDictionary
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim strTmp      As String
    Dim strKey      As String
    Dim aryTmp()    As String
    Dim kk          As Integer
    
    ' Job Group 선택....Sub Toolbar의 내용이 바뀐다.
    intIDX = tabSubMenu.SelectedItem.Index
    
    Set RS = New Recordset
    Set objFrm = New clsDictionary
    objFrm.Clear
    objFrm.FieldInialize "key", "ii"
    Call objFrm.DeleteAll
    
    SSQL = " SELECT * FROM " & T_LAB032 & _
           " WHERE " & _
                     DBW("cdindex=", LC3_HosFrmUsing) & _
           " AND " & DBW("cdval1=", intIDX)
           
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        strTmp = RS.Fields("text1").Value & ""
        aryTmp = Split(strTmp, ";")
        For kk = LBound(aryTmp()) To UBound(aryTmp())
            objFrm.AddNew aryTmp(kk), intIDX
        Next
    End If
'    RS.RsClose
    Set RS = Nothing
    
    ' 올라있던 버튼을 삭제
    For i = tbrSubTool.Buttons.Count To 1 Step -1
        Call tbrSubTool.Buttons.Remove(i)
    Next i
    
    If imlSubList(intIDX - 1).ListImages.Count = 0 Then
        Set objFrm = Nothing
        Exit Sub
    End If
    
    tbrSubTool.ImageList = imlSubList(intIDX - 1)
    kk = 0
    ' 버튼을 다시 그린다.
    For i = 1 To imlSubList(intIDX - 1).ListImages.Count
        strTag = imlSubList(intIDX - 1).ListImages(i).Tag
        If strTag <> "-" Then
            strKey = imlSubList(intIDX - 1).ListImages(i).Key
            If Not objFrm.Exists(strKey) Then
                kk = kk + 1
                If intIDX = 7 Or intIDX = 4 Or intIDX = 1 Or intIDX = 2 Or intIDX = 3 Or intIDX = 9 Then
                    Call tbrSubTool.Buttons.Add(kk, imlSubList(intIDX - 1).ListImages(i).Key, medGetP(medGetP(strTag, 2, "("), 1, ")"), , i)
                Else
                    Call tbrSubTool.Buttons.Add(kk, imlSubList(intIDX - 1).ListImages(i).Key, , , i)
                End If
                tbrSubTool.Buttons(kk).ToolTipText = strTag
                tbrSubTool.Buttons(kk).Tag = strTag
            End If
        Else
            Call tbrSubTool.Buttons.Add(i, , , tbrSeparator, i)
        End If
    Next i
    Set objFrm = Nothing
End Sub


Private Sub SetInvisibleButton(ByVal idx As Long)
    Select Case idx
        Case 1
        Case 2
            tbrSubTool.Buttons(6).Visible = False       '항산성
            tbrSubTool.Buttons(8).Visible = False       'Diff
        Case 3
        Case 4
            tbrSubTool.Buttons(2).Visible = False       '통합결과조회
            tbrSubTool.Buttons(4).Visible = False       '항목 조회
            tbrSubTool.Buttons(10).Visible = False      '과거 결과조회
        Case 5
            tbrSubTool.Buttons(4).Visible = False       'QC 자동처방 New
            tbrSubTool.Buttons(12).Visible = False
            tbrSubTool.Buttons(16).Visible = False
        Case 6
        Case 7
            tbrSubTool.Buttons(7).Visible = False       'WorkLodd
            tbrSubTool.Buttons(8).Visible = False       '그룹통계
            tbrSubTool.Buttons(9).Visible = False       'B/C
        Case 9
'                    tbrSubTool.Buttons(1).Visible = False       'Bypass & POCT
'                    tbrSubTool.Buttons(2).Visible = False       '추가처방
'                    tbrSubTool.Buttons(3).Visible = False       '미채혈사유관리
            tbrSubTool.Buttons(4).Visible = False       '통합채혈
            tbrSubTool.Buttons(5).Visible = False       '산부인과
            tbrSubTool.Buttons(6).Visible = False       'Acting
            tbrSubTool.Buttons(7).Visible = False       '미실시검사
    End Select

        
        '이미지 시스템
    Select Case idx
        Case "2":
            If P_ImageSystem = True Then
                tbrSubTool.Buttons(10).Visible = True
            Else
                tbrSubTool.Buttons(10).Visible = False
            End If
            If p_UseWSBatchRst = False Then tbrSubTool.Buttons(11).Visible = False
            If p_UseInstrBatchRst = False Then tbrSubTool.Buttons(12).Visible = False
        Case "4":
            If P_ImageSystem = True Then
                tbrSubTool.Buttons(12).Visible = True
            Else
                tbrSubTool.Buttons(12).Visible = False
            End If
        Case "7":
            If P_ImageSystem = True Then
                tbrSubTool.Buttons(14).Visible = True
            Else
                tbrSubTool.Buttons(14).Visible = False
            End If
    End Select
End Sub

'-------------------------------'
'   2002-08-06 수정자 : 이상대
'-------------------------------'
'나중에 User Control로 뺄 부분...
Private Sub tbrComTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    ' 공통 Toolbar의 기능
    Select Case Button.Key
        Case "C_HELP":
                frmSysHelp_manager.Left = 2200
                frmSysHelp_manager.Top = 1650
                frmSysHelp_manager.Show , MainFrm
                
                
                Exit Sub
                With diaComDialog
                   .HelpFile = App.HelpFile
                   .HelpCommand = &HB Or &H5&  'HelpCNT Or cdlHelpSetContents
                   .ShowHelp
                End With
                
        Case "C_EXIT":
                Call AppExitRtn
                
        Case "C_READ":  '공지사항읽기 : 아무나...
                Call mnuInform_Click
        
        Case "C_WRITE":
                '공지사항 입력 권한 : Supervisor 또는 Manager 그리구 Developer
                With ObjMyUser
                    If .IsManager Or .IsDeveloper Or .IsSupervisor Then
                        Call mnuWrite_Click
                    Else
                        Call mnuWrite_Click
                    End If
                End With
                
        Case "C_CALCUL":
                If Dir(GetSysDir & "CALC.EXE") = "" Then
                    MsgBox "계산기 프로그램이 설치되지 않았습니다. " & vbCRLF & _
                           "전산실 혹은 임상병리과로 연락 바랍니다. (☎" & ObjSysInfo.HelpLine & ")", vbCritical + vbOKOnly, "Message"
                Else
                    Call Shell(GetSysDir & "CALC.EXE", vbNormalFocus)
                End If
                
        Case "C_SCRLOCK":
                Call mnuScrLock_Click   'Screen Lock
                
        Case "C_DOWNLOAD":
            If MsgBox("새 버전을 받으시겠습니까?", vbExclamation + vbYesNo) = vbYes Then
                Call CheckVersion(False)
            End If
            
        Case "C_PTINFO":
            Call ShowForm(frm210UnverifiedList, "frm210UnverifiedList")
            
    End Select
End Sub

'이 프로젝트에서 전역으로 사용한 모든 개체들을 소멸 시킨다.
Private Sub ClearAllObject()

    Set objS2DSM = Nothing
    Set objMyNote = Nothing
'    Set objMyUser = Nothing
'    Set objSysInfo = Nothing
End Sub

'Application 종료시 확인메세지 후 처리...
'* Coding by 김미경
Public Function AppExitRtn(Optional ByVal blnTerminate As Boolean = False) As VbMsgBoxResult
    Dim Frm As Form
   
    '강제종료
    If Not blnTerminate Then
    
        AppExitRtn = MsgBox(App.LegalTrademarks & "-" & App.FileDescription & " 를 종료하시겠습니까?", _
                            vbYesNo + vbQuestion, "프로그램 종료")
        If AppExitRtn = vbNo Then Exit Function
    
    End If
    
    'medUnloadForms 함수를 사용하면, 에러가 발생합니다....
    '그래서, for 문으로 대체합니다.... wooil
'    medUnloadForms ("medMain")
    For Each Frm In Forms
        If Frm.Name <> Me.Name Then
            Unload Frm
        End If
    Next
    
    'About 창 띄우기
    With ObjSysInfo
        .ProjectId = App.FileDescription
        .Version = App.Major & "." & App.Minor & "." & App.Revision
        .Copyright = App.LegalCopyright
        .LoadAbout True
    End With
    DoEvents
    
    Call DbClose
    'Set DbConn = Nothing
    'Call medSleep(3000)
    
    Call ClearAllObject
    
    'Schweitzer.ini파일이 없는 경우 설정화면 로드할 수 있도록 유도?
    
    End     '******  끝, The End  ******'

End Function

'*****************************************************
'되도록 윗 부분엔 코딩을 삼가해 주십시오.
'해부,혈액,임상 세 시스템의 공통부분입니다.
'*****************************************************

Sub tbrSubTools(ByVal Button As MSComctlLib.Button)
    Button.Key = "LIS501"
    Call tbrSubTool_ButtonClick(Button)
End Sub
' ---------------------------------------------------------------------------------------
'
' 업무 Form을 Load시키는 부분
' 이곳에 추가하십시요.
'
' ---------------------------------------------------------------------------------------
Private Sub tbrSubTool_ButtonClick(ByVal Button As MSComctlLib.Button)
       
    '감염관리
    Call ICSPatientMark
    
    Select Case Button.Key
'채취/접수===========================================================================================================
        Case "LIS201":  Call ShowForm(frm101Order, "frm101Order")                           '처방등록
        Case "LIS214":  Call ShowCollectionForm(Button, "frm161WardCollect")                '병동채혈
        Case "LIS204": Call ShowCollectionForm(Button, "frm154NurCol")                     '간호사채혈
'                        Call ShowCollectionForm(Button, "frm160WardBarReprint")                     '간호사채혈
        
        
        Case "LIS205":  Call ShowForm(frm155Accession, "frm155Accession")                   '일반접수
        Case "LIS206":  Call ShowCollectionForm(Button, "frm165OutCol")                     '외래채혈
        Case "LIS207":  Call ShowForm(frm156Referral, "frm156Referral")                     '외부검사의뢰
        Case "LIS208":  Call ShowForm(frmLisReport, "frmLisReport")                         '바코드재발행
                        If blnFormShow Then Call frmLisReport.LoadReportForm("R001", "Barcode Label 재출력")
        Case "LIS209":  Call ShowForm(frm158AccPtList, "frm158AccPtList")                   '접수대기자병단
        Case "LIS210":  Call ShowForm(frm108AccCancel, "frm108AccCancel")                   '접수취소
        Case "LIS217":  Call ShowCollectionForm(Button, "frm164BatchCol")                   '외래부서채혈(종검):울산(종검,신검)
        Case "LIS212":  Call ShowCollectionForm(Button, "frm159BatchBarReprint")            '외래부서채혈(종검):울산(종검,신검)
        Case "LIS222":  Call ShowCollectionForm(Button, "frm168POCTCol")                    'POCT항목채혈접수
        Case "LIS223":  Call ShowCollectionForm(Button, "frm265BarPrint")                   '특수부재발행

 '===================================================================================================================
                        
'Q.C ================================================================================================================
        Case "LIS6011":     Call ShowQCForm(Button, "frm3011QCControlMaster")       'QC Control마스터
        Case "LIS601":      Call ShowQCForm(Button, "frm301QCMaster")               'QC 마스터
        Case "LIS602":      Call ShowQCForm(Button, "frm302QCReview")               '
        Case "LIS603":      Call ShowQCForm(Button, "frm303QCCalibration")          '
        Case "LIS604":      Call ShowQCForm(Button, "frm304QCEmployee")             '
        Case "LIS605":      Call ShowQCForm(Button, "frm305QCRefrigerator")         '냉장고온도관리
        Case "HIS601":      Call ShowQCForm(Button, "frm601MachHistory")            '장비이력관리
        Case "LIS608":      Call ShowQCForm(Button, "frm308QCPhlebotomist")         '
        Case "LIS609":      Call ShowQCForm(Button, "frm309QCOrder")                '
        Case "LIS610":      Call ShowQCForm(Button, "frm310QCReprint")              '
        Case "LIS610N":     Call ShowQCForm(Button, "frm310QCReprint_N")            'QC자동처방 NEW
        Case "LIS611":      Call ShowQCForm(Button, "frm311QCResultEntry")          '
        Case "LIS612":      Call ShowQCForm(Button, "frm312QCSchedule")             '
        Case "LIS613":      Call ShowQCForm(Button, "frm313QCOutResult")            '
        Case "LIS614":      Call ShowQCForm(Button, "frm314QCMicMaster")            '
        Case "LIS615":      Call ShowQCForm(Button, "frm315QCMicResult")            '
        Case "LIS616":      Call ShowQCForm(Button, "frm316QCBldResult")            '
        Case "LIS630":      Call ShowQCForm(Button, "frm330Calculation")            '
        Case "LIS602N":     Call ShowQCForm(Button, "frm302QCReview_N")             '
        Case "LIS620":      Call ShowQCForm(Button, "frm320Ttest")                  '
'===================================================================================================================
'===================================================================================================================
'QC NEW VERSION 테스트용
        Case "QC01":  Call ShowQCForm(Button, "frm3011QCControlMaster_N")
        Case "QC02":  Call ShowQCForm(Button, "frm301QCMaster_N")
        Case "QC03":  Call ShowQCForm(Button, "frm312QCSchedule_N")
        Case "QC04":  Call ShowQCForm(Button, "frm310QCReprint_N")   'QC자동처방 NEW
        Case "QC05":  Call ShowQCForm(Button, "frm309QCOrder_N")
        Case "QC06":  Call ShowQCForm(Button, "frm311QCResultEntry_N")
        Case "QC07":  Call ShowQCForm(Button, "frm302QCReview_N")   'QC자동처방 NEW
        Case "QC08": '  Call ShowQCForm(Button, "frm330Calculation_N")
        Case "QC09": '  Call ShowQCForm(Button, "frm320Ttest")
        Case "QC10": '  Call ShowQCForm(Button, "frm303QCCalibration_N")
        Case "QC11": '  Call ShowQCForm(Button, "frm305QCRefrigerator_N")
        Case "QC12": '  Call ShowQCForm(Button, "frm331EquipHistory")
        Case "QC13":  Call ShowQCForm(Button, "frm302QCReview_N_ALL")
'===================================================================================================================
'미생물=============================================================================================================
        Case "LIS401":  Call ShowForm(frm251MWS1, "frm251MWS1")                     '미생물 업무 나열서
        Case "LIS402":  Call ShowForm(frm252MBatch, "frm252MBatch")                 'NoGrowth
        Case "LIS403":  Call ShowForm(frm255MStain, "frm255MStain")                 'Stain결과등록
        Case "LIS404":  Call ShowForm(frm259MStainModify, "frm259MStainModify")     'Stain결과수정
        Case "LIS405":  Call ShowForm(frm256MCulture, "frm256MCulture")             '감수성결과등록
        Case "LIS406":  Call ShowForm(frm257MCultureModify, "frm257MCultureModify") '감수성결과수정
        Case "LIS407":  'Call ShowForm(frmMQC)                                      '미생물QC(현재사용않함)
        Case "LIS408":  Call ShowForm(frmLisReport, "frmLisReport")                 '특수검사 업무나열서
                        If blnFormShow Then                                         '
                            Call frmLisReport.LoadReportForm("R005", "기타검사 Worksheet 출력")
                        End If
        Case "LIS409":  Call ShowForm(frm293SpecialTest, "frm293SpecialTest")       '특수검사결과등록
        Case "LIS410":  Call ShowForm(frm253MReading, "frm253MReading")             'Growth
        Case "LIS411":  Call ShowForm(frm264MicBarPrint, "frm264MicBarPrint")       '미생물바코드재발행
'        Case "LIS412":  Call ShowForm(frm456SuscTrand, "frm456SuscTrAND")              '항생제감수성 추이
        Case "LIS413":  Call ShowForm(frmACList, "frmACList")       '환불검사취소내역
'===================================================================================================================
'결과등록===========================================================================================================
        Case "LIS301":  Call ShowForm(frm201WSBuild, "frm201WSBuild")               '업무나열서
        Case "LIS302":  Call ShowForm(frm202AccDataEntry, "frm202AccDataEntry")     '접수번호별
        Case "LIS303":  Call ShowForm(frm203InstDataEntry, "frm203InstDataEntry")   '장비별
        Case "LIS304":  Call ShowForm(frm204WSDataEntry, "frm204WSDataEntry")       '업무나열서별
        Case "LIS305":  Call ShowForm(frm205ItemDataEntry, "frm205ItemDataEntry")   '아이템별
        Case "LIS306":  Call ShowForm(frm206ModifyData, "frm206ModifyData")         '결과수정
        Case "LIS307":  Call ShowForm(frm207WBCDiffCnt, "frm207WBCDiffCnt")         'WBC Diff
        Case "LIS308":  Call ShowForm(frm210UnverifiedList, "frm210UnverifiedList") '미입력리스트
        Case "LIS309": '  Call ShowForm(frm270Tubercle, "frm270Tubercle")             '항산성결과
        Case "LIS310": '  Call ShowForm(frmSlideImage, "frmSlideImage")               '이미지 로드
        Case "LIS311": '  Call ShowForm(frm2301Result, "frm2301Result")               'WS 일괄등록
        Case "LIS312": ' Call ShowForm(frm2302EqpBatch, "frm2302EqpBatch")            '장비일괄등록
        Case "LIS313": Call ShowForm(frmResultReadList, "frmResultReadList")        '판독소견 리스트
'===================================================================================================================
        
'통계===============================================================================================================
        Case "LIS801":  Call ShowStaticForm(Button, "frm451_N")                     '일/월 별 검사건수 통계
        Case "LIS802":  Call ShowStaticForm(Button, "frm452TurnAroundTime")         'Turn Around Time
        Case "LIS803": '  Call ShowStaticForm(Button, "frm464infect")                 '검체군별 배양균 리스트
        Case "LIS804":  Call ShowStaticForm(Button, "frm454SAbnormal")              'Abnormal 리스트
        Case "LIS805":  Call ShowStaticForm(Button, "frm455AnalysisList")           '이상결과 리스트
        Case "LIS806":  Call ShowStaticForm(Button, "frm456SuscTrAND")              '항생제감수성 추이
        Case "LIS807": '  Call ShowStaticForm(Button, "frm453WorkLoad")               'WorkLoad통계
        Case "LIS808":  Call ShowStaticForm(Button, "frm460ItemCnt")                '그룹별 검사건수 통계
        Case "LIS809": '  Call ShowStaticForm(Button, "frm461BldCultureCnt")          '병동별 BloodCulture 통계
        Case "LIS810":  Call ShowStaticForm(Button, "frm459MAccCnt")                '미생물 통계
        Case "LIS811":  Call ShowStaticForm(Button, "frm462CaseStudy")              'Case Study
'추가 EMMALIST
'2011.01.17 온승호

        Case "LIS812": Call ShowStaticForm(Button, "frm463EMMALIST")              'EMMA LIST (TAG 근태관리(근태))
                       
'        Case "LIS812": '
'                       ' If Dir$(INIPath) = "" Then
'                       '     MsgBox "Schweitzer.ini 파일이 없습니다.", vbOKOnly + vbCritical, "Info"
'                       '     Exit Sub
'                       ' End If
'                       ' Call ShowStaticForm(Button, "frm463Statis")                 '근태관리
        Case "LIS813":  Call ShowStaticForm(Button, "frm451AccCnt")                 '검사건수통계
        Case "LIS814": '  Call ShowStaticForm(Button, "frm465ImageCnt")               '이미지통계
        Case "LIS815": '  Call ShowStaticForm(Button, "frm466WorkUnit")               '이미지통계
        Case "LIS816":  Call ShowStaticForm(Button, "frm467TestTAT")                '검사항목별 TAT
        Case "LIS817":  Call ShowStaticForm(Button, "frm500MonthTAT")                '검사항목별 TAT
        Case "LIS818":  Call ShowStaticForm(Button, "frm501AbList")                 '진료/진검 상이리스트
'===================================================================================================================
'조회 및 출력=======================================================================================================
        Case "LIS501":  Call ShowReviewForm(Button, "frm401ResultView")             '처방및 결과조회
        Case "LIS501N": ' Call ShowReviewForm(Button, "frm401ResultView_N")           '결과 조회 통합
        Case "LIS502":  Call ShowReviewForm(Button, "frm402Cumulative")             '누적결과조회
        Case "LIS503": '  Call ShowReviewForm(Button, "frm403SelReview")              '항목별결과조회
        Case "LIS504": '  Call ShowReviewForm(Button, "frm404AllResult")              '전체결과조회
        Case "LIS505": '  Call ShowForm(frmLisVerifyList, "frmLisVerifyList")         '결과보고대기내역
        Case "LIS506":
                        Call MyUnloadForms(frmLisReport.Name)
                        Call ShowForm(frmLisReport, "frmLisReport")                 '출력화면
                        If blnFormShow Then
                            frmLisReport.ZOrder 0
                        End If
        Case "LIS507":  Call ShowReviewForm(Button, "frm408AccResult")              '접수조회
        Case "LIS508": '  Call ShowReviewForm(Button, "frm409MedReport")              '입퇴원결과조회
        Case "LIS509": '  Call ShowReviewForm(Button, "frm410PastResult")             '과거결과조회
        Case "LIS510": '  Call ShowReviewForm(Button, "frm411CumResult_New")          '개별 누적결과
        Case "LIS512": '  Call ShowForm(frmSlideView, "frmSlideView")                 '이미지조회
        Case "LIS514":  Call ShowReviewForm(Button, "frm4NewResultView")                 '이미지조회
'===================================================================================================================
'기타 ==============================================================================================================
        Case "LIS901":  Call ShowForm(frm105Bypass, "frm105Bypass")                 'BYPASS & POCT
        Case "LIS902":  Call ShowForm(frm103AddOrder, "frm103AddOrder")             '추가처방
        Case "LIS903":  Call ShowForm(frm152WardAccession, "frm152WardAccession")   '미채혈사유관리
                        
        
        Case "LIS220":  Call ShowCollectionForm(Button, "frm167CollectionM")        '병동/간호사 통합채혈
        Case "LIS221":  Call ShowCollectionForm(Button, "frm166OgyCollect")         '산부인과 채혈
        Case "LIS906":  'Call ShowForm(frm223Sunap, "frm223Sunap")                   'ACTIONING
        Case "LIS907":  'Call ShowForm(frm222NonAct, "frm222NonAct")                 '병동수납처리
        Case "LIS908":  'Call ShowForm(medSchedule, "medSchedule")                   'Schedule작성
                        Call ShowForm(frm159RoundSchedule, "frm159RoundSchedule")   '아침채혈스케줄작성
        Case "LIS909":  Call ShowForm(medTelephone, "medTelephone")                 'Telephone Information
        Case "LIS910":  Call ShowForm(frmReserve, "frmReserve")                     '검사예약
'        Case "PNTCARE": Call ShowCollectionForm(Button, "frm106PntCare")
 '===================================================================================================================
     
    End Select
      
End Sub


Private Sub ShowCollectionForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    frmLisCollection.ButtonKey = Button.Key
    frmLisCollection.Show
    frmLisCollection.ZOrder 0
    frmLisCollection.ShowThisForm
    lblSubMenu.Caption = medGetP(Button.Tag, 1, "(")

    blnFormShow = True
    Exit Sub

PermissionDenied:

    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
'
End Sub


Private Sub ShowReviewForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = medGetP(Button.Tag, 1, "(")
    
    frmLisReview.ButtonKey = Button.Key
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    blnFormShow = True
    Exit Sub

PermissionDenied:

    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
'
End Sub

Private Sub ShowStaticForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer

'    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
'    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = medGetP(Button.Tag, 1, "(")
    
    frmLisStatistic.ButtonKey = Button.Key
    frmLisStatistic.Show
    frmLisStatistic.ZOrder 0
    frmLisStatistic.ShowThisForm

    blnFormShow = True
    Exit Sub

PermissionDenied:
    
    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
'
End Sub


Private Sub ShowQCForm(ByVal Button As MSComctlLib.Button, ByVal pFrmName As String)

    Dim i As Integer

    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    lblSubMenu.Caption = Button.Tag
    frmLisQC.ButtonKey = Button.Key
    frmLisQC.Show
    frmLisQC.ZOrder 0
    frmLisQC.ShowThisForm

    blnFormShow = True
    Exit Sub

PermissionDenied:
    
    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
'
End Sub

Private Sub MyUnloadForms(ByVal pName As String)
    Dim Frm As Form
    
    For Each Frm In Forms
        If Frm.Name <> pName And Frm.Name <> Me.Name Then
            Unload Frm
        End If
    Next
End Sub
