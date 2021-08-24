VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Begin VB.MDIForm medMain 
   BackColor       =   &H00DEDBDD&
   Caption         =   "SCHWEITZER - LIS 1.0"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "medMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   Picture         =   "medMain.frx":0FEA
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '최대화
   Begin VB.PictureBox picMain 
      Align           =   1  '위 맞춤
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  '단색
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   15180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   15240
      Begin MedControls1.LisLabel lblLocation 
         Height          =   390
         Left            =   14025
         TabIndex        =   3
         Top             =   495
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   688
         BackColor       =   15724767
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
      Begin MSComctlLib.TabStrip tabSubMenu 
         Height          =   360
         Left            =   15
         TabIndex        =   4
         Top             =   600
         Width           =   13050
         _ExtentX        =   23019
         _ExtentY        =   635
         Style           =   2
         Placement       =   1
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   7
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "채혈접수"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "결과등록"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "혈액입출고"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "헌혈 및 Pheresis"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "증서관리"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ABO 검사"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "통계 및 마스터관리"
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
      Begin MSComctlLib.Toolbar tbrSubTool 
         Height          =   525
         Left            =   4185
         TabIndex        =   5
         Top             =   -45
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   926
         ButtonWidth     =   609
         ButtonHeight    =   926
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
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
         Height          =   285
         Left            =   15
         TabIndex        =   6
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape shpSubMenu 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   30
         Top             =   30
         Width           =   4095
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   390
         Left            =   14025
         Picture         =   "medMain.frx":38731
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   45
         Top             =   45
         Width           =   4095
      End
   End
   Begin VB.Timer tmrCheckLock 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4755
      Top             =   3030
   End
   Begin VB.PictureBox picComTool 
      Align           =   4  '오른쪽 맞춤
      Height          =   9465
      Left            =   14640
      ScaleHeight     =   9405
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   1065
      Width           =   600
      Begin MSComctlLib.Toolbar tbrComTool 
         Height          =   570
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "imlComTool"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   10
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  '아래 맞춤
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   10530
      Width           =   15240
      _ExtentX        =   26882
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
            Text            =   "(주)대련 엠티에스"
            TextSave        =   "(주)대련 엠티에스"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlComTool 
      Left            =   4110
      Top             =   3000
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
            Picture         =   "medMain.frx":3ACF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3B5CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3BEAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3C787
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3D063
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3D93F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3E21B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3EAF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3F3D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":3FCAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4058B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":40C87
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog diaComDialog 
      Left            =   6495
      Top             =   3015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save as "
      Filter          =   "Excel worksheet (*.xls)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico|Text (*.txt)|*.txt"
   End
   Begin MSMAPI.MAPISession MAPISess 
      Left            =   5880
      Top             =   2955
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMess 
      Left            =   5265
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   0
      Left            =   210
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":41563
            Key             =   "BBS101"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":429BD
            Key             =   "BBS102"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":43E19
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":45273
            Key             =   "BBS103"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":466CD
            Key             =   "BBS104"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":47B27
            Key             =   "BBS105"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":48F81
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4A3DB
            Key             =   "BBS106"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4B837
            Key             =   "BBS107"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4CC91
            Key             =   "BBS108"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4E0EB
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":4F547
            Key             =   "BBS205"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   2
      Left            =   2460
      Top             =   1080
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
            Picture         =   "medMain.frx":509A3
            Key             =   "BBS301"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":51DFD
            Key             =   "BBS302"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":53257
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":546B1
            Key             =   "BBS303"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":55B0B
            Key             =   "BBS304"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":56F65
            Key             =   "BBS305"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":583BF
            Key             =   "BBS306"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":59819
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5AC73
            Key             =   "BBS307"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5C0CD
            Key             =   "BBS308"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5D527
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5E981
            Key             =   "BBS309"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":5FDDB
            Key             =   "BBS311"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   3
      Left            =   3900
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":61237
            Key             =   "BBS401"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":62691
            Key             =   "BBS402"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":63AEB
            Key             =   "BBS403"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":64F45
            Key             =   "BBS404"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":6639F
            Key             =   "BBS412"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":677FB
            Key             =   "BBS405"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":68C57
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":6A0B1
            Key             =   "BBS406"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":6B50B
            Key             =   "BBS407"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":6C965
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":6DDBF
            Key             =   "BBS310"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   5
      Left            =   6360
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":6F223
            Key             =   "BBS501"
            Object.Tag             =   "일괄등록"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7067D
            Key             =   "BBS502"
            Object.Tag             =   "개별등록"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":71AD9
            Key             =   "BBS503"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":72F33
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7438F
            Key             =   "BBS504"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   1
      Left            =   1260
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":757EB
            Key             =   "BBS201"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":76C45
            Key             =   "BBS202"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7809F
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":794F9
            Key             =   "BBS203"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7A953
            Key             =   "BBS204"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7BDAD
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7D207
            Key             =   "BBS206"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7E661
            Key             =   "BBS207"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":7FABB
            Key             =   "BBS208"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   6
      Left            =   7740
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":80F15
            Key             =   "STATICS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8236F
            Key             =   "MASTER"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   4
      Left            =   5280
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":837C9
            Key             =   "BBS408"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":84C23
            Key             =   "BBS409"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMain.frx":8607F
            Key             =   "BBS410"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuLogon 
         Caption         =   "다른 이름으로 로그온(&L)"
      End
      Begin VB.Menu mnuPasswd 
         Caption         =   "비밀번호변경(&P)"
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersion 
         Caption         =   "&Version"
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
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScrLock 
         Caption         =   "Screen &Lock"
      End
      Begin VB.Menu mnudiv7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBarMaster 
         Caption         =   "바코드 출력양식 설정"
      End
      Begin VB.Menu mnuRegEdit 
         Caption         =   "Registry편집"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "새 프로그램 받기"
      End
      Begin VB.Menu mnuDiv8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormMaster 
         Caption         =   "폼관리"
      End
      Begin VB.Menu mnuEmpMaster 
         Caption         =   "직원등록"
      End
      Begin VB.Menu mnuGroupMaster 
         Caption         =   "그룹등록"
      End
      Begin VB.Menu mnuUserMaster 
         Caption         =   "사용자관리"
      End
      Begin VB.Menu mnuDiv9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWrite 
         Caption         =   "공지사항 쓰기(&W)"
      End
      Begin VB.Menu mnuInform 
         Caption         =   "공지사항 읽기(&R)"
      End
   End
   Begin VB.Menu mnuWins 
      Caption         =   "창(&W)"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(&H)"
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

'Private objS2DMM As clsS2DMM
Private WithEvents objS2DSM As clsS2DSM
Attribute objS2DSM.VB_VarHelpID = -1
Private objMyNote As New clsS2DCU
    

Private MailConfirm As Boolean
Private frmThis As Form
Private hIcons(4) As Long 'Hold Icon Images so we don't have to keep hitting the harddrive
Private LoadS2Code As Boolean

Private blnDownload As Boolean


Private Sub MDIForm_Activate()
    
    tbrComTool.Height = picComTool.Height
    imgLogo.Left = picMain.Width - imgLogo.Width - 70
    lblLocation.Left = picMain.Width - lblLocation.Width - 70
    
End Sub


' 프로그램 기동시 Check 사항 : Splash 창 띄우기, 중복실행Check, DB연결
' - Coding by 김미경
Private Sub MDIForm_Initialize()

    Call GetRegInfo
    
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
        objS2DSM.SetSplashMsg ("중복실행을 Check중입니다.")
        MsgBox App.ProductName & " 이 이미 실행중입니다. " & vbCrLf & _
              "<Ctrl><Alt><Delete> Key를 눌러 확인 후 다시 실행하십시오.", _
              vbOKOnly + vbExclamation, "Schweitzer-" & App.FileDescription
        End
    End If
    
    Call GetDatabase        ' DB연결 및 Server Configuration 설정
    Call CheckVersion       ' 최신버전 Download
    Call LoadBuildingInfo   ' 건물정보 로드
    
    

    Set StatusBar = Me.stsBar
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
        
        IsDBOpen = False
        If .ServerRegistered Then Call DbConnect

        If Not IsDBOpen Then
            .ButtonCheck = "SetDb"
            .LoadDatabaseConfig                     ' DB연결정보 등록 창 로드
            
            If .RegCanceled Then
                If .RunSplash = "1" Then objS2DSM.UnloadSplash
                Call AppExitRtn                     ' 취소했을 경우 Application 종료
            End If
            
            Call DbConnect
            If Not IsDBOpen Then
                MsgBox "Database 연결에 문제가 있습니다. 전산실로 연락바랍니다.", vbCritical + vbOKOnly, "Database 연결오류"
                ClearAllObject
                End
            End If
        End If
        
    End With

End Sub

' 서버에 등록된 최신버전과 현 Application의 버전을 비교하여 Upgrade 프로그램을 실행시킨다.
' ** Coding by 김미경
Private Sub CheckVersion()

    Dim strSql As String
    Dim tmpRs As New DrRecordSet
    Dim strFileServer As String
    Dim strCurVersion As String
    Dim strNewVersion As String
    Dim strAppPath As String
    Dim Resp As VbMsgBoxResult
    
    objS2DSM.SetSplashMsg ("버전을 체크하고 있습니다.")
    
    'Server IP 읽어오기
    strSql = ObjSysInfo.GetFileServer(BC2_File_Server, App.FileDescription)
    tmpRs.RsOpen DBConn, strSql
    If tmpRs.DBerror Then GoTo Err_Trap
    
    If tmpRs.EOF Then
        'strFileServer = "\\127.0.0.1\"
    Else
        strFileServer = Trim(tmpRs.Fields("SvrPath").value & "")
        strNewVersion = Trim(tmpRs.Fields("Version").value & "")
        Call ObjSysInfo.SetFileServer(strFileServer)       'Registry 등록
        tmpRs.RsClose
    End If
    
    blnDownload = True
    strCurVersion = App.Major & "." & App.Minor & "." & App.Revision
    If strNewVersion > strCurVersion Then     '버전비교
        Resp = MsgBox("SCHWEITZER - " & App.FileDescription & " 가 Upgrade되었습니다. " & _
                      "Download 하시겠습니까?", vbQuestion + vbYesNo, "메세지")
        If Resp = vbYes Then
            strAppPath = ObjSysInfo.AppPath
            If Mid(strAppPath, Len(strAppPath)) <> "\" Then strAppPath = strAppPath & "\"
            If Dir(strAppPath & "\GetNewVersion.EXE") <> "" Then       'GetNewVersion 실행
                Call Shell(strAppPath & "\GetNewVersion.EXE " & App.FileDescription)
            Else
                MsgBox "버전관리 프로그램이 설치되지 않았습니다. 전산실에 문의바랍니다.", _
                        vbExclamation + vbOKOnly, "파일누락"
                blnDownload = False
            End If
        Else
            blnDownload = False
        End If
    End If
    
Err_Trap:
    Set tmpRs = Nothing
    
End Sub

'건물정보 설정 창
'* Coding by 김미경
Private Sub LoadBuildingInfo()

    Dim strBldList As String
    
    With ObjSysInfo
        If .UseBuildingInfo = "1" Then      '건물정보를 사용하는 경우
            Set objS2DSM.MyDb = DBConn
            objS2DSM.SetSplashMsg ("건물정보를 로드하고 있습니다.")
            If .BuildingNo = 0 Or .BuildingCd = "" Then
                strBldList = objS2DSM.GetBuildingList(CD2_Buildings)
                .ButtonCheck = "Onlyreg"
                .BuildingList = strBldList
                .LoadBuildingInfo
            End If
        End If
    End With
    
End Sub

Private Sub MDIForm_Load()
    
    Dim ShowAtStartup As Integer
     
    objS2DSM.SetSplashMsg ("메인화면을 로드하고 있습니다.")
    
    Me.Caption = App.LegalTrademarks & " - " & App.FileDescription & " " & _
                 App.Major & "." & App.Minor & "." & App.Revision & " (" & DBConn.Server & ":" & DBConn.dbname & ")"
    
'    If objSysInfo.UseBuildingInfo = "1" Then
        lblLocation.Visible = True
        lblLocation.Caption = ObjSysInfo.BuildingNm         '위치
        lblLocation.ToolTipText = ObjSysInfo.BuildingCd
'    Else
'        lblLocation.Visible = False
'    End If
    App.HelpFile = App.Path & LoadResString(9)          'Help File 지정
    
    Me.Show
    DoEvents
    
    MailConfirm = False
    
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
        Set .MyDb = DBConn
        Call .LoadLogOn
        If Not .SuccessLogIn Then AppExitRtn     '로그온에 실패&취소 했을 경우 종료
    
    
        '// 코드 Dictionary 로드
        Call LoadMasterData
        
    End With
    
    '// Status Bar : 병원명, 사용자, 회사명 Display
    With stsBar
        .Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjMyUser.EmpLngNm
        .Panels(2).Text = "프로그램이 정상적으로 시작되었습니다."
        .Panels(3).Text = App.CompanyName
    End With
    
    tabSubMenu.Tabs(1).Selected = True
    Call tabSubMenu_Click
    
    
    DoEvents
    
    
    '사용등급에 따른 메뉴사용여부 설정-----------
    mnuFormMaster.Visible = ObjMyUser.IsDeveloper
    mnuEmpMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuGroupMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuUserMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
End Sub

Private Sub LoadMasterData()

    '// 코드 Dictionary 로드
    If LoadS2Code = False Then
        Set objBBSComCode = New clsHosComCode
        Call objBBSComCode.setDbConn(DBConn)
        objBBSComCode.ProjectCd = objS2DSM.ProjectId         '프로젝트코드 : APS, BBS, LIS
        objBBSComCode.LoadDept
        objBBSComCode.LoadDoct
        objBBSComCode.LoadEmp
        objBBSComCode.LoadWard
        objBBSComCode.LoadBarcodeInfo
        
        LoadS2Code = True
    
'        Set objInitBBSLibrary = New clsInitBBSLibrary
'        Set objInitBBSLibrary.Database = DbConn
'        Set objInitBBSLibrary.BBSComCode = objBBSComCode
        
'        Set objInitLPFactory = New clsInitLPFactory
'        Set objInitLPFactory.BBSComCode = objBBSComCode
        
'        Set objInitBBSAboTest = New clsInitBBSAboTest
'        Set objInitBBSAboTest.Database = DbConn

'        Set objInitBBSComCode = New clsInitBBSComCode
'        Set objInitBBSComCode.Database = DbConn
'        Set objInitBBSComCode.BBSComCode = objBBSComCode
    End If

End Sub

Private Sub ShowInformAtStart()
        
    '// 시작 시 공지사항 윈도우를 표시할 것인지를 확인한다.
    If ObjSysInfo.ShowAtStartup <> "0" Then
        'Set objMyNote = New clsS2DCU
        With objMyNote
            Set .MyDb = DBConn
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
    With objBBSComCode
        .BarInfo.ProjectCd = ObjSysInfo.ProjectId
        .BarInfo.SetBarConfig
    End With
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
'    picSubMenu.Top = 50
    tbrSubTool.Top = 0
    If tabSubMenu.Width = 4020 Then
        tbrSubTool.Top = 90
    Else
        tbrSubTool.Top = 0
    End If
End Sub

'[메뉴] - 계산기
Private Sub mnuCalcul_Click()
    If Dir(App.Path & "\CALC.EXE") = "" Then
        MsgBox "계산기 프로그램이 설치되지 않았습니다. " & vbCrLf & _
               "전산실로 연락 바랍니다. ", vbCritical + vbOKOnly, "파일누락"
    Else
        Call Shell(App.Path & "\CALC.EXE")
    End If
End Sub

'[메뉴] - 최신버전 받기
Private Sub mnuDownload_Click()
    Dim strAppPath  As String
    strAppPath = ObjSysInfo.AppPath
    If Mid(strAppPath, Len(strAppPath)) <> "\" Then strAppPath = strAppPath & "\"
    If Dir(strAppPath & "\GetNewVersion.EXE") <> "" Then       'GetNewVersion 실행
        Call Shell(strAppPath & "\GetNewVersion.EXE " & App.FileDescription)
    Else
        MsgBox "버전관리 프로그램이 설치되지 않았습니다. 전산실에 문의바랍니다.", _
                vbExclamation + vbOKOnly, "파일누락"
        blnDownload = False
    End If
End Sub

'[메뉴] - 프로그램 종료
Private Sub mnuExit_Click()
    Call AppExitRtn
End Sub

'[메뉴] - 도움말 색인
Private Sub mnuIndex_Click()
    
   With diaComDialog
      .HelpFile = App.HelpFile
      .HelpCommand = &H101&    'cdlHelpIndex
      .ShowHelp
   End With
   
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
        Set .MyDb = DBConn
        Call .LoadLogOn
        Set ObjMyUser = .MyUser
    End With

    '사용등급에 따른 메뉴사용여부 설정-----------
    mnuFormMaster.Visible = ObjMyUser.IsDeveloper
    mnuEmpMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuGroupMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
    mnuUserMaster.Visible = ObjMyUser.IsManager Or ObjMyUser.IsDeveloper
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
        Set .MyDb = DBConn
        .ProjectId = ObjSysInfo.ProjectId
        .TradeMark = App.LegalTrademarks
        .CanDelete = ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisr
        .FormShow (f_ReadNote)
    End With
End Sub

Private Sub mnuVersion_Click()
    
    MsgBox "제품명 : " & App.LegalTrademarks & " " & App.FileDescription & vbNewLine & "버전 : " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation + vbOKOnly, "버전정보"

End Sub

'[메뉴] - 공지사항 쓰기
Private Sub mnuWrite_Click()
    With objMyNote
        Set .MyDb = DBConn
        .EmpId = ObjMyUser.EmpId
        .ProjectId = ObjSysInfo.ProjectId
        .FormShow (f_WriteNote)
    End With
End Sub

'[메뉴] - Log On 화면
Private Sub mnuLogon_Click()
'    Set objS2DSM = New clsS2DSM
    Call ObjSysInfo.ReadRegistryInfo
    Call medUnloadForms(medMain.name)
    With objS2DSM
        .CancelIsEnd = False
        .ProductName = App.ProductName
        .ProjectId = App.FileDescription
        .lockfg = False
        Set .MyDb = DBConn
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
    If Not (ObjMyUser.IsDeveloper Or ObjMyUser.IsManager) Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(2)
End Sub

'[메뉴] - 그룹 관리
Private Sub mnuGroupMaster_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(3)
End Sub

'[메뉴] - 사용자관리
Private Sub mnuUserMaster_Click()
    If Not ObjMyUser.IsDeveloper Then
        MsgBox "사용권한이 없습니다.", vbExclamation + vbOKOnly, "Security Check"
        Exit Sub
    End If
    Call UseS2DSM(4)
End Sub

'툴바를 클릭했을 경우 해당 폼을 띄운다.
Private Sub FormShow(ByVal frmThis As Form)

    Dim i As Integer
    Dim tmpFrm As Object
    
On Error GoTo FormShow_error

    If ObjMyUser(frmThis.name) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(frmThis.name).CanRead Then GoTo PermissionDenied
    
    
    If frmThis.MDIChild = True Then
        frmThis.Show
        frmThis.ZOrder 0
    Else
        frmThis.Show vbModal, Me
    End If
    lblSubMenu.Caption = frmThis.Caption
   
    Exit Sub
    
   
PermissionDenied:
    Unload frmThis
    Set frmThis = Nothing
    
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check"
    Exit Sub
    
FormShow_error:
    MsgBox Err.Description, vbCritical, "오류"
End Sub

'[Event] - Logon 성공 !
Private Sub objS2DSM_LogonSuccess()
    
    Set ObjMyUser = objS2DSM.MyUser
    
    If ObjSysInfo.LogonId <> ObjMyUser.LoginId Then
        
        'Locking의 경우 최근 사용자와 현재 로긴한 사용자가 틀릴경우...
        If objS2DSM.lockfg Then
            Call medUnloadForms(Me.name)
        End If
        
        ObjSysInfo.LogonId = ObjMyUser.LoginId
        ObjSysInfo.EmpId = ObjMyUser.EmpId
        ObjSysInfo.EmpNm = ObjMyUser.EmpLngNm
        Me.stsBar.Panels(1).Text = ObjSysInfo.Hospital & "-" & ObjSysInfo.EmpNm
        
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
        Set .MyDb = DBConn
        Call .FormShow(intCase)
        
    End With

End Sub

Private Sub tabSubMenu_Click()
    'objS2DMM.ShowButtons
    
    Dim Count As Integer, i As Integer
    Dim intIdx As Integer
    Dim tag As String
    Dim btnX As Button
    

    ' Job Group 선택....Sub Toolbar의 내용이 바뀐다.
    intIdx = tabSubMenu.SelectedItem.Index
    lblSubMenu.Caption = medGetP(tabSubMenu.Tabs(intIdx).Caption, 1, "(")
    
    ' 올라있던 버튼을 삭제
    For i = tbrSubTool.Buttons.Count To 1 Step -1
        Call tbrSubTool.Buttons.Remove(i)
    Next i
        
    
    If imlSubList(intIdx - 1).ListImages.Count = 0 Then Exit Sub
    tbrSubTool.ImageList = imlSubList(intIdx - 1)
    
    Count = imlSubList(intIdx - 1).ListImages.Count
    
    ' 버튼을 다시 그린다.
    For i = 1 To Count   ' Step -1
        tag = imlSubList(intIdx - 1).ListImages(i).tag
        If tag <> "-" Then
            Call tbrSubTool.Buttons.Add(i, imlSubList(intIdx - 1).ListImages(i).Key, , , i)
            tbrSubTool.Buttons(i).ToolTipText = tag '  LoadResString(intIdx * 100 + i)
        Else
            Call tbrSubTool.Buttons.Add(i, , , tbrSeparator, i)
        End If
    Next i
End Sub

'나중에 User Control로 뺄 부분...
Private Sub tbrComTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    ' 공통 Toolbar의 기능
    Select Case Button.Index
        Case 1:
                With diaComDialog
                   .HelpFile = App.HelpFile
                   .HelpCommand = &HB Or &H5&  'HelpCNT Or cdlHelpSetContents
                   .ShowHelp
                End With
                
        Case 2:
                Call AppExitRtn
                
        Case 3:  '공지사항읽기 : 아무나...
                Call mnuInform_Click
                
        Case 4:
                '공지사항 입력 권한 : Supervisor 또는 Manager 그리구 Developer
                With ObjMyUser
                   If .IsManager Or .IsDeveloper Or .IsSupervisr Then
                       Call mnuWrite_Click
                   Else
                       MsgBox "이 메뉴를 사용하실 권한이 없습니다.. 전산실에 문의하십시요."
                       Exit Sub
                   End If
                End With
                
        Case 5:
                If Dir(App.Path & "\CALC.EXE") = "" Then
                    MsgBox "계산기 프로그램이 설치되지 않았습니다. " & vbCrLf & _
                           "전산실로 연락 바랍니다. ", vbCritical + vbOKOnly, "Message"
                Else
                    Call Shell(App.Path & "\CALC.EXE")
                End If
                
        Case 6:
                Call mnuScrLock_Click   'Screen Lock
                
        Case 7:
                Dim strAppPath  As String
                strAppPath = ObjSysInfo.AppPath
                If Mid(strAppPath, Len(strAppPath)) <> "\" Then strAppPath = strAppPath & "\"
                If Dir(strAppPath & "\GetNewVersion.EXE") <> "" Then       'GetNewVersion 실행
                    Call Shell(strAppPath & "\GetNewVersion.EXE " & App.FileDescription)
                Else
                    MsgBox "버전관리 프로그램이 설치되지 않았습니다. 전산실에 문의바랍니다.", _
                            vbExclamation + vbOKOnly, "파일누락"
                    blnDownload = False
                End If
    End Select

End Sub


'이 프로젝트에서 전역으로 사용한 모든 개체들을 소멸 시킨다.
Private Sub ClearAllObject()

    Set objS2DSM = Nothing
'    Set objS2DMM = Nothing
    Set ObjSysInfo = Nothing
    Set objMyNote = Nothing
    Set ObjMyUser = Nothing


'    Set objInitStatics = Nothing
'    Set objInitBBSLibrary = Nothing
'    Set objInitLPFactory = Nothing
'    Set objInitBBSAboTest = Nothing
'    Set objInitBBSComCode = Nothing
    
End Sub

'Application 종료시 확인메세지 후 처리...
'* Coding by 김미경
Public Function AppExitRtn(Optional ByVal blnTerminate As Boolean = False) As VbMsgBoxResult
    
    '강제종료
    If Not blnTerminate Then
    
        AppExitRtn = MsgBox(App.LegalTrademarks & "-" & App.FileDescription & " 를 종료하시겠습니까?", _
                            vbYesNo + vbQuestion, "프로그램 종료")
        If AppExitRtn = vbNo Then Exit Function
    
    End If
    
    'About 창 띄우기
    With ObjSysInfo
        .ProjectId = App.FileDescription
        .Version = App.Major & "." & App.Minor & "." & App.Revision
        .Copyright = App.LegalCopyright
        .LoadAbout True
    End With
    DoEvents
    
    Call DbClose
'    Set DbConn = Nothing
'    Call medSleep(3000)
    
    Call ClearAllObject
    
    End     '******  끝, The End  ******'

End Function

'*****************************************************
'되도록 윗 부분엔 코딩을 삼가해 주십시오.
'해부,혈액,임상 세 시스템의 공통부분입니다.
'*****************************************************


' ---------------------------------------------------------------------------------------
'
' 업무 Form을 Load시키는 부분
' 이곳에 추가하십시요.
'
' ---------------------------------------------------------------------------------------
Private Sub tbrSubTool_ButtonClick(ByVal Button As MSComctlLib.Button)

      
      Select Case Button.Key
         Case "BBS101": Call FormShow(frmBBS101)
         Case "BBS102": Call FormShow(frmBBS102)
         Case "BBS103": Call FormShow(frmBBS103)
         Case "BBS104": Call FormShow(frmBBS104)
         Case "BBS105": Call FormShow(frmBBS105)
         Case "BBS106": Call FormShow(frmBBS106)
         Case "BBS107": Call FormShow(frmBBS107)
         Case "BBS108": Call FormShow(frmBBS108)
         
         Case "BBS201": Call FormShow(frmBBS201)
         Case "BBS202": Call FormShow(frmBBS202)
         Case "BBS203": Call FormShow(frmBBS203)
         Case "BBS204": Call FormShow(frmBBS204)
         Case "BBS205": Call FormShow(frmBBS205)
         Case "BBS206": Call FormShow(frmBBS206)
         Case "BBS207": Call FormShow(frmBBS207)
         Case "BBS208": Call FormShow(frmBBS208)
         
         Case "BBS301": Call FormShow(frmBBS301)
         Case "BBS302": Call FormShow(frmBBS302)
         Case "BBS303": Call FormShow(frmBBS303)
         Case "BBS304": Call FormShow(frmBBS304)
         Case "BBS305": Call FormShow(frmBBS305)
         Case "BBS306": Call FormShow(frmBBS306)
         Case "BBS307": Call FormShow(frmBBS307)
         Case "BBS308": Call FormShow(frmBBS308)
         Case "BBS309": Call FormShow(frmBBS309)
         Case "BBS310": Call FormShow(frmBBS310)
         Case "BBS311": Call FormShow(frmBBS311)
         
         Case "BBS401": Call FormShow(frmBBS401)
         Case "BBS402": Call FormShow(frmBBS402)
         Case "BBS403": Call FormShow(frmBBS403)
         Case "BBS404": Call FormShow(frmBBS404)
         Case "BBS405": Call FormShow(frmBBS405)
         Case "BBS406": Call FormShow(frmBBS406)
         Case "BBS407": Call FormShow(frmBBS407)
         Case "BBS408": Call FormShow(frmBBS408)
         Case "BBS409": Call FormShow(frmBBS409)
         Case "BBS410": Call FormShow(frmBBS410)
         Case "BBS412": Call FormShow(frmBBS412)
         
         Case "BBS501": Call FormShow(frmBBS501)
         Case "BBS502": Call FormShow(frmBBS502)
         Case "BBS503": Call FormShow(frmBBS503)
         Case "BBS504": Call FormShow(frmBBS504)
         
         Case "STATICS": Call FormShow(frmStatics)
         Case "MASTER": Call FormShow(frmMaster)
      
      End Select
      
End Sub


'---------------------------------------------------------------------------------------------
'
' 이하 부분은 Custom Menu에 관계된 Function을 모아놓았읍니다.
' 수정하지 마십시요.
' 만일, 수정하게 되면 APS,LIS,BBS모두에 공통으로 적용되게 하여야하며,
' Form medMenuSetting과 연계되어야 합니다.
'
' 제작자 : 이우일
'
'---------------------------------------------------------------------------------------------
Public Function GetResString(ByVal id As Long) As String
    GetResString = LoadResString(id)
End Function











