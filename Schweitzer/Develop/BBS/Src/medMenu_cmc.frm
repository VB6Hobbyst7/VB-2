VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form medMenu 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "메뉴구성"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.TabStrip tabSubMenu 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   0
      Left            =   300
      Top             =   360
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
            Picture         =   "medMenu_cmc.frx":0000
            Key             =   "BBS101"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":145A
            Key             =   "BBS102"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":28B6
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":3D10
            Key             =   "BBS103"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":516A
            Key             =   "BBS104"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":65C4
            Key             =   "BBS105"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":7A1E
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":8E78
            Key             =   "BBS106"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":A2D4
            Key             =   "BBS107"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":B72E
            Key             =   "BBS108"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":CB88
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":DFE4
            Key             =   "BBS205"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   2
      Left            =   2550
      Top             =   360
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
            Picture         =   "medMenu_cmc.frx":F440
            Key             =   "BBS301"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":1089A
            Key             =   "BBS302"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":11CF4
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":1314E
            Key             =   "BBS303"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":145A8
            Key             =   "BBS304"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":15A02
            Key             =   "BBS305"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":16E5C
            Key             =   "BBS306"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":182B6
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":19710
            Key             =   "BBS307"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":1AB6A
            Key             =   "BBS308"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":1BFC4
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":1D41E
            Key             =   "BBS309"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":1E878
            Key             =   "BBS311"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   3
      Left            =   3990
      Top             =   360
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
            Picture         =   "medMenu_cmc.frx":1FCD4
            Key             =   "BBS401"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":2112E
            Key             =   "BBS402"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":22588
            Key             =   "BBS403"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":239E2
            Key             =   "BBS411"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":24E3E
            Key             =   "BBS404"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":26298
            Key             =   "BBS405"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":276F4
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":28B4E
            Key             =   "BBS406"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":29FA8
            Key             =   "BBS407"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":2B402
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":2C85C
            Key             =   "BBS310"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   5
      Left            =   6450
      Top             =   360
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
            Picture         =   "medMenu_cmc.frx":2DCC0
            Key             =   "BBS501"
            Object.Tag             =   "일괄등록"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":2F11A
            Key             =   "BBS502"
            Object.Tag             =   "개별등록"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":30576
            Key             =   "BBS503"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":319D0
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":32E2C
            Key             =   "BBS504"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   1
      Left            =   1350
      Top             =   360
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
            Picture         =   "medMenu_cmc.frx":34288
            Key             =   "BBS201"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":356E2
            Key             =   "BBS202"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":36B3C
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":37F96
            Key             =   "BBS203"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":393F0
            Key             =   "BBS204"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":3A84A
            Key             =   ""
            Object.Tag             =   "-"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":3BCA4
            Key             =   "BBS206"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":3D0FE
            Key             =   "BBS207"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":3E558
            Key             =   "BBS208"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   6
      Left            =   7830
      Top             =   360
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
            Picture         =   "medMenu_cmc.frx":3F9B2
            Key             =   "STATICS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":40E0C
            Key             =   "MASTER"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   4
      Left            =   5370
      Top             =   360
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
            Picture         =   "medMenu_cmc.frx":42266
            Key             =   "BBS408"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":436C0
            Key             =   "BBS409"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "medMenu_cmc.frx":44B1C
            Key             =   "BBS410"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "medMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

