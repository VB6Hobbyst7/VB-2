VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmResult 
   Caption         =   "결과입력"
   ClientHeight    =   7470
   ClientLeft      =   1020
   ClientTop       =   2760
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   11910
   WindowState     =   2  '최대화
   Begin VB.TextBox txtMSeq 
      Height          =   330
      Left            =   4950
      TabIndex        =   62
      Top             =   1035
      Width           =   870
   End
   Begin Threed.SSCommand cmdAdditem 
      Height          =   375
      Left            =   5175
      TabIndex        =   61
      Top             =   495
      Visible         =   0   'False
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "additem"
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin Threed.SSPanel panelDate 
      Height          =   330
      Left            =   1260
      TabIndex        =   60
      Tag             =   "J"
      Top             =   1050
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "접수일자:"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Font3D          =   2
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   420
      Left            =   9270
      TabIndex        =   58
      Top             =   450
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   4630
      _ExtentY        =   741
      _StockProps     =   15
      Caption         =   " BarCode?"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
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
      Begin VB.TextBox txtBarCode 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   945
         TabIndex        =   59
         Text            =   "123456789012345"
         Top             =   45
         Width           =   1590
      End
   End
   Begin VB.TextBox txtRoom 
      BackColor       =   &H00C0E0FF&
      DataMember      =   "&H00C0E0FF&"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10215
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "txtRoom"
      Top             =   1035
      Width           =   645
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   6045
      Left            =   45
      TabIndex        =   14
      Top             =   1440
      Width           =   2220
      _Version        =   65536
      _ExtentX        =   3916
      _ExtentY        =   10663
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand cmdCallHelp 
         Height          =   330
         Left            =   1575
         TabIndex        =   30
         Top             =   225
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "&M"
      End
      Begin VB.TextBox txtMsample 
         Height          =   285
         Left            =   945
         TabIndex        =   29
         Text            =   "M2501"
         Top             =   225
         Width           =   600
      End
      Begin VB.ComboBox cmbWhere 
         Height          =   300
         ItemData        =   "frmResult.frx":0000
         Left            =   225
         List            =   "frmResult.frx":0010
         Style           =   2  '드롭다운 목록
         TabIndex        =   22
         Top             =   855
         Width           =   1905
      End
      Begin VB.ListBox lstMicroList 
         BackColor       =   &H00C0C0C0&
         Height          =   3300
         Left            =   90
         TabIndex        =   15
         Top             =   2610
         Width           =   2040
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   915
         Left            =   1395
         TabIndex        =   21
         Top             =   1530
         Width           =   645
         _Version        =   65536
         _ExtentX        =   1138
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "View"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "바탕체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "frmResult.frx":0032
      End
      Begin Threed.SSPanel panelWhere 
         Height          =   1275
         Left            =   225
         TabIndex        =   23
         Top             =   1170
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   2249
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelInner      =   1
         Begin VB.CheckBox chkWhere 
            Caption         =   "접수중"
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   28
            Tag             =   "R"
            Top             =   45
            Width           =   1050
         End
         Begin VB.CheckBox chkWhere 
            Caption         =   "부분결과"
            Height          =   225
            Index           =   1
            Left            =   45
            TabIndex        =   27
            Tag             =   "P"
            Top             =   270
            Width           =   1050
         End
         Begin VB.CheckBox chkWhere 
            Caption         =   "결과완료"
            Height          =   225
            Index           =   2
            Left            =   45
            TabIndex        =   26
            Tag             =   "C"
            Top             =   495
            Width           =   1050
         End
         Begin VB.CheckBox chkWhere 
            Caption         =   "외부의뢰"
            Height          =   225
            Index           =   3
            Left            =   45
            TabIndex        =   25
            Tag             =   "W"
            Top             =   720
            Width           =   1050
         End
         Begin VB.CheckBox chkWhere 
            Caption         =   "미확인"
            Height          =   225
            Index           =   4
            Left            =   45
            TabIndex        =   24
            Tag             =   "U"
            Top             =   945
            Width           =   1050
         End
      End
      Begin VB.Label Label3 
         Caption         =   "검체구분"
         Height          =   195
         Left            =   135
         TabIndex        =   31
         Top             =   270
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":0914
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":0C30
            Key             =   "Diff"
            Object.Tag             =   "Diff"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":0F54
            Key             =   "FuncS"
            Object.Tag             =   "FuncS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":13A8
            Key             =   "Clear"
            Object.Tag             =   "Clear"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":2B3C
            Key             =   "FuncG"
            Object.Tag             =   "FuncG"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":2F90
            Key             =   "SLip"
            Object.Tag             =   "SLip"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":4234
            Key             =   "Virus"
            Object.Tag             =   "Virus"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":4B10
            Key             =   "UrineCup"
            Object.Tag             =   "UrineCup"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":4E2C
            Key             =   "Urine"
            Object.Tag             =   "Urine"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":5148
            Key             =   "QryPt"
            Object.Tag             =   "QtyPt"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResult.frx":559C
            Key             =   "micro"
            Object.Tag             =   "micro"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      ButtonWidth     =   1984
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "결과입력종료"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "Clear"
            Object.ToolTipText     =   "화면정리"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "GramSt."
            Key             =   "Diff"
            Object.ToolTipText     =   "DiffCount"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "뇨검경"
            Key             =   "Urine"
            Object.ToolTipText     =   "뇨검경List"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "FuncS"
            Key             =   "FuncS"
            Object.ToolTipText     =   "Function Text Setting"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "FuncG"
            Key             =   "FuncG"
            Object.ToolTipText     =   "Funtion Text Get"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "상병조회"
            Key             =   "Virus"
            Object.ToolTipText     =   "상병조회"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "환자조회"
            Key             =   "QryPt"
            Object.ToolTipText     =   "환자별 조회"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "미생물"
            Key             =   "micro"
            Object.ToolTipText     =   "미생물검체접수"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Enrol"
                  Text            =   "검체접수"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sheet1"
                  Text            =   "Sheet"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Stool1"
                  Text            =   "Stool1"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Stool2"
                  Text            =   "Stool2"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "GS"
                  Text            =   "GramSt."
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtStatus 
      Height          =   330
      Left            =   10935
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1035
      Width           =   915
   End
   Begin Threed.SSCommand cmdHelp 
      Height          =   330
      Left            =   4635
      TabIndex        =   9
      Top             =   1035
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&H"
   End
   Begin VB.TextBox txtDr 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   9495
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "txtDr"
      Top             =   1035
      Width           =   735
   End
   Begin VB.TextBox txtDept 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "txtDept"
      Top             =   1035
      Width           =   1005
   End
   Begin VB.TextBox txtAge 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   8055
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "txtAge"
      Top             =   1035
      Width           =   375
   End
   Begin VB.TextBox txtSex 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   7695
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "txtSex"
      Top             =   1035
      Width           =   330
   End
   Begin VB.TextBox txtSname 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "txtSname"
      Top             =   1035
      Width           =   825
   End
   Begin VB.TextBox txtPtno 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   5850
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "txtPtno"
      Top             =   1035
      Width           =   960
   End
   Begin MSComCtl2.DTPicker dtJeobsu 
      Height          =   330
      Left            =   2340
      TabIndex        =   2
      Top             =   1035
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24772611
      CurrentDate     =   36501
   End
   Begin VB.TextBox txtSLipno2 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   3780
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1035
      Width           =   825
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   465
      Left            =   765
      TabIndex        =   1
      Top             =   450
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   820
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      MouseIcon       =   "frmResult.frx":5E78
      Begin VB.ComboBox cmbSLip 
         Height          =   300
         Left            =   1575
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   90
         Width           =   2670
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmResult.frx":711A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label9 
         Caption         =   "검사종목:"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   12
         Top             =   135
         Width           =   780
      End
   End
   Begin Threed.SSPanel panelMain 
      Height          =   6045
      Left            =   2295
      TabIndex        =   33
      Top             =   1440
      Width           =   9600
      _Version        =   65536
      _ExtentX        =   16933
      _ExtentY        =   10663
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin Threed.SSPanel panelGs 
         Height          =   4695
         Left            =   4095
         TabIndex        =   63
         Top             =   405
         Visible         =   0   'False
         Width           =   5145
         _Version        =   65536
         _ExtentX        =   9075
         _ExtentY        =   8281
         _StockProps     =   15
         Caption         =   " GramStain 결과입력Box"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         RoundedCorners  =   0   'False
         Alignment       =   0
         Begin VB.TextBox txtRange 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   3780
            TabIndex        =   69
            Text            =   "txtRange"
            Top             =   405
            Width           =   1095
         End
         Begin VB.TextBox txtResult 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   1890
            TabIndex        =   68
            Text            =   "txtResult"
            Top             =   405
            Width           =   1905
         End
         Begin VB.ListBox lstRange 
            Height          =   2940
            Left            =   3780
            TabIndex        =   67
            Top             =   810
            Width           =   1095
         End
         Begin VB.ListBox lstResult 
            Height          =   2940
            Left            =   1890
            TabIndex        =   66
            Top             =   810
            Width           =   1905
         End
         Begin VB.TextBox txtResult1 
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   135
            TabIndex        =   65
            Text            =   "txtResult1"
            Top             =   405
            Width           =   1770
         End
         Begin VB.ListBox lstitem 
            Height          =   2940
            Left            =   135
            TabIndex        =   64
            Top             =   810
            Width           =   1770
         End
         Begin MSForms.CommandButton cmdSelect 
            Height          =   420
            Left            =   1980
            TabIndex        =   71
            Top             =   4050
            Width           =   1365
            Caption         =   "선택"
            Size            =   "2408;741"
            FontName        =   "굴림체"
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdHideGs 
            Height          =   420
            Left            =   3420
            TabIndex        =   70
            Top             =   4050
            Width           =   1365
            Caption         =   "닫기"
            Size            =   "2408;741"
            FontName        =   "굴림체"
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
      End
      Begin FPSpreadADO.fpSpread sprSLip 
         Height          =   5205
         Left            =   45
         TabIndex        =   48
         Top             =   45
         Width           =   9510
         _Version        =   196608
         _ExtentX        =   16775
         _ExtentY        =   9181
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   18
         MaxRows         =   150
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "frmResult.frx":83AC
         UserResize      =   0
         VisibleCols     =   17
         VisibleRows     =   150
         Appearance      =   2
         TextTip         =   2
      End
      Begin VB.TextBox txtCount 
         BackColor       =   &H00C0E0FF&
         Height          =   295
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5685
         Width           =   950
      End
      Begin VB.TextBox txtGeomsaCm 
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000005&
         Height          =   735
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   34
         Top             =   5265
         Width           =   4695
      End
      Begin Threed.SSCommand cmdVerify 
         Height          =   735
         Left            =   8820
         TabIndex        =   36
         Top             =   5265
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Verify"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "frmResult.frx":933B
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   735
         Left            =   8100
         TabIndex        =   37
         Top             =   5265
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "선택All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "frmResult.frx":AAFD
      End
      Begin Threed.SSPanel panelText 
         Height          =   4470
         Left            =   4365
         TabIndex        =   38
         Top             =   495
         Visible         =   0   'False
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5900
         _ExtentY        =   7885
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
         BevelInner      =   2
         Begin VB.ListBox lstText 
            Appearance      =   0  '평면
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3150
            Left            =   180
            TabIndex        =   39
            Top             =   585
            Width           =   2985
         End
         Begin MSForms.CommandButton cmdHide 
            Height          =   420
            Left            =   180
            TabIndex        =   40
            Top             =   135
            Width           =   1455
            Caption         =   "Exit"
            PicturePosition =   327683
            Size            =   "2566;741"
            FontName        =   "굴림"
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
      End
      Begin Threed.SSPanel panelUrine 
         Height          =   4875
         Left            =   4185
         TabIndex        =   41
         Top             =   270
         Visible         =   0   'False
         Width           =   5235
         _Version        =   65536
         _ExtentX        =   9234
         _ExtentY        =   8599
         _StockProps     =   15
         ForeColor       =   -2147483634
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "바탕체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         Alignment       =   0
         Begin VB.CommandButton cmdCancel 
            Caption         =   "취소"
            Height          =   450
            Left            =   3750
            TabIndex        =   46
            Top             =   4260
            Width           =   1185
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "확인"
            Height          =   450
            Left            =   3750
            TabIndex        =   45
            Top             =   3810
            Width           =   1185
         End
         Begin VB.TextBox txtCastResult 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   135
            TabIndex        =   44
            Top             =   360
            Width           =   4800
         End
         Begin VB.ListBox lstGrade 
            BackColor       =   &H00FFC0C0&
            Height          =   2940
            ItemData        =   "frmResult.frx":B3D7
            Left            =   3705
            List            =   "frmResult.frx":B3D9
            TabIndex        =   43
            Top             =   705
            Width           =   1230
         End
         Begin VB.ListBox lstCast 
            BackColor       =   &H00C0FFFF&
            Height          =   4020
            ItemData        =   "frmResult.frx":B3DB
            Left            =   150
            List            =   "frmResult.frx":B3DD
            TabIndex        =   42
            Top             =   675
            Width           =   3540
         End
         Begin VB.Label Label2 
            BackColor       =   &H00004000&
            Caption         =   "뇨검경 결과"
            ForeColor       =   &H8000000E&
            Height          =   240
            Left            =   180
            TabIndex        =   47
            Top             =   90
            Width           =   1230
         End
      End
      Begin MSForms.CommandButton cmdRemark 
         Height          =   420
         Left            =   45
         TabIndex        =   49
         Top             =   5265
         Width           =   1185
         Caption         =   "Remark"
         PicturePosition =   327683
         Size            =   "2090;741"
         Picture         =   "frmResult.frx":B3DF
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin Threed.SSPanel panelBonemarrow 
      Height          =   6045
      Left            =   2295
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   9600
      _Version        =   65536
      _ExtentX        =   16933
      _ExtentY        =   10663
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox txtAspiration 
         BackColor       =   &H00C0E0FF&
         Height          =   2535
         Left            =   135
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   55
         Top             =   3105
         Width           =   5550
      End
      Begin VB.TextBox txtLength1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2745
         Width           =   1140
      End
      Begin VB.TextBox txtSmear 
         BackColor       =   &H00FFC0C0&
         Height          =   1950
         Left            =   180
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   51
         Top             =   540
         Width           =   5550
      End
      Begin VB.TextBox txtLength 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4590
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
      End
      Begin VB.Frame frDiffc 
         Caption         =   "DiffCount [500Count]"
         BeginProperty Font 
            Name            =   "바탕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5595
         Left            =   5805
         TabIndex        =   17
         Top             =   90
         Width           =   3660
         Begin FPSpreadADO.fpSpread sprDiffc 
            Height          =   4650
            Left            =   90
            TabIndex        =   18
            Top             =   270
            Width           =   3525
            _Version        =   196608
            _ExtentX        =   6218
            _ExtentY        =   8202
            _StockProps     =   64
            BackColorStyle  =   1
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   4
            MaxRows         =   18
            ScrollBars      =   0
            SpreadDesigner  =   "frmResult.frx":B6F9
            UserResize      =   1
            Appearance      =   1
         End
         Begin MSForms.CommandButton cmdPrint 
            Height          =   465
            Left            =   1755
            TabIndex        =   19
            Top             =   4950
            Width           =   1815
            Caption         =   "출력"
            PicturePosition =   327683
            Size            =   "3201;820"
            FontName        =   "굴림체"
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdBMVerify 
            Height          =   465
            Left            =   90
            TabIndex        =   20
            Top             =   4950
            Width           =   1680
            Caption         =   "결과입력"
            PicturePosition =   327683
            Size            =   "2963;820"
            FontName        =   "굴림체"
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Bone marrow aspiration"
         BeginProperty Font 
            Name            =   "바탕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   57
         Top             =   2835
         Width           =   2625
      End
      Begin MSForms.CommandButton cmdRmkAspi 
         Height          =   330
         Left            =   3420
         TabIndex        =   56
         Top             =   2745
         Width           =   1095
         Caption         =   "GetText"
         Size            =   "1931;582"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label4 
         Caption         =   "Peripheral blood smear"
         BeginProperty Font 
            Name            =   "바탕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   53
         Top             =   270
         Width           =   2535
      End
      Begin MSForms.CommandButton cmdRmkPbsmear 
         Height          =   330
         Left            =   3420
         TabIndex        =   52
         Top             =   180
         Width           =   1140
         Caption         =   "GetText"
         PicturePosition =   327683
         Size            =   "2011;582"
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim UrCode(1 To 30)                 As String * 6
Dim UrResult(1 To 30, 1 To 30)      As String * 30

Dim UrCastGrade(0 To 9)             As String * 10
Dim UrCellGrade(0 To 9)             As String * 10
Dim UrCastGubun(0 To 28)            As String * 1
Dim UrCastResult(0 To 28)           As String * 30


Public Sub Init_Result_Data()

    UrCastGrade(0) = " 1 ~  2 "                  ' cast 등급
    UrCastGrade(1) = " 3 ~  5 "
    UrCastGrade(2) = " 6 ~ 10 "
    UrCastGrade(3) = "11 ~ 15 "
    UrCastGrade(4) = "16 ~ 20 "
    UrCastGrade(5) = "   > 20 "
    
    
    UrCellGrade(0) = "A few   "                   ' crystall(그외) 등급
    UrCellGrade(1) = "Some    "
    UrCellGrade(2) = "Many    "
    UrCellGrade(3) = "Positive"
    UrCellGrade(4) = "Rare    "

    '--------------------------------------------
    '  CAST:0   CRYSTAL(cell):1
    '--------------------------------------------
    UrCastGubun(0) = "1"
    UrCastGubun(1) = "1"
    UrCastGubun(2) = "1"
    UrCastGubun(3) = "1"
    UrCastGubun(4) = "1"
    UrCastGubun(5) = "1"
    UrCastGubun(6) = "1"
    UrCastGubun(7) = "1"
    UrCastGubun(8) = "1"
    UrCastGubun(9) = "1"
    UrCastGubun(10) = "1"
    UrCastGubun(11) = "0"
    UrCastGubun(12) = "0"
    UrCastGubun(13) = "0"
    UrCastGubun(14) = "1"
    UrCastGubun(15) = "0"
    UrCastGubun(16) = "1"
    UrCastGubun(17) = "1"
    UrCastGubun(18) = "0"
    UrCastGubun(19) = "1"
    UrCastGubun(20) = "1"
    UrCastGubun(21) = "1"
    UrCastGubun(22) = "1"
    UrCastGubun(23) = "1"
    UrCastGubun(24) = "1"
    UrCastGubun(25) = "1"
    UrCastGubun(26) = "0"
    UrCastGubun(27) = "1"

    
    
    UrCastResult(0) = "Ammonium urate crystal"         '검경리스트"
    UrCastResult(1) = "Amorphous phosphate crystal"
    UrCastResult(2) = "Amorphous urate crystal"
    UrCastResult(3) = "Bacteria""              "
    UrCastResult(4) = "Bilirubin crystal"
    UrCastResult(5) = "Calcium carbonate crystal"
    UrCastResult(6) = "Calcium oxalate crystal"
    UrCastResult(7) = "Calcium phosphate crystal"
    UrCastResult(8) = "Calcium urate crystal"
    UrCastResult(9) = "Cholesterol crystal"
    UrCastResult(10) = "Cystine crystal"
    UrCastResult(11) = "Epithelial cell "
    UrCastResult(12) = "Fat granule"
    UrCastResult(13) = "Granular cast"
    UrCastResult(14) = "Hippuric acid"
    UrCastResult(15) = "Hyaline cast"
    UrCastResult(16) = "Leucine crystal"
    UrCastResult(17) = "Mucous thread"
    UrCastResult(18) = "RBC cast"
    UrCastResult(19) = "Sodium urate crystal"
    UrCastResult(20) = "Spermatozoa"
    UrCastResult(21) = "Sulfa crystal"
    UrCastResult(22) = "Trichomonas vaginalis"
    UrCastResult(23) = "Triple phosphate crystal"
    UrCastResult(24) = "Tyrosine crystal"
    UrCastResult(25) = "Uric acid"
    UrCastResult(26) = "WBC cast"
    UrCastResult(27) = "Yeast-like organism"
    
    
    lstCast.Clear
    For i = 0 To 27
        lstCast.AddItem UrCastResult(i)
    Next i
    
End Sub


Private Sub cmbSLip_Change()
    
    Select Case Left(cmbSLip.Text, 2)
        Case 42:   panelBonemarrow.Visible = False
        Case 15:   GoSub BonMarrow_Set
        Case Else: panelBonemarrow.Visible = False
    End Select
    
    
    If Left(cmbSLip.Text, 1) = "4" Then
        Toolbar1.Buttons(16).Visible = True
        txtMSeq.Visible = True
    Else
        Toolbar1.Buttons(16).Visible = False
        txtMSeq.Visible = False
    End If
    
    
    Exit Sub
    

BonMarrow_Set:
    panelBonemarrow.Top = 1440
    panelBonemarrow.Left = 2295
    panelBonemarrow.Height = 6045
    panelBonemarrow.Width = 9600
    panelBonemarrow.Visible = True
    panelBonemarrow.ZOrder 0

    Return
    
    
End Sub

Private Sub cmbSLip_Click()
    
    GoSub SCREEN_Clear_Sub
    
    
    Select Case Left(cmbSLip.Text, 2)
        Case 42:   panelBonemarrow.Visible = False
        Case 15:   GoSub BonMarrow_Set
        Case Else: panelBonemarrow.Visible = False
    End Select
    
    If Left(cmbSLip.Text, 1) = "4" Then
        Toolbar1.Buttons(16).Visible = True
        txtMSeq.Visible = True
    Else
        Toolbar1.Buttons(16).Visible = False
        txtMSeq.Visible = False
    End If
    
    
    Exit Sub
    

BonMarrow_Set:
    panelBonemarrow.Top = 1440
    panelBonemarrow.Left = 2340
    panelBonemarrow.Height = 5955
    panelBonemarrow.Width = 9555
    panelBonemarrow.Visible = True
    panelBonemarrow.ZOrder 0

    Return
    
SCREEN_Clear_Sub:
    mdiMain.stbMain.Panels(1).Text = ""
    cmdAll.Caption = "선택All"
    
    GoSub Spread_Clear_SLip

    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
    
    txtMsample.Text = ""
    
    cmbWhere.ListIndex = 3
    chkWhere(0).Value = "1"
    chkWhere(1).Value = "0"
    chkWhere(2).Value = "0"
    chkWhere(3).Value = "0"
    chkWhere(4).Value = "0"
    
    lstMicroList.Clear
    'dtJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Call SpreadSetClear(sprDiffc)
    
    If Left(cmbSLip.Text, 2) = "15" Then
        panelBonemarrow.Visible = True
        panelBonemarrow.ZOrder 0
    Else
        'Call SetComboBox(cmbSLip, GiExamNumb, 2)
        'txtSLipno2.SetFocus
    End If
    
    
    Return
    
Spread_Clear_SLip:
    Call SSInitialize(sprSLip)

    sprSLip.Row = 1: sprSLip.Row2 = sprSLip.MaxRows
    sprSLip.Col = 1: sprSLip.Col2 = sprSLip.MaxCols
    sprSLip.BlockMode = True
    sprSLip.BackColor = RGB(235, 245, 235)
    sprSLip.BlockMode = False
    
    sprSLip.Row = 1: sprSLip.Row2 = sprSLip.MaxRows
    sprSLip.Col = 1: sprSLip.Col2 = 2
    sprSLip.BlockMode = True
    sprSLip.CellType = CellTypeStaticText
    sprSLip.BlockMode = False
    
    Return

End Sub

Private Sub cmbWhere_Click()

    
    Select Case cmbWhere.ListIndex
        Case 0: '응급
                 panelWhere.Visible = True
                 chkWhere(0).Value = "1"
        Case 1: 'ABNormal
                 panelWhere.Visible = False
        Case 2: '전체
                 panelWhere.Visible = False
        Case 3: '조건별
                 panelWhere.Visible = True
                 chkWhere(0).Value = "1"
    End Select
    
    
    
End Sub

Private Sub cmdAdditem_Click()
    
    If Trim(txtSLipno2.Text) <> "" Then
        frmitemadd.Show vbModal
        Call txtSLipno2_KeyDown(vbKeyReturn, 1)
    End If
    
End Sub

Private Sub cmdAll_Click()
        
    Dim i           As Integer
    
    If cmdAll.Caption = "선택All" Then
        cmdAll.Caption = "해제All"
        cmdAll.ForeColor = RGB(255, 0, 0)
        For i = 1 To sprSLip.DataRowCnt
            sprSLip.Row = i
            sprSLip.Col = 7
            sprSLip.Text = "1"
        Next i
    Else
        cmdAll.Caption = "선택All"
        cmdAll.ForeColor = RGB(0, 0, 255)
        For i = 1 To sprSLip.DataRowCnt
            sprSLip.Row = i
            sprSLip.Col = 7
            sprSLip.Text = "0"
        Next i
    End If

End Sub

Private Sub cmdBMVerify_Click()
    
    
    GoSub BM_PBSmear_Update
    GoSub BM_ASpiration_Update
    GoSub BM_Diffc_Update
    
    GoSub BM_General_Update
    
    Exit Sub
    

    
    
BM_PBSmear_Update:
    Dim sJeobsuDt       As String
    sJeobsuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")
    
    
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General_Sub"
    strSql = strSql & " SET    Chamgo = '" & txtSmear.Text & "'"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    SLipno2  = " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    ItemCD   = '1505011'"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    
    Return
    
BM_ASpiration_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General_Sub"
    strSql = strSql & " SET    Chamgo = '" & txtAspiration.Text & "'"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    SLipno2  = " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    ItemCD   = '1505012'"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
    
BM_Diffc_Update:
    Dim sItemCd     As String
    Dim sCount      As String
    Dim sPercent    As String
    
    For i = 1 To 18
        sprDiffc.Row = i
        sprDiffc.Col = 1: sItemCd = sprDiffc.Text
        sprDiffc.Col = 3: sCount = sprDiffc.Text
        sprDiffc.Col = 4: sPercent = sprDiffc.Text
        
        strSql = ""
        strSql = strSql & " UPDATE TWEXAM_General_Sub"
        strSql = strSql & " SET    Result1  = '" & sPercent & "',"
        strSql = strSql & "        Result2  = '" & sCount & "',"
        strSql = strSql & "        Verify   = 'Y'"
        strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
        strSql = strSql & " AND    SLipno1  = " & Val(Left(cmbSLip.Text, 2))
        strSql = strSql & " AND    SLipno2  = " & Val(txtSLipno2.Text)
        strSql = strSql & " AND    ItemCD   = '" & sItemCd & "'"
        
        adoConnect.BeginTrans
        If adoExec(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
        End If
    Next
    Return
    

BM_General_Update:
    Dim sGeomsaDt       As String
    Dim sGeomsaT1       As String
    Dim sGeomsaT2       As String
    
    sGeomsaDt = Dual_Date_Get("yyyy-MM-dd")
    sGeomsaT1 = Dual_Date_Get("hh")
    sGeomsaT2 = Dual_Date_Get("mi")
    

    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_GENERAL "
    strSql = strSql & " SET    GeomsaDt   =   TO_DATE('" & sGeomsaDt & "','YYYY-MM-DD'),"
    strSql = strSql & "        GeomsaT1   =    " & sGeomsaT1 & ","
    strSql = strSql & "        GeomsaT2   =    " & sGeomsaT2 & ","
    strSql = strSql & "        Geomsaja   =   '" & GstrIdnumber & "',"
    strSql = strSql & "        Status     =   'C'"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    SLipno2  = " & Val(txtSLipno2.Text)
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

End Sub

Private Sub cmdCallHelp_Click()
    
    frmQrySample.Show vbModal

End Sub



Private Sub CmdCancel_Click()

    panelUrine.Visible = False
    
End Sub

Private Sub cmdHelp_Click()
    
    
    frmLabList.Show vbModal
    
    
End Sub

Private Sub cmdHide_Click()
    
    panelText.Visible = False
    
End Sub


Private Sub cmdHideGs_Click()
    
    panelGs.Visible = False
    txtResult1.Text = ""
    txtResult.Text = ""
    txtRange.Text = ""
    
    
End Sub

Private Sub cmdOk_Click()

    sprSLip.Row = sprSLip.ActiveRow
    sprSLip.Col = 2
    sprSLip.Text = Trim(txtCastResult.Text)
    panelUrine.Visible = False

End Sub

Private Sub cmdPrint_Click()
    Dim sPrSmear        As String
    Dim sPrAspiration   As String
    Dim sPrDiffCount    As String
    
    Dim sDiffCode       As String * 10
    Dim sDiffName       As String * 30
    Dim sDiffCount      As String * 6
    Dim sDiffPercent    As String * 5
    Dim sJeobsuDt       As String
    
    sJeobsuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")
    
    
    GoSub Select_Data_General_Sub
    GoSub Print_Data_General_Sub
    Exit Sub
    
    
Select_Data_General_Sub:
    strSql = ""
    strSql = strSql & " SELECT a.*, b.ItemNM"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_ITEMML      b "
    strSql = strSql & " WHERE  a.JeobsuDt =  TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  =  15 "
    strSql = strSql & " AND    a.SLipno2  =  " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    a.RoutinCd =  '150050'"        'Bone marrow RoutineCode
    strSql = strSql & " AND    a.ItemCd   =  b.Codeky(+)"
    strSql = strSql & " ORDER  BY a.iTemCd"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        
        Select Case Trim(adoSet.Fields("iTemCD").Value & "")
            Case "1505011":
                sPrSmear = adoSet.Fields("Chamgo").Value & ""
            Case "1505012":
                sPrAspiration = adoSet.Fields("Chamgo").Value & ""
            Case "15050201" To "15050218":
                sDiffCode = adoSet.Fields("ItemCD").Value & ""
                sDiffName = adoSet.Fields("ItemNM").Value & ""
                sDiffCount = adoSet.Fields("Result2").Value & ""
                sDiffPercent = adoSet.Fields("Result1").Value & ""
                
                sPrDiffCount = sPrDiffCount & sDiffName & sDiffCount & sDiffPercent & "%" & vbCrLf
            Case Else
        End Select
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    Return
    


Print_Data_General_Sub:
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 20
    Printer.FontBold = True
    Printer.FontItalic = False
    
    
    Printer.Print ""
    Printer.Print Space(12) & "Bone marrow study report"
    Printer.Print ""
    
    Printer.FontName = "바탕체":  Printer.FontSize = 12:  Printer.FontBold = False
    Printer.Print "환자정보 : " & txtPtno.Text & " " & txtSname.Text & " " & txtSex.Text & "/" & txtAge.Text & "  " & _
                                  "(" & Left(cmbSLip.Text, 2) & "-" & txtSLipno2.Text & ")"
    
    Printer.Print ""
    Printer.FontName = "바탕체":  Printer.FontSize = 12:  Printer.FontBold = True
    Printer.Print "Peripheral blood smear"
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Printer.FontName = "바탕체":  Printer.FontSize = 10:  Printer.FontBold = False
    Printer.Print sPrSmear
    Printer.Print ""
    
    
    Printer.FontName = "바탕체":  Printer.FontSize = 12:  Printer.FontBold = True
    Printer.Print "Bone marrow aspiration"
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Printer.FontName = "바탕체":  Printer.FontSize = 10:  Printer.FontBold = False
    Printer.Print sPrAspiration
    Printer.Print ""
    Printer.Print ""
    
    Printer.FontName = "바탕체":  Printer.FontSize = 12:  Printer.FontBold = False
    Printer.Print "DiffCount (500 cells are counted)"
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Printer.FontName = "굴림체":  Printer.FontSize = 10:  Printer.FontBold = False
    Printer.Print sPrDiffCount
    Printer.Print ""
    
    Printer.EndDoc
    
    Return
    

End Sub

Private Sub cmdRemark_Click()
    Dim sTmpTEXT        As String
    Dim sTmpSLip        As String

    
    Select Case Left(Me.cmbSLip, 2)
        Case "11", "22", "31" To "34":
            If Trim(txtGeomsaCm.Text) <> "" Then
                sTmpTEXT = txtGeomsaCm.Text
            End If
    End Select
    
    
    hWndReturn = Me.txtGeomsaCm.hwnd
    gSRmkSLipno = Left(Me.cmbSLip.Text, 2)
    clpRemark.Show vbModal
    
    Select Case gSRmkSLipno
        Case "11", "22", "31" To "34":
                If Trim(sTmpTEXT) <> "" Then
                    If vbYes = MsgBox("이전 Remark 사항과 합치겠습니까?", vbYesNo + vbQuestion, "Remark 합치기") Then
                        txtGeomsaCm.Text = sTmpTEXT & vbCrLf & txtGeomsaCm.Text
                    End If
                End If
    End Select
    
    
    gSRmkSLipno = ""
    sTmpTEXT = ""


End Sub

Private Sub cmdRmkAspi_Click()
    
    hWndReturn = Me.txtAspiration.hwnd
    gSRmkSLipno = Left(Me.cmbSLip.Text, 2)
    clpRemark.Show vbModal
    gSRmkSLipno = ""


End Sub

Private Sub cmdRmkPbsmear_Click()
    
    hWndReturn = Me.txtSmear.hwnd
    gSRmkSLipno = Left(Me.cmbSLip.Text, 2)
    clpRemark.Show vbModal
    gSRmkSLipno = ""

End Sub

Private Sub cmdSelect_Click()
    Dim sGsItem     As String
    Dim sRet        As String
    
    sRet = Trim(txtResult1.Text) & " " & Trim(txtResult.Text) & " " & Trim(txtRange.Text)
    
    sprSLip.Row = sprSLip.ActiveRow
    sprSLip.Col = 2: sprSLip.Text = Trim(sRet)
    
    
    
End Sub

Private Sub cmdVerify_Click()
    
    Dim i
    Dim LiRecallCnt     As Integer
    Dim LiVerifyCnt     As Integer
    Dim LiNoVerifyCnt   As Integer
    Dim LiPos           As Integer
    Dim LiLen           As Integer
    Dim LiSlipNo1       As Integer
    Dim LiSlipNo2       As Integer
    Dim LsStatus        As String
    Dim sJeobsuDt       As String
    Dim sRowID          As String
    Dim iABNormalCnt    As Integer
    Dim sPanic          As String * 1
    Dim sDelta          As String * 1
    
    
    GoSub GeomsaCm_Update_Sub
    
    For i = 1 To sprSLip.DataRowCnt
        sprSLip.Row = i
        sprSLip.Col = 9:  sDelta = sprSLip.Text
        sprSLip.Col = 18: sPanic = sprSLip.Text
        If Trim(sDelta) = "D" Or Trim(sPanic) = "P" Then
            iABNormalCnt = iABNormalCnt + 1
        End If
    Next
    
    
    For i = 1 To sprSLip.DataRowCnt
        sprSLip.Row = i
        sprSLip.Col = 2
        If sprSLip.BackColor = RGB(250, 250, 225) Then
            LiNoVerifyCnt = LiNoVerifyCnt + 1
        End If
        
        sprSLip.Col = 7
        If sprSLip.Text = "1" Then
            LiVerifyCnt = LiVerifyCnt + 1
            
            sprSLip.Col = 2
            If sprSLip.BackColor = RGB(250, 250, 225) Then
                 LiNoVerifyCnt = LiNoVerifyCnt - 1
            End If
            
            sprSLip.Col = 12
            If sprSLip.Text <> "" Then
                sprSLip.Col = 17
                'If sprSLip.Text <> "S" Then       'Culture Sensitivity then SKip
                    GoSub VERIFY_OK_RTN
                'End If
            End If
        End If
        
       sprSLip.Col = 10     'Spread Title = 'T'
       If sprSLip.Text = "1" Then
           LiRecallCnt = LiRecallCnt + 1
           GoSub VERIFY_RECALL_RTN
       End If
    Next i

    If LiVerifyCnt = 0 And LiRecallCnt = 0 Then         'Verify Check 된것이 하나도 없을때
        GoSub Micro_Senstivity
        Exit Sub
    End If
    

    If LiNoVerifyCnt >= 1 Or LiRecallCnt >= 1 Then
        LsStatus = "P"
    ElseIf LiRecallCnt = sprSLip.DataRowCnt Then
        LsStatus = "R"
    Else
        LsStatus = "C"
    End If
    
    sMsg = "Panic Or Delta 결과 가 있습니다!....." & vbCrLf & "abnormal Data 로 등록하여 관리하시겠습니까?"
    If iABNormalCnt > 0 Then
        If vbYes = MsgBox(sMsg, vbYesNo + vbQuestion, "abnormal 결과 등록") Then LsStatus = "X"
    End If
    
    
    GoSub PARTorALL_VERIFY_OK_RTN
    
    sprSLip.Col = 7
    sprSLip.Row = 0
    sprSLip.Text = "V"
    
    GoSub Micro_Senstivity
    
    If LsStatus = "X" Then
        mdiMain.stbMain.Panels(1).Text = "이상 Data로 Verify 되었습니다!..."
    Else
        mdiMain.stbMain.Panels(1).Text = "Verify 확인 되었습니다!..."
    End If
    
    
    GoSub Right_Celar_Sub
    
    Exit Sub
    
    
    
Right_Celar_Sub:
    
    cmdAll.Caption = "선택All"
    txtPtno.Text = ""
    txtSname.Text = ""
    txtSex.Text = ""
    txtAge.Text = ""
    txtDept.Text = ""
    txtDr.Text = ""
    txtRoom.Text = ""
    txtBarCode.Text = ""
    txtStatus.Text = ""
    txtGeomsaCm.Text = ""
    txtSLipno2.SetFocus
    
    sprSLip.ReDraw = False
        
    sprSLip.MaxRows = 0
    sprSLip.MaxRows = 120
    sprSLip.RowHeight(-1) = 11
    
    sprSLip.ReDraw = True
    
    dtJeobsu.Tag = ""
    
    
    Return
    
    
Micro_Senstivity:
    For i = 1 To sprSLip.DataRowCnt
        sprSLip.Row = i
        sprSLip.Col = 17
        If Trim(sprSLip.Text) = "S" Then
            frmMicroClass.Show vbModal
            Exit For
        End If
    Next
    Return
    
    
GeomsaCm_Update_Sub:
    
    LiSlipNo1 = Val(Left(cmbSLip.Text, 2))
    LiSlipNo2 = Val(txtSLipno2.Text)
    
    
    sJeobsuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")

    
    gStrSql = ""
    gStrSql = gStrSql & " UPDATE TWEXAM_GENERAL "
    gStrSql = gStrSql & " SET    GeomsaCm   =   '" & Quot_Conv(Trim(txtGeomsaCm.Text)) & "'"
    gStrSql = gStrSql & " WHERE  JeobsuDt    =   TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    gStrSql = gStrSql & " AND    SlipNo1     =   " & LiSlipNo1
    gStrSql = gStrSql & " AND    SlipNo2     =   " & LiSlipNo2
    adoConnect.BeginTrans
    If adoExec(gStrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    

    
VERIFY_OK_RTN:
    Dim sResult1
    
    sprSLip.Col = 17:
    If Trim(sprSLip.Text) = "B" Then     'Bonemarrow 건너뜀
        Return
    End If
    
    sprSLip.Col = 12: sRowID = sprSLip.Text
    sprSLip.Col = 2:  sResult1 = Trim(sprSLip.Text)
    
    LiPos = InStr(Time, ":") + 1
    
    gStrSql = ""
    gStrSql = gStrSql & " UPDATE TWEXAM_GENERAL_SUB "
    gStrSql = gStrSql & " SET    Verify     =   'Y',"
    gStrSql = gStrSql & "        Result1    =   '" & Quot_Conv(sResult1) & "'"
    gStrSql = gStrSql & " WHERE  ROWID      =   '" & sRowID & "'"
    
    Call adoExec(gStrSql)

    Return


PARTorALL_VERIFY_OK_RTN:
    Dim cGeomsaDt       As String
    Dim cGeomsat1       As Integer
    Dim cGeomsat2       As Integer
    
    LiSlipNo1 = Val(Left(cmbSLip.Text, 2))
    LiSlipNo2 = Val(txtSLipno2.Text)
    sJeobsuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")
    
    
    cGeomsaDt = Dual_Date_Get("yyyy-MM-dd")
    cGeomsat1 = Format(Dual_Date_Get("hh24"), "00")
    cGeomsat2 = Format(Dual_Date_Get("mi"), "00")
    
    gStrSql = ""
    gStrSql = gStrSql & " UPDATE TWEXAM_GENERAL "
    gStrSql = gStrSql & " SET    GeomsaDt   =   TO_DATE('" & cGeomsaDt & "','YYYY-MM-DD'),"
    gStrSql = gStrSql & "        GeomsaT1   =    " & cGeomsat1 & ","
    gStrSql = gStrSql & "        GeomsaT2   =    " & cGeomsat2 & ","
    gStrSql = gStrSql & "        Geomsaja   =   '" & GstrIdnumber & "',"
    gStrSql = gStrSql & "        Status     =   '" & LsStatus & "'"
    gStrSql = gStrSql & " WHERE JeobsuDt    =   TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    gStrSql = gStrSql & " AND   SlipNo1     =   " & LiSlipNo1
    gStrSql = gStrSql & " AND   SlipNo2     =   " & LiSlipNo2
    
    adoConnect.BeginTrans
    If adoExec(gStrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
VERIFY_RECALL_RTN:
    
    sprSLip.Col = 12: sRowID = sprSLip.Text
    
    gStrSql = ""
    gStrSql = gStrSql & " UPDATE  TWEXAM_GENERAL_SUB "
    gStrSql = gStrSql & " SET     Verify     =   'N'"
    gStrSql = gStrSql & " WHERE   ROWID      =   '" & sRowID & "'"
    
    adoConnect.BeginTrans
    If adoExec(gStrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If

   
   '-------------------------------------------------
    LiSlipNo1 = Val(Left(cmbSLip.Text, 2))
    LiSlipNo2 = Val(txtSLipno2.Text)
    sJeobsuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")
    
    gStrSql = ""
    gStrSql = gStrSql & " UPDATE  TWEXAM_GENERAL "
    gStrSql = gStrSql & " SET     Status     =   'P'"
    gStrSql = gStrSql & " WHERE   JeobsuDt   =   TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')  "
    gStrSql = gStrSql & " AND     SlipNo1    =   " & LiSlipNo1
    gStrSql = gStrSql & " AND     SlipNo2    =   " & LiSlipNo2
    adoConnect.BeginTrans
    If adoExec(gStrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If

    
    Return

End Sub


Private Sub Form_Load()
    
        
    GoSub FORMClear
    dtJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    
    GoSub SLip_Select
    
    GiExamNumb = Val(GetSetting("CP", "CPRESULT", "SLip"))
    
    Call SetComboBox(cmbSLip, GiExamNumb, 2)
    cmbWhere.ListIndex = 3
    chkWhere(0).Value = "1"
    
    
    
    Exit Sub
    
'/--------------------------------------------------------

FORMClear:
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    Return
    
    
SLip_Select:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky < '90'"
    strSql = strSql & " ORDER  BY Codeky"
    
    cmbSLip.Clear
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return

    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim sSLip       As String
    
    
    'sSLip = Left(cmbSLip.Text, 2)
    'Call SaveSetting("CP", "CPRESULT", "SLip", sSLip)
    
    
End Sub


Private Sub Label1_Click()

End Sub

Private Sub lstCast_Click()
    Dim i           As Integer
    
    lstGrade.Clear
    
    txtCastResult = Trim(UrCastResult(lstCast.ListIndex))
    
    If UrCastGubun(lstCast.ListIndex) = "1" Then
        
        For i = 0 To 9
            lstGrade.AddItem UrCellGrade(i)
        Next i
    
    ElseIf UrCastGubun(lstCast.ListIndex) = "0" Then
        
        For i = 0 To 9
            lstGrade.AddItem UrCastGrade(i)
        Next i
    
    End If


End Sub

Private Sub lstGrade_Click()
    If UrCastGubun(lstCast.ListIndex) = "1" Then
        
        txtCastResult = Trim(UrCastResult(lstCast.ListIndex)) & " " & UrCellGrade(lstGrade.ListIndex)
    
    ElseIf UrCastGubun(lstCast.ListIndex) = "0" Then
        
        txtCastResult = Trim(UrCastResult(lstCast.ListIndex)) & " " & UrCastGrade(lstGrade.ListIndex)
    
    End If

End Sub

Private Sub lstitem_Click()
    
    txtResult1.Text = lstitem.List(lstitem.ListIndex)
    txtResult.Text = ""
    txtRange.Text = ""
    
    
    
    GoSub Set_Range
    
    Select Case lstitem.ListIndex
        Case 0:  'WBC
                 lstResult.Clear
                 GoSub Set_Range
        Case 1:  'Epithelial Cells
                 lstResult.Clear
                 GoSub Set_Range
        Case 2:  'Gram(+)
                 lstResult.Clear
                 lstResult.AddItem "cocci"
                 lstResult.AddItem "cocci in pair"
                 lstResult.AddItem "cocci in chain"
                 lstResult.AddItem "cocci in cluster"
                 lstResult.AddItem "bacilli"
                 lstResult.AddItem "bacilli, large"
                 lstResult.AddItem "bacilli, coryneform"
                 lstResult.AddItem "bacilli, filamentous"
                 GoSub Set_Range
        Case 3:  'Gram(-)
                 lstResult.Clear
                 lstResult.AddItem "cocci"
                 lstResult.AddItem "coccoid bacilli"
                 lstResult.AddItem "diplococci"
                 lstResult.AddItem "bacilli"
                 lstResult.AddItem "bacilli filamentous"
                 lstResult.AddItem "bacilli small pleomorphic"
                 GoSub Set_Range
        Case 4:  'Other
                 lstResult.Clear
                 lstRange.Clear
                 txtResult.Text = ""
                 txtRange.Text = ""
        Case 5, 6, 7: lstResult.Clear
                      GoSub Set_Range
        
    End Select
    
    Exit Sub
    
Set_Range:
    lstRange.Clear
    lstRange.AddItem "Rare(1~4 immersion oil field)"
    lstRange.AddItem "Few (5~9 immersion oil field)"
    lstRange.AddItem "A Few(10~15 immersion oil field)"
    lstRange.AddItem "Some(15~24 immersion oil field)"
    lstRange.AddItem "Many( >24 immersion oil field)"
    Return
    
End Sub

Private Sub lstMicroList_Click()
    
    txtSLipno2.Text = Left(lstMicroList.Text, 5)
    
    If Trim(txtSLipno2.Text) <> "" Then
        DoEvents: Call txtSLipno2_KeyDown(vbKeyReturn, 1)
    End If

End Sub

Private Sub lstRange_Click()
    
    Select Case lstitem.ListIndex
        Case 2, 3: txtRange.Text = lstRange.List(lstRange.ListIndex)
        Case 5, 6: txtRange.Text = lstRange.List(lstRange.ListIndex)
        Case Else: txtResult1.Text = ""
                   txtResult.Text = ""
                   txtRange.Text = lstRange.List(lstRange.ListIndex)

    End Select
    
    
    
End Sub

Private Sub lstResult_Click()
    
    txtResult.Text = lstResult.List(lstResult.ListIndex)
    txtRange.Text = ""
    
End Sub

Private Sub lstText_DblClick()
    
    sprSLip.Row = sprSLip.ActiveRow
    sprSLip.Col = 2
    sprSLip.Text = Trim(lstText.List(lstText.ListIndex))

End Sub

Private Sub panelDate_DblClick()
    
'    If panelDate.Caption = "접수일자:" Then
'        panelDate.Caption = "검사일자:"
'        panelDate.Tag = "G"
'        panelDate.ForeColor = RGB(0, 0, 255)
'    Else
'        panelDate.Caption = "접수일자:"
'        panelDate.Tag = "J"
'        panelDate.ForeColor = RGB(0, 0, 0)
'    End If
    
    
End Sub

Private Sub sprDiffc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If sprDiffc.ActiveCol = 3 Then
            GoSub Calcurate_Sub
        End If
    End If
    Exit Sub
    
Calcurate_Sub:
    Dim sngCount        As Single
    Dim sngPercent      As Single
    
    
    sprDiffc.Row = sprDiffc.ActiveRow
    sprDiffc.Col = 3:
    If Trim(sprDiffc.Text) = "" Then
        sprDiffc.Col = 4: sprDiffc.Text = ""
        Return
    Else
        sprDiffc.Col = 3: sngCount = CSng(sprDiffc.Text)
        sprDiffc.Col = 4: sprDiffc.Text = (sngCount / 500) * 100
    End If
    
    Return
    
End Sub

Private Sub sprSLip_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim sItemCode       As String
    
    If Col = 1 Then
        sprSLip.Row = Row
        sprSLip.Col = 17
        Select Case Trim(sprSLip.Text)
            Case "S"       'Culture & Senstivity Check
                    frmSens.Show vbModal
        End Select
        
    End If

End Sub

Public Sub sprSLip_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lngOrderno        As Long
    
    
    mdiMain.stbMain.Panels(1).Text = ""
    
    GoSub Select_CmDoctor
    GoSub Select_Sample
    
    Exit Sub
    
    
Select_CmDoctor:
    sprSLip.Row = Row
    sprSLip.Col = 16: lngOrderno = Val(sprSLip.Text)
    
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Order"
    strSql = strSql & " WHERE  Ptno = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    Orderno = " & lngOrderno
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    mdiMain.stbMain.Panels(1).Text = Trim(adoSet.Fields("CmDoctor").Value & "")
    
    Call adoSetClose(adoSet)
    Return
    
Select_Sample:
    Dim sCompCode       As String
    Dim sRowID          As String
    
    Dim sSampleC        As String
    Dim sSampleN        As String
    
    sSampleC = "": sSampleN = ""
    
    sprSLip.Row = Row
    sprSLip.Col = 11: sCompCode = Trim(sprSLip.Text)
    sprSLip.Col = 12: sRowID = sprSLip.Text
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Code, b.Codenm, b.Class2 "
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_Sample      b"
    strSql = strSql & " WHERE  a.ROWID    = '" & sRowID & "'"
    strSql = strSql & " AND    a.GeomchCd = b.Code(+)"
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    sSampleC = Trim(adoSet.Fields("Code").Value & "")
    sSampleN = Trim(adoSet.Fields("Codenm").Value & "")
    
    Call adoSetClose(adoSet)
    
    If Trim(sSampleC) <> "" Then
        mdiMain.stbMain.Panels(1).Text = mdiMain.stbMain.Panels(1).Text & "  =>" & _
                                         sSampleC & ":" & sSampleN
    End If
    
    Return
    
    
    
End Sub

Private Sub sprSLip_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then Exit Sub
    If Row > sprSLip.DataRowCnt Then Exit Sub
    
    If Col = 1 Then frmHistory.Show vbModal
    
End Sub

Private Sub sprSLip_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode < 112 Or KeyCode > 123 Then Exit Sub
    
    sprSLip.Row = sprSLip.ActiveRow
    sprSLip.Col = 1
    If Trim(sprSLip.Text) = "" Then Exit Sub
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT * "
    gStrSql = gStrSql & " FROM   TWEXAM_SPECODE    "
    gStrSql = gStrSql & " WHERE  CodeGu  =  '19'    "
    gStrSql = gStrSql & " AND    CodeKy  =  '" & Left(cmbSLip.Text, 2) & KeyCode & "' "
    If False = adoSetOpen(gStrSql, adoSet) Then Exit Sub
    sprSLip.Col = 2
    sprSLip.Text = adoSet.Fields("CodeNm").Value & ""
    Call adoSetClose(adoSet)

End Sub

Private Sub sprSLip_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim LsJeobSuDt          As String
    Dim LiSlipNo1           As Integer
    Dim LiSlipNo2           As Integer

    
    If Row = 0 Then Exit Sub
    
    If Col = 2 Then
        sprSLip.Col = Col
        sprSLip.Row = Row
        
        If sprSLip.Text = "" Then
            sprSLip.Col = 9
            sprSLip.Text = ""
            sprSLip.BackColor = RGB(235, 245, 235)
            Exit Sub
        End If
        
        
        LsJeobSuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")
        LiSlipNo1 = Val(Left(cmbSLip.Text, 2))
        LiSlipNo2 = Val(txtSLipno2.Text)
        
        If Left(cmbSLip.Text, 1) <> "4" Then   '미생물실이 아니면.........
            GoSub READ_PREVIOUS_DATA
        End If
        
        sprSLip.Col = 11
        If sprSLip.Text = "510101" Then    '(혈액형 Data 는 대문자로 치환 (무식하게 소문자가 있더라구요))
            sprSLip.Text = UCase(sprSLip.Text)
        End If
        
    End If
    
    
    Exit Sub
    
    
'-------------------------------------------------------------------------------
READ_PREVIOUS_DATA:
    Dim jj          As Integer
    Dim LsChkDate   As String
    Dim LsItemCD    As String
    Dim LsQC        As String * 1
    
    Dim LiResult
    Dim LiMinCham
    Dim LiMaxCham
    Dim LiPanicMin
    Dim LiPanicMax
    Dim LiDeltaMin
    Dim LiDeltaMax
    Dim LiCurVal
    Dim LiPreVal
    
    
    LsChkDate = UCase(Format$(LsJeobSuDt, "YYYY-MM-DD"))
    
    sprSLip.Row = Row
    sprSLip.Col = 2:  LiResult = Val(sprSLip.Text)
    sprSLip.Col = 11: LsItemCD = sprSLip.Text
            
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT  TO_CHAR(a.JeobsuDT,'YYYY-MM-DD') JeobsuDt,"
    gStrSql = gStrSql & "         a.Result1, a.Slipno2, a.ItemCD,  b.PanicMin, b.PanicMax"
    gStrSql = gStrSql & " FROM    TWEXAM_GENERAL_SUB a,"
    gStrSql = gStrSql & "         TWEXAM_ITEMML      b "
    gStrSql = gStrSql & " WHERE  a.Ptno      =  '" & txtPtno.Text & "'"
    gStrSql = gStrSql & " AND    a.ItemCD    =  '" & LsItemCD & "'"
    gStrSql = gStrSql & " AND    a.Verify    =  'Y'"
    gStrSql = gStrSql & " AND    a.JeobsuDt <=  TO_DATE('" & LsJeobSuDt & "','YYYY-MM-DD')"
    gStrSql = gStrSql & " AND    a.Itemcd    =  b.Codeky(+)"
    gStrSql = gStrSql & " ORDER  BY JeobsuDt DESC, SLipno2 ASC"
    
    If adoSetOpen(gStrSql, adoSet) Then
        LiPanicMin = Val(adoSet.Fields("PanicMin").Value & "")
        LiPanicMax = Val(adoSet.Fields("PanicMax").Value & "")
        If LsChkDate = adoSet.Fields("Jeobsudt").Value And LiSlipNo2 = adoSet.Fields("slipno2").Value Then
            '
        Else
            sprSLip.Col = 3:     sprSLip.Text = Trim(adoSet.Fields("Result1").Value & "")
            Call adoSetClose(adoSet)
        End If
        If Not adoSet Is Nothing Then Call adoSetClose(adoSet)
    Else
        sprSLip.Col = 3:    sprSLip.Text = ""
        
        strSql = " SELECT PanicMin, PanicMax FROM TWEXAM_ITEMML WHERE Codeky = '" & LsItemCD & "'"
        If adoSetOpen(strSql, adoSet) Then
            LiPanicMin = Val(adoSet.Fields("PanicMin").Value & "")
            LiPanicMax = Val(adoSet.Fields("PanicMax").Value & "")
            Call adoSetClose(adoSet)
        End If
    End If
    
    
    '---------------------------'
    '  PANIC VALUE CHECK        '
    '---------------------------'
    If LiResult > 0 Then                             '결과DATA 가 있을때
        If LiPanicMin = 0 And LiPanicMax = 0 Then        'Panic 기초Data 가 없을때 Check 안함...
                sprSLip.Col = 8:    sprSLip.Text = ""
                sprSLip.BackColor = RGB(235, 245, 235)
        Else
            If (LiResult < LiPanicMin) Or (LiResult > LiPanicMax) Then      'Panic...
                sprSLip.Col = 8:    sprSLip.Text = " "
                sprSLip.BackColor = RGB(250, 250, 0)
                sprSLip.Col = 18: sprSLip.Text = "P"
            Else
                sprSLip.Col = 8:    sprSLip.Text = ""
                sprSLip.BackColor = RGB(235, 245, 235)
            End If
        End If
    End If
    
    


    '---------------------------'
    '  DELTA VALUE CHECK        '
    '---------------------------'
    sprSLip.Col = 9
    If sprSLip.Text = "" Then
        sprSLip.Text = ""
        sprSLip.BackColor = RGB(235, 245, 235)
    End If
    
    sprSLip.Col = 2:    LiCurVal = Val(sprSLip.Text)
    sprSLip.Col = 3:    LiPreVal = Val(sprSLip.Text)
    sprSLip.Col = 13:   LsQC = sprSLip.Text
    sprSLip.Col = 14:   LiDeltaMin = Val(sprSLip.Text)
    sprSLip.Col = 15:   LiDeltaMax = Val(sprSLip.Text)

    If LiPreVal <> 0 And LiCurVal <> 0 Then       '양쪽(계산할 2개의 Data가 모두 있을때...)
        If LsQC = "1" Or LsQC = "2" Or LsQC = "3" Or LsQC = "4" Then
            If DeltaCheck(LiCurVal, LiPreVal, LsQC) < LiDeltaMin Or _
               DeltaCheck(LiCurVal, LiPreVal, LsQC) > LiDeltaMax Then
                sprSLip.Col = 9
                If Trim(sprSLip.Text) = "" Then
                    sprSLip.Text = "D"
                    sprSLip.BackColor = RGB(250, 0, 0)
                End If
            Else
                sprSLip.Col = 9
                sprSLip.Text = ""
                sprSLip.BackColor = RGB(235, 245, 235)
            End If
        End If
    End If
    
Return

End Sub

Private Sub SSCommand1_Click()
    Dim sJeobsuDt       As String
    Dim sWhere          As String
    
    
    sJeobsuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")
    lstMicroList.Clear
    mdiMain.stbMain.Panels(1).Text = ""
    
    GoSub Right_Clear_Sub
    
    
    strSql = ""
    sWhere = ""
    
    
    Select Case cmbWhere.ListIndex
        Case 0: '응급
                 sWhere = strSql & " AND  (RTRIM(a.DeptCode) = 'ER' OR a.GBER = 'E' )"
                 sWhere = sWhere & Set_CheckBox_SqlSum(chkWhere, " a.Status")
        Case 1: 'ABNormal
                 sWhere = strSql & " AND  a.Status = 'X'"
        Case 2: '전체
                 sWhere = ""
        Case 3: '조건별
                 sWhere = Set_CheckBox_SqlSum(chkWhere, " a.Status")
                 If chkWhere(3).Value = "1" Then
                    sWhere = strSql & " AND  a.ReporCd = 'W'"
                 End If
    End Select
        
        
    'strSql = ""
    'strSql = strSql & " SELECT a.SLipno2, b.Sname"
    'strSql = strSql & " FROM   TWEXAM_General a,"
    'strSql = strSql & "        TWEXAM_IDNOMST b,"
    'strSql = strSql & "        TWEXAM_Order   c "
    'strSql = strSql & " WHERE  a.JeobsuDt  = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    'strSql = strSql & " AND    a.SLipno1   = " & Val(Left(cmbSLip.Text, 2))
    'strSql = strSql & " AND    a.GBCH      = 'Y'"
    'strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    'strSql = strSql & " AND    a.JeobsuDt  = c.CollDate(+)"
    'STRSQL = STRSQL & " AND    a.Ptno      = c.Ptno(+)
    'strsql = strsql & " AND    a.SLipno1   = c.SLipno1(+)
    'strsql = strsql & " AND    a.jeobsuT1  = c.CollHH(+)
    'strsql = strsql & " AND    a.JeobsuT2  = C.CollMM(+)
    'strSql = strSql & " AND    a.Matchno   = c.Matchno(+)"
    'strSql = strSql & " AND    c.JeobsuYn != '#'"     'Match No = Unique
        
    strSql = ""
    strSql = strSql & " SELECT a.SLipno2, b.Sname"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TWEXAM_IDNOMST b"
    strSql = strSql & " WHERE  a.JeobsuDt  = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.GBCH      = 'Y'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    
    If Trim(txtMsample.Text) <> "" Then
        strSql = strSql & " AND a.GeomchCd  = '" & txtMsample.Text & "'"      '검체TextBox
    End If
    
    
    strSql = strSql & sWhere
    strSql = strSql & " GROUP  BY a.SLipno2, b.Sname"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        lstMicroList.AddItem Format(adoSet.Fields("SLipno2").Value & "", "00000") & "  " & _
                                    adoSet.Fields("Sname").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Exit Sub
    



Right_Clear_Sub:
    GoSub Spread_Clear_SLip
    txtBarCode.Text = ""
    Me.txtSLipno2.Text = ""
    txtPtno.Text = ""
    txtSname.Text = ""
    txtSex.Text = ""
    txtAge.Text = ""
    txtDept.Text = ""
    txtDr.Text = ""
    txtRoom.Text = ""
    txtCount.Text = ""
    txtGeomsaCm.Text = ""
    txtStatus.Text = ""
    
    Return
    
    
Spread_Clear_SLip:
    sprSLip.ReDraw = False
        
    sprSLip.MaxRows = 0
    sprSLip.MaxRows = 120
    sprSLip.RowHeight(-1) = 11
    
    sprSLip.ReDraw = True
    
    Return
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1:  Unload Me
        Case 2:
        Case 3:  GoSub Clear_Button_Sub
        Case 4:
        Case 5:  ' If Left(cmbSLip.Text, 2) = "11" Then frmDiffc.Show vbModal
                 GoSub Set_GramStain_ResultData
                 panelGs.Visible = True
                 panelGs.ZOrder 0
        Case 6:
        Case 7:  GoSub Cast_Set
        Case 8:
        Case 9:  frmFunc.Show vbModal
        Case 10: GoSub Get_FunctionText
        Case 11:
        Case 12: gOiLLQryPtno = txtPtno.Text
                 If Trim(gOiLLQryPtno) <> "" Then frmOills.Show vbModal
        Case 14: frmWhere.Show vbModal
        
        'Case 16: frmMicroEnrol.Show
        '         frmMicroEnrol.ZOrder 0

    End Select
    Exit Sub
    

Set_GramStain_ResultData:
    lstitem.Clear
    
    lstitem.AddItem "WBC"
    lstitem.AddItem "Epithelial Cells"
    lstitem.AddItem "Gram(+)"
    lstitem.AddItem "Gram(-)"
    lstitem.AddItem "Other"
    
    lstitem.AddItem "Budding Yeast like cell, pseudohyphae"
    lstitem.AddItem "Clue Cell"
    lstitem.AddItem "not found"
    
    Return

Cast_Set:
    If Left(cmbSLip.Text, 2) = "13" Or _
       Left(cmbSLip.Text, 2) = "24" Then
        Call Init_Result_Data
        panelUrine.Visible = True
        panelUrine.ZOrder 0
    End If
    
    Return
    
    
    
Clear_Button_Sub:
    mdiMain.stbMain.Panels(1).Text = ""
    cmdAll.Caption = "선택All"
    
    GoSub Spread_Clear_SLip

    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
    
    txtMsample.Text = ""
    
    cmbWhere.ListIndex = 3
    chkWhere(0).Value = "1"
    chkWhere(1).Value = "0"
    chkWhere(2).Value = "0"
    chkWhere(3).Value = "0"
    chkWhere(4).Value = "0"
    
    lstMicroList.Clear
    dtJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Call SpreadSetClear(sprDiffc)
    
    If Left(cmbSLip.Text, 2) = "15" Then
        panelBonemarrow.Visible = True
        panelBonemarrow.ZOrder 0
    Else
        Call SetComboBox(cmbSLip, GiExamNumb, 2)
        txtSLipno2.SetFocus
    End If
    
    panelDate.Caption = "접수일자:"
    panelDate.Tag = "J"
    panelDate.ForeColor = RGB(0, 0, 0)
    Return
    


Spread_Clear_SLip:
    sprSLip.ReDraw = False
        
    sprSLip.MaxRows = 0
    sprSLip.MaxRows = 120
    sprSLip.RowHeight(-1) = 11
    
    sprSLip.ReDraw = True
    
    Return
    
    
    
Get_FunctionText:
    
    Dim sCodeky     As String * 8
    
    panelText.Visible = True
    panelText.ZOrder 0
    lstText.Clear
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '19'"
    strSql = strSql & " AND    Codeky Like '" & Left(cmbSLip.Text, 2) & "%'"
    strSql = strSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    lstText.Clear
    Do Until adoSet.EOF
        If Trim(adoSet.Fields("Codenm").Value & "") <> "" Then
            lstText.AddItem Trim(adoSet.Fields("Codenm").Value & "")
        End If
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    
    Select Case ButtonMenu.Index
        Case 1:  frmMicroEnrol.Show
                 frmMicroEnrol.ZOrder 0
        Case 2:  frmMEnrol.Show
                 frmMEnrol.ZOrder 0
        Case 3:  frmSheetStool.Show
                 frmSheetStool.ZOrder 0
        Case 4:  frmSheetStool2.Show
                 frmSheetStool2.ZOrder 0
        Case 5:  GoSub Set_GramStain_ResultData
                 panelGs.Visible = True
                 panelGs.ZOrder 0
                 
        Case Else:
    End Select
    Exit Sub
    
Set_GramStain_ResultData:
    lstitem.Clear
    
    lstitem.AddItem "WBC"
    lstitem.AddItem "Epithelial Cells"
    lstitem.AddItem "Gram(+)"
    lstitem.AddItem "Gram(-)"
    lstitem.AddItem "Other"
    
    lstitem.AddItem "Budding Yeast like cell, pseudohyphae"
    lstitem.AddItem "Clue Cell"
    lstitem.AddItem "not found"
    
    Return
    
    
End Sub

Private Sub txtAspiration_Change()
    txtLength1.Text = Len(txtAspiration.Text) & "/500"

End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    Dim iSLipno2        As Integer
    
    
    If KeyCode = vbKeyReturn Then
        txtBarCode.Tag = txtBarCode.Text
        GoSub SetForm_Clear
        txtBarCode.Text = txtBarCode.Tag
        
        Select Case Len(Trim(txtBarCode.Text))
            Case 12
                sJeobsuDt = convLabnoToExpand(Left(txtBarCode.Text, 5))
                iSLipno1 = Val(Mid(txtBarCode.Text, 6, 2))
                iSLipno2 = Val(Mid(txtBarCode.Text, 8, 5))
            Case 15
                sJeobsuDt = Left(txtBarCode.Text, 8)
                iSLipno1 = Val(Mid(txtBarCode.Text, 9, 2))
                iSLipno2 = Val(Mid(txtBarCode.Text, 11, 5))
            'Case Is < 6
            '    txtSLipno2.Text = Format(txtBarCode.Text, "00000")
            '    Call txtSLipno2_KeyDown(vbKeyReturn, 1)
            '    Exit Sub
        End Select
        sJeobsuDt = Left(sJeobsuDt, 4) & "-" & Mid(sJeobsuDt, 5, 2) & "-" & Mid(sJeobsuDt, 7, 2)
        
        dtJeobsu.Value = Format(sJeobsuDt, "yyyy-MM-dd")
        Call SetComboBox(cmbSLip, iSLipno1, 2)
        txtSLipno2.Text = iSLipno2
        Call txtSLipno2_KeyDown(vbKeyReturn, 1)
    End If
    Exit Sub
    
    
    
    
    
SetForm_Clear:
    mdiMain.stbMain.Panels(1).Text = ""
    cmdAll.Caption = "선택All"
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
    
    txtMsample.Text = ""
    
    
    chkWhere(0).Value = "1"
    chkWhere(1).Value = "0"
    chkWhere(2).Value = "0"
    chkWhere(3).Value = "0"
    chkWhere(4).Value = "0"
    cmbWhere.ListIndex = 2
    
    lstMicroList.Clear
    'dtJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Call SpreadSetClear(sprDiffc)
    
    txtSLipno2.SetFocus

    Return
End Sub

Private Sub txtGeomsaCm_Change()
    
    txtCount.Text = GetWindowTextLength(txtGeomsaCm.hwnd) & " / 600"
    
End Sub

Private Sub txtMsample_DblClick()
    
    frmQrySample.Show vbModal
    If Trim(txtMsample.Text) <> "" Then
         frmQrySample.Show vbModal

    End If
    
End Sub

Private Sub txtMSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtMSeq.Text) = "" Then Exit Sub
        gSMicroCheck = "Micro"
        GoSub Micro_Data_Select
    End If
    gSMicroCheck = ""
    Exit Sub
    
    
Micro_Data_Select:
    Dim sMicroDate      As String
    
    sMicroDate = Format(dtJeobsu.Value, "yyyy-MM")
    
    strSql = ""
    strSql = strSql & " SELECT a.*, TO_CHAR(JeobsuDt,'yyyy-MM-dd') JeobsuDt"
    strSql = strSql & " FROM   TWEXAM_General_Sub a"
    strSql = strSql & " WHERE  TO_CHAR(a.MDate,'yyyy-MM') = '" & sMicroDate & "'"
    strSql = strSql & " AND    a.MSeq  = " & Val(txtMSeq.Text)
    strSql = strSql & " Order  By a.MDate Desc"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    
    txtMSeq.Tag = txtMSeq.Text
    GiExamNumb = Val(adoSet.Fields("SLipno1").Value & "")
    Call SetComboBox(cmbSLip, GiExamNumb, 2)
    txtMSeq.Text = txtMSeq.Tag
    txtSLipno2.Text = Format(adoSet.Fields("SLipno2").Value & "")
    dtJeobsu.Value = Format(adoSet.Fields("JeobsuDt").Value & "", "yyyy-MM-dd")
    Call adoSetClose(adoSet)
    
    Call txtSLipno2_KeyDown(vbKeyReturn, 1)
    
    Return

End Sub

Private Sub txtPtno_DblClick()
    
    frmWhere.Show vbModal
    
End Sub

Private Sub txtSLipno2_GotFocus()
    
    txtSLipno2.SelStart = 0
    txtSLipno2.SelLength = Len(txtSLipno2.Text)
    mdiMain.stbMain.Panels(1).Text = "Labno 를 입력하고 Enter Key 를 치십시오!.."
    
    
End Sub

Public Sub txtSLipno2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    Dim iMinCham        As Integer
    Dim iMaxCham        As Integer
    Dim sItemCd         As String
    Dim nRetAge         As Integer
    Dim sRetSex         As String
    
    Dim iResult         As Long
    Dim iPanicMin       As Long
    Dim iPanicMax       As Long

        
    If KeyCode = vbKeyReturn Then
        
        Screen.MousePointer = vbHourglass
        mdiMain.stbMain.Panels(1).Text = ""
        cmdAll.Caption = "선택All"

        
        DoEvents: GoSub Check_SLipno             'Accept SLipno2 Check
        DoEvents: GoSub Get_HJData_Select        '환자기본 ID & 검사 Comment
        If Left(cmbSLip.Text, 2) = "15" Then
            GoSub GETData_BoneMarrow
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        DoEvents: GoSub Set_Clear_Reset          'Spread REset
        
        DoEvents: GoSub Get_ResultData           '실제 결과 Data
        
            
        If Left(cmbSLip.Text, 1) = "4" Then
            sprSLip.Row = 0
            sprSLip.Col = 3: sprSLip.Text = "MSeq"
            sprSLip.Col = 4: sprSLip.Text = "검체코드"
            sprSLip.Col = 5: sprSLip.Text = "검체명"
            DoEvents: GoSub Select_Sample      'SampleData Select
        Else
            sprSLip.Row = 0
            sprSLip.Col = 3: sprSLip.Text = "Prev"
            sprSLip.Col = 4: sprSLip.Text = "Min"
            sprSLip.Col = 5: sprSLip.Text = "Max"
            DoEvents: GoSub READ_PREVIOUS_DATA       '이전 결과 Data ( Delta Check용)
        End If
        
        DoEvents: GoSub Special_Code_Check
        If sprSLip.DataRowCnt > 0 Then
            sprSLip.Row = 1
            sprSLip.Col = 2: sprSLip.Action = ActionActiveCell
        End If
        gSMicroCheck = ""   '만약에 Micro 에서 결과라면 초기화시킴
        Screen.MousePointer = vbDefault
    End If
    
    Exit Sub
    
    
    
    
Micro_Data_Select:
    Dim sMicroDate      As String
    
    sMicroDate = Format(dtJeobsu.Value, "yyyy-MM")
    
    strSql = ""
    strSql = strSql & " SELECT a.*, TO_CHAR(JeobsuDt,'yyyy-MM-dd') JeobsuDt"
    strSql = strSql & " FROM   TWEXAM_General_Sub a"
    strSql = strSql & " WHERE  TO_CHAR(a.MDate,'yyyy-MM') = '" & sMicroDate & "'"
    strSql = strSql & " AND    a.MSeq  = " & Val(txtSLipno2.Text)
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    GiExamNumb = Val(adoSet.Fields("SLipno1").Value & "")
    Call SetComboBox(cmbSLip, GiExamNumb, 2)
    txtMSeq.Text = txtSLipno2.Text
    txtSLipno2.Text = Format(adoSet.Fields("SLipno2").Value & "")
    dtJeobsu.Value = Format(adoSet.Fields("JeobsuDt").Value & "", "yyyy-MM-dd")
    Call adoSetClose(adoSet)
    
    Return
    
    
    
Set_Clear_Reset:
    sprSLip.ReDraw = False
        
    sprSLip.MaxRows = 0
    sprSLip.MaxRows = 120
    sprSLip.RowHeight(-1) = 11
    
    sprSLip.ReDraw = True
    Return
    

Get_HJData_Select:
'O  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, e.DeptNamek, f.Drname, b.AgeYY age,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt"
    strSql = strSql & " FROM   TWEXAM_General  a,"
    strSql = strSql & "        TWEXAM_IDNOMST  b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      e,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    f "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Slipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SlipNo2  = " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    a.GBCh     = 'Y'"
    strSql = strSql & " AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode = e.Deptcode(+)"
    strSql = strSql & " AND    a.Drcode   = f.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub LOCal_Clear_Screen
        Screen.MousePointer = vbDefault
        Exit Sub
        Return
    End If
    
    
    dtJeobsu.Tag = adoSet.Fields("JeobsuDt").Value & ""
    txtPtno.Text = adoSet.Fields("Ptno").Value & ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtAge.Text = adoSet.Fields("age").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtDept.Text = adoSet.Fields("DeptNamek").Value & ""
    txtDr.Text = adoSet.Fields("Drname").Value & ""
    Me.txtGeomsaCm.Text = adoSet.Fields("GeomsaCM").Value & ""
    
    If adoSet.Fields("GBIO").Value & "" = "I" Then
        txtRoom.Text = adoSet.Fields("RoomCode").Value & ""
    Else
        txtRoom.Text = "OPD"
    End If
    
    
    txtStatus.Text = ""
    Select Case Trim(adoSet.Fields("Status").Value & "")
        Case "R": txtStatus.Text = "접수중"
        Case "P": txtStatus.Text = "부분결과"
        Case "U": txtStatus.Text = "미확인"
        Case "C": txtStatus.Text = "결과완료"
        Case "X": txtStatus.Text = "이상Data"
    End Select
    
    Call adoSetClose(adoSet)
    
    Return


Check_SLipno:
    txtSLipno2.Text = Format(txtSLipno2.Text, "00000")
    
    If IsNumeric(txtSLipno2.Text) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    sJeobsuDt = Format(dtJeobsu.Value, "yyyy-MM-dd")
    
    If cmbSLip.ListIndex > -1 Then iSLipno1 = Val(Left(cmbSLip.Text, 2))
    Return
    
    
    
Get_ResultData:
    strSql = ""
    strSql = strSql & "  SELECT c.m_min, c.m_max, c.f_min, c.f_max, "
    strSql = strSql & "         a.Rowid RWID, a.SLipno1, a.SLipno2, a.ItemCD, a.Result1, a.Result2, a.Verify,"
    strSql = strSql & "         a.Orderno, a.Routincd,  b.ItemNM,"
    strSql = strSql & "         TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "         b.Resultw, b.DeltaQC, b.DeltaMin, b.DeltaMax, b.Danwi, b.PanicMin, b.PanicMax,"
    strSql = strSql & "         a.MSeq"
    strSql = strSql & "  FROM   TWEXAM_GENERAL_SUB  a,"
    strSql = strSql & "         TWEXAM_ITEMML       b,"
    strSql = strSql & "         TWEXAM_REFDATA      c "
    strSql = strSql & "  WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.Slipno1  = " & iSLipno1
    strSql = strSql & "  AND    a.SlipNo2  = " & Val(txtSLipno2.Text)
    
    If Left(cmbSLip.Text, 1) = "4" Then
        If gSMicroCheck = "Micro" Then
            strSql = strSql & "  AND    ( a.MSeq     = " & Val(Me.txtMSeq.Text)
            strSql = strSql & "  OR       a.MSeq    IS NOT NULL )"
        End If
    End If
    
    strSql = strSql & "  AND    a.ItemCD   =  b.Codeky(+) "
    strSql = strSql & "  AND    a.ItemCD   =  c.ItemCode(+)"
    strSql = strSql & "  AND    a.AgeYY   >=  c.AgeMin(+)"
    strSql = strSql & "  AND    a.AgeYY   <=  c.AgeMax(+)"
    strSql = strSql & "  AND    NVL(c.appdate,SysDate) = "
    strSql = strSql & "            (Select NVL(MAX(APPDATE), SysDate)"
    strSql = strSql & "             From   TWEXAM_REFDATA d"
    strSql = strSql & "             Where  d.ItemCode = a.ItemCD"
    strSql = strSql & "             And    d.AgeMin  <= a.AgeYY"
    strSql = strSql & "             And    d.AgeMax  >= a.AgeYY)"
    strSql = strSql & "  ORDER  BY  a.ItemCd  "

    If False = adoSetOpen(strSql, adoSet) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Do Until adoSet.EOF
        sprSLip.Row = sprSLip.DataRowCnt + 1
        sprSLip.Col = 1:  sprSLip.Text = adoSet.Fields("ItemNM").Value & ""
        
        If Val(adoSet.Fields("SLipno1").Value & "") = 31 Then
            sprSLip.Col = 2:  sprSLip.Text = Get_CutOFFData(Trim(adoSet.Fields("ItemCd").Value & ""), _
                                                            Trim(adoSet.Fields("Result1").Value & ""))
        Else
            sprSLip.Col = 2:  sprSLip.Text = convResultFormat(Trim(adoSet.Fields("Result1").Value & ""))
        End If
        
        If Trim(txtSex.Text) = "M" Then
            sRefDataMin = Trim(adoSet.Fields("M_MIN").Value & "")
            sRefDataMax = Trim(adoSet.Fields("M_MAX").Value & "")
        Else
            sRefDataMin = Trim(adoSet.Fields("F_MIN").Value & "")
            sRefDataMax = Trim(adoSet.Fields("F_MAX").Value & "")
        End If
        
        If gSMicroCheck = "Micro" Then
            sprSLip.Col = 3: sprSLip.Text = adoSet.Fields("MSeq").Value & ""
        Else
            sprSLip.Col = 3:  '이전 결과 Data (READ_PREVIOUS_DATA)
        End If
        sprSLip.Col = 4:  sprSLip.Text = convResultFormat(sRefDataMin)
                          If Val(sprSLip.Text) = 0 Then sprSLip.Text = ""
        sprSLip.Col = 5:  sprSLip.Text = convResultFormat(sRefDataMax)
                          If Val(sprSLip.Text) = 0 Then sprSLip.Text = ""

        sprSLip.Col = 6:  sprSLip.Text = adoSet.Fields("Danwi").Value & ""
        
        If Trim(adoSet.Fields("Verify").Value & "") = "Y" Then
            sprSLip.Col = 7: sprSLip.Text = "1"
        End If
        
        sprSLip.Col = 8:  'Panic  Data Display
        sprSLip.Col = 9:  'Delta  Data Display
        sprSLip.Col = 10: '접수취소시 Flag
        
        sprSLip.Col = 11: sprSLip.Text = adoSet.Fields("ItemCd").Value & ""
        GoSub RET_Setting_Init1
        
        sprSLip.Col = 12: sprSLip.Text = adoSet.Fields("RWID").Value & ""
        sprSLip.Col = 13: sprSLip.Text = adoSet.Fields("DeltaQC").Value & ""
        sprSLip.Col = 14: sprSLip.Text = adoSet.Fields("DeltaMin").Value & ""
        sprSLip.Col = 15: sprSLip.Text = adoSet.Fields("DeltaMax").Value & ""
        sprSLip.Col = 16: sprSLip.Text = adoSet.Fields("Orderno").Value & ""
        sprSLip.Col = 17: sprSLip.Text = adoSet.Fields("ResultW").Value & ""
        
        iPanicMin = Val(adoSet.Fields("PanicMin").Value & "")
        iPanicMax = Val(adoSet.Fields("PanicMax").Value & "")
        
        
        '---------------------------'
        '  PANIC VALUE CHECK        '
        '---------------------------'
        
        If Trim(adoSet.Fields("ResultW").Value & "") = "N" Then    'Numeric Result
            iResult = Val(adoSet.Fields("Result1").Value & "")
            If iResult > 0 Then                                    '결과 DATA 가 있을때...
                If iPanicMin = 0 And iPanicMax = 0 Then            'Panic 기초Data 가 없을때 Check 안함...
                    sprSLip.Col = 8:    sprSLip.Text = ""
                    sprSLip.BackColor = RGB(235, 245, 235)
                Else
                    If (iResult < iPanicMin) Or (iResult > iPanicMax) Then      'Panic...
                        sprSLip.Col = 8:    sprSLip.Text = " "
                        sprSLip.BackColor = RGB(250, 250, 0)
                        sprSLip.Col = 18: sprSLip.Text = "P"
                    Else
                        sprSLip.Col = 8:    sprSLip.Text = ""
                        sprSLip.BackColor = RGB(235, 245, 235)
                    End If
                End If
            End If
        End If
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    Return
    
RET_Setting_Init1:
    Dim sResult     As String
    
    sprSLip.Col = 2:
    sprSLip.CellType = SS_CELL_TYPE_EDIT
    sprSLip.TypeHAlign = SS_CELL_H_ALIGN_LEFT
    sprSLip.TypeEditCharCase = SS_CELL_EDIT_CASE_NO_CASE
    sprSLip.TypeEditMultiLine = False
    sprSLip.TypeEditLen = 50

    sprSLip.Col = 11: sItemCd = sprSLip.Text
    sResult = Get_Result_Text(sItemCd)
    If Trim(sResult) <> "" Then
        sprSLip.Col = 2
        sprSLip.CellType = CellTypeComboBox
        sprSLip.TypeComboBoxList = sResult
        sprSLip.TypeComboBoxEditable = True
    End If
    
    Return



READ_PREVIOUS_DATA:
    Dim nPreDateCnt     As Integer
    Dim sPreDate        As String
    Dim LiCurVal        As Long
    Dim LiPreVal        As Long
    Dim LsQC            As String
    Dim LiDeltaMin      As Long
    Dim LiDeltaMax      As Long
    Dim adoPRet         As ADODB.Recordset
    
    
    nPreDateCnt = 0
    sPreDate = ""
    
    For i = 1 To sprSLip.DataRowCnt
        sprSLip.Row = i
        sprSLip.Col = 11: sItemCd = sprSLip.Text

'/-----------------------------------------------------------------------------------------------
' 당일 포함하여 가장 최근의 Data 를 Select 함
        strSql = ""
        strSql = strSql & " SELECT TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt,"
        strSql = strSql & "        a.Result1, a.Slipno2, a.ItemCD"
        strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a"
        strSql = strSql & " WHERE  a.Ptno      =  '" & txtPtno.Text & "'"
        strSql = strSql & " AND    a.ItemCD    =  '" & sItemCd & "'"
        strSql = strSql & " AND    a.Verify    =  'Y'"
        strSql = strSql & " AND    TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') || a.SLipno2  ="
        strSql = strSql & "                      ( SELECT Max(TO_CHAR(b.JeobsuDt,'yyyy-MM-dd') || b.SLipno2)"
        strSql = strSql & "                        FROM   TWEXAM_GENERAL_SUB b"
        strSql = strSql & "                        WHERE  TO_CHAR(b.JeobsuDt,'YYYY-MM-DD') || TO_CHAR(b.SLipno2,'00000') < "
        strSql = strSql & "                               '" & sJeobsuDt & "' || ' ' || '" & txtSLipno2.Text & "'"
        strSql = strSql & "                        AND    b.Ptno     =  '" & txtPtno.Text & "'"
        strSql = strSql & "                        AND    b.ItemCD   = '" & sItemCd & "'"
        strSql = strSql & "                        AND    b.Verify   = 'Y')"
        strSql = strSql & " ORDER  BY a.JeobsuDt DESC, a.SLipno2 DESC"
        
        'jeobsudt || slipno2 < sjeobsudt || slipno2
        
        If adoSetOpen(strSql, adoPRet) Then
            sprSLip.Col = 3: sprSLip.Text = Trim(adoPRet.Fields("Result1").Value & "")
            nPreDateCnt = nPreDateCnt + 1
            sPreDate = adoPRet.Fields("JeobsuDt").Value & ""
            Call adoSetClose(adoPRet)
        End If

        
        '---------------------------'
        '  DELTA VALUE CHECK        '
        '---------------------------'
        
        sprSLip.Col = 9:    sprSLip.Text = ""
        sprSLip.BackColor = RGB(235, 245, 235)

        sprSLip.Col = 2:    LiCurVal = Val(sprSLip.Text)
        sprSLip.Col = 3:    LiPreVal = Val(sprSLip.Text)
        sprSLip.Col = 13:   LsQC = Trim(sprSLip.Text)
        sprSLip.Col = 14:   LiDeltaMin = Val(sprSLip.Text)
        sprSLip.Col = 15:   LiDeltaMax = Val(sprSLip.Text)


        If LiPreVal <> 0 And LiCurVal <> 0 Then       '양쪽(계산할 2개의 Data가 모두 있을때...)
            If LsQC = "1" Or LsQC = "2" Or LsQC = "3" Or LsQC = "4" Then  'QC 방법이 ItemML 에 Setting 되어있는것
                If DeltaCheck(LiCurVal, LiPreVal, LsQC) < LiDeltaMin Or _
                   DeltaCheck(LiCurVal, LiPreVal, LsQC) > LiDeltaMax Then
                    sprSLip.Col = 9
                    If Trim(sprSLip.Text) = "" Then
                        sprSLip.Text = "D"
                        sprSLip.BackColor = RGB(250, 0, 0)
                    End If
                End If
            End If
        End If
    Next i
    
    If nPreDateCnt > 0 Then
        sprSLip.Row = 0
        sprSLip.Col = 3
        sprSLip.Text = "Prev" & "[" & sPreDate & "]"
    Else
        sprSLip.Text = ""
        sprSLip.Row = 0
        sprSLip.Col = 3:  sprSLip.Text = "Prev"
    End If
    
    Return
    
    
    
Special_Code_Check:
    Dim sText22     As String * 22
    
    For i = 1 To sprSLip.DataRowCnt
        sprSLip.Row = i
        sprSLip.Col = 17
        Select Case Trim(sprSLip.Text)
            Case "S": GoSub ISRcode_Check_Sub                   'Culture & Sensitivity 검사 Check
                      sprSLip.Col = 1
                      sprSLip.CellType = CellTypeButton
                      sprSLip.CellBorderColor = RGB(192, 192, 192)
                      sprSLip.ShadowColor = RGB(192, 192, 192)
                      sprSLip.TypeButtonLightColor = RGB(192, 192, 192)
                      sprSLip.CellBorderStyle = CellBorderStyleSolid
                      sText22 = sprSLip.Text
                      sprSLip.TypeButtonText = sText22
                      sprSLip.Lock = False
                      sprSLip.AllowCellOverflow = True
                      GoSub Default_Data_Check
            Case Else
        End Select
    Next
    
    Return
    
Default_Data_Check:
    sprSLip.Col = 2
    'If Trim(sprSLip.Text) = "" Then
    '    If InStr(1, UCase(txtSamplename.Text), "STOOL", vbTextCompare) > 0 Then
    '        sprSLip.Col = 2
    '        sprSLip.Text = "No growth for Salmonella,Shigella,Vibrio,Campylobacter"
    '    ElseIf InStr(1, UCase(txtSamplename.Text), "URINE", vbTextCompare) > 0 Then
    '        sprSLip.Col = 2
    '        sprSLip.Text = "less then 1000 CFU/mL"
    '    End If
    'End If
    Return
    
    
ISRcode_Check_Sub:
    Dim nCount      As Integer
    Dim sSensItemCd As String
    
    
    sprSLip.Col = 11: sSensItemCd = sprSLip.Text
    
    strSql = ""
    strSql = strSql & " SELECT Rcode1,Rcode2,Rcode3,Rcode4,Rcode5, Result1"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    Slipno1  =  " & Val(iSLipno1)
    strSql = strSql & " AND    Slipno2  =  " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    ItemCd   = '" & sSensItemCd & "'"
    
    nCount = 0
    If adoSetOpen(strSql, adoSet) Then
    
        If Trim(adoSet.Fields("Rcode1").Value & "") <> "" Then
            nCount = nCount + 1: End If
        If Trim(adoSet.Fields("Rcode2").Value & "") <> "" Then
            nCount = nCount + 1: End If
        If Trim(adoSet.Fields("Rcode3").Value & "") <> "" Then
            nCount = nCount + 1: End If
        If Trim(adoSet.Fields("Rcode4").Value & "") <> "" Then
            nCount = nCount + 1: End If
        If Trim(adoSet.Fields("Rcode5").Value & "") <> "" Then
            nCount = nCount + 1: End If
        
        sprSLip.Col = 2: sprSLip.Text = nCount
        
        If nCount = 0 Then
            If Trim(adoSet.Fields("Result1").Value & "") <> "" Then
                sprSLip.Col = 2: sprSLip.Text = Trim(adoSet.Fields("Result1").Value & "")
            End If
        End If
        Call adoSetClose(adoSet)
    End If
    Return
    
    
GETData_BoneMarrow:
    sJeobsuDt = Format(frmResult.dtJeobsu.Value, "yyyy-MM-dd")
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.ItemNM,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TWEXAM_ITEMML      b "
    strSql = strSql & " WHERE    a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  =  " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.SLipno2  =  " & Val(txtSLipno2.Text)
    strSql = strSql & " AND    a.ROUTINCD =  '150050'"        'Bone marrow RoutineCode
    strSql = strSql & " AND    a.ItemCd   =  b.Codeky(+)"
    strSql = strSql & " ORDER  BY a.iTemCd"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        
        
        Select Case Trim(adoSet.Fields("iTemCD").Value & "")
            Case "1505011":
                txtSmear.Text = adoSet.Fields("Chamgo").Value & ""
            Case "1505012":
                txtAspiration.Text = adoSet.Fields("Chamgo").Value & ""
            Case "15050201" To "15050218":
                sprDiffc.Row = Val(Mid(Trim(adoSet.Fields("ItemCd").Value & ""), 7, 2))
                sprDiffc.Col = 1: sprDiffc.Text = adoSet.Fields("ItemCd").Value & ""
                sprDiffc.Col = 2: sprDiffc.Text = adoSet.Fields("ItemNM").Value & ""
                sprDiffc.Col = 3: sprDiffc.Text = adoSet.Fields("Result2").Value & ""
                sprDiffc.Col = 4: sprDiffc.Text = adoSet.Fields("Result1").Value & ""
            Case Else
        End Select
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return
    
    
Select_Sample:
    Dim sCompCode       As String
    Dim sRowID          As String
    Dim iRow            As Integer
    
    Dim sSampleC        As String
    Dim sSampleN        As String
    Dim sTempSample     As String
    
    sSampleC = "": sSampleN = ""
    
    
    For iRow = 1 To Me.sprSLip.DataRowCnt
        sprSLip.Row = iRow
        sprSLip.Col = 11: sCompCode = Trim(sprSLip.Text)
        sprSLip.Col = 12: sRowID = sprSLip.Text
        
        
        strSql = ""
        strSql = strSql & " SELECT a.*, b.Code, b.Codenm, b.Class2 "
        strSql = strSql & " FROM   TWEXAM_General_Sub a,"
        strSql = strSql & "        TWEXAM_Sample      b "
        strSql = strSql & " WHERE  a.ROWID    = '" & sRowID & "'"
        strSql = strSql & " AND    a.GeomchCd = b.Code(+)"
        If False = adoSetOpen(strSql, adoSet) Then Return
        
        If sTempSample <> Trim(adoSet.Fields("Code").Value & "") Then
            sprSLip.Col = 4: sprSLip.Text = Trim(adoSet.Fields("Code").Value & "")
            sprSLip.Col = 5: sprSLip.Text = Trim(adoSet.Fields("Codenm").Value & "")
        End If
        
        sTempSample = Trim(adoSet.Fields("Code").Value & "")
        Call adoSetClose(adoSet)
        
        If Trim(sTempSample) = "M2308" Or Trim(sTempSample) = "M2804" Then  'Other
            GoSub Get_Cmdoctor
        End If
    Next
        
    Return
    
Get_Cmdoctor:
    Dim lngOrderno      As Long
    Dim sTempName       As String
    
    sprSLip.Row = iRow
    sprSLip.Col = 16: lngOrderno = Val(sprSLip.Text)
    
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Order"
    strSql = strSql & " WHERE  Ptno = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    Orderno = " & lngOrderno
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    If Trim(adoSet.Fields("CMDoctor").Value & "") <> "" Then
        If sTempName <> Trim(adoSet.Fields("CMDoctor").Value & "") Then
            sprSLip.Col = 5: sprSLip.Text = Trim(adoSet.Fields("CMDoctor").Value & "")
        End If
    End If
    
    sTempName = Trim(adoSet.Fields("CMDoctor").Value & "")
    Call adoSetClose(adoSet)
    
    Return
    
    
LOCal_Clear_Screen:

    'Bone marrow Screen
    txtLength.Text = ""
    txtSmear.Text = ""
    txtLength1.Text = ""
    txtAspiration.Text = ""
    Call SpreadSetClear(sprDiffc)
    
    'Normal Screen
    txtPtno.Text = ""
    txtSname.Text = ""
    txtSex.Text = ""
    txtAge.Text = ""
    txtDept.Text = ""
    txtDr.Text = ""
    txtRoom.Text = ""
    txtStatus.Text = ""
    
    sprSLip.ReDraw = False
    sprSLip.MaxRows = 0
    sprSLip.MaxRows = 60
    sprSLip.RowHeight(-1) = 10.91
    sprSLip.ReDraw = True
    
    'txtSLipno2.SetFocus
    Return
    
End Sub

Private Sub txtSLipno2_LostFocus()
    
    mdiMain.stbMain.Panels(1).Text = ""
    
End Sub

Private Sub txtSmear_Change()
        txtLength.Text = Len(txtSmear.Text) & "/500"

End Sub
