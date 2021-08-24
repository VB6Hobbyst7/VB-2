VERSION 5.00
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{9255B445-567E-4A7A-9DCD-987EFAE369A8}#2.0#0"; "XCheckbutton.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmAreaSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "프로그램 공용설정"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4290
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin BHButton.BHImageButton cmdSave 
      Height          =   375
      Left            =   930
      TabIndex        =   10
      Top             =   3780
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "저 장"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentPicture=   "frmAreaSet.frx":0000
      BackColor       =   12632319
      ImgOutLineSize  =   3
   End
   Begin XLibrary_XGroupBox.XGroupBox XGroupBox2 
      Height          =   375
      Left            =   1590
      Top             =   1290
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BackColor       =   16777215
      BorderColor     =   16744576
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TextPosition    =   0
      TextCustomMargin=   4
      GroupBoxStyle   =   0
      TextBarColor1   =   12757903
      TextBarStyle    =   3
      TextBarColor2   =   11767328
      TextBarSymbol   =   0   'False
      TextBarSymbolColor=   16777215
      TextBarHeightMargin=   10
      MouseCursor     =   0
      TextBarMouseCursor=   0
      IconandTextMargin=   4
      BodyColor       =   16777215
      Enabled         =   -1  'True
      Begin Threed.SSOption optChanggo 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   262144
         BackStyle       =   1
         Caption         =   "자동불출"
      End
      Begin Threed.SSOption optChanggo 
         Height          =   285
         Index           =   1
         Left            =   1350
         TabIndex        =   7
         Top             =   60
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   262144
         BackStyle       =   1
         Caption         =   "수기불출"
      End
   End
   Begin XLibrary_XGroupBox.XGroupBox XGroupBox1 
      Height          =   375
      Left            =   1590
      Top             =   1710
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BackColor       =   16777215
      BorderColor     =   16744576
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TextPosition    =   0
      TextCustomMargin=   4
      GroupBoxStyle   =   0
      TextBarColor1   =   12757903
      TextBarStyle    =   3
      TextBarColor2   =   11767328
      TextBarSymbol   =   0   'False
      TextBarSymbolColor=   16777215
      TextBarHeightMargin=   10
      MouseCursor     =   0
      TextBarMouseCursor=   0
      IconandTextMargin=   4
      BodyColor       =   16777215
      Enabled         =   -1  'True
      Begin Threed.SSOption optArea 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   60
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   262144
         BackStyle       =   1
         Caption         =   "지부"
      End
      Begin Threed.SSOption optArea 
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   5
         Top             =   60
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   262144
         BackStyle       =   1
         Caption         =   "중앙검사본부"
      End
   End
   Begin XLibrary_XTextBox.XTextBox txtCd 
      Height          =   315
      Left            =   1590
      TabIndex        =   0
      Top             =   540
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BackColor       =   16777215
      BorderColor     =   16744576
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderTextMargin=   4
      PasswordChar    =   ""
      MaxLength       =   0
      MouseCursor     =   4
      TextColor       =   0
      ToolTipOpacity  =   100
      ToolTipIcon     =   2
      ToolTipPopupTime=   -1
      ToolTipHoverTime=   -1
      ToolTipBackColor=   16777215
      ToolTipForeColor=   0
      ToolTipStyle    =   3
      ToolTipCentered =   0   'False
      ToolTipTitleText=   ""
      ToolTipBodyText =   ""
      Locked          =   0   'False
      Mask            =   0
      PromptChar      =   "_"
      WrongSound      =   0
      CustomSound     =   ""
      MaskShow        =   0   'False
      MaskColor       =   33023
      CustomMask      =   ""
      TextAlign       =   0
      Enabled         =   -1  'True
   End
   Begin XLibrary_XTextBox.XTextBox txtNm 
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   930
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BackColor       =   16777215
      BorderColor     =   16744576
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderTextMargin=   4
      PasswordChar    =   ""
      MaxLength       =   0
      MouseCursor     =   4
      TextColor       =   0
      ToolTipOpacity  =   100
      ToolTipIcon     =   2
      ToolTipPopupTime=   -1
      ToolTipHoverTime=   -1
      ToolTipBackColor=   16777215
      ToolTipForeColor=   0
      ToolTipStyle    =   3
      ToolTipCentered =   0   'False
      ToolTipTitleText=   ""
      ToolTipBodyText =   ""
      Locked          =   0   'False
      Mask            =   0
      PromptChar      =   "_"
      WrongSound      =   0
      CustomSound     =   ""
      MaskShow        =   0   'False
      MaskColor       =   33023
      CustomMask      =   ""
      TextAlign       =   0
      Enabled         =   -1  'True
   End
   Begin XLibrary_XTextBox.XTextBox txtWrtdt 
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Top             =   2910
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BackColor       =   14737632
      BorderColor     =   16744576
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderTextMargin=   4
      PasswordChar    =   ""
      MaxLength       =   0
      MouseCursor     =   4
      TextColor       =   0
      ToolTipOpacity  =   100
      ToolTipIcon     =   2
      ToolTipPopupTime=   -1
      ToolTipHoverTime=   -1
      ToolTipBackColor=   16777215
      ToolTipForeColor=   0
      ToolTipStyle    =   3
      ToolTipCentered =   0   'False
      ToolTipTitleText=   ""
      ToolTipBodyText =   ""
      Locked          =   0   'False
      Mask            =   0
      PromptChar      =   "_"
      WrongSound      =   0
      CustomSound     =   ""
      MaskShow        =   0   'False
      MaskColor       =   33023
      CustomMask      =   ""
      TextAlign       =   0
      Enabled         =   0   'False
   End
   Begin XLibrary_XTextBox.XTextBox txtModdt 
      Height          =   315
      Left            =   1590
      TabIndex        =   3
      Top             =   3300
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BackColor       =   14737632
      BorderColor     =   16744576
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderTextMargin=   4
      PasswordChar    =   ""
      MaxLength       =   0
      MouseCursor     =   4
      TextColor       =   0
      ToolTipOpacity  =   100
      ToolTipIcon     =   2
      ToolTipPopupTime=   -1
      ToolTipHoverTime=   -1
      ToolTipBackColor=   16777215
      ToolTipForeColor=   0
      ToolTipStyle    =   3
      ToolTipCentered =   0   'False
      ToolTipTitleText=   ""
      ToolTipBodyText =   ""
      Locked          =   0   'False
      Mask            =   0
      PromptChar      =   "_"
      WrongSound      =   0
      CustomSound     =   ""
      MaskShow        =   0   'False
      MaskColor       =   33023
      CustomMask      =   ""
      TextAlign       =   0
      Enabled         =   0   'False
   End
   Begin XLibrary_XTextBox.XTextBox txtPswd 
      Height          =   315
      Left            =   1590
      TabIndex        =   4
      Top             =   2550
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BackColor       =   16777215
      BorderColor     =   16744576
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderTextMargin=   4
      PasswordChar    =   "*"
      MaxLength       =   0
      MouseCursor     =   4
      TextColor       =   0
      ToolTipOpacity  =   100
      ToolTipIcon     =   2
      ToolTipPopupTime=   -1
      ToolTipHoverTime=   -1
      ToolTipBackColor=   16777215
      ToolTipForeColor=   0
      ToolTipStyle    =   3
      ToolTipCentered =   0   'False
      ToolTipTitleText=   ""
      ToolTipBodyText =   ""
      Locked          =   0   'False
      Mask            =   0
      PromptChar      =   "_"
      WrongSound      =   0
      CustomSound     =   ""
      MaskShow        =   0   'False
      MaskColor       =   33023
      CustomMask      =   ""
      TextAlign       =   0
      Enabled         =   -1  'True
   End
   Begin XLibrary_XCheckButton.XCheckButton chkEntAuto 
      Height          =   315
      Left            =   1590
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      CbBackColor1    =   15132390
      CbBorderColor   =   8409372
      CbBorderStyle   =   0
      Text            =   "ERP자동입고확인"
      TextColor       =   0
      CbTextMargin    =   4
      CbBackStyle     =   1
      CbGDirection    =   0
      CbBackColor2    =   16777215
      CheckColor      =   2203937
      CheckCustomColor=   2998317
      Value           =   0   'False
      CbOverEffect    =   -1  'True
      CbOverEffectGDtn=   0
      CbOverColor1    =   10280958
      CbOverColor2    =   3388664
      MouseCursor     =   0
      ToolTipOpacity  =   100
      ToolTipIcon     =   1
      ToolTipPopupTime=   -1
      ToolTipHoverTime=   -1
      ToolTipBackColor=   14811135
      ToolTipForeColor=   0
      ToolTipStyle    =   2
      ToolTipCentered =   -1  'True
      ToolTipTitleText=   ""
      ToolTipBodyText =   ""
      Enabled         =   -1  'True
      EnabledAutoStyle=   -1  'True
      EnCbBackColor   =   14215660
      EnCbBorderColor =   10070188
      EnCheckColor    =   10070188
      EnTextColor     =   10070188
      CheckStyle      =   0
      ControlType     =   0
      AutoSize        =   0   'False
   End
   Begin BHButton.BHImageButton cmdClose 
      Height          =   375
      Left            =   2190
      TabIndex        =   11
      Top             =   3780
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "닫 기"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentPicture=   "frmAreaSet.frx":17C2
      BackColor       =   12632319
      ImgOutLineSize  =   3
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   345
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   609
      _Version        =   262144
      BackColor       =   16249839
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "☞ 프로그램 공통 환경설정 ☜"
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Left            =   150
      TabIndex        =   14
      Top             =   930
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "지부명칭"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   315
      Left            =   150
      TabIndex        =   15
      Top             =   1350
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "창고불출"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Top             =   1770
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "사용자유형"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   315
      Left            =   150
      TabIndex        =   17
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "자동입고확인"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   315
      Left            =   150
      TabIndex        =   18
      Top             =   2550
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "관리자비밀번호"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel8 
      Height          =   315
      Left            =   150
      TabIndex        =   19
      Top             =   2910
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "등록일시"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel9 
      Height          =   315
      Left            =   150
      TabIndex        =   20
      Top             =   3300
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "수정일시"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   315
      Left            =   150
      TabIndex        =   13
      Top             =   540
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   262144
      BackColor       =   16311512
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "지부코드(ERP)"
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "frmAreaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cPis999 As clsPis999

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdSave_Click()

    With cPis999
        .areacd = Trim(txtCd.Text)
        .areanm = Trim(txtNm.Text)
        .changgofg = IIf(optChanggo(1).Value, "1", "0")
        .areatype = IIf(optArea(1).Value, "1", "0")
        .pswd = Trim(txtPswd.Text)
        .autoentfg = IIf(chkEntAuto.Value, "1", "0")
        
        If .cfSave Then
            gAreaCd = .areacd
            frmMain.stsBar.Panels(1).Text = .areanm
            gWorkArea = optArea(1).Value
            gChangGoMng = optChanggo(1).Value
            
            MsgBox "저장되었습니다.!", vbInformation
        End If
    End With

End Sub

Private Sub Form_Activate()

    If InputBox("관리자비밀번호를 입력하세요.!", "관리자확인") <> Trim(txtPswd.Text) Then
        MsgBox "비밀번호가 틀립니다.!", vbCritical
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    Set cPis999 = New clsPis999
    With cPis999
        If .cfSeek Then
            txtCd.Text = .areacd
            txtNm.Text = .areanm
            optChanggo(Val(.changgofg)).Value = True
            optArea(Val(.areatype)).Value = True
            txtWrtdt.Text = .wrtdt
            txtModdt.Text = .moddt
            txtPswd.Text = .pswd
            chkEntAuto.Value = (.autoentfg = "1")
        Else
            txtCd.Text = ""
            txtNm.Text = ""
            optChanggo(0).Value = True
            optArea(0).Value = True
            txtWrtdt.Text = ""
            txtModdt.Text = ""
            txtPswd.Text = ""
            chkEntAuto.Value = False
        End If
    End With

End Sub
