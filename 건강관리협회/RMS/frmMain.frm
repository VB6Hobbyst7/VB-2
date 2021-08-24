VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "중앙검사센터 검사의뢰시스템"
   ClientHeight    =   11775
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17460
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  '소유자 가운데
   Begin Threed.SSPanel SSPanel9 
      Align           =   1  '위 맞춤
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17460
      _ExtentX        =   30798
      _ExtentY        =   1032
      _Version        =   262144
      BackColor       =   16777215
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   30
         Picture         =   "frmMain.frx":1272
         ScaleHeight     =   435
         ScaleWidth      =   4545
         TabIndex        =   9
         Top             =   90
         Width           =   4545
         Begin VB.Label lblMenu 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00794444&
            Height          =   315
            Left            =   1110
            TabIndex        =   10
            Top             =   120
            Width           =   3345
         End
      End
      Begin VB.PictureBox picNode 
         BackColor       =   &H00FFFFFF&
         Height          =   14505
         Left            =   0
         ScaleHeight     =   14445
         ScaleWidth      =   3870
         TabIndex        =   6
         Top             =   585
         Width           =   3930
         Begin VB.CommandButton cmdNode 
            Caption         =   "◀"
            Height          =   14445
            Left            =   3570
            TabIndex        =   7
            Top             =   0
            Width           =   315
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   14445
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   25479
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlSubList(1)"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   0
         Left            =   4680
         TabIndex        =   1
         Top             =   30
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "검사접수"
         PictureAlignment=   1
      End
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   1
         Left            =   6900
         TabIndex        =   2
         Top             =   30
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "검사결과"
      End
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   2
         Left            =   9120
         TabIndex        =   3
         Top             =   30
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "처리내역"
      End
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   3
         Left            =   11340
         TabIndex        =   4
         Top             =   30
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "검사편람"
      End
      Begin XLibrary_XButton.XButton cmdClose 
         Height          =   435
         Left            =   16170
         TabIndex        =   5
         Top             =   30
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   767
         BackColor1      =   16777215
         BackColor2      =   16777215
         BackColorEx     =   14737632
         BackGradientStyle=   2
         BackStyle       =   4
         BevelHeight     =   5
         BackGradientExPercent=   80
         BackGlassColorStyle=   1
         BackGradientAutoValue=   40
         BackGlassAutoValue=   70
         BackLightShadowShadowValue=   -30
         BackLightShadowLightValue=   30
         BorderStyle     =   0
         BorderWidth     =   1
         BorderColor     =   16744576
         MaskColor       =   13828096
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "종료"
         TextWidthPos    =   2
         TextHeightPos   =   2
         TextWidthMargin =   5
         TextHeightMargin=   5
         TextColor       =   128
         IconPosition    =   2
         IconAndTextMargin=   0
         IconMaskColor   =   13828096
         MouseOverMargin =   2
         MouseOverEffectAutoValue=   -20
         MouseDownBorderEffectValue=   -40
         MouseDownDefaultValue=   20
         FocusDefaultMargin=   3
         FocusColor1     =   16777152
         FocusColor2     =   16777088
         FocusColorStyle =   1
         FocusColorMargin=   2
         FocusEffectAutoValue=   -20
         ToolTipBodyText =   "XBUTTON 2"
         ToolTipTitleText=   ""
         ToolTipCentered =   -1  'True
         ToolTipBackColor=   12648447
         ToolTipExBackColor1=   12648447
         ToolTipExHoverTime=   1000
         ToolTipExPopupTime=   10000
         ToolTipExPopupPos=   0
         ToolTipExArrowWidth=   10
         ToolTipExArrowHeight=   15
         ToolTipExBorderRoundNum=   0
         ToolTipExPopupPosWMargin=   5
         ToolTipExPopupPosHMargin=   5
         ToolTipExBackColor2=   16777215
         ToolTipExBorderColor=   4210752
         ToolTipExTitleText=   "Title"
         ToolTipExIconAndTitleMargin=   5
         ToolTipExTitleAlign=   2
         BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTopMargin=   5
         ToolTipExBottomMargin=   5
         ToolTipExLeftMargin=   5
         ToolTipExRightMargin=   5
         ToolTipExBodyText=   "Body Text"
         ToolTipExBodyTextColor=   4210752
         BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTitleLineColor=   4210752
         ToolTipExTitleAndLineMargin=   5
         ToolTipExPostScriptText=   "PostScript"
         ToolTipExIconAndPostScriptMargin=   5
         ToolTipExPostScriptLineColor=   4210752
         BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTitleLineShadow=   -1  'True
         ToolTipExTitleLine=   -1  'True
         ToolTipExTitleLineLeftMargin=   5
         ToolTipExTitleLineRightMargin=   5
         ToolTipExPostScriptLineShadow=   -1  'True
         ToolTipExPostScriptLine=   -1  'True
         ToolTipExPostScriptLineLeftMargin=   5
         ToolTipExPostScriptLineRightMargin=   5
         ToolTipExTitleAndBodyMargin=   5
         ToolTipExBodyAndPostScriptMargin=   5
         ToolTipExTitleTextBackColor=   16777215
         ToolTipExTitleIconMaskColor=   13828096
         ToolTipExTitleIconAndTextAlign=   2
         ToolTipExTitleIconAndTextMargin=   5
         ToolTipExPopupAutoPos=   -1  'True
         ToolTipExPostScriptAndLineMargin=   5
         ToolTipExPostScriptIconPos=   1
         ToolTipExPostScriptIconAndTextMargin=   5
         ToolTipExPostScriptIconAndTextAlign=   2
         ToolTipExPostScriptIconMaskColor=   13828096
         ToolTipExBodyTextBackColor=   16761024
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   435
      Left            =   0
      TabIndex        =   11
      Top             =   11340
      Width           =   17460
      _ExtentX        =   30798
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3810
            MinWidth        =   3809
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14994
            MinWidth        =   14994
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "2015-04-29"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오전 11:38"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "한국건강관리협회"
            TextSave        =   "한국건강관리협회"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu munMenu01 
      Caption         =   "메뉴"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 해당 폼을 띄운다.
Sub ShowForm(ByVal frmThis As Form, ByVal strFrmNm As String)
    
    Screen.MousePointer = vbHourglass
    If frmThis.MDIChild = True Then
        
        frmThis.Show
        frmThis.ZOrder 0
        lblMenu.Caption = strFrmNm
    Else
        frmThis.Show , Me
    End If
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub ssMenu_Click(Index As Integer, Value As Integer)
    
    
    
    If Index = 0 Then
        ssMenu(0).BackColor = &H80000013
        ssMenu(1).BackColor = &H80000018
        ssMenu(2).BackColor = &H80000018
        ssMenu(3).BackColor = &H80000018
        
        Call ShowForm(frmRMSReg, frmRMSReg.Caption)
        
    ElseIf Index = 1 Then
        ssMenu(0).BackColor = &H80000018
        ssMenu(1).BackColor = &H80000013
        ssMenu(2).BackColor = &H80000018
        ssMenu(3).BackColor = &H80000018
    
        Call ShowForm(frmRMSRst, frmRMSRst.Caption)
    
    ElseIf Index = 2 Then
        ssMenu(0).BackColor = &H80000018
        ssMenu(1).BackColor = &H80000018
        ssMenu(2).BackColor = &H80000013
        ssMenu(3).BackColor = &H80000018
    
        Call ShowForm(frmRMSRpt, frmRMSRpt.Caption)
    
    ElseIf Index = 3 Then
        ssMenu(0).BackColor = &H80000018
        ssMenu(1).BackColor = &H80000018
        ssMenu(2).BackColor = &H80000018
        ssMenu(3).BackColor = &H80000013
    
        Call ShowForm(frmRMSMst, frmRMSRpt.Caption)
    
    End If
    
    
    
End Sub


