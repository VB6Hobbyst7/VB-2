VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Begin VB.Form frmRMSReg 
   BackColor       =   &H00FFFFFF&
   Caption         =   "중앙검사센터 검사접수"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18180
   Icon            =   "frmRMSReg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   18180
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel2 
      Height          =   11085
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   18465
      _ExtentX        =   32570
      _ExtentY        =   19553
      _Version        =   262144
      BackColor       =   16777215
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      FloodShowPct    =   -1  'True
      Begin VB.CheckBox chkOrder 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   180
         TabIndex        =   19
         Top             =   810
         Width           =   225
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   1191
         _Version        =   262144
         BackColor       =   -2147483629
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   1500
            TabIndex        =   2
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   21299201
            CurrentDate     =   40248
         End
         Begin XLibrary_XButton.XButton XButton15 
            Height          =   405
            Left            =   4500
            TabIndex        =   3
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "조회"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
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
         Begin XLibrary_XButton.XButton XButton16 
            Height          =   405
            Left            =   5760
            TabIndex        =   4
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "의뢰등록"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            TextColor       =   16711680
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
         Begin Threed.SSFrame SSFrame2 
            Height          =   495
            Left            =   7350
            TabIndex        =   5
            Top             =   90
            Width           =   5565
            _ExtentX        =   9816
            _ExtentY        =   873
            _Version        =   262144
            BackColor       =   -2147483629
            Begin XLibrary_XTextBox.XTextBox XTextBox7 
               Height          =   285
               Left            =   900
               TabIndex        =   6
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox8 
               Height          =   285
               Left            =   2670
               TabIndex        =   7
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox9 
               Height          =   285
               Left            =   4470
               TabIndex        =   8
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '투명
               Caption         =   "검사건수"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   3690
               TabIndex        =   11
               Top             =   180
               Width           =   720
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '투명
               Caption         =   "검체건수"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   1875
               TabIndex        =   10
               Top             =   180
               Width           =   720
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '투명
               Caption         =   "의뢰건수"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   9
               Top             =   180
               Width           =   720
            End
         End
         Begin XLibrary_XButton.XButton XButton17 
            Height          =   405
            Left            =   13080
            TabIndex        =   12
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "출력"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
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
         Begin XLibrary_XButton.XButton XButton18 
            Height          =   405
            Left            =   14160
            TabIndex        =   13
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Excel"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
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
         Begin XLibrary_XButton.XButton XButton19 
            Height          =   405
            Left            =   15690
            TabIndex        =   14
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "화면지움"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
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
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   315
            Left            =   3090
            TabIndex        =   15
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   21299201
            CurrentDate     =   40248
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2940
            TabIndex        =   17
            Top             =   210
            Width           =   60
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "접수일자"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Width           =   720
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            X1              =   15450
            X2              =   15450
            Y1              =   180
            Y2              =   510
         End
      End
      Begin FPSpreadADO.fpSpread spdRcp 
         CausesValidation=   0   'False
         Height          =   8685
         Left            =   30
         TabIndex        =   18
         Tag             =   "20001"
         Top             =   750
         Width           =   17055
         _Version        =   524288
         _ExtentX        =   30083
         _ExtentY        =   15319
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
         GrayAreaBackColor=   16777215
         MaxCols         =   15
         MaxRows         =   10
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRMSReg.frx":014A
         VisibleCols     =   10
         VisibleRows     =   10
         TextTip         =   2
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "frmRMSReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
