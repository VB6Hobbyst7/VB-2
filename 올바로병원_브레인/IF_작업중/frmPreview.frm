VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPreview 
   BorderStyle     =   0  '없음
   Caption         =   "출력"
   ClientHeight    =   10050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9525
   Icon            =   "frmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin FPSpreadADO.fpSpreadPreview spdResultPreview 
      Height          =   8985
      Left            =   60
      TabIndex        =   4
      Top             =   990
      Width           =   9345
      _Version        =   524288
      _ExtentX        =   16484
      _ExtentY        =   15849
      _StockProps     =   96
      BorderStyle     =   1
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00BF8B59&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   9525
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      Begin HSCotrol.CButton cmdClose 
         Height          =   405
         Left            =   7770
         TabIndex        =   1
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   714
         BackColor       =   15698777
         Caption         =   " 닫    기"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmPreview.frx":1272
      End
      Begin HSCotrol.CButton cmdRsltPrint 
         Height          =   405
         Left            =   6360
         TabIndex        =   3
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   714
         BackColor       =   15698777
         Caption         =   " 출    력"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmPreview.frx":180C
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "출력 미리보기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   2
         Top             =   360
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRsltPrint_Click()
    
    With frmResult
        If .optPrtOri(0).Value = True Then
            .spdResult.PrintOrientation = PrintOrientationPortrait       '세로출력
        Else
            .spdResult.PrintOrientation = PrintOrientationLandscape      '가로출력
        End If
        .spdResult.Action = 13
    End With
                
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

'        ' Set background color to yellow, RGB(255, 255, 0)
'        spdResultPreview.GrayAreaColor = &HFFFF&
'        ' Set gray area margins to 180 twips
'        spdResultPreview.GrayAreaMarginH = 180
'        spdResultPreview.GrayAreaMarginV = 180
'        ' Show pages reflecting actual size
'        spdResultPreview.GrayAreaMarginType = GrayAreaMarginTypeActual
'        ' Show multiple pages in the control
'        spdResultPreview.PageViewType = PageViewTypeMultiplePages
    
    ' Display three pages across and two pages down
    spdResultPreview.PageMultiCntH = 3
    spdResultPreview.PageMultiCntV = 2

End Sub

