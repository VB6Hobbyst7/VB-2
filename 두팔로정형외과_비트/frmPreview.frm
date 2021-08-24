VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPreview 
   BorderStyle     =   1  '단일 고정
   Caption         =   "출력"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9525
   Icon            =   "frmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   9525
   StartUpPosition =   1  '소유자 가운데
   Begin FPSpreadADO.fpSpreadPreview spdResultPreview 
      Height          =   9315
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   9345
      _Version        =   524288
      _ExtentX        =   16484
      _ExtentY        =   16431
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
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9525
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      Begin HSCotrol.CButton cmdRsltPrint 
         Height          =   375
         Left            =   6930
         TabIndex        =   3
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "출력"
         ForeColor       =   16744576
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmPreview.frx":1272
      End
      Begin HSCotrol.CButton cmdClose 
         Height          =   375
         Left            =   8040
         TabIndex        =   4
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "닫기"
         ForeColor       =   0
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmPreview.frx":13CC
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "출력 미리보기"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   1
         Top             =   150
         Width           =   1140
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

'    ' Set background color to yellow, RGB(255, 255, 0)
'    spdResultPreview.GrayAreaColor = &HFFFF&
'    ' Set gray area margins to 180 twips
'    spdResultPreview.GrayAreaMarginH = 180
'    spdResultPreview.GrayAreaMarginV = 180
'    ' Show pages reflecting actual size
'    spdResultPreview.GrayAreaMarginType = GrayAreaMarginTypeActual
'    ' Show multiple pages in the control
'    spdResultPreview.PageViewType = PageViewTypeMultiplePages

    ' Display three pages across and two pages down
    spdResultPreview.PageMultiCntH = 3
    spdResultPreview.PageMultiCntV = 2

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.ScaleHeight = 0 Then Exit Sub
    picHeader.WIDTH = Me.WIDTH
    spdResultPreview.WIDTH = Me.WIDTH - 200
    spdResultPreview.HEIGHT = Me.HEIGHT - picHeader.HEIGHT - 700
End Sub
