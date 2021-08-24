VERSION 5.00
Object = "{38B18A4D-67F2-4F9B-B495-7ABA033953BB}#2.0#0"; "XProgressBar.ocx"
Begin VB.Form frmProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   120
   End
   Begin XLibrary_XProgressBar.XProgress XProgress1 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   609
      BackColor       =   16777215
      BorderColor     =   14737632
      BorderWidth     =   3
      ProgressColor1  =   8454016
      ProgressColor2  =   8454016
      ProgressStyle   =   0
      Min             =   0
      Max             =   100
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextFontColor   =   0
      TextFontBackColor=   0
      TextFontBackStyle=   1
      TextAlign       =   2
      TextAlignMargin =   0
      GradientStyle   =   4
      GradientPosition=   0
      BevelStyle      =   0
      BevelHeight     =   1
      PictureStyle    =   0
      BoxWidth        =   6
      BoxWidthMargin  =   1
      BoxHeightMargin =   1
      Text            =   ""
      BorderStyle     =   2
      MouseCursor     =   0
      Enabled         =   -1  'True
      rImgWidth       =   0
      rImgHeight      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "50 / 100"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   510
      Width           =   8265
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
