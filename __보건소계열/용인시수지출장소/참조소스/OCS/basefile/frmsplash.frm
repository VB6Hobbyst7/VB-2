VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   2145
   ClientLeft      =   2580
   ClientTop       =   3420
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2130
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   7290
      Begin VB.Label LabelTitle 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "LabelTitle"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   300
         Width           =   4785
      End
      Begin VB.Image Image1 
         Height          =   1710
         Left            =   180
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "프로그램을 로딩중입니다. 잠시만 기다리십시요...."
         Height          =   255
         Left            =   2940
         TabIndex        =   1
         Top             =   1740
         Width           =   4125
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
