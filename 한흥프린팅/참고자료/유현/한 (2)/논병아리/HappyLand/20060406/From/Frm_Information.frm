VERSION 5.00
Begin VB.Form Frm_Information 
   BackColor       =   &H00FFFFFF&
   Caption         =   "도움말 정보"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   Icon            =   "Frm_Information.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Cmd_Close 
      Caption         =   "닫기"
      Height          =   450
      Left            =   6510
      TabIndex        =   3
      Top             =   4695
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   5220
      Left            =   0
      Picture         =   "Frm_Information.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1470
   End
   Begin VB.Image Image2 
      Height          =   1740
      Left            =   5730
      Picture         =   "Frm_Information.frx":228AC
      Stretch         =   -1  'True
      Top             =   375
      Width           =   2370
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   1815
      Picture         =   "Frm_Information.frx":6B4AE
      Stretch         =   -1  'True
      Top             =   1365
      Width           =   3705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Frm_Information.frx":89C70
      Height          =   945
      Left            =   1860
      TabIndex        =   2
      Top             =   3105
      Width           =   5145
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Programmed by D. H. KIM"
      Height          =   180
      Left            =   5730
      TabIndex        =   1
      Top             =   2580
      Width           =   2235
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "경고: 이 프로그램의 전부나 일부를 무단으로 복제하거나 배포하는 경우에는 저작권의 침해가 되므로 주의하시기 바랍니다."
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1785
      TabIndex        =   0
      Top             =   4365
      Width           =   6180
   End
End
Attribute VB_Name = "Frm_Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************************************************************
'***  Description   :  폼 닫기 이벤트
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Cmd_Close_Click()
Unload Me
End Sub

