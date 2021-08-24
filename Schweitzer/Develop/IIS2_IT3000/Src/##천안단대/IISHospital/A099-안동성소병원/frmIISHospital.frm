VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIISHospital 
   Caption         =   "IISHospital"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.ImageList imlHospital 
      Left            =   60
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0000
            Key             =   "XE-2100"
            Object.Tag             =   "XE-2100,XE-2100"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":145A
            Key             =   "XS-1000i"
            Object.Tag             =   "XS-1000i,XS-1000i"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":28B4
            Key             =   "Stago"
            Object.Tag             =   "Stago,Stago"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3D0E
            Key             =   "RapidLab348"
            Object.Tag             =   "Rapid348,Rapid348"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":5168
            Key             =   "RapidLab744"
            Object.Tag             =   "Rapid744,Rapid744"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":65C2
            Key             =   "Hitachi7180"
            Object.Tag             =   "Hit7180,Hit7180"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":7A1C
            Key             =   "Vitros250"
            Object.Tag             =   "Vitros250,Vitros250"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":8E76
            Key             =   "Centaur"
            Object.Tag             =   "Centaur,Centaur"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":A2D0
            Key             =   "Axsym"
            Object.Tag             =   "Axsym,Axsym"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":B72A
            Key             =   "UrinScanPro"
            Object.Tag             =   "UrinScan,UrinScan"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":CB84
            Key             =   "PFA100"
            Object.Tag             =   "PFA100,PFA100"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":DFDE
            Key             =   "MultiReader"
            Object.Tag             =   "MultiReader,MultiReader"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIISHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISHospital.frm (안동성소병원)
'   작성자  : 오세원
'   내  용  : 병원별로 사용장비의 아이콘을 관리하는 폼
'   작성일  : 2008-07-07
'   메  모  :
'       1.imlHospital에 이미지 추가시에
'         Key : 해당 장비키 (되도록 전체이름 입력)
'         Tag : 툴바에 표시되는 캡션,메뉴바(툴팁)에 표시되는 캡션
'         예) Key:Hitachi 7600
'             Tag:H7600,Hitachi 7600
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISHospital = Nothing
End Sub


