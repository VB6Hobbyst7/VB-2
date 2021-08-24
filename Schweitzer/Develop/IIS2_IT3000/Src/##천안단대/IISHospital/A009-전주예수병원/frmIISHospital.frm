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
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0000
            Key             =   "Uriscan Pro-1"
            Object.Tag             =   "Uriscan-1,Uriscan Pro-1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0E52
            Key             =   "Uriscan Pro-2"
            Object.Tag             =   "Uriscan-2,Uriscan Pro-2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1CA4
            Key             =   "Dimension RXL"
            Object.Tag             =   "D-RXL,Dimension RXL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2876
            Key             =   "RapidLab 865"
            Object.Tag             =   "R-865,RapidLab 865"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3AF8
            Key             =   "Stks-1"
            Object.Tag             =   "Stks-1,Stks-1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":46CA
            Key             =   "Stks-2"
            Object.Tag             =   "Stks-2,Stks-2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":529C
            Key             =   "Hitachi 7600"
            Object.Tag             =   "H7600,Hitachi 7600"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":62EE
            Key             =   "Axsym"
            Object.Tag             =   "Axsym,Axsym"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":7140
            Key             =   "SE-9000"
            Object.Tag             =   "SE-9000,SE-9000"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":7D12
            Key             =   "Variant II"
            Object.Tag             =   "Variant II,Variant II"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":85EC
            Key             =   "CX3 Delta"
            Object.Tag             =   "CX3 Delta, CX3 Delta"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":943E
            Key             =   "LPIA-NV7"
            Object.Tag             =   "LPIA,LPIA-NV7"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":A490
            Key             =   "Thrombolyzer compact"
            Object.Tag             =   "Compact,Thrombolyzer compact"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":B4E2
            Key             =   "Vitek"
            Object.Tag             =   "Vitek,Vitek"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":C334
            Key             =   "RapidLab 850"
            Object.Tag             =   "R-850,RapidLab 850"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":CF06
            Key             =   "RapidLab 860"
            Object.Tag             =   "R-860,RapidLab 860"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":DAD8
            Key             =   "Vitek II"
            Object.Tag             =   "Vitek II,Vitek II"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":E92A
            Key             =   "Thrombolyzer RackRotor"
            Object.Tag             =   "RackRotor,Thrombolyzer RackRotor"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":F239
            Key             =   "XT-1800i"
            Object.Tag             =   "1800i,XT-1800i"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":FB13
            Key             =   "CA-500"
            Object.Tag             =   "CA500,CA-500"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":107ED
            Key             =   "CA-1500"
            Object.Tag             =   "CA1500,CA-1500"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1163F
            Key             =   "CA-1500ER"
            Object.Tag             =   "CA1500ER,CA-1500ER"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":11F19
            Key             =   "D-RXLM"
            Object.Tag             =   "D-RXLM,D-RXLM"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":12BF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1404D
            Key             =   "Architect"
            Object.Tag             =   "Architect,Architect"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":14927
            Key             =   "ESR"
            Object.Tag             =   "ESR,ESR"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":15201
            Key             =   "PCX1"
            Object.Tag             =   "PCX1,PCX1"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":15ADB
            Key             =   "Gem3000"
            Object.Tag             =   "Gem3000,Gem3000"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":16F35
            Key             =   "D-RXLM2"
            Object.Tag             =   "D-RXLM2,D-RXLM2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1838F
            Key             =   "XE2100"
            Object.Tag             =   "XE2100,XE2100"
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
'   파일명  : frmIISHospital.frm (전주예수병원)
'   작성자  : 이상대
'   내  용  : 병원별로 사용장비의 아이콘을 관리하는 폼
'   작성일  : 2004-10-05
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


