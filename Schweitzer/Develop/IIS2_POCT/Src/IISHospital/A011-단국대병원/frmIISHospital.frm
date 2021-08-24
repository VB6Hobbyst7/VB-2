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
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0000
            Key             =   "XE-2100"
            Object.Tag             =   "XE-2100,XE-2100"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0E52
            Key             =   "ACL9000"
            Object.Tag             =   "ACL9000,ACL9000"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1CA4
            Key             =   "TEST-1"
            Object.Tag             =   "TEST-1,TEST-1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2AF6
            Key             =   "Modular DP"
            Object.Tag             =   "Modular DP,Modular DP"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":33D0
            Key             =   "CX3 Delta"
            Object.Tag             =   "CX3 Delta,CX3 Delta"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":4222
            Key             =   "BEP III"
            Object.Tag             =   "BEP III, BEP III"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":5074
            Key             =   "BN II"
            Object.Tag             =   "BN II,BN II"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":5EC6
            Key             =   "Axsym"
            Object.Tag             =   "Axsym,Axsym"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":6D18
            Key             =   "Elecsys 2010"
            Object.Tag             =   "E2010,Elecsys 2010"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":8822
            Key             =   "Vidas"
            Object.Tag             =   "Vidas,Vidas"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":90FC
            Key             =   "Variant II"
            Object.Tag             =   "Variant II,Variant II"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":AC06
            Key             =   "XE-2100(응급)"
            Object.Tag             =   "XE2100,XE-2100(응급)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":BA58
            Key             =   "KX-21"
            Object.Tag             =   "KX-21,KX-21"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":CAAA
            Key             =   "ACL9000(응급)"
            Object.Tag             =   "ACL9000,ACL9000(응급)"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":D8FC
            Key             =   "Modular P"
            Object.Tag             =   "Modular P,Modular P"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":E1D6
            Key             =   "Integra 700"
            Object.Tag             =   "Integra700,Integra 700"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":EEB0
            Key             =   "ABL 555"
            Object.Tag             =   "ABL555,ABL 555"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":FD02
            Key             =   "ABL 520"
            Object.Tag             =   "ABL520,ABL 520"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":10D54
            Key             =   "ABL 50"
            Object.Tag             =   "ABL50,ABL 50"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":11DA6
            Key             =   "Miditron"
            Object.Tag             =   "Miditron,Miditron"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":12680
            Key             =   "Miditron(응급)"
            Object.Tag             =   "Miditron,Miditron(응급)"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":12F5A
            Key             =   "Vitek"
            Object.Tag             =   "Vitek,Vitek"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":13DAC
            Key             =   "OC-SensorII"
            Object.Tag             =   "SensorII,OC-SensorII"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":14146
            Key             =   "Dimension RXL"
            Object.Tag             =   "RXL,Dimension RXL"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":14F98
            Key             =   "Osmometer"
            Object.Tag             =   "Osmometer,Osmometer"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":15872
            Key             =   "Integra 800"
            Object.Tag             =   "Integra 800,Integra 800"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1614C
            Key             =   "ABL 830"
            Object.Tag             =   "ABL 830,ABL 830"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":16A26
            Key             =   "ABL 835"
            Object.Tag             =   "ABL 835,ABL 835"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":17300
            Key             =   "BEP2000"
            Object.Tag             =   "BEP2000,BEP2000"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":17BDA
            Key             =   "iQ200"
            Object.Tag             =   "iQ200,iQ200"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":184E9
            Key             =   "GEM3000"
            Object.Tag             =   "GEM3000,GEM3000"
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
'   파일명  : frmIISHospital.frm (단국대병원)
'   작성자  : 이상대
'   내  용  : 병원별로 사용장비의 아이콘을 관리하는 폼
'   작성일  : 2005-06-23
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


