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
   StartUpPosition =   3  'Windows �⺻��
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
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0000
            Key             =   "ABL555"
            Object.Tag             =   "ABL555,ABL555"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0E52
            Key             =   "Uriscan Pro+"
            Object.Tag             =   "Uriscan,Uriscan Pro+"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1CA4
            Key             =   "ADVIA 120-1"
            Object.Tag             =   "ADVIA 1,ADVIA 120-1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2AF6
            Key             =   "AU1000"
            Object.Tag             =   "AU1000,AU1000"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3948
            Key             =   "RapidLab 850"
            Object.Tag             =   "R850,RapidLab 850"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":479A
            Key             =   "Dimension AR"
            Object.Tag             =   "D-AR,Dimension AR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":55EC
            Key             =   "Dimension RXL"
            Object.Tag             =   "D-RXL,Dimension RXL"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":643E
            Key             =   "Elecsys 1010"
            Object.Tag             =   "E1010, Elecsys 1010"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":7290
            Key             =   "BN100"
            Object.Tag             =   "BN100,BN100"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":7E62
            Key             =   "ADVIA 120-2"
            Object.Tag             =   "ADVIA 2,ADVIA 120-2"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":8A34
            Key             =   "ADVIA Centaur"
            Object.Tag             =   "Centaur,ADVIA Centaur"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":A53E
            Key             =   "Elecsys 2010"
            Object.Tag             =   "E2010,Elecsys 2010"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":C048
            Key             =   "VIDAS"
            Object.Tag             =   "VIDAS,VIDAS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":C922
            Key             =   "ACL6000"
            Object.Tag             =   "A6000,ACL6000"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":D774
            Key             =   "ACL100"
            Object.Tag             =   "ACL100,ACL100"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":E04E
            Key             =   "Vitek"
            Object.Tag             =   "Vitek,Vitek"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":E95D
            Key             =   "BN II"
            Object.Tag             =   "BN II,BN II"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":EDAF
            Key             =   "ACL9000"
            Object.Tag             =   "A9000,ACL9000"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":FC01
            Key             =   "ABL835"
            Object.Tag             =   "ABL835,ABL835"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":104DB
            Key             =   "Exicycler"
            Object.Tag             =   "Exicycler,Exicycler"
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
'   ���ϸ�  : frmIISHospital.frm (���ֵ����뺴��)
'   �ۼ���  : �̻��
'   ��  ��  : �������� �������� �������� �����ϴ� ��
'   �ۼ���  : 2004-05-28
'   ��  ��  :
'   ��  ��  :
'       1.imlHospital�� �̹��� �߰��ÿ�
'         Key : �ش� ���Ű (�ǵ��� ��ü�̸� �Է�)
'         Tag : ���ٿ� ǥ�õǴ� ĸ��,�޴���(����)�� ǥ�õǴ� ĸ��
'         ��) Key:Hitachi 7600
'             Tag:H7600,Hitachi 7600
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISHospital = Nothing
End Sub

