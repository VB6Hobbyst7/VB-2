VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
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
      Left            =   45
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0000
            Key             =   "Hitachi 7600"
            Object.Tag             =   "H7600,Hitachi 7600"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":090F
            Key             =   "Hitachi 7180"
            Object.Tag             =   "H7180,Hitachi 7180"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":11E9
            Key             =   "LH750"
            Object.Tag             =   "LH750,LH750"
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
'   ���ϸ�  : frmIISHospital.frm (�����ھֺ���)
'   �ۼ���  : �̻��
'   ��  ��  : �������� �������� �������� �����ϴ� ��
'   �ۼ���  : 2005-03-11
'   ��  ��  :
'       1. 1.1.0: �̻��(2005-05-08)
'          - Hitachi 7180 ����߰�
'       2. 1.2.0: �̻��(2005-05-17)
'          - Coulter LH750 ����߰�
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


