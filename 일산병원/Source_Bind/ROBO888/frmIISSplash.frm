VERSION 5.00
Begin VB.Form frmIISSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   ClientHeight    =   4635
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Image Image1 
      Height          =   4590
      Left            =   4530
      Picture         =   "frmIISSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2295
   End
   Begin VB.Label lblProjectNm 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  '����
      Caption         =   "ROBO 888 System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   210
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  '����
      Caption         =   "Interface System�� �ε��ϰ� �����ϴ�......"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   0
      Top             =   4335
      Width           =   4725
   End
End
Attribute VB_Name = "frmIISSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISSplash.frm
'   �ۼ���  :
'   ��  ��  : Splash Form
'   �ۼ���  : 2003-12-04
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISSplash = Nothing
End Sub

