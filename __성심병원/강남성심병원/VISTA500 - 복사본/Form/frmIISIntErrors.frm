VERSION 5.00
Begin VB.Form frmIISIntErrors 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "��������"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.TextBox txtDetail 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   10845
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   9630
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   6030
      Width           =   1215
   End
   Begin VB.Label lblDetail 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   60
      TabIndex        =   2
      Top             =   3180
      Width           =   10815
   End
End
Attribute VB_Name = "frmIISIntErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
