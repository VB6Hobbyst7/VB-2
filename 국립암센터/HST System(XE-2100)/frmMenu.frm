VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  '���� ����
   Caption         =   "�޴�"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdSearch 
      Caption         =   "�˻�"
      Height          =   855
      Left            =   4680
      TabIndex        =   5
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdIPU 
      Caption         =   "IPU"
      Height          =   855
      Left            =   3540
      TabIndex        =   4
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "WorkList"
      Height          =   855
      Left            =   2400
      TabIndex        =   3
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�"
      Height          =   855
      Left            =   5820
      TabIndex        =   2
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdCode 
      Caption         =   "�ڵ弳��"
      Height          =   855
      Left            =   1260
      TabIndex        =   1
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdComSetup 
      Caption         =   "��ż���"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCode_Click()
    Connect_Local
    
    frmCode.Show 1
    
    DisConnect_Local
End Sub

Private Sub cmdComSetup_Click()
    frmConfig.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdWorkList_Click()

End Sub

Private Sub Form_Load()
    'gEquip = "XE2100"
End Sub
