VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FSM0201 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "Database Serial No."
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "FSM0201.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin Threed.SSCommand cmdOK 
      Height          =   645
      Left            =   660
      TabIndex        =   1
      Top             =   1140
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "DB.SerialNo Ȯ��"
   End
   Begin VB.TextBox txtSerial 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   390
      Width           =   3945
   End
End
Attribute VB_Name = "FSM0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Dim bRetVal As Boolean
    
    If txtSerial = "" Then
            
    Else
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\SemiLIS\Program Config\System.Manager", "DB.SerialNo", txtSerial)
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        Unload Me
    End If
End Sub
