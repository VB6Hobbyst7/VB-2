VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmConfirm 
   BorderStyle     =   1  '���� ����
   Caption         =   "����� Ȯ��"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5220
   StartUpPosition =   2  'ȭ�� ���
   Begin Threed.SSPanel SSPanel3 
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _Version        =   65536
      _ExtentX        =   9128
      _ExtentY        =   2937
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.CommandButton cmdCancel 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3473
         TabIndex        =   6
         Top             =   930
         Width           =   1245
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1613
         TabIndex        =   3
         Top             =   420
         Width           =   1695
      End
      Begin VB.TextBox txtPwd 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  '��� ����
         Left            =   1613
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   930
         Width           =   1695
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Ȯ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3473
         TabIndex        =   1
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   503
         TabIndex        =   5
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��й�ȣ"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   503
         TabIndex        =   4
         Top             =   990
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    gIDConfirm = -1
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If UCase(txtID) = "HITACHI" Then
        txtPwd.SetFocus
    Else
        txtID = ""
        txtPwd = ""
        MsgBox "���̵� �߸��Ǿ����ϴ�", vbInformation, "�˸�"
        txtID.SetFocus
        Exit Sub
    End If
    If UCase(txtPwd) = "7600" Then
        cmdOK.SetFocus
    Else
        txtPwd = ""
        MsgBox "��й�ȣ�� �߸��Ǿ����ϴ�", vbInformation, "�˸�"
        txtPwd.SetFocus
        Exit Sub
    End If
    gIDConfirm = 1
    Unload Me
End Sub

Private Sub Form_Load()
    gIDConfirm = -1
End Sub

Private Sub txtID_GotFocus()
    SelectFocus txtID
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If UCase(txtID) = "HITACHI" Then
            txtPwd.SetFocus
        Else
            MsgBox "���̵� �߸��Ǿ����ϴ�", vbInformation, "�˸�"
            txtID = ""
            txtID.SetFocus
        End If
    End If
End Sub

Private Sub txtPwd_GotFocus()
    SelectFocus txtPwd
End Sub

Private Sub txtPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If UCase(txtPwd) = "7600" Then
            cmdOK.SetFocus
        Else
            MsgBox "��й�ȣ�� �߸��Ǿ����ϴ�", vbInformation, "�˸�"
            txtPwd = ""
            txtPwd.SetFocus
        End If
    End If
End Sub
