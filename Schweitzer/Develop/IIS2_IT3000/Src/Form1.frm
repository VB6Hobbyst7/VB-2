VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����Ȯ��"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1275
   ScaleWidth      =   3915
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   330
      Left            =   2835
      TabIndex        =   3
      Top             =   855
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   330
      Left            =   1755
      TabIndex        =   2
      Top             =   855
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   450
      Width           =   3750
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BackStyle       =   0  '����
      Caption         =   "�����Ͻ÷��� ��й�ȣ�� �Է��ϼ���."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   6660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strLogOn     As Boolean

Public Sub Command1_Click()
    Me.Show
    If Text1.Text = "1234" Then
        strLogOn = True
        Call mdiIISMain.MDIForm_Unload(0)
        Unload Me
    Else
        If Text1.Text = "" Then
            Call MsgBox("��й�ȣ�� �Է��ϼ���. Ȯ���ϼ���.", vbExclamation, App.Title)
        Else
            Call MsgBox("�߸� �� ��й�ȣ�Դϴ�. Ȯ���ϼ���.", vbExclamation, App.Title)
            Text1.Text = ""
        End If
        Text1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    Call mdiIISMain.MDIForm_QueryUnload(1, 1)
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Text1) > 0 Then
            Call Command1_Click
        End If
    End If
End Sub
