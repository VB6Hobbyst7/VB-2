VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChatClient 
   Caption         =   "���ϼ���"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   510
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdClearTxt 
      Caption         =   "C"
      Height          =   315
      Left            =   4740
      TabIndex        =   10
      Top             =   7110
      Width           =   765
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "�����"
      Height          =   285
      Left            =   4740
      TabIndex        =   9
      Top             =   7770
      Width           =   765
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "����"
      Height          =   315
      Left            =   4740
      TabIndex        =   8
      Top             =   7440
      Width           =   765
   End
   Begin VB.TextBox txtPort 
      Height          =   405
      Left            =   840
      TabIndex        =   5
      Top             =   7560
      Width           =   2955
   End
   Begin VB.TextBox txtIP 
      Height          =   405
      Left            =   840
      TabIndex        =   4
      Top             =   7110
      Width           =   2955
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����"
      Height          =   435
      Left            =   3930
      TabIndex        =   3
      Top             =   7620
      Width           =   765
   End
   Begin VB.CommandButton cmdConnetion 
      Caption         =   "����"
      Height          =   465
      Left            =   3930
      TabIndex        =   2
      Top             =   7110
      Width           =   765
   End
   Begin MSWinsockLib.Winsock ctrclient 
      Left            =   2160
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtChat 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   1
      Top             =   5520
      Width           =   5475
   End
   Begin RichTextLib.RichTextBox rtfView 
      Height          =   5445
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   9604
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"frmChatClient.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   7
      Top             =   7560
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   6
      Top             =   7140
      Width           =   375
   End
End
Attribute VB_Name = "frmChatClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ŭ���̾�Ʈ �ڵ�
'������� ��ī 1. rich text box control 6.0
'              2. winsock control 6.0
'rich text box�� �Ӽ� ���� 'Locked'üũǥ��

Private Sub cmdClear_Click()
    rtfView.Text = ""
    
End Sub

Private Sub cmdClearTxt_Click()
    txtChat.Text = ""
End Sub

Private Sub cmdClose_Click()
    ctrclient.Close

End Sub

Private Sub cmdConnetion_Click()
    Dim ipAdd
'    ipAdd = InputBox("������ IP�� ������", "����", "118.36.79.132")
    ipAdd = txtIP.Text '"3.45.82.188"
    ctrclient.Close
    ctrclient.Connect ipAdd, CInt(txtPort.Text) '5150
    MsgBox "������ ���ӵǾ����ϴ�."
End Sub

Private Sub cmdExit_Click()
End
End Sub

'�����κ����� ����Ÿ ������ �̺�Ʈ ����
Private Sub ctrclient_dataarrival(ByVal bytestotal As Long)
    Dim strtext As String
    Dim strtmp As String
    
    ctrclient.GetData strtext 'strtext�� ����Ÿ�� �Է�
    strtmp = rtfView.Text + Chr(13) + Chr(10) + strtext '���� ����Ÿ�� ÷��
    rtfView.Text = strtmp 'rich text �ڽ��� ���
End Sub
'���� ���� �߻��� ����ڿ��� �˸��� �ڵ� error�̺�Ʈ ���
Private Sub ctrclient_error(ByVal number As Integer, description As String, ByVal scode As Long, ByVal source As String, ByVal helpfile As String, ByVal helpcontext As Long, canceldisplay As Boolean)
    MsgBox description, vbOKOnly, "����"
End Sub



Private Sub Form_Load()
    'Call cmdConnetion_Click
End Sub

'Ű���带 ������ �̺�Ʈ �߻��ϰ� �߻������� ������ ����
Private Sub txtchat_keypress(keyascii As Integer)
    Dim strtext As String
    
    If keyascii = vbKeyReturn Then
        strtext = "Ŭ���̾�Ʈ:" + txtChat.Text
        ctrclient.SendData txtChat.Text 'strtext
        Debug.Print txtChat.Text
        rtfView.Text = rtfView.Text + Chr(13) + Chr(10) + strtext
        'txtChat.Text = ""
    End If
End Sub
'������ ������ �������� ǥ��
Private Sub ctrclient_close()
    MsgBox "�������� ������ ���������ϴ�."
    ctrclient.Close
End Sub
