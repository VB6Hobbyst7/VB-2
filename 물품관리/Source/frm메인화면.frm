VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frm����ȭ�� 
   BackColor       =   &H8000000C&
   Caption         =   "��ǰ�����ý���"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   19080
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'ȭ�� ���
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  '�� ����
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   1191
      ButtonWidth     =   1455
      ButtonHeight    =   1138
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���ſ�û"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imgTree 
         Left            =   12840
         Top             =   120
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
               Picture         =   "frm����ȭ��.frx":0000
               Key             =   "close"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm����ȭ��.frx":08DA
               Key             =   "open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm����ȭ��.frx":11B4
               Key             =   "choice"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgMenu 
         Left            =   11820
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm����ȭ��.frx":1EA6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  '�Ʒ� ����
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   10875
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2725
            MinWidth        =   2716
            Text            =   "(��)��������Ƽ"
            TextSave        =   "(��)��������Ƽ"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25241
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "ȫ�浿"
            TextSave        =   "ȫ�浿"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "2012-07-16"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu�����ڷ� 
      Caption         =   "[&1]�����ڷ�  "
      Begin VB.Menu mnu�����ڵ��� 
         Caption         =   "�����ڵ���"
      End
      Begin VB.Menu mnu��ü�����ڷ� 
         Caption         =   "��ü�����ڷ�"
      End
      Begin VB.Menu mnu��ǰ�����ڷ� 
         Caption         =   "��ǰ�����ڷ�"
      End
      Begin VB.Menu mnu�������ڷ� 
         Caption         =   "�������ڷ�"
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu�˻纰�ҿ䷮ 
         Caption         =   "�˻��׸� �ҿ䷮"
      End
      Begin VB.Menu mnu��񺰿�ҿ䷮ 
         Caption         =   "��񺰿 �ҿ䷮"
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu mnu����ڱ����ڷ� 
         Caption         =   "����ڱ����ڷ�"
      End
   End
   Begin VB.Menu mnu��ǰ��û 
      Caption         =   "[&2]��ǰ��û  "
      Begin VB.Menu mnu��ǰ��û���Ϲ� 
         Caption         =   "��ǰ��û�� �ۼ�(�Ϲ�)"
      End
      Begin VB.Menu mnu��ǰ��û���з� 
         Caption         =   "��ǰ��û�� �ۼ�(�з�)"
      End
      Begin VB.Menu line20 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu���ֱ��� 
      Caption         =   "[&3]���ֱ���  "
      Begin VB.Menu mnu���ּ��Ϲ� 
         Caption         =   "���ּ� �ۼ�"
      End
      Begin VB.Menu mnu���ּ���ü 
         Caption         =   "���ּ� �ۼ�(��ü)"
      End
      Begin VB.Menu mnu���ּ���û 
         Caption         =   "���ּ� �ۼ�(��û)"
      End
      Begin VB.Menu line30 
         Caption         =   "-"
      End
      Begin VB.Menu mnu�����԰� 
         Caption         =   "�����԰� ó��"
      End
      Begin VB.Menu mnu�Ϲݱ��ż� 
         Caption         =   "�Ϲݱ��ż� �ۼ�"
      End
   End
   Begin VB.Menu mnu��ǰ��� 
      Caption         =   "[&4]��ǰ���  "
      Begin VB.Menu mnu��ǰ��� 
         Caption         =   "��ǰ��� �ۼ�"
      End
      Begin VB.Menu line40 
         Caption         =   "-"
      End
      Begin VB.Menu mnu������ 
         Caption         =   "�������"
      End
      Begin VB.Menu mnu�����˻��� 
         Caption         =   "�����˻���"
      End
      Begin VB.Menu line41 
         Caption         =   "-"
      End
      Begin VB.Menu mnu�ϸ��� 
         Caption         =   "���ϸ����۾�"
      End
   End
   Begin VB.Menu mnu��ǰ��� 
      Caption         =   "[&5]��ǰ���  "
      Begin VB.Menu mnuǰ�������Ȳ 
         Caption         =   "ǰ�� �����Ȳ"
      End
      Begin VB.Menu mnu������Ȳ 
         Caption         =   "ǰ�� ������Ȳ"
      End
      Begin VB.Menu line50 
         Caption         =   "-"
      End
      Begin VB.Menu mnu��������� 
         Caption         =   "���������"
      End
      Begin VB.Menu mnu���⵵�̿� 
         Caption         =   "���⵵ �̿�"
      End
   End
   Begin VB.Menu mnuȯ�漳�� 
      Caption         =   "[&6]ȯ�漳��  "
      Begin VB.Menu mnu����ڼ��� 
         Caption         =   "����� ȯ�漳��"
      End
      Begin VB.Menu line60 
         Caption         =   "-"
      End
      Begin VB.Menu mnu�������� 
         Caption         =   "�������� ����"
      End
      Begin VB.Menu mnu���� 
         Caption         =   "���α׷� ����"
      End
   End
End
Attribute VB_Name = "frm����ȭ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()

    Me.Height = 12000
    Me.Width = 19200

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If MsgBox("���α׷��� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo) <> vbYes Then
        Cancel = 1
    Else
        End
    End If

End Sub

Private Sub mnu�˻纰�ҿ䷮_Click()

    With frm�˻��׸񺰽þ����
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu�����ڵ���_Click()

    With frm�����ڵ�
        Call psFormCenter(frm�����ڵ�)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu��ǰ�����ڷ�_Click()

    With frm��ǰ�����ڷ�
        Call psFormCenter(frm��ǰ�����ڷ�)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu��ǰ��û���з�_Click()

    With frm��ǰ��û���з�
        Call psFormCenter(frm��ǰ��û���з�)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu��ǰ��û���Ϲ�_Click()

    With frm��ǰ��û���Ϲ�
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub psFormCenter(ByVal brForm As Form)

    brForm.Top = (frm����ȭ��.ScaleHeight - brForm.Height) / 2
    brForm.Left = (frm����ȭ��.ScaleWidth - brForm.Width) / 2
    
    brForm.Height = brForm.Height - 120

End Sub

Private Sub mnu��ǰ�����Ȳ_Click()

End Sub

Private Sub mnu��ǰ���_Click()

    With frm����Ϲ�
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu���ּ���ü_Click()

    With frm���ּ���ü
        Call psFormCenter(frm���ּ���ü)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu���ּ���û_Click()

    With frm���ּ���û
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu���ּ��Ϲ�_Click()

    With frm���ּ��Ϲ�
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu�����԰�_Click()

    With frm���ż�����
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu����ڱ����ڷ�_Click()

    With frm����ڱ����ڷ�
        Call psFormCenter(frm����ڱ����ڷ�)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu����ڼ���_Click()

    frm�����ͺ��̽�.Show vbModal
    Call gsRegisterApply

End Sub

Private Sub mnu��ü�����ڷ�_Click()

    With frm��ü�����ڷ�
        Call psFormCenter(frm��ü�����ڷ�)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu�Ϲݱ��ż�_Click()

    With frm���ż��Ϲ�
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu�������ڷ�_Click()

    With frm�������ڷ�
        Call psFormCenter(frm�������ڷ�)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu��񺰿�ҿ䷮_Click()

    With frm��񺰿�þ����
        Call psFormCenter(frm��񺰿�þ����)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu������_Click()

    With frm���������
        Call psFormCenter(frm���������)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu����_Click()

    Unload Me
    
End Sub

Private Sub mnuǰ�������Ȳ_Click()

    With frmǰ�������Ȳ
        Call psFormCenter(frmǰ�������Ȳ)
        .Show
        .ZOrder 0
    End With

End Sub
