VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "�Ƿڰ������"
   ClientHeight    =   7860
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   9960
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows �⺻��
   WindowState     =   2  '�ִ�ȭ
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "���������ں���"
            TextSave        =   "���������ں���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1765
            MinWidth        =   1765
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2010-09-06"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "���� 6:32"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "�޵����Ʈ ��(051)462-1751"
            TextSave        =   "�޵����Ʈ ��(051)462-1751"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuPatList 
      Caption         =   "�Ƿ�ȯ�ڸ���Ʈ(SCL)"
   End
   Begin VB.Menu mnuSCLRes 
      Caption         =   "SCL �������"
   End
   Begin VB.Menu mnuNeodinRes 
      Caption         =   "�׿��� �������"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    '��������
    gConnect = False

    If gConnect = False Then
        If Connect = False Then
            Exit Sub
        Else
            gConnect = True
        End If
    End If

    mvbFrm.Mvb1.MServer = "CN_IPTCP:211.57.171.3[6001]"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If gConnect = True Then
        DisConnect
        gConnect = False
    End If
    
    Call KillProcess("�Ƿڰ������.exe")
    
    End
End Sub

Private Sub mnuPatList_Click()
    Call fnActiveFormIsAppoint(frmPatList.hwnd)
End Sub

Private Sub mnuSCLRes_Click()
    Call fnActiveFormIsAppoint(frmSCLTrans.hwnd)
End Sub
