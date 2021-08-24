VERSION 5.00
Begin VB.Form frm����_SerialDataViewer 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "Serial Data Viewer"
   ClientHeight    =   4965
   ClientLeft      =   7560
   ClientTop       =   3495
   ClientWidth     =   10095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm����_SerialDataViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "�ݱ�(&Q)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8820
      TabIndex        =   1
      Top             =   4500
      Width           =   1215
   End
   Begin VB.TextBox txtSerialDataView 
      Appearance      =   0  '���
      Height          =   3675
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Text            =   "frm����_SerialDataViewer.frx":9F8A
      Top             =   600
      Width           =   9975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   10020
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Serial Data Viewer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   180
      TabIndex        =   3
      Top             =   60
      Width           =   1935
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   90
      TabIndex        =   2
      Top             =   3420
      Width           =   60
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   495
      Index           =   1
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   9975
   End
End
Attribute VB_Name = "frm����_SerialDataViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call cmdQuit_Click
    End Select
End Sub

Private Sub Form_Load()
    Me.Height = 5445
    Me.Width = 10215
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    txtSerialDataView = ""
End Sub
