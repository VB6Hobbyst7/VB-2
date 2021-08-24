VERSION 5.00
Begin VB.Form frm����_Set_DataBase 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "DataBase Info Setting"
   ClientHeight    =   1815
   ClientLeft      =   3255
   ClientTop       =   2550
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
   Icon            =   "frm����_Set_DataBase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
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
      Left            =   7560
      TabIndex        =   1
      Top             =   1380
      Width           =   1215
   End
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
      TabIndex        =   2
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txtDBConStr 
      Appearance      =   0  '���
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   9975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   10020
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "DataBase Connect String"
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
      TabIndex        =   4
      Top             =   60
      Width           =   2595
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
      TabIndex        =   3
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
Attribute VB_Name = "frm����_Set_DataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtDBConStr) = "" Then
        MsgBox "DB Connect String�� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��"
        txtDBConStr.SetFocus
        Exit Sub
    Else
        If OpenDB(txtDBConStr) = True Then
            Call CloseDB
            
            Call SaveSetting(REG_MAKER & "\" & REG_PRODUCT, REG_DB_INFO, REG_DB_CONSTR, txtDBConStr)
            
            MsgBox "ȯ�漳���� ����Ǿ����ϴ�." & vbCrLf & _
                   "���α׷��� �� �����Ͻʽÿ�!", vbInformation, "���α׷� ����"
            
            End
        Else
            MsgBox "DB Connect String�� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��"
        End If
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call cmdCancle_Click
    End Select
End Sub

Private Sub Form_Load()
    Me.Height = 2295
    Me.Width = 10215
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
'''    Call SaveSetting(REG_MAKER & "\" & REG_PRODUCT, REG_DB_INFO, REG_DB_CONSTR, "Provider=msdaora;Data Source=phis;User Id=Phis_lis;Password=Phis_lis;")
    
    txtDBConStr = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_DB_INFO, REG_DB_CONSTR) '/RemoteHost
End Sub

Private Sub txtPasswd_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtPasswd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtServer_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtServer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtUser_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDBConStr_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtDBConStr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
