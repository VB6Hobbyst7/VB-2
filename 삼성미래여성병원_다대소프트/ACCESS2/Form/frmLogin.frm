VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  '���
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  '����
   Caption         =   "�α���"
   ClientHeight    =   3390
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   6795
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin BHButton.BHImageButton cmdCancel 
      Height          =   375
      Left            =   5325
      TabIndex        =   10
      Top             =   2850
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "���"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdOK 
      Height          =   375
      Left            =   4065
      TabIndex        =   9
      Top             =   2850
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Ȯ��"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1140
      TabIndex        =   8
      Top             =   2685
      Width           =   5610
   End
   Begin VB.TextBox txtUserID 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      Height          =   270
      IMEMode         =   8  '����
      Left            =   4245
      TabIndex        =   3
      Top             =   1350
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      Height          =   270
      IMEMode         =   3  '��� ����
      Left            =   4245
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1975
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      Height          =   270
      Left            =   4245
      TabIndex        =   1
      Top             =   1665
      Width           =   2325
   End
   Begin VB.Timer Timer1 
      Left            =   2310
      Top             =   2880
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Ver."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   3090
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   ":: m2i soft Lab Management"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Top             =   360
      Width           =   3525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Interface System"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4350
      TabIndex        =   7
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����� ID(&U) :"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   3015
      TabIndex        =   6
      Top             =   1395
      Width           =   1155
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ȣ(&P) :"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   3420
      TabIndex        =   5
      Top             =   2025
      Width           =   750
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����� �̸� :"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   3090
      TabIndex        =   4
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����� ID�� �Է� �Ͻʽÿ�."
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3030
      TabIndex        =   0
      Top             =   2400
      Width           =   2205
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   2730
      Picture         =   "frmLogin.frx":030A
      Top             =   2370
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   2730
      Picture         =   "frmLogin.frx":0454
      Top             =   2370
      Width           =   240
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   2730
      Picture         =   "frmLogin.frx":059E
      Top             =   2370
      Width           =   240
   End
   Begin VB.Image imgNet4 
      Height          =   240
      Left            =   2730
      Picture         =   "frmLogin.frx":06E8
      Top             =   2370
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   3330
      Left            =   30
      Picture         =   "frmLogin.frx":0832
      Stretch         =   -1  'True
      Top             =   30
      Width           =   6750
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldUid          As String
Private OldPwd          As String
Private MsgFg           As Boolean
Private OldUser         As UserInfo

Public CancelIsEnd      As Boolean
Public LoginSucceeded   As Boolean

Private adoRS As ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   Call ReleaseCapture
   Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub cmdCancel_Click()
    If MainForm Is Nothing Then
        Call Unload(Me)
        Set frmLogin = Nothing
        End
    Else
        CurrUser = OldUser
        Call Unload(Me)
        Set frmLogin = Nothing
    End If

End Sub

Private Sub cmdOk_Click()
   Dim ShowAtStartup As Variant

    Timer1.Enabled = False
    imgNet4.ZOrder
    If txtPassword = CurrUser.CuUserPW Then
        If CurrUser.CuPower = Authority.ELVEL_NOT Then
            MsgBox "���� ������ �����ϴ�. �����ڿ��� ���� �ϼ���. ", vbOKOnly + vbExclamation
            Exit Sub
        End If
        Call Unload(Me)
        
        If MainForm Is Nothing Then
            Set MainForm = New MDIMain
            MainForm.Show
            MainForm.stbMain.Panels(1).text = CurrUser.CuUserNM
        Else
            MainForm.stbMain.Panels(1).text = CurrUser.CuUserNM
        End If
        
        Call Load_From(frmComm) 'frmComm
        
      Else
         MsgBox "��й�ȣ�� Ʋ���ϴ�. ��й�ȣ�� Ȯ���ϼ���. ", vbOKOnly + vbExclamation
         txtPassword.SetFocus
         txtPassword.SelStart = 0
         txtPassword.SelLength = Len(txtPassword)
      End If
      

End Sub


Private Sub Load_From(ByVal frm As Form)
    
    Dim hMenu       As Long
    Dim lngStyle    As Long
    
    With frm
        .Show
        .SetFocus
    End With
    
End Sub
Private Sub Form_Activate()
    
    lblVersion.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
    txtUserID.SetFocus

End Sub

Private Sub Form_Load()

    imgNet1.ZOrder 0
    Timer1.Interval = 500
    Timer1.Enabled = True
    
    If Not MainForm Is Nothing Then
        OldUser = CurrUser
    End If
    
End Sub

Private Sub Timer1_Timer()
    DoEvents

    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If

End Sub

Private Sub txtPassword_GotFocus()
   With txtPassword
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        Call cmdOk_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserID_Change()
   txtUserName.text = ""
End Sub

Private Sub txtUserID_GotFocus()
   With txtUserID
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
        Call txtUserID_LostFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtUserID_LostFocus()
    Dim Ret As Boolean

    Dim objUserInf As clsCommon
    On Error GoTo ErrorTrap

    If ActiveControl.Name = "cmdCancel" Then Exit Sub

        If txtUserID.text = "" Then
            MsgFg = True
            MsgBox "�α׿� ID�� �Է��ϼ���. ", vbOKOnly + vbExclamation
            MsgFg = False
            txtUserID.SetFocus
            Exit Sub
        End If

        labMsg.Caption = "����Ÿ ���̽��� ������ ...."
        Screen.MousePointer = vbArrowHourglass

        Set objUserInf = New clsCommon
        With objUserInf
            .SetAdoCn AdoCn_Jet
            Set AdoRs_Jet = .Get_UserInfo(txtUserID)
            If AdoRs_Jet Is Nothing Then
                MsgBox "��ϵ��� ���� ID�Դϴ�. �α��� ID�� Ȯ���ϼ���. ", vbOKOnly + vbExclamation
                With txtUserID
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
                Set objUserInf = Nothing
            End If
        End With

        Screen.MousePointer = vbDefault
        labMsg.Caption = "����Ÿ ���̽��� ���� �Ǿ����ϴ�."

        If AdoRs_Jet.EOF Then
            MsgBox "��ϵ��� ���� ID�Դϴ�. �α��� ID�� Ȯ���ϼ���. ", vbOKOnly + vbExclamation
            Set AdoRs_Jet = Nothing
            Set objUserInf = Nothing
            With txtUserID
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
        Else
            Timer1.Enabled = False
            With CurrUser
                .CuUserID = AdoRs_Jet.Fields("EMP_ID") & ""
                .CuUserNM = AdoRs_Jet.Fields("EMP_NM") & ""
                .CuUserPW = AdoRs_Jet.Fields("PASSWD") & ""
                .CuPower = AdoRs_Jet.Fields("POWERS") & ""
                txtUserName = .CuUserNM
            End With
            imgNet4.ZOrder 0
            txtPassword.SetFocus
            AdoRs_Jet.Close
        End If

ErrorTrap:
    Set AdoRs_Jet = Nothing
    Set objUserInf = Nothing
End Sub
