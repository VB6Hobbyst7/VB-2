VERSION 5.00
Begin VB.Form FSM0101 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '����
   Caption         =   "LogIn"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FSM0101.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FSM0101.frx":030A
   ScaleHeight     =   4320
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtUserCd 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFF0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2010
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2460
      Width           =   1065
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFF0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  '��� ����
      Left            =   2010
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2850
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4860
      Picture         =   "FSM0101.frx":D015
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   2850
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ȯ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3900
      MaskColor       =   &H000000FF&
      Picture         =   "FSM0101.frx":D5F2
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   2850
      Width           =   975
   End
   Begin VB.Image imgCancel 
      Height          =   420
      Left            =   5250
      MousePointer    =   99  '����� ����
      ToolTipText     =   "���"
      Top             =   2880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgOK 
      Height          =   420
      Left            =   4650
      MousePointer    =   99  '����� ����
      ToolTipText     =   "Ȯ��"
      Top             =   2880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   120
      Picture         =   "FSM0101.frx":DB59
      Top             =   330
      Width           =   1950
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '����
      Caption         =   "���� �ӻ󺴸���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "�����ȣ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   810
      TabIndex        =   6
      Top             =   2520
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "��й�ȣ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   675
      TabIndex        =   5
      Top             =   2910
      Width           =   1290
   End
   Begin VB.Label lblUserNm 
      Appearance      =   0  '���
      BackColor       =   &H80000018&
      BackStyle       =   0  '����
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3060
      TabIndex        =   4
      Top             =   2460
      Width           =   2745
   End
End
Attribute VB_Name = "FSM0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bAdminLogIn As Boolean
Dim sPassword As String

Public Sub SetImg(ByRef imgControl As Object, ByVal strImgName As String)
    imgControl.Picture = LoadResPicture(strImgName, vbResBitmap)
    imgControl.MouseIcon = LoadResPicture("Point", vbResCursor)
End Sub

Private Function ConfirmUserCd() As Integer
    On Error GoTo ErrHandler
    
    Dim i%
    Dim bRetVal As Boolean
    Dim sBuf$
    Dim CUser As DCB0101
    
    ConfirmUserCd = 0
    bAdminLogIn = False
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\System.Manager", "AdminID")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\System.Manager", "AdminID", "JCE SA")

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        sBuf = "JCE SA"
    End If
    
    If sBuf = txtUserCd Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\System.Manager", "AdminPWD")
        
        If sBuf = "" Then
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                        "Software\SemiLIS\Program Config\System.Manager", "AdminPWD", "sksmssk")
    
            If bRetVal = True Then
            Else
                MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            End If
            
            sBuf = "sksmssk"
        End If
        
        ConfirmUserCd = 1
        
        sPassword = sBuf
        
        bAdminLogIn = True
        
        txtPassword.SetFocus
        
        Exit Function
    End If
    
    Set CUser = New DCB0101
    
    CUser.Get_Login_Info 0, txtUserCd, ""
    
    i = CUser.CurItemCnt
    
    If i = 0 Then
         MsgBox "�������� �ʴ� �����ȣ �Դϴ�." & Chr(10) & _
        "Ȯ���Ͽ� �ֽʽÿ�!!", vbOKOnly + vbInformation

'        MsgBox "�������� �ʴ� �����ȣ �Դϴ�. Ȯ���Ͽ� �ֽʽÿ�!!"
        Set CUser = Nothing
        Exit Function
    ElseIf i > 1 Then
        MsgBox "User������ ������ �ֽ��ϴ�"
        Set CUser = Nothing
        Exit Function
    End If
    
    ConfirmUserCd = 1
    
    sPassword = CUser.TotField01
    lblUserNm.Caption = CUser.TotField02
    gsDefaultPartCd = Left$(CUser.TotField03, 1)
    
    For i = 1 To giPartCnt
        If gPartTable(i).sPartInit = gsDefaultPartCd Then
            gsDefaultPartNm = gPartTable(i).sPartName
            Exit For
        End If
    Next
    
    gsDefaultSlipCd = CUser.TotField03
    gsDefaultSlipNm = CUser.TotField04
    gsDefaultSpecimenCd = CUser.TotField05
    gsDefaultSpecimenNm = CUser.TotField06
    gsDefaultSchOpt = CUser.TotField07
    
    Set CUser = Nothing
    
    Call WriteRegistry
    
    Exit Function

ErrHandler:
    Set CUser = Nothing
End Function

Private Sub WriteRegistry()
    Dim bRetVal As Boolean
    '----- Registry ������ Log In ������ ���� Cur.Cfg�� ���� -------
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "PartCd", gsDefaultPartCd)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If

    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "PartNm", gsDefaultPartNm)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SlipCd", gsDefaultSlipCd)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SlipNm", gsDefaultSlipNm)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SpecimenCd", gsDefaultSpecimenCd)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SpecimenNm", gsDefaultSpecimenNm)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserCd", txtUserCd)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserNm", lblUserNm.Caption)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserSchOpt", gsDefaultSchOpt)

    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
End Sub

Private Sub cmdCancel_Click()

    End
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        End
    End If
    
End Sub

Private Sub imgCancel_Click()
    
    End

End Sub

Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgCancel, "Cancel_d")
End Sub

Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgCancel, "Cancel")
End Sub

Private Sub imgOK_Click()

    Dim UserCd As String
    Dim ConnectOk As Integer
    Dim sBuf As String
    Dim bRetVal As Boolean
    Dim i%
    
    On Error GoTo cmdOk_ERROR
    
    UserCd = Trim(txtUserCd)
    If UserCd = "" Or Trim(txtPassword) = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    ConnectOk = False
    
    If txtPassword = sPassword Then
        ConnectOk = True
    Else
        MsgBox "Password�� ��Ȯ���� �ʽ��ϴ�. Ȯ���Ͽ� �ֽʽÿ�!!"
        Call Txt_Highlight(txtPassword)
    End If
             
    If ConnectOk Then
        If bAdminLogIn = False Then
            ViewUserNm lblUserNm
        Else
            Screen.MousePointer = vbDefault
            
            sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\System.Manager", "DB.SerialNo")

            If sBuf = "" Then
                Load FSM0201
                FSM0201.Show vbModal, FSM0101
                
                For i = 1 To 6
                    FGM0101.mnuB(i).Visible = True
                Next
                
                For i = 9 To 10
                    FGM0101.mnuB(i).Visible = False
                Next
                
                FGM0101.mnuJR00.Visible = False
                FGM0101.mnuS00.Visible = False
                FGM0101.mnuO00.Visible = False
                FGM0101.mnuT00.Visible = False
                FGM0101.mnuI00.Visible = False
            End If
            
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserCd", "SA")

            If bRetVal = True Then
            Else
                MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            End If
            
        End If
        Unload Me
    End If
    
    Screen.MousePointer = vbDefault
    
cmdOk_ERROR:
End Sub

Private Sub imgOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgOK, "OK_d")
End Sub

Private Sub imgOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgOK, "OK")
End Sub
Private Sub cmdOK_Click()

    Dim UserCd As String
    Dim ConnectOk As Integer
    Dim sBuf As String
    Dim bRetVal As Boolean
    Dim i%
    
    On Error GoTo cmdOk_ERROR
    
    UserCd = Trim(txtUserCd)
    If UserCd = "" Or Trim(txtPassword) = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    ConnectOk = False
    
    If txtPassword = sPassword Then
        ConnectOk = True
    Else
        MsgBox "Password�� ��Ȯ���� �ʽ��ϴ�." & Chr(10) & _
               "Ȯ���Ͽ� �ֽʽÿ�!!", vbOKOnly + vbInformation
        Call Txt_Highlight(txtPassword)
    End If
             
    If ConnectOk Then
        If bAdminLogIn = False Then
            ViewUserNm lblUserNm
        Else
            Screen.MousePointer = vbDefault
            
            sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\System.Manager", "DB.SerialNo")

            If sBuf = "" Then
                Load FSM0201
                FSM0201.Show vbModal, FSM0101
                
                For i = 1 To 6
                    FGM0101.mnuB(i).Visible = True
                Next
                
                For i = 9 To 10
                    FGM0101.mnuB(i).Visible = False
                Next
                
                FGM0101.mnuJR00.Visible = False
                FGM0101.mnuS00.Visible = False
                FGM0101.mnuO00.Visible = False
                FGM0101.mnuT00.Visible = False
                FGM0101.mnuI00.Visible = False
            End If
            
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserCd", "SA")

            If bRetVal = True Then
            Else
                MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            End If
            
        End If
        Unload Me
    End If
    
    Screen.MousePointer = vbDefault
    
cmdOk_ERROR:
            
End Sub


Private Sub Form_Load()
    Dim bRetVal As Boolean
    
    FSM0101.lblTitle.Caption = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\LogIn.Title", "")
    
    If FSM0101.lblTitle.Caption = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\LogIn.Title", "", "Lab.")

        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        FSM0101.lblTitle.Caption = "Lab."
    End If
    
    '������Ʈ���� Ÿ��Ʋ �Է�
    RegEditCurFrmTitle FSM0101.Caption
    
    Call SetImg(imgCancel, "Cancel")
    Call SetImg(imgOK, "OK")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    '������Ʈ������ Ÿ��Ʋ ����
    Call InitRegCurFrmTitle
End Sub


Private Sub txtPassword_Change()

    If txtPassword.SelStart = txtPassword.MaxLength Then cmdOK.SetFocus
    
End Sub

Private Sub txtPassword_GotFocus()

    Call Txt_Highlight(txtPassword)
    
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdOK_Click
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtUserCd_Change()
    Dim iRetVal As Integer
    
    If lblUserNm <> "" Then lblUserNm = ""
    
    If Len(txtUserCd) = txtUserCd.MaxLength Then
        iRetVal = ConfirmUserCd
        
        If iRetVal = 1 Then
            txtPassword.SetFocus
        Else
            Call Txt_Highlight(txtUserCd)
        End If
        
    End If
    
End Sub

Private Sub txtUserCd_GotFocus()
    
    Call Txt_Highlight(txtUserCd)
    
End Sub

Private Sub txtUserCd_KeyPress(KeyAscii As Integer)
    Dim iRetVal%
    
    If KeyAscii = 13 Then
        iRetVal = ConfirmUserCd
        
        If iRetVal = 1 Then
            txtPassword.SetFocus
        Else
            Call Txt_Highlight(txtUserCd)
        End If
        
        KeyAscii = 0
    End If
    
    '�׻� �빮�ڷ� ������
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtUserCd_Validate(Cancel As Boolean)
    Dim iRetVal%
    
    If txtUserCd = "" Then
    Else
        iRetVal = ConfirmUserCd
        
        If iRetVal = 1 Then
            txtPassword.SetFocus
        Else
            Call Txt_Highlight(txtUserCd)
        End If
    End If
End Sub
