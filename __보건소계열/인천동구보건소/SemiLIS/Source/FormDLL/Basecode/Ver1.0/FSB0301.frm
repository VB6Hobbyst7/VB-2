VERSION 5.00
Begin VB.Form FSB0301 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "LogIn"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FSB0301.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FSB0301.frx":030A
   ScaleHeight     =   4095
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtUserCd 
      BackColor       =   &H00FFFFF0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2310
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2700
      Width           =   1065
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFF0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  '사용 못함
      Left            =   2310
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3090
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5340
      Picture         =   "FSB0301.frx":2C66
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4380
      MaskColor       =   &H000000FF&
      Picture         =   "FSB0301.frx":3243
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      TabIndex        =   7
      Top             =   2130
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "사원번호 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1110
      TabIndex        =   6
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   975
      TabIndex        =   5
      Top             =   3150
      Width           =   1290
   End
   Begin VB.Label lblUserNm 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3360
      TabIndex        =   4
      Top             =   2700
      Width           =   2955
   End
End
Attribute VB_Name = "FSB0301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bAdminLogIn As Boolean
Dim sPassword As String

Private Function ConfirmUserCd() As Integer
    On Error GoTo ErrHandler
    
    Dim CUser As DCB0101
    Dim i%
    Dim bRetVal As Boolean
    Dim sBuf$
    
    ConfirmUserCd = 0
    bAdminLogIn = False
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\System.Manager", "AdminID")
    
    If sBuf = txtUserCd Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\System.Manager", "AdminPWD")
                    
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
        MsgBox "존재하지 않는 사원번호 입니다. 확인하여 주십시요!!"
        Set CUser = Nothing
        Exit Function
    ElseIf i > 1 Then
        MsgBox "User설정에 문제가 있습니다"
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
    '----- Registry 현재의 Log In 정보에 대해 Cur.Cfg에 쓴다 -------
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "PartCd", gsDefaultPartCd)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If

    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "PartNm", gsDefaultPartNm)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SlipCd", gsDefaultSlipCd)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SlipNm", gsDefaultSlipNm)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SpecimenCd", gsDefaultSpecimenCd)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "SpecimenNm", gsDefaultSpecimenNm)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserCd", txtUserCd)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserNm", lblUserNm.Caption)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Cur.Cfg", "UserSchOpt", gsDefaultSchOpt)

    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    Dim UserCd As String
    Dim ConnectOk As Integer
    
    On Error GoTo cmdOk_ERROR
    
    UserCd = Trim(txtUserCd)
    If UserCd = "" Or Trim(txtPassword) = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    ConnectOk = False
    
    If txtPassword = sPassword Then
        ConnectOk = True
    Else
        MsgBox "Password가 정확하지 않습니다. 확인하여 주십시요!!"
        Call Txt_Highlight(txtPassword)
    End If
             
    If ConnectOk Then
        If bAdminLogIn = False Then
            ViewUserNm lblUserNm
        Else
            
        End If
        Unload Me
    End If
    
    Screen.MousePointer = vbDefault
    
cmdOk_ERROR:
            
End Sub


Private Sub Form_Load()
    Dim bRetVal As Boolean
    
    lblTitle.Caption = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\LogIn.Title", "")
    
    If lblTitle.Caption = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\LogIn.Title", "", "Lab.")

        If bRetVal = True Then
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
    End If
    
    '레지스트리에 타이틀 입력
    RegEditCurFrmTitle FSB0301.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '레지스트리에서 타이틀 제거
    Call InitRegCurFrmTitle
End Sub

Private Sub txtPassword_Change()

    If txtPassword.SelStart = txtPassword.MaxLength Then cmdOk.SetFocus
    
End Sub

Private Sub txtPassword_GotFocus()

    Call Txt_Highlight(txtPassword)
    
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdOk_Click
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
