VERSION 5.00
Begin VB.Form Frm기본_공정코드입력2006 
   BorderStyle     =   1  '단일 고정
   Caption         =   "공정코드 입력"
   ClientHeight    =   1995
   ClientLeft      =   1290
   ClientTop       =   4785
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm기본_공정코드입력2006.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6135
   Begin VB.TextBox Txt공정코드 
      Height          =   315
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "1234"
      Top             =   540
      Width           =   465
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "저장(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4320
      Picture         =   "Frm기본_공정코드입력2006.frx":000C
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5220
      Picture         =   "Frm기본_공정코드입력2006.frx":08D6
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin VB.Frame Fra입력 
      Height          =   975
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   6015
      Begin VB.ComboBox Cbo사용여부 
         Height          =   300
         Left            =   1260
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox Txt공정명 
         Height          =   315
         IMEMode         =   10  '한글 
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "12345678901234567890123456789012345678901234567890"
         Top             =   240
         Width           =   4665
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "공정명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "사용여부"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   780
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "공정코드"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   600
      Width           =   780
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   6060
      Y1              =   900
      Y2              =   900
   End
End
Attribute VB_Name = "Frm기본_공정코드입력2006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDCK_CANCEL()
    Txt공정코드 = ""             '/공정코드
    Call MDCK_KEY_CLEAR
End Sub

Private Sub MDCK_DELETE()

End Sub

Private Sub MDCK_INITIAL()
    '/사용여부
    Cbo사용여부.Clear
    Cbo사용여부.AddItem "1.사용함" & Space(100) & "1"
    Cbo사용여부.AddItem "2.사용안함" & Space(100) & "2"
    
    Call MDCK_CANCEL
End Sub

Private Sub MDCK_KEY_CLEAR()
    Fra입력.Enabled = False
    
    Txt공정명 = ""           '/공정명
    Cbo사용여부.ListIndex = 0 '/사용여부
End Sub

Private Sub MDCK_PRINT()

End Sub

Private Sub MDCK_SAVE()
    loAdoCnn.BeginTrans

    StrSQL = "         SELECT * "
    StrSQL = StrSQL & "  FROM BAG_GONGJUNG "
    StrSQL = StrSQL & " WHERE GONGJUNG_CODE = '" & Trim(Txt공정코드) & "' "
    If ReadADO(StrSQL, 0) = True Then
        Call CloseADO(ARS(0))
        
        StrSQL = "         UPDATE BAG_GONGJUNG SET "
        StrSQL = StrSQL & "       GONGJUNG_NAME = '" & Trim(TEXT_LSET(Trim(Txt공정명), 50)) & "', "
        StrSQL = StrSQL & "       GONGJUNG_USE  = '" & Trim(Right(Cbo사용여부, 10)) & "' "
        StrSQL = StrSQL & " WHERE GONGJUNG_CODE = '" & Trim(Txt공정코드) & "' "
        If RunADO(StrSQL) = False Then Exit Sub
    Else
        StrSQL = " INSERT INTO BAG_GONGJUNG "
        StrSQL = StrSQL & " (GONGJUNG_CODE, GONGJUNG_NAME, GONGJUNG_USE) "
        StrSQL = StrSQL & " VALUES "
        StrSQL = StrSQL & "  ('" & Trim(TEXT_LSET(Trim(Txt공정코드), 4)) & "', "
        StrSQL = StrSQL & "   '" & Trim(TEXT_LSET(Trim(Txt공정명), 50)) & "', "
        StrSQL = StrSQL & "   '" & Trim(Right(Cbo사용여부, 10)) & "') "
        If RunADO(StrSQL) = False Then Exit Sub
    End If
    
    loAdoCnn.CommitTrans
    GstrInputUpdateYN = "1"
End Sub

Private Sub MDCK_VIEW()
    If Trim(Txt공정코드) = "" Then Exit Sub
    
    StrSQL = "         SELECT * "
    StrSQL = StrSQL & "  FROM BAG_GONGJUNG "
    StrSQL = StrSQL & " WHERE GONGJUNG_CODE  = '" & Trim(Txt공정코드) & "' "
    If ReadADO(StrSQL, 0) = True Then
        If GstrInputUpdate = "2" Then '/1.Input, 2.Update
            Fra입력.Enabled = True
            
            Txt공정명 = Trim(ARS(0)!GONGJUNG_NAME & "") '/공정명
            Call SET_COMBO_DATA_R(Trim(ARS(0)!GONGJUNG_USE & ""), Cbo사용여부) '/사용여부
        Else
            MsgBox "기존자료가 있습니다!", vbInformation, "확인"
        End If
        Call CloseADO(ARS(0))
    Else
        If GstrInputUpdate = "1" Then '/1.Input, 2.Update
            Fra입력.Enabled = True
            Txt공정명.SetFocus
        End If
    End If
End Sub

Private Sub Cbo사용여부_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub CmdQuit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If Trim(Txt공정코드) = "" Then MsgBox "공정코드를 (재)입력하십시오!", vbInformation, "확인": Txt공정코드.SetFocus: Exit Sub
    
    If MsgBox("저장하겠습니까?", vbQuestion + vbOKCancel, "저장질의") = vbCancel Then Exit Sub
    
    Call MDCK_SAVE
    
    MsgBox "저장되었습니다!", vbInformation, "확인"
    
    Call MDCK_CANCEL
    
    If GstrInputUpdate = "1" Then '/1.Input, 2.Update
        Txt공정코드.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Height = 2505
    Me.Width = 6255
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call MDCK_INITIAL

    If GstrInputUpdate = "2" Then '/1.Input, 2.Update
        Txt공정코드 = GstrArgTemp1
        Txt공정코드.BackColor = RGB(255, 255, 240)
        Txt공정코드.Enabled = False
        Call MDCK_VIEW
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Frm기본_공정코드입력2006 = Nothing
End Sub

Private Sub Txt공정코드_Change()
    Call MDCK_KEY_CLEAR
End Sub

Private Sub Txt공정코드_GotFocus()
    Call TEXTSELECT
End Sub

Private Sub Txt공정코드_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MDCK_VIEW
End Sub

Private Sub Txt공정명_GotFocus()
    Call TEXTSELECT
End Sub

Private Sub Txt공정명_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
