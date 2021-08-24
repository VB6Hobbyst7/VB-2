VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "(주) 대련엠티에스"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":16CA
   ScaleHeight     =   4335
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6030
      Top             =   1110
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5970
      Top             =   450
   End
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   1380
   End
   Begin VB.Frame fraWait 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   660
      TabIndex        =   2
      Top             =   2460
      Width           =   5445
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   285
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   660
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   360
      Top             =   2670
      Width           =   615
   End
   Begin VB.Label LblStatus 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   1140
      TabIndex        =   4
      Top             =   2790
      Width           =   1545
   End
   Begin VB.Image imgOK 
      Height          =   420
      Left            =   5520
      MousePointer    =   99  '사용자 정의
      ToolTipText     =   "확인"
      Top             =   3180
      Width           =   555
   End
   Begin VB.Image imgCancel 
      Height          =   420
      Left            =   4920
      MousePointer    =   99  '사용자 정의
      ToolTipText     =   "취소"
      Top             =   3180
      Width           =   555
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   9
      Left            =   2040
      Picture         =   "frmSplash.frx":E518
      Top             =   4950
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   8
      Left            =   1560
      Picture         =   "frmSplash.frx":E822
      Top             =   4950
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   7
      Left            =   1080
      Picture         =   "frmSplash.frx":EB2C
      Top             =   4950
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   6
      Left            =   600
      Picture         =   "frmSplash.frx":EE36
      Top             =   4950
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   5
      Left            =   90
      Picture         =   "frmSplash.frx":F140
      Top             =   4950
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   4
      Left            =   2025
      Picture         =   "frmSplash.frx":F44A
      Top             =   4470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   3
      Left            =   1530
      Picture         =   "frmSplash.frx":F754
      Top             =   4470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   2
      Left            =   1035
      Picture         =   "frmSplash.frx":FA5E
      Top             =   4470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   1
      Left            =   540
      Picture         =   "frmSplash.frx":FD68
      Top             =   4470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   0
      Left            =   90
      Picture         =   "frmSplash.frx":10072
      Top             =   4470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   900
      TabIndex        =   1
      Top             =   3180
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "SemiLis 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   0
      Top             =   2100
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   390
      Picture         =   "frmSplash.frx":1037C
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   0
      Picture         =   "frmSplash.frx":10C46
      Top             =   0
      Width           =   1950
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_iStep As Integer
Private m_iIconPtr As Integer

' DB Connect
Public adoConn As adoDb.Connection
Public adoRs   As adoDb.Recordset


Public Sub SetImg(ByRef imgControl As Object, ByVal strImgName As String)
    imgControl.Picture = LoadResPicture(strImgName, vbResBitmap)
    imgControl.MouseIcon = LoadResPicture("Point", vbResCursor)
End Sub

Private Sub Command1_Click()
    Unload Me
    frmSplash2.Show

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
        
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
    Me.Height = 0: Me.Width = 0
    
    Call SetImg(imgCancel, "Cancel")
    Call SetImg(imgOK, "OK")

'    fraWait.Visible = False

    Timer1.Enabled = True
    Timer1.Interval = 500

    Timer2.Enabled = True
    Timer2.Interval = 1000
        
    Timer3.Enabled = True
    Timer3.Interval = 100
    
    m_iIconPtr = 0

    'lblMsg = "Connecting Configure Database(Semilis.mdb)..."
    lblMsg = "데이터베이스에 연결하고 있습니다(Semilis.mdb)..."
    'lblRight = "※본 프로그램을 무단으로 복제하거나 배포하면 법적처벌을 받습니다.   본프로그램은 주식회사 대련엠티에스에 권한이 있습니다."
    DoEvents

End Sub

Private Sub imgCancel_Click()
    Unload Me
End Sub

Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgCancel, "Cancel_d")
End Sub

Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgCancel, "Cancel")
End Sub

Private Sub imgOK_Click()

On Error GoTo Error_Hnd
    
    Set adoRs = New Recordset


    adoRs.Close

    If Not adoRs Is Nothing Then
        Set adoRs = Nothing
    End If
    
    Exit Sub

Error_Hnd:
    
    MsgBox "Error 발  생" & Chr(13) & _
           "Error 번  호 : " & Err.Number & Chr(13) & _
           "Error 메세지 : " & Err.Source, vbOKOnly, Me.Caption
           
    If Not adoRs Is Nothing Then
        Set adoRs = Nothing
    End If

End Sub

Private Sub imgOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgOK, "OK_d")
End Sub

Private Sub imgOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetImg(imgOK, "OK")
End Sub

Private Sub Timer1_Timer()
    Static bFlag As Boolean
    Dim sSymbols As String
    If bFlag Then
        sSymbols = " ▷▶ "
    Else
        sSymbols = " ▶▷ "
    End If
    
    Me.Caption = "☞ " & CStr(prgProgress.Value) & sSymbols & CStr(prgProgress.Max)
    bFlag = Not bFlag
                
   #If Win16 Then
      Me.Icon = imgPtr(m_iIconPtr + 10)
   #ElseIf Win32 Then
      Me.Icon = imgPtr(m_iIconPtr)
   #End If
    
'    LblStatus.Caption = "☞ " & CStr(prgProgress.Value) & sSymbols & CStr(prgProgress.Max)
'
'   #If Win16 Then
'      Image3.Picture = imgPtr(m_iIconPtr + 10)
'   #ElseIf Win32 Then
'      Image3.Picture = imgPtr(m_iIconPtr)
'   #End If
   
   m_iIconPtr = (m_iIconPtr + 1) Mod 10


End Sub

Private Sub Timer2_Timer()
    m_iStep = m_iStep + 1
    
    If prgProgress.Value + m_iStep >= 100 Then
        prgProgress.Value = 100
        prgProgress.Enabled = False
            
        Me.Caption = "☞ " & CStr(prgProgress.Value) & " ▶▷ " & CStr(prgProgress.Max)

        MsgBox "DB에 성공적으로 연결되었습니다", vbOKOnly, Me.Caption
        Me.Hide
        
        Timer1.Enabled = False
        Timer2.Enabled = False
        End
        'Exit Sub
    End If
    
    prgProgress.Value = prgProgress.Value + m_iStep
    If prgProgress.Value >= prgProgress.Max Then
        prgProgress.Enabled = False
        Me.Hide
    End If

End Sub

Private Sub Timer3_Timer()
    
    If Me.Width < 6840 Then
        Me.Width = Me.Width + 600
    Else
        Me.Width = 6840
        Timer3.Enabled = False
        Timer4.Enabled = True
        Timer4.Interval = 100
    End If

End Sub

Private Sub Timer4_Timer()
    
    If Me.Height < 4605 Then
        Me.Height = Me.Height + 600
    Else
        Me.Height = 4605
        Timer4.Enabled = False
    End If

End Sub
