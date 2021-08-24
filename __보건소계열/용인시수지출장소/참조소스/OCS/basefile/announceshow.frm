VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmAnnounce 
   BorderStyle     =   3  '고정 대화 상자
   Caption         =   "공지사항 확인"
   ClientHeight    =   4725
   ClientLeft      =   675
   ClientTop       =   1065
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AnnounceShow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOK 
      Caption         =   "확인 (&O)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5250
      TabIndex        =   12
      Top             =   150
      Width           =   1125
   End
   Begin VB.CommandButton CmdScroll 
      Height          =   435
      Index           =   3
      Left            =   4500
      Picture         =   "AnnounceShow.frx":0442
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   150
      Width           =   435
   End
   Begin VB.CommandButton CmdScroll 
      Height          =   435
      Index           =   2
      Left            =   4050
      Picture         =   "AnnounceShow.frx":0B44
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   150
      Width           =   435
   End
   Begin VB.CommandButton CmdScroll 
      Height          =   435
      Index           =   1
      Left            =   3600
      Picture         =   "AnnounceShow.frx":1246
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   150
      Width           =   435
   End
   Begin VB.CommandButton CmdScroll 
      Height          =   435
      Index           =   0
      Left            =   3150
      Picture         =   "AnnounceShow.frx":1948
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   150
      Width           =   435
   End
   Begin Threed.SSPanel Panel 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   4260
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "  현재 공지사항 : 1"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   2
   End
   Begin VB.PictureBox PicLabel 
      Height          =   795
      Index           =   1
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   2865
      TabIndex        =   2
      Top             =   120
      Width           =   2925
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         Caption         =   "공지대상 : ALL  (전체)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   5
         Top             =   510
         Width           =   1980
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         Caption         =   "입 력 자 : 홍길동"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Width           =   1530
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         Caption         =   "입력일자 : 1998-01-01  17:35"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   2520
      End
   End
   Begin VB.CheckBox ChkShow 
      Caption         =   "확인된 공지사항 내용 다시 안보기"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   3150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   690
      Value           =   1  '선택함
      Width           =   3195
   End
   Begin VB.TextBox TxtAnnounce 
      Height          =   3165
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Width           =   6225
   End
   Begin Threed.SSPanel Panel 
      Height          =   285
      Index           =   1
      Left            =   3300
      TabIndex        =   7
      Top             =   4260
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "  공지사항 총수 : 3"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   2
   End
End
Attribute VB_Name = "FrmAnnounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, k                 As Integer
Dim nShowCount              As Integer
Dim nCurrentCount           As Integer

Dim saShowGroup()           As String
Dim saShowPerson()          As String

Private Sub Insert_Announce_Set()
    
    strSql = "INSERT  INTO TWBAS_ANNOUNCESET " & _
             "       (AnnounceDate, IDnumber, GbRetry)   " & _
             "VALUES (TRUNC(SYSDATE), " & Val(GstrPassIDnumber) & ", 'N' )"
             
    RdoDB.BeginTrans
    Result = ExecRDO(strSql)
    
    If Result = False Then
        RdoDB.RollbackTrans
        MsgBox "공지사항 확인관리 TABLE INSERT ERROR!"
        Exit Sub
    End If
    
    RdoDB.CommitTrans
    
End Sub

Private Sub Memo_Show(ArgInx As Integer)
    
    If saShowGroup(ArgInx) = "" Then
        nShowCount = nShowCount + 1
        Select Case GsaAnnounceGroup(ArgInx)
            Case "ALL ":    saShowGroup(ArgInx) = "ALL  (전체)"
            Case "OCS ":    saShowGroup(ArgInx) = "OCS  (진료부분)"
            Case "ADM ":    saShowGroup(ArgInx) = "ADM  (관리부분)"
            Case "PMPA":    saShowGroup(ArgInx) = "PMPA (원무부분)"
            Case "DEPT":    saShowGroup(ArgInx) = "DEPT (과목별)"
            Case "PERS":    saShowGroup(ArgInx) = "PERS (개인별)"
            Case Else:      saShowGroup(ArgInx) = GsaAnnounceGroup(ArgInx)
        End Select
        
        strSql = "SELECT NAME FROM TWBAS_PASS " & _
                 " WHERE ProgramID = ' '  AND  IDnumber = " & GnaAnnouncePerson(ArgInx)
        
        If OpenRDO(strSql, 0) Then
            saShowPerson(ArgInx) = GnaAnnouncePerson(ArgInx) & "  " & _
                                   RdoSet(0).rdoColumns("Name")
            RdoSet(0).Close
        End If
    End If
    
    Labels(0).Caption = "입력일자 : " & GsaAnnounceDateTime(ArgInx)
    Labels(1).Caption = "입 력 자 : " & saShowPerson(ArgInx)
    Labels(2).Caption = "공지대상 : " & saShowGroup(ArgInx)
    Panel(0).Caption = "현재 공지 사항 : " & ArgInx
    Panel(1).Caption = "공지 사항 총수 : " & GnAnnounceGetCount
    TxtAnnounce.Text = GsaAnnounceMemos(ArgInx)
    
End Sub

Private Sub CmdOK_Click()
    Dim nReturn     As Integer
    
    If nShowCount <> GnAnnounceGetCount Then
        nReturn = MsgBox("공지사항 내용을 모두 확인하지 않으셨습니다." & vbCrLf & _
                         "공지사항을 종료하시겠습니까 ? ", vbOKCancel, "확인")
        If nReturn = vbCancel Then Exit Sub
    End If
    
    If ChkShow.Value = 1 Then Call Insert_Announce_Set
    
    Unload Me
    
End Sub

Private Sub CmdScroll_Click(Index As Integer)
    
    Select Case Index
        Case 0: nCurrentCount = 1
        Case 1: nCurrentCount = nCurrentCount - 1
        Case 2: nCurrentCount = nCurrentCount + 1
        Case 3: nCurrentCount = GnAnnounceGetCount
    End Select
    
    If nCurrentCount < 1 Then nCurrentCount = 1
    If nCurrentCount > GnAnnounceGetCount Then nCurrentCount = GnAnnounceGetCount
    
    If nCurrentCount = 1 Then
        CmdScroll(0).Enabled = False
        CmdScroll(1).Enabled = False
    Else
        CmdScroll(0).Enabled = True
        CmdScroll(1).Enabled = True
    End If
    
    If nCurrentCount = GnAnnounceGetCount Then
        CmdScroll(2).Enabled = False
        CmdScroll(3).Enabled = False
    Else
        CmdScroll(2).Enabled = True
        CmdScroll(3).Enabled = True
    End If
    
    Call Memo_Show(nCurrentCount)
    
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2 - 200
    Me.Left = (Screen.Width - Me.Width) / 2
    
    TxtAnnounce.Text = ""
    Labels(0).Caption = ""
    Labels(1).Caption = ""
    Labels(2).Caption = ""
    Panel(0).Caption = ""
    Panel(1).Caption = ""
    
    If GnAnnounceGetCount < 2 Then
        CmdScroll(0).Enabled = False
        CmdScroll(1).Enabled = False
        CmdScroll(2).Enabled = False
        CmdScroll(3).Enabled = False
    Else
        CmdScroll(0).Enabled = False
        CmdScroll(1).Enabled = False
    End If
    
    If GnAnnounceGetCount < 1 Then
        Unload Me
        Exit Sub
    End If
    
    ReDim saShowGroup(GnAnnounceGetCount)
    ReDim saShowPerson(GnAnnounceGetCount)
    
    nShowCount = 0
    nCurrentCount = 1
    Call Memo_Show(nCurrentCount)
    
End Sub
