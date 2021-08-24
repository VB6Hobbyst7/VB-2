VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmAnnounce 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "공지사항 확인"
   ClientHeight    =   5970
   ClientLeft      =   1980
   ClientTop       =   3000
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "announ16.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptInsert 
      Caption         =   "공지사항 입력시 누르세요"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   960
      Width           =   2895
   End
   Begin VB.PictureBox PicLabel 
      BackColor       =   &H00C0FFFF&
      Height          =   945
      Index           =   1
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   1725
      TabIndex        =   13
      Top             =   6750
      Width           =   1785
      Begin VB.Image ImageOFF 
         Height          =   360
         Index           =   0
         Left            =   30
         Picture         =   "announ16.frx":0442
         Top             =   480
         Width           =   360
      End
      Begin VB.Image ImageOFF 
         Height          =   360
         Index           =   1
         Left            =   450
         Picture         =   "announ16.frx":0B44
         Top             =   480
         Width           =   360
      End
      Begin VB.Image ImageOFF 
         Height          =   360
         Index           =   2
         Left            =   870
         Picture         =   "announ16.frx":1246
         Top             =   480
         Width           =   360
      End
      Begin VB.Image ImageOFF 
         Height          =   360
         Index           =   3
         Left            =   1320
         Picture         =   "announ16.frx":1948
         Top             =   480
         Width           =   360
      End
      Begin VB.Image ImageON 
         Height          =   360
         Index           =   0
         Left            =   30
         Picture         =   "announ16.frx":204A
         Top             =   60
         Width           =   360
      End
      Begin VB.Image ImageON 
         Height          =   360
         Index           =   1
         Left            =   450
         Picture         =   "announ16.frx":274C
         Top             =   60
         Width           =   360
      End
      Begin VB.Image ImageON 
         Height          =   360
         Index           =   2
         Left            =   870
         Picture         =   "announ16.frx":2E4E
         Top             =   60
         Width           =   360
      End
      Begin VB.Image ImageON 
         Height          =   360
         Index           =   3
         Left            =   1320
         Picture         =   "announ16.frx":3550
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CheckBox ChkShow 
      Caption         =   "확인된 공지사항 내용 다시 안보기"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   3630
      TabIndex        =   2
      Top             =   1290
      Value           =   1  '확인
      Width           =   4395
   End
   Begin VB.TextBox TxtAnnounce 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Top             =   1620
      Width           =   7785
   End
   Begin VB.CommandButton CmdOK 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   6570
      Picture         =   "announ16.frx":3C52
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   150
      Width           =   1365
   End
   Begin VB.PictureBox PicLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   0
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   3345
      TabIndex        =   3
      Top             =   240
      Width           =   3405
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         Caption         =   "공지대상 : ALL  (전체)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   6
         Top             =   870
         Width           =   2310
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         Caption         =   "입 력 자 : 홍길동"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   510
         Width           =   1785
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         Caption         =   "입력일자 : 1998-01-01   17:35"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   3045
      End
   End
   Begin Threed.SSPanel Panel 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   5580
      Width           =   3765
      _Version        =   65536
      _ExtentX        =   6641
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "현재 공지 사항 : 1"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   8
   End
   Begin Threed.SSPanel Panel 
      Height          =   285
      Index           =   1
      Left            =   3900
      TabIndex        =   8
      Top             =   5580
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7064
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "공지 사항 총수 : 3"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   8
   End
   Begin Threed.SSCommand CmdScroll 
      Height          =   555
      Index           =   3
      Left            =   5655
      TabIndex        =   12
      Top             =   270
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   979
      _StockProps     =   78
   End
   Begin Threed.SSCommand CmdScroll 
      Height          =   555
      Index           =   2
      Left            =   4980
      TabIndex        =   11
      Top             =   270
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   979
      _StockProps     =   78
   End
   Begin Threed.SSCommand CmdScroll 
      Height          =   555
      Index           =   1
      Left            =   4305
      TabIndex        =   10
      Top             =   270
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   979
      _StockProps     =   78
   End
   Begin Threed.SSCommand CmdScroll 
      Height          =   555
      Index           =   0
      Left            =   3630
      TabIndex        =   9
      Top             =   270
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   979
      _StockProps     =   78
   End
End
Attribute VB_Name = "FrmAnnounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, K                 As Integer
Dim nShowCount              As Integer
Dim nCurrentCount           As Integer

Dim saShowGroup()           As String
Dim saShowPerson()          As String
Dim rs                      As ADODB.Recordset
    
Private Sub CmdOK_Click()
    Dim nReturn     As Integer
    
    If OptInsert.Value = True Then
'B        Call Insert_Announce
        Exit Sub
    End If
    
    If nShowCount <> GnAnnounceGetCount Then
        nReturn = MsgBox("공지사항 내용을 모두 확인하지 않으셨습니다." & vbCrLf & _
                         "공지사항을 종료하시겠습니까 ? ", vbOKCancel, "확인")
        If nReturn = vbCancel Then Exit Sub
    End If
    If GnAnnounceGetCount <> 0 Then
        If ChkShow.Value = 1 Then Call Insert_Announce_Set
    End If
    Unload Me
    
    If MfrmMain.Caption = "외래OCS" Then
        MfrmMain.PictureM.Visible = False
        strPassOk = "OK"
        Load FrmViewSlips       'SLIP View 시 Perform을 위해 미리 Load
        FrmOrders.Show
    End If
    
End Sub


Private Sub Form_Initialize()

'    Call DbConnect1("TW_MIS_PMPA", "HOSPITAL", "av3600")
    
End Sub

Private Sub Form_Load()
    
'B    Me.Top = (Screen.Height - Me.Height) / 2 - 200
'B    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Width = 8200
    Me.Height = 6300
    Me.Show
    TxtAnnounce.Text = ""
    Labels(0).Caption = ""
    Labels(1).Caption = ""
    Labels(2).Caption = ""
    Panel(0).Caption = ""
    Panel(1).Caption = ""
    PicLabel(1).Visible = False
    
    If GnAnnounceGetCount < 2 Then
        CmdScroll(0).Enabled = False:   CmdScroll(0).Picture = ImageOFF(0).Picture
        CmdScroll(1).Enabled = False:   CmdScroll(1).Picture = ImageOFF(1).Picture
        CmdScroll(2).Enabled = False:   CmdScroll(2).Picture = ImageOFF(2).Picture
        CmdScroll(3).Enabled = False:   CmdScroll(3).Picture = ImageOFF(3).Picture
    Else
        CmdScroll(0).Enabled = False:   CmdScroll(0).Picture = ImageOFF(0).Picture
        CmdScroll(1).Enabled = False:   CmdScroll(1).Picture = ImageOFF(1).Picture
        CmdScroll(2).Enabled = True:    CmdScroll(2).Picture = ImageON(2).Picture
        CmdScroll(3).Enabled = True:    CmdScroll(3).Picture = ImageON(3).Picture
    End If
    
    If GnAnnounceGetCount < 1 Then
'B        Unload Me
        Exit Sub
    End If
    
    ReDim saShowGroup(GnAnnounceGetCount)
    ReDim saShowPerson(GnAnnounceGetCount)
    
    nShowCount = 0
    nCurrentCount = 1
    Call Memo_Show(nCurrentCount)
    
End Sub

Private Sub Insert_Announce_Set()
     
    Dim i, j            As Integer
    Dim nReturn         As Integer
    Dim strMarNo        As String           '아래 두개 배주리가 추가
    Dim strGnRetry      As String
    
    j = 0
    If ChkShow.Value = 1 Then
        For i = 1 To GnAnnounceGetCount
            If i > nShowCount Then
                Exit For
            Else
'B                GlueSetString "cMgrNo", j, GnaAnnounceMgrNOs(i)
'B                GlueSetString "cGbRetry", j, "N"
                  strMarNo = GnaAnnounceMgrNOs(i)
                  strGnRetry = "N"
            End If
            
            j = j + 1
        Next i
        adoConnect.BeginTrans
        
        strSQL = "INSERT  INTO TWOCS_ANNOUNCESET "
        strSQL = strSQL & "       (AnnounceDate, IDnumber, GbRetry, MgrNo)   "
        strSQL = strSQL & "  VALUES (TRUNC(SYSDATE ) "
        strSQL = strSQL & "          , " & VarToStr(GstrPassIDnumber)
        strSQL = strSQL & "          , " & VarToStr(strGnRetry)
        strSQL = strSQL & "          , " & VarToStr(strMarNo)        'TRUNC(SYSDATE)
        strSQL = strSQL & "          ) "
'B                 "VALUES (CONVERT(VARCHAR(10), GETDATE(), 120), " & GstrPassIDnumber & ", :cGbRetry:, :cMgrNo:)"        'TRUNC(SYSDATE)
       Result = AdoExecute(strSQL)
        
        If Result = -1 Then
            adoConnect.RollbackTrans
'B          Result = dosql("Rollback")
            MsgBox "공지사항 확인관리 TABLE INSERT ERROR!"
            Exit Sub
        End If
        adoConnect.CommitTrans
'B        Result = dosql("Commit")
        
        CmdOK.Enabled = False
        ChkShow.Enabled = False
    End If
        
End Sub

Private Sub Memo_Show(ArgInx As Integer)
    Dim strName         As String
    
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
    
    strSQL = "SELECT  NAME " & _
             " FROM TW_MIS_PMPA.TWBAS_PASS " & _
             " WHERE ProgramID = ' '        " & _
             "   AND  IDnumber = " & GstrDrCode & ""            '배주리가 변경
'b             "   AND  IDnumber = " & GnaAnnouncePerson(ArgInx)
    Result = AdoOpenSet(rs, strSQL)
    
    If rowindicator > 0 Then
        saShowPerson(ArgInx) = GnaAnnouncePerson(ArgInx) & "  " & AdoGetString(rs, "name", 0)
                                   'GlueGetString("Name", 0)
    End If

   End If
    
    Labels(0).Caption = "입력일자 : " & GsaAnnounceDateTime(ArgInx)
    Labels(1).Caption = "입 력 자 : " & saShowPerson(ArgInx)
    Labels(2).Caption = "공지대상 : " & saShowGroup(ArgInx)
    Panel(0).Caption = "현재 공지 사항 : " & ArgInx
    Panel(1).Caption = "공지 사항 총수 : " & GnAnnounceGetCount
    TxtAnnounce.Text = GsaAnnounceMemos(ArgInx)
    
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
        CmdScroll(0).Enabled = False:   CmdScroll(0).Picture = ImageOFF(0).Picture
        CmdScroll(1).Enabled = False:   CmdScroll(1).Picture = ImageOFF(1).Picture
    Else
        CmdScroll(0).Enabled = True:    CmdScroll(0).Picture = ImageON(0).Picture
        CmdScroll(1).Enabled = True:    CmdScroll(1).Picture = ImageON(1).Picture
    End If
    
    If nCurrentCount = GnAnnounceGetCount Then
        CmdScroll(2).Enabled = False:   CmdScroll(2).Picture = ImageOFF(2).Picture
        CmdScroll(3).Enabled = False:   CmdScroll(3).Picture = ImageOFF(3).Picture
    Else
        CmdScroll(2).Enabled = True:    CmdScroll(2).Picture = ImageON(2).Picture
        CmdScroll(3).Enabled = True:    CmdScroll(3).Picture = ImageON(3).Picture
    End If
    
    Call Memo_Show(nCurrentCount)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    
End Sub


Private Sub OptInsert_Click()
    CmdScroll(0).Enabled = False:   CmdScroll(0).Picture = ImageOFF(0).Picture
    CmdScroll(1).Enabled = False:   CmdScroll(1).Picture = ImageOFF(1).Picture
    CmdScroll(2).Enabled = False:   CmdScroll(2).Picture = ImageOFF(2).Picture
    CmdScroll(3).Enabled = False:   CmdScroll(3).Picture = ImageOFF(3).Picture
    
    strSQL = " SELECT TO_CHAR(SYSDATE, 'YYYY-MM-DD HH24:MI:SS') SYSDATE1 FROM DUAL"
    Result = AdoOpenSet(rs, strSQL)
    Labels(0).Caption = AdoGetString(rs, "SYSDATE1", 0)
    Labels(1).Caption = GstrPassName
    Labels(0).Caption = " ALL (전체) "
    TxtAnnounce.Text = ""
    
End Sub

