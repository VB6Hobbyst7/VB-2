VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form INTface41_OLD 
   BorderStyle     =   0  '없음
   Caption         =   "해당일의 검사결과 받아 보기"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   1200
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7500
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread spdface 
      Height          =   5970
      Left            =   0
      OleObjectBlob   =   "MACHINE41_old.frx":0000
      TabIndex        =   2
      Top             =   600
      Width           =   11925
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   645
      Left            =   4380
      TabIndex        =   23
      Top             =   6630
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "삭                제"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "MACHINE41_old.frx":046E
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   495
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   512
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin Threed.SSCommand cmdclose 
      Height          =   645
      Left            =   7380
      TabIndex        =   1
      Top             =   6630
      Width           =   3060
      _Version        =   65536
      _ExtentX        =   5397
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "닫        기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "MACHINE41_old.frx":1348
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   645
      Left            =   1350
      TabIndex        =   0
      Top             =   6630
      Visible         =   0   'False
      Width           =   3060
      _Version        =   65536
      _ExtentX        =   5397
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "결과       등록"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "MACHINE41_old.frx":179A
   End
   Begin Threed.SSCommand cmdOrder 
      Height          =   645
      Left            =   1350
      TabIndex        =   11
      Top             =   6630
      Width           =   3060
      _Version        =   65536
      _ExtentX        =   5397
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "Order       내리기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "MACHINE41_old.frx":2264
   End
   Begin Threed.SSFrame fmeGuide 
      Height          =   585
      Left            =   9720
      TabIndex        =   21
      Top             =   0
      Width           =   2085
      _Version        =   65536
      _ExtentX        =   3678
      _ExtentY        =   1032
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtguide 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   165
         Width           =   1845
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   585
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5625
      _Version        =   65536
      _ExtentX        =   9922
      _ExtentY        =   1032
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "검사 인터페이스 작업을 수행합니다!!"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1590
         TabIndex        =   10
         Top             =   240
         Width           =   3585
      End
      Begin VB.Label LblMMDD 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
   End
   Begin Threed.SSFrame fmeSlipSeq 
      Height          =   585
      Left            =   5610
      TabIndex        =   3
      Top             =   0
      Width           =   4125
      _Version        =   65536
      _ExtentX        =   7276
      _ExtentY        =   1032
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txslipno 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   150
         Width           =   795
      End
      Begin VB.TextBox txsapno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3060
         TabIndex        =   5
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "일련번호"
         Height          =   210
         Left            =   2160
         TabIndex        =   7
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "Slip No."
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   255
         Width           =   945
      End
   End
   Begin Threed.SSFrame fmeOrder 
      Height          =   5460
      Left            =   30
      TabIndex        =   12
      Top             =   600
      Width           =   6360
      _Version        =   65536
      _ExtentX        =   11218
      _ExtentY        =   9631
      _StockProps     =   14
      Caption         =   "WORKLIST  작성"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdWorklist2 
         Height          =   4590
         Left            =   6630
         OleObjectBlob   =   "MACHINE41_old.frx":2D2E
         TabIndex        =   20
         Top             =   420
         Width           =   3165
      End
      Begin FPSpread.vaSpread spdWorklist1 
         Height          =   4590
         Left            =   2760
         OleObjectBlob   =   "MACHINE41_old.frx":10709
         TabIndex        =   13
         Top             =   420
         Width           =   3195
      End
      Begin VB.TextBox txtSId 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   16
         Top             =   1035
         Width           =   1110
      End
      Begin VB.TextBox txtPos 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1290
         MaxLength       =   2
         TabIndex        =   15
         Top             =   765
         Width           =   825
      End
      Begin VB.TextBox txtRackNo 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   14
         Top             =   495
         Width           =   825
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   19
         Top             =   795
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rack No."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   18
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "SampleID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   17
         Top             =   1065
         Width           =   1005
      End
   End
End
Attribute VB_Name = "INTface41_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private transcfg        As commset
Dim errfound            As Integer
Dim Porttag             As Integer
Dim RcvBuffer           As String
Dim TestNameTable(30)   As TestNameTbl
Dim nowcount            As Integer

Dim EmerCnt             As Integer
Dim sampleCnt           As Integer
Dim EditCnt             As Integer  'Express처럼 한 슬립이 두 프레임 이상 거쳐 나올 경우 사용
Dim PrevCnt             As Integer
Dim Startslip           As String

Dim StartBCol           As Integer
Dim EndBCol             As Integer
Dim StartBRow           As Integer
Dim EndBRow             As Integer
Dim identbOpenKey       As Integer
Dim IdList              As String
Dim BlockKey            As Integer
Dim Errkey              As Integer

'public에 slipno As String 선언
Dim phase               As Integer
Dim bufcnt              As Integer
Dim wkbuf               As String
Dim ix1                 As Integer
Dim PrevReq             As Integer

Dim Test_OpenFlag        As Integer

Sub add_db_identb(sample As String, slip As String)
   
   
   If RecordExist(identb, "PrimaryKey", sample) Then
      identb.Edit
   Else
      identb.AddNew
      identb!seq_no = sample
   End If
   
   identb!slip_no = slip
   identb.Update
   identb.MoveLast

End Sub

Sub add_db_resulttb(sample2 As String, tcd As String, trt As String)
   
   tcd = Right$(tcd, 2)
   If RecordExists(resulttb, "PrimaryKey", sample2, tcd) Then
      resulttb.Edit
   Else
      resulttb.AddNew
      resulttb!seq_no = sample2
      resulttb!testcode = tcd
   End If
   resulttb!TestResult = trt
   resulttb.Update
   resulttb.MoveLast

End Sub

Function FindIdList(position As Integer) As Integer
     FindIdList = InStr(position + 1, RcvBuffer, Chr(13))
     IdList = Mid$(RcvBuffer, FindIdList + 1, 1)
End Function

Sub PhaseCfg_Protocol()
    
    Dim wkdat As String
    Dim ix1 As Integer
    
    For ix1 = 1 To Len(wkbuf)

        wkdat = Mid$(wkbuf, ix1, 1)

            Select Case phase

                Case 1
                    Select Case Asc(wkdat)
                        Case 2
                            phase = 2
                    End Select

                Case 2
                    Select Case Asc(wkdat)
                        Case 3
                            Call edit_data
                            RcvBuffer = ""
                            phase = 1

                        Case Else
                            RcvBuffer = RcvBuffer & wkdat
                            phase = 2
                    End Select

            End Select

    Next ix1
        
End Sub

Private Function RecordExist(Tb As Recordset, IndexName As String, Samp As String) As Integer
         
         Dim CurrRecord As Variant

         If Tb.RecordCount < 1 Or Tb.BOF Or Tb.EOF Then
            RecordExist = False
            Exit Function
         End If

         '''CurrRecord = Tb.Bookmark
         Tb.MoveFirst
         Tb.Index = IndexName
         Tb.Seek "=", Samp

         If Tb.NoMatch Then
            '''Tb.Bookmark = CurrRecord
            RecordExist = False
         Else
            RecordExist = True
         End If
         
End Function
Private Function RecordExists(Tb As Recordset, IndexName As String, samp2 As String, tcd2 As String) As Integer
         Dim CurrRecord As Variant

         If Tb.RecordCount < 1 Or Tb.BOF Or Tb.EOF Then
            RecordExists = False
            Exit Function
         End If

         '''CurrRecord = Tb.Bookmark
         Tb.MoveFirst
         Tb.Index = IndexName
         Tb.Seek "=", samp2, tcd2

         If Tb.NoMatch Then
            '''Tb.Bookmark = CurrRecord
            RecordExists = False
         Else
             RecordExists = True
         End If

End Function

Sub edit_data()
        
'Dim slipno      As String
    
    Dim SeqNo  As String
    Dim tresult As String
    Dim tcode   As String
    Dim i       As Integer
    Dim Colldx  As Integer
    Dim TestNm  As String
    Dim TestNo  As Integer
    Dim ix1     As Integer

    ' Get Test Name
    Colldx = InStr(RcvBuffer, "{")
    ' Get Result
    If Mid$(RcvBuffer, Colldx + 1, 1) = "R" Then
       slipno = Trim$(Mid$(RcvBuffer, Colldx + 10, 12))        '  SLipNo.
       slipno = Right$("0000" + slipno, 4)
       
       nowcount = nowcount + 1
       SeqNo = Format(nowcount, "0000")
       
       ' Get Test Count
       TestNo = Val(Trim$(Mid$(RcvBuffer, Colldx + 58, 2)))
       
       Call Row_Plus(spdface)
       
       For ix1 = 1 To TestNo
       
           TestNm = Trim(Mid$(RcvBuffer, (ix1 - 1) * 43 + Colldx + 61, 4))
           tresult = Trim$(Mid$(RcvBuffer, (ix1 - 1) * 43 + 67, 7))
        
           tcode = ""
           For i = 1 To 30
                If Trim$(TestNameTable(i).Name) = Trim$(TestNm) And Trim$(TestNm) <> "" Then
                    tcode = Format$(i, "00")
                    Exit For
                End If
           Next i

           If tcode <> "" Then
             add_db_identb SeqNo, slipno
             add_db_resulttb SeqNo, tcode, tresult
           End If
           'RowPtr = RowPtr + 1                   ' Row Position Increament
           identbOpenKey = True
           
           Call spdsettext(spdface, Val(tcode) + 1, spdface.Row, tresult & "")
       
       Next
       
        txsapno.Text = SeqNo
        txslipno.Text = slipno
        txtguide.Text = "TX Data!!"
        
        Call spdsettext(spdface, 1, spdface.Row, slipno & "")
        
    End If

End Sub
Sub Test()
    
    Dim rv%
    
    Test_OpenFlag = 1
    
    Open "c:\Lab\interfac\Minos\mindmp.log" For Input As #3
    Test_OpenFlag = 2
    
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop
    
    Close #3
    
    Call PhaseCfg_Protocol
    
End Sub

Private Sub cmdclose_Click()

        Unload Me
        FrmFlag = 0
End Sub

Private Sub cmdDelete_Click()
    Dim rv As Integer
    Dim i  As Integer
    Dim CurrentTbRows As Integer
    Dim ExistTxtKey As Integer
    Dim tmpSlip
    
    If StartBRow = -1 And EndBRow = -1 Then
        StartBRow = 1
        EndBRow = nowcount - PrevCnt
    End If
    
    For i = StartBRow To EndBRow
        rv = spdface.GetText(1, i, tmpSlip)
        If tmpSlip = "" Then
            ExistTxtKey = False
            Exit For
        Else
            ExistTxtKey = True
        End If
    Next
    
    If identbOpenKey = True And ExistTxtKey = True Then
        If StartBCol = -1 And EndBCol = -1 And BlockKey = True Then
            
            rv = MsgBox("블록으로 지정된 Slip을 삭제하시겠습니까?", 4, Title & "  " & "Slip No. 삭제 확인!!")
            If rv = 7 Then
                BlockKey = False
                spdface.EditMode = True
                spdface.EditMode = False
                cmdclose.SetFocus
                Exit Sub
            End If
            
            identb.Index = "primarykey"
            resulttb.Index = "seq_no"
            
            identb.MoveLast
            CurrentTbRows = sampleCnt
            
            For i = Val(StartBRow) To Val(EndBRow)
                identb.Seek "=", Format$(i + PrevCnt, "0000")
                identb.Delete
                       
                resulttb.Seek "=", Format$(i + PrevCnt, "0000")
                
                If resulttb.NoMatch = False Then
                   Do Until resulttb.EOF
                       If resulttb!seq_no <> Format$(i + PrevCnt, "0000") Then Exit Do
                       
                       resulttb.Delete
                            
                       resulttb.MoveNext
                   Loop
                End If
            Next
            
            If EndBRow <> CurrentTbRows Then
                For i = EndBRow + 1 To CurrentTbRows
                    
                    identb.Seek "=", Format$(i + PrevCnt, "0000")
                    identb.Edit
                    identb!seq_no = Format$(i + PrevCnt - (EndBRow - StartBRow + 1), "0000")
                    identb.Update
                    
                    resulttb.Seek "=", Format(i + PrevCnt, "0000")
                    
                    If resulttb.NoMatch = False Then
                       Do Until resulttb.EOF
                           If resulttb!seq_no <> Format(i + PrevCnt, "0000") Then Exit Do
                           
                           resulttb.Edit
                           resulttb!seq_no = Format$(i + PrevCnt - (EndBRow - StartBRow + 1), "0000")
                           resulttb.Update
                           
                           resulttb.MoveNext
                       Loop
                    End If
                    
                Next
            End If
            
        '삭제하는 Spread 라인의 텍스트를 지움.
            spdface.BlockMode = True
            spdface.Col = -1
            spdface.Col2 = -1
            spdface.Row = StartBRow
            spdface.Row2 = EndBRow
            spdface.Action = SS_ACTION_DELETE_ROW
            spdface.BlockMode = False

        '1st Column(SlipNo)의 색깔을 노란색
            spdface.BlockMode = True
            spdface.Col = 1
            spdface.Col2 = 1
            spdface.Row = -1
            spdface.Row2 = -1
            spdface.BackColor = &HC0FFFF
            spdface.BlockMode = False
            
            nowcount = nowcount - (EndBRow - StartBRow + 1)   '삭제한 후에 다시 삭제할 경우를 대비
            sampleCnt = sampleCnt - (EndBRow - StartBRow + 1)   '삭제한 후에 다시 삭제할 경우를 대비
            txslipno = ""
            txsapno = Format(nowcount, "0000")
            txtguide = "Data 삭제!!"
            
        Else
        
            MsgBox "잘못된 삭제 방법입니다." & Chr(10) & "왼쪽의 회색빛 헤더부분을 클릭하거나 끌어서 해당줄의 전체가 어두워지게 한 후," & Chr(10) & "삭제를 하십시요!!"
        
        End If
   Else
   
        MsgBox "데이터가 없거나 검사 결과 전송을 받지 않으셨습니다!!"
        
   End If

   BlockKey = False
   spdface.EditMode = True
   spdface.EditMode = False
   
'현재의 Row를 점검
   spdface.Row = sampleCnt
   
End Sub
Private Sub cmdOrder_Click()
    
    Call Send_Order
    
    fmeOrder.Visible = False
    cmdOrder.Visible = False
    
    cmdReg.Visible = True
    cmdDelete.Visible = True
    spdface.Visible = True
    fmeSlipSeq.Visible = True
    
    fmeGuide.Left = 9720
    fmeGuide.Top = 0
    fmeGuide.Width = 2115
    fmeGuide.Height = 585

    txtguide.Left = 150
    txtguide.Top = 180
    txtguide.Width = 1815
    txtguide.Height = 315
    
    txtguide.Text = "RX Results!!"
    
End Sub

Private Sub cmdReg_Click()
    
    Dim i%, rt%, seqnoVar, slipnoVar, tcode$, tresult$
    Dim tmpSlip
    Dim ExistTxtKey As Integer
    
    If StartBRow = -1 And EndBRow = -1 Then
        StartBRow = 1
        EndBRow = nowcount - PrevCnt
    End If
    
    For i = StartBRow To EndBRow
        rt = spdface.GetText(1, i, tmpSlip)
        If tmpSlip = "" Then
            ExistTxtKey = False
            Exit For
        Else
            ExistTxtKey = True
        End If
    Next
    
    If identbOpenKey = True And ExistTxtKey = True Then
        If StartBCol = -1 And EndBCol = -1 And BlockKey = True Then
            identb.Index = "primarykey"
            identb.MoveFirst
                                  
            Screen.MousePointer = 11
            'pnlResult.Visible = True
            'Timer1.Enabled = True
            'rotateflag = 1
            
            'DoEvents
           
            Open "C:\Lab\interfac\" & machstr & "\" & machstr & ".txt" For Append As #7
                 
            For i = Val(StartBRow) To Val(EndBRow)
                
                       
                'DoEvents
                
                   'rt = spdface.GetText(1, i, seqnoVar)
                   rt = spdface.GetText(1, i, slipnoVar)
                                
                   identb.Index = "primarykey"
                   identb.Seek "=", Format(i + PrevCnt, "0000")
                          
                   If identb.NoMatch = False Then
                       identb.Edit
                       identb!chkresult = "*"
                       identb.Update
                       spdface.Col = 1
                       spdface.Col2 = 1
                       spdface.Row = i
                       spdface.Row2 = i
                       spdface.BlockMode = True
                       spdface.BackColor = &HFFFFC0
                       spdface.BlockMode = False
                   End If

' SQL문을 사용(Order By)하여 결과 등록할 때의 구문

'''      SQLQ = "SELECT testcode, testresult"
'''      SQLQ = SQLQ & " FROM sp_result"
'''      SQLQ = SQLQ & " WHERE seq_no = " & """" & idnum & """"
'''      SQLQ = SQLQ & " ORDER BY testcode"
'''      Set Db5 = OpenDatabase(filename & "comm\" & strmmdd + ".mdb")
'''      Set snapData = Db5.OpenRecordset(SQLQ, dbOpenSnapshot) '"select TestCode from cm_result where SampleNo = ""0001""")
'''      If snapData.EOF = False Then
'''         snapData.MoveFirst
'''         Do While snapData.EOF = False
                   
                   resulttb.Index = "seq_no"
                   resulttb.Seek "=", Format(i + PrevCnt, "0000")
                   
                   If resulttb.NoMatch = False Then
                       Do Until resulttb.EOF
                           If resulttb!seq_no <> identb!seq_no Then Exit Do
                           tcode = TestNameTable(Val(resulttb!testcode)).code
                           tresult = resulttb!TestResult
            
            ''LAB MANAGER에게 텍스트파일 형태로 결과등록
                           If Trim$(slipnoVar) <> "" Then
                                Print #7, slipnoVar & tcode & Left(Format(Now, "YYYYMMDD"), 4) & "-" & Left$(textmmdd, 2) & "-" & Right$(textmmdd, 2) & tresult & Chr(13) & Chr(10);
                           End If
            
                           resulttb.MoveNext
                       Loop
                   End If
                             
               Screen.MousePointer = 0
           Next
           
           'pnlResult.Visible = False
           'Timer1.Enabled = False
           'rotateflag = 0
           Close #7
        
        Else
        
            MsgBox "잘못된 결과 등록 방법입니다." & Chr(10) & "왼쪽의 회색빛 헤더부분을 클릭하거나 끌어서 해당줄의 전체가 어두워지게 한 후," & Chr(10) & "결과 등록을 하십시요!!"
        
        End If
   Else
   
        MsgBox "데이터가 없거나 검사 결과를 전송받지 않으셨습니다!!"
        
   End If
   
   BlockKey = False
   spdface.EditMode = True
   spdface.EditMode = False
   
'현재의 Row를 점검
   spdface.Row = sampleCnt
   
End Sub

Private Sub Comm1_OnComm()
    
    Screen.MousePointer = 11
    'Dim wkbuf As String
    
    
    Select Case Comm1.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                               ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
                            
            wkbuf = Comm1.Input
                                            
            Print #1, wkbuf;    'Test
           
            Call PhaseCfg_Protocol
                    
            txtguide.Text = "Data 전송 완료!!"
                
        Case MSCOMM_EV_CTS      'j
        Case MSCOMM_EV_DSR      ' Change in the DSR line.
        Case MSCOMM_EV_CD       ' Change in the CD line.
        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
        ' Errors
        Case MSCOMM_ER_BREAK    ' A Break was received.
        ' Code to handle a BREAK goes here, and so on.
        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
        Case MSCOMM_ER_FRAME    ' Framing Error.
        Case MSCOMM_ER_OVERRUN  ' Data Lost.
        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
        Case MSCOMM_ER_RXPARITY ' Parity Error.
        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
    End Select
           
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    'form을 가운데에 위치
    Me.Top = 0
    Me.Left = 0
    Me.Height = INTmain00.Height - INTmain00.pnlMain.Height - 500
    Me.Width = INTmain00.Width - 200
    
    Dim tablerows As Integer
    Dim sRow As Integer
    Dim i As Integer
    Dim TestItemNo As Integer
    
        
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")
    
    tbcode.MoveLast
    tablerows = tbcode.RecordCount
        
    tbcode.MoveFirst
   
    For sRow = 1 To 30
        TestNameTable(sRow).Name = tbcode!Name & ""
        TestNameTable(sRow).code = tbcode!code & ""
        If TestNameTable(sRow).code <> "" Then
            TestItemNo = TestItemNo + 1
        End If
        tbcode.MoveNext
    Next
    
    If TestItemNo >= 17 Then
        spdface.MaxCols = TestItemNo + 1
        For i = 1 To TestItemNo
            Call spdsettext(spdface, i + 1, 0, TestNameTable(i).Name)
        Next
    Else
        spdface.MaxCols = 18
        For i = 1 To TestItemNo
            Call spdsettext(spdface, i + 1, 0, TestNameTable(i).Name)
        Next
        
        For i = TestItemNo + 1 To spdface.MaxCols - 1
            Call spdsettext(spdface, i + 1, 0, "-")
        Next
    End If
    
'1st Column(SlipNo)의 색깔을 노란색
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.Col2 = 1
    spdface.Row = -1
    spdface.Row2 = -1
    spdface.BackColor = &HC0FFFF
    spdface.BlockMode = False

'Interface Result를 일단 Lock
    spdface.BlockMode = True
    spdface.Col = 1
    spdface.Col2 = spdface.MaxCols
    spdface.Row = 1
    spdface.Row2 = spdface.MaxRows
    spdface.Lock = True
    spdface.BlockMode = False

'Spread.Row Initialization
    spdface.Row = 0
    
    tbcode.Close
    dbcode.Close
    
    LblMMDD.Caption = Val(Left$(textmmdd, 2)) & "월" & " " & Val(Right$(textmmdd, 2)) & "일"
    
        
    If OrderKey = True Then
        fmeOrder.Visible = True
        
        spdface.Visible = False
        cmdReg.Visible = False
        cmdDelete.Visible = False
        fmeSlipSeq.Visible = False
        
        For i = 1 To 30
                       
            If i < 16 Then
                fmeOrder.Width = 6360
                Call spdChksettext(spdWorklist1, 1, i, TestNameTable(i).code)
                Call spdsettext(spdWorklist1, 2, i, TestNameTable(i).Name)
            Else
                If TestNameTable(i).code <> "" Or TestNameTable(i).Name <> "" Then
                    fmeOrder.Width = 10260
                    Call spdChksettext(spdWorklist2, 1, i - 15, TestNameTable(i).code)
                    Call spdsettext(spdWorklist2, 2, i - 15, TestNameTable(i).Name)
                End If
            End If
        
        Next i
        
        fmeGuide.Left = 5610
        fmeGuide.Top = 0
        fmeGuide.Width = 6240
        fmeGuide.Height = 585
        
        txtguide.Width = 5000
        txtguide.Text = "Worklist & Order Process"
        
    Else
        fmeOrder.Visible = False
        cmdOrder.Visible = False
        
        txtguide.Text = "RX Results!!"
    End If

    
    On Error GoTo PortOpenErr:
    errfound = False
    
    Set dbcomm = OpenDatabase(filename & commstr)
    Set tbcomm = dbcomm.OpenRecordset("cfgcomm")

    tbcomm.MoveFirst
        
    With transcfg
        .Port = tbcomm!Port
        .data_bit = tbcomm!data_bit
        .stop_bit = tbcomm!stop_bit
        .baud_rate = tbcomm!baud_rate
        .parity = tbcomm!parity
        .blocksize = tbcomm!blocksize
    End With
    
    tbcomm.Close
    dbcomm.Close
    
    With Comm1
        .CommPort = transcfg.Port
        .Settings = transcfg.baud_rate & "," & transcfg.parity & "," & transcfg.data_bit & "," & transcfg.stop_bit
        .PortOpen = True
        .RTSEnable = True
        .RThreshold = 1
    End With
          
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
    
    Porttag = 1     'PortOpen에 성공을 나타냄
    identbOpenKey = False
    
    Set DB = OpenDatabase(filename & "comm\" & strmmdd + ".mdb")
    Set identb = DB.OpenRecordset("sp_identify")
    Set resulttb = DB.OpenRecordset("sp_result")
    
    identbOpenKey = True
    Porttag = 2     'OpenDB에 성공을 나타냄

'해당일의 이전까지의 샘플의 갯수를 db.identb에서 읽어 옴
    nowcount = identb.RecordCount
    PrevCnt = nowcount
        
    phase = 1
    EditCnt = 0
    sampleCnt = 0
    Test_OpenFlag = 0
    
    Open filename & machstr & ".log" For Output As #1
    
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
    
    Open filename & machstr & "Buff.log" For Output As #2
          
    Porttag = 3     'OpenSamFile에 성공을 나타냄
    FrmFlag = 41
    
    'Call Test
          
    If errfound = True Then
        Me.MousePointer = 0
        interfacfrm.Show
        Unload interfacfrm
        Exit Sub
    End If
          
Exit Sub

PortOpenErr:
    
    errfound = True
    
    MsgBox "통신 구성 에러!!  통신구성을 다시 설정해 주십시요!!"
    
    If Porttag = 2 Then     'OpenDB까지 성공, OpenSamFile에서 실패
        identb.Close
        resulttb.Close
        DB.Close
        Close #1
    End If
    
    Porttag = 0
    
    If Test_OpenFlag = 1 Then    'Sub Test에서 Open시 에러발생하면 Test_OpenFlag = 1가 되고,
        Close #3                   '완전히 Open되면 Test_OpenFlag = 2
        Porttag = 3
    End If
    
    Resume Next
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Porttag = 3 Then
        Comm1.PortOpen = False
        identb.Close
        resulttb.Close
        DB.Close
        identbOpenKey = False
        Close #1
        Close #2
    End If
End Sub

Private Sub spdface_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    StartBCol = CInt(BlockCol)
    StartBRow = CInt(BlockRow)
    EndBCol = CInt(BlockCol2)
    EndBRow = CInt(BlockRow2)
    BlockKey = True
End Sub
