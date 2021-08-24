VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS409 
   BackColor       =   &H00DBE6E6&
   Caption         =   "헌혈 증서 조회"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "frmBBS409.frx":0000
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.TextBox txtBldYY 
      Alignment       =   2  '가운데 맞춤
      Height          =   315
      Left            =   12405
      TabIndex        =   6
      Top             =   1035
      Width           =   615
   End
   Begin VB.TextBox txtBldSrc 
      Alignment       =   2  '가운데 맞춤
      Height          =   315
      Left            =   11865
      TabIndex        =   5
      Top             =   1035
      Width           =   555
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   10740
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   7320
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   12060
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   7320
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   4605
      TabIndex        =   4
      Top             =   1020
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "센터별 사용량 정보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1020
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "수령 일자"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2940
      Left            =   1680
      TabIndex        =   8
      Top             =   1260
      Width           =   2910
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회(&Q)"
         Height          =   420
         Left            =   1410
         Style           =   1  '그래픽
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin VB.ListBox lstRcvDt 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   165
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   630
         Width           =   2565
      End
      Begin VB.Label lblRcvDt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   165
         TabIndex        =   20
         Tag             =   "103"
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   3060
      Left            =   1680
      TabIndex        =   9
      Top             =   4110
      Width           =   2910
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   135
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   855
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "총수량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   15
         Left            =   135
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "미사용"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   135
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "사용량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   135
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1935
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "반납량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   0
         Left            =   300
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "사용량정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblNoUseCnt 
         Height          =   330
         Left            =   1215
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblUseCnt 
         Height          =   330
         Left            =   1215
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReturnCnt 
         Height          =   330
         Left            =   1215
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1935
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTotCnt 
         Height          =   330
         Left            =   1215
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   855
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   4170
      Left            =   4605
      TabIndex        =   23
      Top             =   1260
      Width           =   8790
      Begin FPSpread.vaSpread tblBloodPaper 
         Height          =   3630
         Left            =   165
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   300
         Width           =   8430
         _Version        =   196608
         _ExtentX        =   14870
         _ExtentY        =   6403
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   8
         MaxRows         =   13
         OperationMode   =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS409.frx":076A
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   1815
      Left            =   4605
      TabIndex        =   21
      Top             =   5355
      Width           =   8790
      Begin VB.TextBox txtRemark 
         Height          =   1155
         Left            =   225
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   1
         Top             =   540
         Width           =   8355
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   225
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   195
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "수령 Remark"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private objProgress As clsProgress



Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Call Query
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    SetLstRcvDt
    ClearAll
End Sub

Private Sub lstRcvDt_Click()
    If lblRcvDt.Caption <> lstRcvDt.Text Then
        lblRcvDt.Caption = lstRcvDt.Text
        Clear
        cmdQuery.Enabled = True
    End If
End Sub







Private Sub ClearAll()
    lblRcvDt.Caption = ""
    lstRcvDt.ListIndex = -1
    Clear
    
    cmdQuery.Enabled = False
End Sub

Private Sub Clear()
    txtBldSrc = ""
    txtBldYY = ""

    lblTotCnt.Caption = ""
    lblUseCnt.Caption = ""
    lblNoUseCnt.Caption = ""
    lblReturnCnt.Caption = ""
    tblBloodPaper.MaxRows = 0
End Sub

Private Sub SetLstRcvDt()
    Dim objBDP As clsBloodDonationPaper
    Dim astrRcvDt() As String
    Dim Cnt As Long
    Dim i As Long
    
    
    '과거에 입고처리된 일자리스트
    Set objBDP = New clsBloodDonationPaper
    Cnt = objBDP.GetRcvDtList(astrRcvDt)
    lstRcvDt.Clear
    For i = 0 To Cnt - 1
        lstRcvDt.AddItem Format(astrRcvDt(i), "####-##-##")
    Next i
    Set objBDP = Nothing
End Sub

Private Sub Query()
    Dim i As Long
    Dim r As Long
    Dim frno As Long
    Dim tono As Long
    Dim centernm As String
    Dim statusnm As String
    Dim statuscd As String
    Dim RS As Recordset
    Dim objBDP As clsBloodDonationPaper
    
    Dim totcnt As Long
    Dim usecnt As Long
    Dim nousecnt As Long
    Dim returncnt As Long
    
    
   
    Set RS = New Recordset
    
    Set objBDP = New clsBloodDonationPaper
    Call RS.Open(objBDP.GetBloodPaper(Format(lblRcvDt, PRESENTDATE_FORMAT)), DBConn)
    Set objBDP = Nothing
    
    Set objProgress = New clsProgress
    
'    Set objProgress.StatusBar = medMain.stsBar
    objProgress.Container = MainFrm.stsBar
    
    objProgress.Min = 1
    objProgress.Max = Val(RS.RecordCount)
    objProgress.value = 0
    
    totcnt = 0
    usecnt = 0
    nousecnt = 0
    returncnt = 0
    
    Clear
    
    With tblBloodPaper
        .MaxRows = 0
        For i = 1 To RS.RecordCount
        
            txtBldSrc = RS.Fields("bldsrc").value & ""
            txtBldYY = RS.Fields("bldyy").value & ""
            
            objProgress.value = objProgress.value + 1
            
            totcnt = totcnt + 1
            
            If Trim(RS.Fields("usedt").value & "" & "") = "" And Trim(RS.Fields("returndt").value & "") = "" Then
                statuscd = "0"
                statusnm = "미사용"
                nousecnt = nousecnt + 1
            ElseIf Trim(RS.Fields("usedt").value & "") <> "" And Trim(RS.Fields("returndt").value & "") = "" Then
                statuscd = "1"
                statusnm = "사용"
                usecnt = usecnt + 1
            Else
                statuscd = "2"
                statusnm = "반납"
                returncnt = returncnt + 1
            End If
            
            If .MaxRows = 0 Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                
                If RS.Fields("divcd").value & "" = "0" Then
                    centernm = GetCenterNm(RS.Fields("centercd").value & "")
                Else
                    centernm = GetBranchNm(RS.Fields("centercd").value & "")
                End If
                
                .Col = 1: .value = centernm
                .Col = 2: .value = RS.Fields("bldno").value & ""
                .Col = 3: .value = RS.Fields("bldno").value & ""
                .Col = 4: .value = 1
                .Col = 5: .value = statusnm
                .Col = 6: .value = RS.Fields("centercd").value & ""
                .Col = 7: .value = RS.Fields("divcd").value & ""
                .Col = 8: .value = statuscd
            Else
                r = FindRow(RS, statuscd)
                If r > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = r
                    
                    If RS.Fields("divcd").value & "" = "0" Then
                        centernm = GetCenterNm(RS.Fields("centercd").value & "")
                    Else
                        centernm = GetBranchNm(RS.Fields("centercd").value & "")
                    End If
                    
                    .Col = 1: .value = centernm
                    .Col = 2: .value = RS.Fields("bldno").value & ""
                    .Col = 3: .value = RS.Fields("bldno").value & ""
                    .Col = 4: .value = 1
                    .Col = 5: .value = statusnm
                    .Col = 6: .value = RS.Fields("centercd").value & ""
                    .Col = 7: .value = RS.Fields("divcd").value & ""
                    .Col = 8: .value = statuscd
                Else
                    .Row = r
                    .Col = 2: If Val(RS.Fields("bldno")) < Val(.value) Then .value = RS.Fields("bldno").value & ""
                    .Col = 3: If Val(RS.Fields("bldno")) > Val(.value) Then .value = RS.Fields("bldno").value & ""
                End If
            End If
            
            .Col = 2: frno = Val(.value)
            .Col = 3: tono = Val(.value)
            .Col = 4: .value = tono - frno + 1
            
            RS.MoveNext
        Next i
        
        
    End With
    
    
    lblTotCnt.Caption = IIf(totcnt = 0, "", totcnt)
    lblUseCnt.Caption = IIf(usecnt = 0, "", usecnt)
    lblNoUseCnt.Caption = IIf(nousecnt = 0, "", nousecnt)
    lblReturnCnt.Caption = IIf(returncnt = 0, "", returncnt)
    
    Set objProgress = Nothing
    Set RS = Nothing
End Sub

Private Function FindRow(ByVal DrRS As Recordset, ByVal status As String) As Long
    Dim r As Long

    Dim divcd As String
    Dim CenterCd As String
    Dim statuscd As String
    Dim frno As Long
    Dim tono As Long
    
    With tblBloodPaper
        For r = 1 To .MaxRows
            .Row = r
            .Col = 6: CenterCd = .value
            .Col = 7: divcd = .value
            .Col = 8: statuscd = .value
            .Col = 2: frno = Val(.value)
            .Col = 3: tono = Val(.value)
            
            If CenterCd = DrRS.Fields("centercd").value & "" And divcd = DrRS.Fields("divcd").value & "" And statuscd = status Then
                If DrRS.Fields("bldno").value & "" >= (frno - 1) And DrRS.Fields("bldno").value & "" <= (tono + 1) Then
                    FindRow = r
                    Exit Function
                End If
            End If
        Next r
        
        FindRow = .MaxRows + 1
    End With
    
End Function

