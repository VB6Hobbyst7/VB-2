VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS410 
   BackColor       =   &H00DBE6E6&
   Caption         =   "헌혈 증서 반납"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "frmBBS410.frx":0000
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.TextBox txtRemark 
      BackColor       =   &H00DBE6E6&
      Height          =   1350
      Left            =   4605
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1350
      Width           =   8805
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">>"
      Height          =   1770
      Left            =   8235
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   3045
      Width           =   480
   End
   Begin VB.TextBox txtBldSrc 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7215
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2700
      Width           =   495
   End
   Begin VB.CommandButton cmdDelLine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<"
      Height          =   2040
      Left            =   8265
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   4815
      Width           =   450
   End
   Begin VB.CheckBox chkArrangement 
      BackColor       =   &H00DBE6E6&
      Caption         =   "입력시 번호정리"
      Height          =   255
      Left            =   1620
      TabIndex        =   11
      Top             =   8520
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CommandButton cmdArrangement 
      BackColor       =   &H00E0E0E0&
      Caption         =   "번호정리"
      Height          =   420
      Left            =   180
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   8580
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "반납(&S)"
      Height          =   510
      Left            =   10740
      Style           =   1  '그래픽
      TabIndex        =   1
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
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   9420
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   7320
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "반납취소"
      Height          =   510
      Left            =   6900
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   7320
      Width           =   1320
   End
   Begin VB.TextBox txtBldYY 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2700
      Width           =   495
   End
   Begin FPSpread.vaSpread tblNoUse 
      Height          =   3870
      Left            =   1680
      TabIndex        =   5
      Top             =   3030
      Width           =   6540
      _Version        =   196608
      _ExtentX        =   11536
      _ExtentY        =   6826
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
      MaxCols         =   6
      MaxRows         =   5
      OperationMode   =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS410.frx":076A
   End
   Begin FPSpread.vaSpread tblReturn 
      Height          =   3870
      Left            =   8760
      TabIndex        =   8
      Top             =   3030
      Width           =   4635
      _Version        =   196608
      _ExtentX        =   8176
      _ExtentY        =   6826
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
      MaxCols         =   3
      MaxRows         =   5
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS410.frx":0BEA
   End
   Begin MedControls1.LisLabel lblReturnSum 
      Height          =   330
      Left            =   10005
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6930
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
   Begin MedControls1.LisLabel lblReturnDt 
      Height          =   315
      Left            =   2520
      TabIndex        =   14
      Top             =   8820
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
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
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   8760
      TabIndex        =   17
      Top             =   2700
      Width           =   4635
      _ExtentX        =   8176
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
      Caption         =   "반납 헌혈증서 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   1680
      TabIndex        =   18
      Top             =   2700
      Width           =   6555
      _ExtentX        =   11562
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
      Caption         =   "미사용 헌혈증서 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   4605
      TabIndex        =   20
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
      Caption         =   "수령 REMARK"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Left            =   1680
      TabIndex        =   21
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
      Height          =   1455
      Left            =   1680
      TabIndex        =   22
      Top             =   1245
      Width           =   2910
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
         Height          =   840
         Left            =   75
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   555
         Width           =   2790
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회(&Q)"
         Height          =   345
         Left            =   1560
         Style           =   1  '그래픽
         TabIndex        =   0
         Top             =   195
         Width           =   1290
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
         Height          =   330
         Left            =   75
         TabIndex        =   24
         Tag             =   "103"
         Top             =   195
         Width           =   1440
      End
   End
   Begin MedControls1.LisLabel lbldt 
      Height          =   330
      Index           =   4
      Left            =   1680
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6945
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "합계"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblNoUseSum 
      Height          =   330
      Left            =   2925
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6945
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
   Begin MedControls1.LisLabel lbldt 
      Height          =   330
      Index           =   0
      Left            =   8775
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6930
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "합계"
      Appearance      =   0
   End
   Begin VB.Label Label4 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "반납일자 :"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1620
      TabIndex        =   16
      Tag             =   "103"
      Top             =   8880
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "합계 :"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10860
      TabIndex        =   15
      Tag             =   "103"
      Top             =   7320
      Width           =   540
   End
End
Attribute VB_Name = "frmBBS410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private objProgress As clsProgress
Private AlreadyReturn As Boolean

Private Sub cmdAdd_Click()
    Dim i As Long
    Dim r As Long
    Dim sum As Long
    Dim frno As Long
    Dim tono As Long
    Dim Row As Long
    Dim Col As Long
    
    
    'If tblReturn.MaxRows < 1 Then Exit Sub
    If tblNoUse.MaxRows < 1 Then Exit Sub
    Row = tblNoUse.ActiveRow
    Col = tblNoUse.ActiveCol
    
    With tblNoUse
        .Row = Row
        .Col = 2: frno = .value
        .Col = 3: tono = .value
    End With
    
    With tblReturn
        .MaxRows = .DataRowCnt + 1
        .Row = .MaxRows
        
        .Col = 1: .value = frno
        .Col = 2: .value = tono
        .Col = 3: .value = tono - frno + 1
        
        sum = 0
        For i = 1 To .MaxRows
            .Row = i
            .Col = 3: sum = sum + Val(.value)
        Next i
        
        lblReturnSum.Caption = IIf(sum = 0, "", sum)
        
        '.MaxRows = .MaxRows + 1
        
        cmdSave.Enabled = True
        
        If chkArrangement.value = 1 Then Call ArrangementTblReturn
    End With

End Sub

Private Sub cmdArrangement_Click()
    ArrangementTblReturn
End Sub

Private Sub cmdCancel_Click()
    Dim strMsg As String
    Dim objBDP As clsBloodDonationPaper
    
    Set objBDP = New clsBloodDonationPaper
    
    If objBDP.GetReturnCnt(Format(lblRcvDt, PRESENTDATE_FORMAT), ">") > 0 Then
        MsgBox "반납취소를 할 수 없습니다.", vbCritical, Me.Caption
    Else
    
        strMsg = "취소작업을 다시 취소할 수는 없습니다." & vbNewLine & _
                 "계속 하시겠읍니까?"
        If MsgBox(strMsg, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        
            If objBDP.CancelReturn(txtBldSrc, Format(lblRcvDt, PRESENTDATE_FORMAT), Format(lblReturnDt.Caption, PRESENTDATE_FORMAT)) = True Then
                SetLstRcvDt
                Call Query
            Else
                'dbconn.DisplayErrors
            End If
        
        End If
    
    End If
    
    Set objBDP = Nothing
    
    
    
    
End Sub

Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdDelLine_Click()
    If AlreadyReturn = True Then Exit Sub
    
    With tblReturn
        If .ActiveRow < 1 Then Exit Sub
        
        .Row = .ActiveRow
        .Action = ActionDeleteRow
        .MaxRows = .MaxRows - 1
        
        .MaxRows = .DataRowCnt + 1
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Call Query
End Sub

Private Sub cmdSave_Click()
    If chkValid = False Then Exit Sub
    
    Set objProgress = New clsProgress
    
    If DoReturn = True Then
        SetLstRcvDt
        Call Query
    End If
    
    Set objProgress = Nothing
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
'    With Me
'        .BackColor = medMain.BackColor
'        .Picture = medMain.Picture
'    End With
    SetLstRcvDt
    ClearAll
End Sub

Private Sub lstRcvDt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdQuery.Enabled = True
    If lstRcvDt.Text <> lblRcvDt Then
        lblRcvDt = lstRcvDt.Text
        Clear
    End If
End Sub

Private Sub tblNoUse_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim i As Long
'    Dim r As Long
'    Dim sum As Long
'    Dim frno As Long
'    Dim tono As Long
'
'    If tblReturn.MaxRows < 1 Then Exit Sub
'    '
'    With tblNoUse
'        .Row = Row
'        .Col = 2: frno = .value
'        .Col = 3: tono = .value
'    End With
'
'    With tblReturn
'        .Row = .MaxRows
'
'        .Col = 1: .value = frno
'        .Col = 2: .value = tono
'        .Col = 3: .value = tono - frno + 1
'
'        sum = 0
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 3: sum = sum + Val(.value)
'        Next i
'
'        lblReturnSum.Caption = IIf(sum = 0, "", sum)
'
'        .MaxRows = .MaxRows + 1
'
'        cmdSave.Enabled = True
'
'        If chkArrangement.value = 1 Then Call ArrangementTblReturn
'    End With
End Sub

Private Sub tblNoUse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call tblNoUse_DblClick(tblNoUse.ActiveCol, tblNoUse.ActiveRow)
    End If
End Sub

Private Sub tblReturn_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal mode As Integer, ByVal ChangeMade As Boolean)
    Static bfValue As String
    Dim i As Long
    Dim sum As Long
    Dim frno As Long
    Dim tono As Long
    
    If mode = 1 Then Exit Sub
    
    With tblReturn
        .Row = Row
        
        .Col = 1: frno = Val(.value)
        .Col = 2: tono = Val(.value)
        
        If frno = 0 And tono = 0 Then
            '라인삭제
            If Row <> .MaxRows Then
                .Row = Row
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                
                .Col = 1: .Row = .MaxRows
                .Action = ActionActiveCell
            End If
        Else
            If Row = .MaxRows Then .MaxRows = .MaxRows + 1
            
            If frno = 0 Then
                frno = tono
                .Col = 1: .value = frno
            End If
            If tono = 0 Then
                tono = frno
                .Col = 2: .value = tono
            End If
            
            sum = tono - frno + 1
            .Col = 3: .value = sum
        
            cmdSave.Enabled = (sum > 0)
        End If
        
        sum = 0
        For i = 1 To .MaxRows
            .Row = i
            .Col = 3: sum = sum + Val(.value)
        Next i
        
        lblReturnSum.Caption = IIf(sum = 0, "", sum)
    
        If chkArrangement.value = 1 Then Call ArrangementTblReturn
    End With
End Sub






Private Sub ClearAll()
    lstRcvDt.ListIndex = -1
    lblRcvDt = ""
    
    Clear
    cmdQuery.Enabled = False
End Sub

Private Sub Clear()
    Dim i As Long
    
    txtBldSrc = ""
    
    tblNoUse.MaxRows = 0
    tblReturn.MaxRows = 0
    lblNoUseSum.Caption = ""
    lblReturnSum.Caption = ""
    
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
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

Private Function FindRowNoUse(ByVal DrRS As Recordset) As Long
    Dim r As Long

    Dim divcd As String
    Dim CenterCd As String
    Dim frno As Long
    Dim tono As Long
    
    With tblNoUse
        For r = 1 To .MaxRows
            .Row = r
            .Col = 2: frno = Val(.value)
            .Col = 3: tono = Val(.value)
            .Col = 5: CenterCd = .value
            .Col = 6: divcd = .value
            
            If CenterCd = DrRS.Fields("centercd").value & "" And divcd = DrRS.Fields("divcd").value & "" Then
                If DrRS.Fields("bldno").value & "" >= (frno - 1) And DrRS.Fields("bldno").value & "" <= (tono + 1) Then
                    FindRowNoUse = r
                    Exit Function
                End If
            End If
        Next r
        
        FindRowNoUse = .MaxRows + 1
    End With
    
End Function

Private Sub SetTblNoUse(ByVal Rs As Recordset)
    Dim r As Long
    Dim sum As Long
    Dim frno As Long
    Dim tono As Long
    Dim centernm As String
    
    With tblNoUse
        If .MaxRows = 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            If Rs.Fields("divcd").value & "" = "0" Then
                centernm = GetCenterNm(Rs.Fields("centercd").value & "")
            Else
                centernm = GetBranchNm(Rs.Fields("centercd").value & "")
            End If
            
            .Col = 1: .value = centernm
            .Col = 2: .value = Rs.Fields("bldno").value & ""
            .Col = 3: .value = Rs.Fields("bldno").value & ""
            .Col = 5: .value = Rs.Fields("centercd").value & ""
            .Col = 6: .value = Rs.Fields("divcd").value & ""
        Else
            r = FindRowNoUse(Rs)
            If r > .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = r
                
                If Rs.Fields("divcd") = "0" Then
                    centernm = GetCenterNm(Rs.Fields("centercd").value & "")
                Else
                    centernm = GetBranchNm(Rs.Fields("centercd").value & "")
                End If
                
                .Col = 1: .value = centernm
                .Col = 2: .value = Rs.Fields("bldno").value & ""
                .Col = 3: .value = Rs.Fields("bldno").value & ""
                .Col = 5: .value = Rs.Fields("centercd").value & ""
                .Col = 6: .value = Rs.Fields("divcd").value & ""
            Else
                .Row = r
                .Col = 2: If Val(Rs.Fields("bldno").value & "") < Val(.value) Then .value = Rs.Fields("bldno").value & ""
                .Col = 3: If Val(Rs.Fields("bldno").value & "") > Val(.value) Then .value = Rs.Fields("bldno").value & ""
            End If
        End If
        
        .Col = 2: frno = Val(.value)
        .Col = 3: tono = Val(.value)
        .Col = 4: .value = tono - frno + 1
    End With

End Sub

Private Function FindRowReturn(ByVal DrRS As Recordset) As Long
    Dim r As Long

    Dim frno As Long
    Dim tono As Long
    
    With tblReturn
        For r = 1 To .MaxRows
            .Row = r
            .Col = 1: frno = Val(.value)
            .Col = 2: tono = Val(.value)
            
            If DrRS.Fields("bldno").value & "" >= (frno - 1) And DrRS.Fields("bldno").value & "" <= (tono + 1) Then
                FindRowReturn = r
                Exit Function
            End If
        Next r
        
        FindRowReturn = .MaxRows + 1
    End With
    
End Function

Private Sub SetTblReturn(ByVal DrRS As Recordset)
    Dim r As Long
    Dim sum As Long
    Dim frno As Long
    Dim tono As Long
    
    With tblReturn
        If .MaxRows = 0 Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .value = DrRS.Fields("bldno").value & ""
            .Col = 2: .value = DrRS.Fields("bldno").value & ""
        Else
            r = FindRowReturn(DrRS)
            If r > .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = r
                
                .Col = 1: .value = DrRS.Fields("bldno").value & ""
                .Col = 2: .value = DrRS.Fields("bldno").value & ""
            Else
                .Row = r
                .Col = 1: If Val(DrRS.Fields("bldno").value & "") < Val(.value) Then .value = DrRS.Fields("bldno").value & ""
                .Col = 2: If Val(DrRS.Fields("bldno").value & "") > Val(.value) Then .value = DrRS.Fields("bldno").value & ""
            End If
        End If
        
        .Col = 1: frno = Val(.value)
        .Col = 2: tono = Val(.value)
        .Col = 3: .value = tono - frno + 1
    End With
End Sub

Private Sub Query()
    Dim i As Long
    Dim nousesum As Long
    Dim returnsum As Long
    Dim returndt As String
    Dim DrRS As Recordset
    Dim objBDP As clsBloodDonationPaper
    
    Set objBDP = New clsBloodDonationPaper
    
    Set DrRS = New Recordset
    
    Call DrRS.Open(objBDP.GetBloodPaper(Format(lblRcvDt, PRESENTDATE_FORMAT)), DBConn)
    If DrRS.EOF Then
'        'dbconn.DisplayErrors
        Set DrRS = Nothing
        Set objBDP = Nothing
        Exit Sub
    End If
    If DrRS.RecordCount < 1 Then
        MsgBox "조회할 내역이 없습니다.", vbInformation, Me.Caption
        Set DrRS = Nothing
        Set objBDP = Nothing
        Exit Sub
    End If
    
    Set objProgress = New clsProgress
    
'    Set objProgress.StatusBar = medMain.stsBar
    objProgress.Container = MainFrm.stsBar
    
    objProgress.Min = 1
    objProgress.Max = Val(DrRS.RecordCount)
    objProgress.value = 0
    
    Clear
    
    AlreadyReturn = False
    
    returndt = ""
    nousesum = 0
    returnsum = 0
    For i = 1 To DrRS.RecordCount
    
        txtBldSrc = DrRS.Fields("bldsrc").value & ""
        txtBldYY = DrRS.Fields("bldyy").value & ""
        
        objProgress.value = objProgress.value + 1
        
        If Trim(DrRS.Fields("usedt").value & "" & "") = "" And Trim(DrRS.Fields("returndt").value & "") = "" Then
            nousesum = nousesum + 1
            Call SetTblNoUse(DrRS)
            
        ElseIf Trim(DrRS.Fields("usedt").value & "") = "" And Trim(DrRS.Fields("returndt").value & "") <> "" Then
            AlreadyReturn = True
            returndt = DrRS.Fields("returndt").value & ""
            returnsum = returnsum + 1
            Call SetTblReturn(DrRS)
        End If
        
        DrRS.MoveNext
    Next i
    
    lblNoUseSum.Caption = IIf(nousesum = 0, "", nousesum)
    lblReturnSum.Caption = IIf(returnsum = 0, "", returnsum)
    
    '이미 반납처리되었으면 반납취소할 수 있다.
    If AlreadyReturn = True Then
        cmdCancel.Enabled = objBDP.GetReturnCnt(Format(lblRcvDt, PRESENTDATE_FORMAT), ">") = 0
        tblReturn.Enabled = False
    Else
        cmdCancel.Enabled = False
        '신규반납처리할 수 있는 기준
        If objBDP.IsExistUseable(Format(lblRcvDt, PRESENTDATE_FORMAT), "<") = False Then
            If objBDP.IsExistUseable(Format(lblRcvDt, PRESENTDATE_FORMAT), ">") = True Then
                tblReturn.MaxRows = 1
            Else
                tblReturn.MaxRows = 0
            End If
        Else
            tblReturn.MaxRows = 0
        End If
        tblReturn.Enabled = True
    End If
    
    If returndt <> "" Then
        lblReturnDt.Caption = Format(returndt, "####-##-##")
    Else
        lblReturnDt.Caption = Format(GetSystemDate, "YYYY-MM-DD")
    End If

    Set objBDP = Nothing
    
    Set objProgress = Nothing
End Sub

Private Sub ArrangementTblReturn()
    Dim r As Long
    Dim sum As Long
    Dim frno As Long
    Dim tono As Long
    Dim fg As Boolean
    Dim bfFrno As Long
    Dim bfTono As Long
    Dim newFrno As Long
    Dim newTono As Long
    
    With tblReturn
        '정렬
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .SortBy = SortByRow
        .SortKey(1) = 1
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Action = ActionSort
        .BlockMode = False
        
        r = 1
        Do
            If r > .DataRowCnt Then Exit Do
            
            .Row = r
            .Col = 1: frno = Val(.value)
            .Col = 2: tono = Val(.value)
            
            If r = 1 Then
                bfFrno = frno
                bfTono = tono
                r = r + 1
            Else
                fg = False
                If frno >= (bfFrno - 1) And frno <= (bfTono + 1) Then fg = True
                If tono >= (bfFrno - 1) And tono <= (bfTono + 1) Then fg = True
                If bfFrno >= (frno - 1) And bfFrno <= (tono + 1) Then fg = True
                If bfTono >= (frno - 1) And bfTono <= (tono + 1) Then fg = True
                
                If fg = False Then
                    r = r + 1
                    bfFrno = frno
                    bfTono = tono
                Else
                    newFrno = IIf(bfFrno < frno, bfFrno, frno)
                    newTono = IIf(bfTono > tono, bfTono, tono)
                    sum = newTono - newFrno + 1
                    
                    .Row = r - 1
                    
                    .Col = 1: .value = newFrno
                    .Col = 2: .value = newTono
                    .Col = 3: .value = sum
                    
                    .Row = r
                    .Action = ActionDeleteRow
                    
                    .MaxRows = .MaxRows - 1
                    
                    
                    bfFrno = newFrno
                    bfTono = newTono
                End If
            End If
        Loop
        
        sum = 0
        For r = 1 To .MaxRows
            .Row = r
            .Col = 3: sum = sum + Val(.value)
        Next r
        
        lblReturnSum.Caption = IIf(sum = 0, "", sum)
        
        .MaxRows = .DataRowCnt + 1
    End With
End Sub

Private Function chkValid() As Boolean
    Dim r As Long
    Dim no As Long
    Dim frno As Long
    Dim tono As Long
    Dim colNoUse As Collection
    Dim objBDP As clsBloodDonationPaper
    
    Set objBDP = New clsBloodDonationPaper
    
    ' ----------------------------------------------------------
    ' step 1 : 현재 반환 처리하려는 혈액증서들 보다 먼저 받은
    '          내역중 아직 미사용으로 남아있는 것이 있는지 검사
    ' ----------------------------------------------------------
    If objBDP.IsExistNotUsed(Format(lblRcvDt, PRESENTDATE_FORMAT)) Then
        MsgBox "먼저 받은 혈액증서들의 반납처리를 하십시요.", vbInformation, Me.Caption
        chkValid = False
        Set objBDP = Nothing
        Exit Function
    End If
    ' ----------------------------------------------------------
    ' step 2 : 반환처리하고 나서도 사용할 수 있는
    '          혈액증서가 있는지 검사
    ' ----------------------------------------------------------
'    If objBDP.IsExistUseable(Format(lblRcvDt, PRESENTDATE_FORMAT), ">") = False Then
'        MsgBox "먼저 사용할 수 있는 혈액증서의 입고처리를 하십시요.", vbInformation, Me.Caption
'        chkValid = False
'        Set objBDP = Nothing
'        Exit Function
'    End If
    ' ----------------------------------------------------------
    ' step 3 : 센터에 할당된 증서중 미사용 증서가 모두
    '          반납처리 되는지 검사
    ' ----------------------------------------------------------
    Set colNoUse = New Collection
    With tblNoUse
        For r = 1 To .MaxRows
            .Row = r
            .Col = 6
            If .value = "0" Then
                .Col = 2: frno = Val(.value)
                .Col = 3: tono = Val(.value)
                For no = frno To tono
                    colNoUse.Add no, "K" & no
                Next no
            End If
        Next r
    End With
    
On Error Resume Next

    With tblReturn
        For r = 1 To .DataRowCnt
            .Row = r
            .Col = 1: frno = Val(.value)
            .Col = 2: tono = Val(.value)
            For no = frno To tono
                Call colNoUse.Remove("K" & no)
            Next no
        Next r
    End With
    If colNoUse.Count > 0 Then
        MsgBox "센터에 할당된 미사용 증서는 모두 반납처리하여야합니다.", vbCritical, Me.Caption
        Set colNoUse = Nothing
        chkValid = False
        Exit Function
    End If
    
    Set colNoUse = Nothing
    
    chkValid = True
    
    Set objBDP = Nothing
End Function

Private Function DoReturn() As Boolean
    Dim r As Long
    Dim frno As Long, tono As Long
    Dim no As Long
    Dim objBloodDonationPaper As clsBloodDonationPaper
    Dim returndt As String
    Dim returnid As String
    
    
    returndt = Format(lblReturnDt.Caption, PRESENTDATE_FORMAT)
    returnid = ObjMyUser.EmpId
    
    objProgress.Min = 0
    objProgress.Max = Val(lblReturnSum.Caption)
    objProgress.value = 0
    
On Error GoTo DoReturn_error

    DBConn.BeginTrans
    
    Set objBloodDonationPaper = New clsBloodDonationPaper
    With objBloodDonationPaper
        .BldSrc = txtBldSrc
        .BldYY = Mid(Format(lblRcvDt, PRESENTDATE_FORMAT), 3, 2)
        .rcvdt = Format(lblRcvDt, PRESENTDATE_FORMAT)
        .returndt = returndt
        .returnid = returnid
    
        For r = 1 To tblReturn.DataRowCnt
            tblReturn.Row = r
            tblReturn.Col = 1: frno = Val(tblReturn.value)
            tblReturn.Col = 2: tono = Val(tblReturn.value)
            
            For no = frno To tono
                .BldNo = no
                If .DoReturn() = False Then GoTo DoReturn_error
                
                objProgress.value = objProgress.value + 1
                
            Next no
        Next r
        
        '자병원의 반환처리 되지 않은 혈액증서를 사용상태로 셋팅
        If .DoReturnAll(txtBldSrc, Format(lblRcvDt, PRESENTDATE_FORMAT), returndt, returnid) = False Then GoTo DoReturn_error
    End With
    
    Set objBloodDonationPaper = Nothing
    
    DBConn.CommitTrans
    DoReturn = True
    
    Exit Function
    
DoReturn_error:
    Set objBloodDonationPaper = Nothing
    DBConn.RollbackTrans
        
    Set objProgress = Nothing
    MsgBox Err.Description, vbCritical, Me.Caption
    DoReturn = False
End Function

Private Function chkValidCancel() As Boolean
    Dim objBDP As clsBloodDonationPaper
    
    Set objBDP = New clsBloodDonationPaper
    
    If objBDP.GetReturnCnt(Format(lblRcvDt, PRESENTDATE_FORMAT), ">") > 0 Then
        MsgBox "반납취소를 할 수 없습니다.", vbCritical, Me.Caption
        chkValidCancel = False
    Else
        chkValidCancel = True
    End If
    
    Set objBDP = Nothing
End Function
