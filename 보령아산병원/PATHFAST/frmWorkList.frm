VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BorderStyle     =   1  '단일 고정
   Caption         =   "워크리스트"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10230
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkWAll 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   660
      TabIndex        =   11
      Top             =   900
      Width           =   225
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10005
      Begin VB.CommandButton cmdSendOrder 
         Appearance      =   0  '평면
         Caption         =   "오더전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   13
         Top             =   150
         Width           =   1155
      End
      Begin VB.CommandButton cmdPatDelete 
         Caption         =   "선택제외"
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
         Left            =   4920
         TabIndex        =   10
         Top             =   180
         Width           =   945
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7410
         TabIndex        =   8
         Text            =   "1"
         Top             =   180
         Width           =   585
      End
      Begin VB.CheckBox chkSaveAll 
         Caption         =   "저장포함"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6090
         TabIndex        =   3
         Top             =   210
         Width           =   825
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "워크조회"
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
         Left            =   3840
         TabIndex        =   2
         Top             =   180
         Width           =   1065
      End
      Begin VB.CommandButton cmdWorkPrint 
         Caption         =   "워크출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16290
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   345
         Left            =   2340
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63569921
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   345
         Left            =   690
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63569921
         CurrentDate     =   40248
      End
      Begin VB.Label Label1 
         Caption         =   "SEQ"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6900
         TabIndex        =   9
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "기간"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label12 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   330
         Width           =   105
      End
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   7065
      Left            =   90
      TabIndex        =   12
      Top             =   840
      Width           =   9945
      _Version        =   393216
      _ExtentX        =   17542
      _ExtentY        =   12462
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      ColsFrozen      =   16
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   19
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":144A
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub chkWAll_Click()
    Dim iRow As Long
    
    With vasID
        If chkWAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 1
            Next iRow
        ElseIf chkWAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 0
            Next iRow
        End If
    End With
    
End Sub

Private Sub cmdPatDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCheckBox
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                j = j + 1
            End If
        Next
    End With
    
End Sub

Private Sub cmdSearch_Click()

    Call GetWorkList_ASAN(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))

End Sub

Private Sub cmdSendOrder_Click()
    Dim i       As Integer
    
    With frmInterface.vasID
        For i = 1 To vasID.DataRowCnt
            vasID.Row = i
            vasID.Col = colCheckBox
            If vasID.Value = "1" Then
                .MaxRows = .MaxRows + 1
                Call .SetText(colHOSPDATE, .MaxRows, GetText(vasID, i, colHOSPDATE))
                Call .SetText(colCHARTNO, .MaxRows, GetText(vasID, i, colCHARTNO))
                Call .SetText(colPNAME, .MaxRows, GetText(vasID, i, colPNAME))
                Call .SetText(colPSEX, .MaxRows, GetText(vasID, i, colPSEX))
                Call .SetText(colBARCODE, .MaxRows, GetText(vasID, i, colBARCODE))
                Call .SetText(colPID, .MaxRows, GetText(vasID, i, colPID))
                Call .SetText(colSN, .MaxRows, GetText(vasID, i, colSN))
                
                'vasID.Row = i
                'vasID.Col = colCheckBox
                'vasID.Value = "0"
                
                vasID.Row = i
                vasID.Action = ActionDeleteRow
                vasID.MaxRows = vasID.MaxRows - 1
                
            End If
        Next
    End With
    
End Sub

Private Sub Form_Load()
    
    vasID.MaxRows = 0
      
    Call cmdSearch_Click
    
End Sub


Private Sub GetWorkList_ASAN(ByVal pFrDt As String, ByVal pToDt As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    
    blnSame = False
    vasID.ReDraw = False
    

    SQL = ""    ', OKFL as Y/N결과확정여부, OKDT as 결과확정DATE, OKID as 확정자
    SQL = SQL & "SELECT DISTINCT HONM, OKDT as 접수일자, SPNO as 바코드번호, WKNO as 챠트번호, PAID as 내원번호, PANM as 이름, SEXS as 성별, SJDT as 생년월일" & vbCrLf
    SQL = SQL & "  FROM V_AFISINFO " & vbCrLf
    SQL = SQL & " WHERE OKFL = 'N'" & vbCrLf
    SQL = SQL & "   AND ORCD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & "   AND OKDT BETWEEN  TO_DATE('" & pFrDt & "','yyyymmdd') AND TO_DATE('" & pToDt & "','yyyymmdd')+1" & vbCrLf
    SQL = SQL & "   AND HOCD = '" & gEquipCode & "'" & vbCrLf
    SQL = SQL & " ORDER BY OKDT "
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
                        blnSame = True
                    End If
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("성별")) & "", .MaxRows, colPSEX
                    
'                    SetText vasID, "1", .MaxRows, colCheckBox
'                    SetText vasID, "20160831", .MaxRows, colHOSPDATE
'                    SetText vasID, "1234567890", .MaxRows, colBARCODE
'                    SetText vasID, "11111111", .MaxRows, colCHARTNO
'                    SetText vasID, "22222222", .MaxRows, colPID
'                    SetText vasID, "홍길동1", .MaxRows, colPNAME
'                    SetText vasID, "남", .MaxRows, colPSEX
                    
                    SetText vasID, txtSeq.Text, .MaxRows, colSN
                    
                    txtSeq.Text = txtSeq.Text + 1

                End If
                blnSame = False
            End With
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        MsgBox "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    With frmInterface.vasID
        .MaxRows = .MaxRows + 1
        
        Call .SetText(colHOSPDATE, .MaxRows, GetText(vasID, Row, colHOSPDATE))
        Call .SetText(colCHARTNO, .MaxRows, GetText(vasID, Row, colCHARTNO))
        Call .SetText(colPNAME, .MaxRows, GetText(vasID, Row, colPNAME))
        Call .SetText(colPSEX, .MaxRows, GetText(vasID, Row, colPSEX))
        Call .SetText(colBARCODE, .MaxRows, GetText(vasID, Row, colBARCODE))
        Call .SetText(colPID, .MaxRows, GetText(vasID, Row, colPID))
        Call .SetText(colSN, .MaxRows, GetText(vasID, Row, colSN))
        
        vasID.Row = Row
        vasID.Action = ActionDeleteRow
        vasID.MaxRows = vasID.MaxRows - 1
    End With

End Sub
