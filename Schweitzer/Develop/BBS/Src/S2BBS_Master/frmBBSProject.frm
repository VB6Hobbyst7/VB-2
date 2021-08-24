VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form frmBBSProject 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "혈액은행 프로젝트 옵션관리"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9465
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8070
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8145
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8070
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   6810
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8070
      Width           =   1320
   End
   Begin DRcontrol1.DrFrame fraIndex 
      Height          =   7200
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   12700
      Title           =   "색 인 관 리"
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox PicBoolean 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4650
         TabIndex        =   13
         Top             =   2355
         Width           =   4680
         Begin VB.OptionButton optBo 
            BackColor       =   &H80000005&
            Caption         =   "False"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2910
            TabIndex        =   15
            Top             =   75
            Width           =   930
         End
         Begin VB.OptionButton optBo 
            BackColor       =   &H80000005&
            Caption         =   "True"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1200
            TabIndex        =   14
            Top             =   75
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.Label lblSK 
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  '투명
            Caption         =   "색  인  값"
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   90
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  '단색
            Height          =   345
            Index           =   5
            Left            =   0
            Shape           =   4  '둥근 사각형
            Top             =   30
            Width           =   1020
         End
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Boolean"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3540
         TabIndex        =   12
         Top             =   1455
         Width           =   1170
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Long"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2490
         TabIndex        =   11
         Top             =   1455
         Width           =   1050
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "String"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   10
         Top             =   1455
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.PictureBox picString 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4650
         TabIndex        =   7
         Top             =   1920
         Width           =   4680
         Begin VB.TextBox txtSValue 
            BackColor       =   &H00F1F5F4&
            Height          =   360
            Left            =   1215
            TabIndex        =   8
            Top             =   15
            Width           =   3450
         End
         Begin VB.Label lblSK 
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  '투명
            Caption         =   "색  인  값"
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   90
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  '단색
            Height          =   345
            Index           =   6
            Left            =   15
            Shape           =   4  '둥근 사각형
            Top             =   15
            Width           =   1020
         End
      End
      Begin VB.TextBox txtMemo 
         BackColor       =   &H00F1F5F4&
         Height          =   3690
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3225
         Width           =   4590
      End
      Begin VB.TextBox txtKeyCode 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1305
         TabIndex        =   5
         Top             =   540
         Width           =   2070
      End
      Begin VB.TextBox txtKeyName 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1305
         TabIndex        =   4
         Top             =   960
         Width           =   3510
      End
      Begin VB.Label lblSK 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "유       형"
         Height          =   225
         Index           =   2
         Left            =   225
         TabIndex        =   20
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "Memo"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   2895
         Width           =   1335
      End
      Begin VB.Label lblSK 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "색인 코드"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   18
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lblSK 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "색  인  명"
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   17
         Top             =   1035
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   350
         Index           =   2
         Left            =   120
         Shape           =   4  '둥근 사각형
         Top             =   960
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   350
         Index           =   0
         Left            =   120
         Shape           =   4  '둥근 사각형
         Top             =   540
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   345
         Index           =   4
         Left            =   120
         Shape           =   4  '둥근 사각형
         Top             =   1410
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   345
         Index           =   3
         Left            =   120
         Shape           =   4  '둥근 사각형
         Top             =   2805
         Width           =   705
      End
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   7770
      Left            =   270
      TabIndex        =   21
      Top             =   840
      Width           =   5595
      _Version        =   196608
      _ExtentX        =   9869
      _ExtentY        =   13705
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   5
      MaxRows         =   30
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14411494
      SpreadDesigner  =   "frmBBSProject.frx":0000
   End
   Begin VB.Label lblSubMenu 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "임상병리 프로젝트 옵션관리"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   300
      TabIndex        =   22
      Top             =   330
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H005E957E&
      BorderWidth     =   3
      FillColor       =   &H00A7C9B9&
      FillStyle       =   0  '단색
      Height          =   495
      Index           =   1
      Left            =   285
      Top             =   195
      Width           =   3915
   End
End
Attribute VB_Name = "frmBBSProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblCol
    enIndex = 1
    enName
    enValue
    enKind
    enMemo
End Enum

Private Sub cmdClear_Click()
    Call ClearData
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call ClearData
    Call GetIniDSP
End Sub
Private Sub ClearData()
    txtKeyCode.Text = "": txtKeyName.Text = "": txtMemo.Text = "": txtSValue.Text = ""
    Call optDiv_Click(0)
End Sub

Private Sub optDiv_Click(Index As Integer)
    picString.Enabled = False: PicBoolean.Enabled = False
    If Index = 2 Then
        PicBoolean.Enabled = True
    Else
        picString.Enabled = True
    End If
End Sub

Private Sub txtKeyCode_Change()
    txtKeyName.Text = "": txtMemo.Text = "": txtSValue.Text = ""
End Sub

Private Sub txtKeyCode_GotFocus()
    txtKeyCode.SelStart = 0
    txtKeyCode.SelLength = Len(txtKeyCode.Text)
End Sub

Private Sub txtKeyCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub GetIniDSP()
    Dim strTmp      As String
    Dim blnMaster   As Boolean
    Dim sFile       As String
    
    sFile = INIPath
    
    If Dir(INIPath) = "" Then
        MsgBox "파일이 존재하지 않습니다.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    Open sFile For Input As #1
    On Error Resume Next
    
    With tblList
        .MaxRows = 0
        Do While Not EOF(1)
            Line Input #1, strTmp
            If InStr(1, strTmp, "[") > 0 And strTmp = "[BBS_CONST]" Then
                blnMaster = True
            ElseIf InStr(1, strTmp, "[") > 0 And strTmp <> "[BBS_CONST]" Then
                blnMaster = False
            End If
            
            If blnMaster = True And strTmp <> "[BBS_CONST]" Then
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = tblCol.enIndex: .Value = Trim(medGetP(strTmp, 1, "="))
                .Col = tblCol.enKind: .Value = Trim(medGetP(medGetP(strTmp, 2, "="), 2, LINE_DIV))
                .Col = tblCol.enValue: .Value = Trim(medGetP(medGetP(strTmp, 2, "="), 1, LINE_DIV))
                .Col = tblCol.enName: .Value = Trim(medGetP(medGetP(strTmp, 2, "="), 3, LINE_DIV))
                .Col = tblCol.enMemo: .Value = Trim(medGetP(medGetP(strTmp, 2, "="), 4, LINE_DIV))
            End If
        Loop
        Call tblList_Click(1, 1)
    End With
    Close #1
End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    Call ClearData
    With tblList
        If .DataRowCnt < 1 Then Exit Sub
        .Row = Row: .Col = 1: If .Value = "" Then Exit Sub
        .Col = tblCol.enIndex:  txtKeyCode.Text = .Value
        .Col = tblCol.enName:   txtKeyName.Text = .Value
        .Col = tblCol.enKind:   optDiv(CLng(.Value)).Value = True
            Select Case .Value
                Case "0", "1"
                    .Col = tblCol.enValue: txtSValue.Text = .Value
                Case "2"
                    .Col = tblCol.enValue
                    If UCase(.Value) = "TRUE" Then
                        optBo(0).Value = True
                    Else
                        optBo(1).Value = True
                    End If
            End Select
        .Col = tblCol.enMemo: txtMemo.Text = .Value
    End With
    
End Sub

Private Sub cmdSave_Click()
    Dim strIndex    As String
    Dim strValue    As String
    Dim strKind     As String
    Dim strName     As String
    Dim strMemo     As String
    
    strIndex = Trim(txtKeyCode.Text)
    strName = Trim(txtKeyName.Text)
    strKind = IIf(optDiv(0).Value = True, "0", IIf(optDiv(1).Value = True, "1", "2"))
    
    Select Case strKind
        Case "0", "1": strValue = Trim(txtSValue.Text)
        Case "2": strValue = IIf(optBo(0).Value = True, "TRUE", "FALSE")
    End Select
    strMemo = txtMemo.Text
    
    Call medSetINI("BBS_CONST", strIndex, strValue & LINE_DIV & strKind & LINE_DIV & strName & LINE_DIV & strMemo, INIPath)
    Call GetIniDSP
End Sub

