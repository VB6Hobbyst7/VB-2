VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form frm340ShowIndex 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "Lis340.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00EBF3ED&
      Height          =   675
      Left            =   5325
      ScaleHeight     =   615
      ScaleWidth      =   5355
      TabIndex        =   8
      Top             =   7800
      Width           =   5415
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   1350
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00DBE6E6&
         Caption         =   "화면지움(&C)"
         Height          =   510
         Left            =   2670
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   3990
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   555
      Left            =   5325
      TabIndex        =   2
      Top             =   300
      Width           =   5415
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00DBE6E6&
         Height          =   315
         Left            =   1380
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   4
         Top             =   160
         Width           =   3795
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00DBE6E6&
            Caption         =   "템플릿"
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   7
            ToolTipText     =   "COM002"
            Top             =   40
            Width           =   1215
         End
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00DBE6E6&
            Caption         =   "공통코드1"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "COM002"
            Top             =   40
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optSelect 
            BackColor       =   &H00DBE6E6&
            Caption         =   "공통코드2"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   5
            ToolTipText     =   "COM003"
            Top             =   40
            Width           =   1155
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFF9E3&
         Caption         =   "코드관리 선택"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   225
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFF9E3&
         FillStyle       =   0  '단색
         Height          =   330
         Index           =   4
         Left            =   60
         Shape           =   4  '둥근 사각형
         Top             =   150
         Width           =   1260
      End
   End
   Begin FPSpread.vaSpread ssTableKey 
      Height          =   7605
      Left            =   75
      TabIndex        =   0
      Top             =   900
      Width           =   5205
      _Version        =   196608
      _ExtentX        =   9181
      _ExtentY        =   13414
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   2
      MaxRows         =   5
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14411494
      SpreadDesigner  =   "Lis340.frx":08CA
   End
   Begin VB.PictureBox fraIndex 
      BackColor       =   &H00DBE6E6&
      Enabled         =   0   'False
      Height          =   6855
      Left            =   5325
      ScaleHeight     =   6795
      ScaleWidth      =   5355
      TabIndex        =   12
      Top             =   900
      Width           =   5415
      Begin VB.TextBox txtKeyName 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1530
         TabIndex        =   17
         Top             =   900
         Width           =   3510
      End
      Begin VB.TextBox txtKeyCode 
         BackColor       =   &H00F1F5F4&
         Height          =   330
         Left            =   1530
         TabIndex        =   16
         Top             =   480
         Width           =   2070
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00F1F5F4&
         Height          =   2355
         Left            =   345
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   4305
         Width           =   4560
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00DBE6E6&
         Height          =   2475
         Left            =   345
         ScaleHeight     =   2415
         ScaleWidth      =   4605
         TabIndex        =   13
         Top             =   1335
         Width           =   4665
         Begin FPSpread.vaSpread tblFields 
            Height          =   2385
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   4605
            _Version        =   196608
            _ExtentX        =   8123
            _ExtentY        =   4207
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
            EditEnterAction =   2
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
            MaxCols         =   3
            MaxRows         =   8
            ScrollBars      =   2
            ShadowColor     =   14411494
            ShadowDark      =   12632256
            SpreadDesigner  =   "Lis340.frx":0C3E
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "색 인 관 리"
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
         Left            =   2145
         TabIndex        =   21
         Top             =   90
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   195
         X2              =   5145
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblSK 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "색  인  명"
         Height          =   225
         Index           =   0
         Left            =   450
         TabIndex        =   20
         Top             =   975
         Width           =   975
      End
      Begin VB.Label lblSK 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "색인 코드"
         Height          =   225
         Index           =   1
         Left            =   435
         TabIndex        =   19
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "색인 주석"
         Height          =   195
         Left            =   420
         TabIndex        =   18
         Top             =   4005
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   345
         Index           =   0
         Left            =   345
         Shape           =   4  '둥근 사각형
         Top             =   480
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   345
         Index           =   2
         Left            =   345
         Shape           =   4  '둥근 사각형
         Top             =   900
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   345
         Index           =   3
         Left            =   345
         Shape           =   4  '둥근 사각형
         Top             =   3900
         Width           =   1020
      End
   End
   Begin VB.Label lblSubMenu 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "임상병리 공통 코드관리"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   75
      TabIndex        =   1
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
      Left            =   120
      Shape           =   4  '둥근 사각형
      Top             =   210
      Width           =   3915
   End
End
Attribute VB_Name = "frm340ShowIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objsSQL As clsLISSqlCodeMaster

Private Rkey As String
Private fTName As String            ' Access Table
Private fRowPos As Integer          ' Current Row

Private NewFg As Boolean

Private Sub cmdClear_Click()
    
    txtKeyCode.Text = ""
    Call ClearRtn
    Call LoadIndexList
    
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    Dim SSQL(1) As String
    Dim i As Integer
    Dim strFields As String
    Dim blnInsert As Boolean
    Dim RS As New Recordset

    strFields = ""
    With tblFields
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 2: strFields = strFields & .Value & ":"
            .Col = 3: strFields = strFields & .Value & ";"
        Next
    End With

    Set objsSQL = New clsLISSqlCodeMaster
    
    With objsSQL
        RS.Open .GetComCdIndex(Rkey, txtKeyCode.Text, "0"), DBConn
    End With
    
    If RS.RecordCount > 0 Then
        blnInsert = False
    Else
        blnInsert = True
    End If
    
    Set RS = Nothing
    
    If blnInsert Then
        SSQL(0) = objsSQL.SetComCdIndex(False, Rkey, txtKeyCode.Text, "0", txtKeyName.Text, , , , , _
                                       strFields, txtRemark.Text)
    Else
        SSQL(0) = objsSQL.SetComCdIndex(True, Rkey, txtKeyCode.Text, "0", txtKeyName.Text, , , , , _
                                        strFields, txtRemark.Text)
    End If
    
    DBConn.BeginTrans
    If InsertData(SSQL, False) Then
        DBConn.CommitTrans
        txtKeyCode.Text = ""
        Call ClearRtn
        Call LoadIndexList
    Else
        DBConn.RollbackTrans
        MsgBox Err.Description, vbExclamation
    End If
    
    Set objsSQL = Nothing
    
End Sub


Private Sub Form_Load()
    
    fRowPos = -1
    LoadIndexList
    LoadTableFields
    Rkey = "LC1"
    
End Sub

Private Sub LoadIndexList()

    Dim RS As Recordset
    Dim SSQL As String
    
    If Rkey = "" Then
        Rkey = "LC1"
    End If
    
    Set objsSQL = New clsLISSqlCodeMaster
    SSQL = objsSQL.GetComCdIndex(Rkey) & " and cdindex <> cdval1"
    With objsSQL
        Set RS = New Recordset
        RS.Open SSQL, DBConn
    End With
    
    With ssTableKey
        .MaxRows = 0
        .MaxRows = RS.RecordCount
        .Row = 0
        While (Not RS.EOF)
            .Row = .Row + 1
            .Col = 1: .Value = "" & RS.Fields("cdval1").Value
            .Col = 2: .Value = "" & RS.Fields("field1").Value
            RS.MoveNext
        Wend
    End With
    
    Set RS = Nothing
    Set objsSQL = Nothing
    
End Sub


Private Sub LoadTableFields()
    
    Dim i As Integer, SSQL As String
    Dim dsTableKey As Recordset
    Dim strFields As String
    
    Set objsSQL = New clsLISSqlCodeMaster
    With objsSQL
        Set dsTableKey = New Recordset
        dsTableKey.Open .GetComCdIndex(Rkey, Rkey), DBConn
    End With
    
    tblFields.MaxRows = 0
    
    If dsTableKey.EOF Then
        Set dsTableKey = Nothing
        Set objsSQL = Nothing
        Exit Sub
    End If
    
    strFields = "" & dsTableKey.Fields("text1").Value
    
    While (Trim(strFields) <> "")
        tblFields.MaxRows = tblFields.MaxRows + 1
        tblFields.Row = tblFields.MaxRows
        tblFields.Col = 1
        tblFields.Text = medShift(strFields, ";")
    Wend
    
    Set dsTableKey = Nothing
    Set objsSQL = Nothing
    
End Sub

Private Sub ClearRtn()

    txtKeyName.Text = ""
    With tblFields
        .Row = 1: .Row2 = .MaxRows
        .Col = 2: .Col2 = 3
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    txtRemark.Text = ""
    
End Sub


Private Sub lblTableName_Change()
    LoadIndexList
    LoadTableFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objsSQL = Nothing
End Sub

Private Sub optSelect_Click(Index As Integer)
    
    Dim FKey As String
    
    FKey = Rkey
    '
    If Index = 0 Then
        Rkey = "LC1"
    ElseIf Index = 1 Then
        Rkey = "LC2"
    Else
        Rkey = "LC4"
    End If
    '
    If FKey <> Rkey Then
        txtKeyCode = ""
        txtKeyName = ""
        txtRemark = ""
        tblFields.MaxRows = 0
        ssTableKey.MaxRows = 0
        Label3.Caption = "색 인 관 리"
        lblTableName_Change
        DoEvents
    End If
    '
End Sub

Private Sub ssTableKey_Click(ByVal Col As Long, ByVal Row As Long)

    Dim i As Integer, SSQL As String
    Dim dsTableKey As Recordset
    Dim strFields As String
    Dim strIndex As String

    If Row < 1 Then Exit Sub
    
    Call ClearRtn
    
    ssTableKey.Row = Row
    ssTableKey.Col = 2
    Label3.Caption = ssTableKey.Value
    '
    ssTableKey.Col = 1
    strIndex = ssTableKey.Value
    If strIndex Like "C1*" Then
        Rkey = "LC1"
    ElseIf strIndex Like "C2*" Then
        Rkey = "LC2"
    Else
        Rkey = "LC4"
    End If
    
    Set objsSQL = New clsLISSqlCodeMaster
    
    With objsSQL
        Set dsTableKey = New Recordset
        dsTableKey.Open .GetComCdIndex(Rkey, Trim(ssTableKey.Value)), DBConn
    End With
    
    If dsTableKey.EOF Then
        Set dsTableKey = Nothing
        Set objsSQL = Nothing
        Exit Sub
    End If
    
    txtKeyCode.Text = "" & dsTableKey.Fields("cdval1").Value
    txtKeyName.Text = "" & dsTableKey.Fields("field1").Value
    
    
    Dim aryFld() As String
    Dim ii As Long
    
    With tblFields
    
        .Row = 0
        strFields = "" & dsTableKey.Fields("text1").Value
        aryFld = Split(strFields, ";")
        
        For ii = LBound(aryFld) To UBound(aryFld)
            If .MaxRows <= .Row Then .MaxRows = .MaxRows + 1
            .Row = .Row + 1
            .Col = 2    'Field명
            .Value = Trim(medGetP(aryFld(ii), 1, ":"))
            If .Value <> "" Then
                .Col = 3    'Length
                .Value = Val(medGetP(aryFld(ii), 2, ":"))
            End If
        Next ii

        txtRemark.Text = "" & dsTableKey.Fields("text2").Value
        
        Set dsTableKey = Nothing
        Set objsSQL = Nothing
    
    End With
    
End Sub

Private Sub tblFields_Advance(ByVal AdvanceNext As Boolean)
    If AdvanceNext Then tblFields.MaxRows = tblFields.MaxRows + 1
End Sub

Private Sub txtKeyCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKeyCode_LostFocus()

    Dim i As Integer, SSQL As String
    Dim dsTableKey As Recordset
    Dim strFields As String

    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdOk.Name Then Exit Sub
    
    Set objsSQL = New clsLISSqlCodeMaster
    With objsSQL
        Set dsTableKey = New Recordset
        dsTableKey.Open .GetComCdIndex(Rkey, Trim(txtKeyCode.Text)), DBConn
    End With
        
    If dsTableKey.EOF Then
        NewFg = True
        Set dsTableKey = Nothing
        Set objsSQL = Nothing
        If tblFields.MaxRows = 0 Then tblFields.MaxRows = 1
        Call ClearRtn
        Exit Sub
    End If
        
    NewFg = False
    
    txtKeyCode.Text = "" & dsTableKey.Fields("cdval1").Value
    txtKeyName.Text = "" & dsTableKey.Fields("field1").Value
    
    With tblFields
        .Row = 0
        strFields = "" & dsTableKey.Fields("text1").Value
        While (Trim(strFields) <> "")
            .Row = .Row + 1
            .Col = 2
            .Value = medShift(strFields, ";")
        Wend
    End With
            
    txtRemark.Text = "" & dsTableKey.Fields("text1").Value
    
    Set dsTableKey = Nothing
    Set objsSQL = Nothing
    
End Sub
