VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmTmpBMReport 
   BackColor       =   &H00E8EEEE&
   Caption         =   "Bone Marrow Report"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "frmTmpBMReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H00FEF5F3&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   4725
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   7830
      Width           =   1320
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00FEF5F3&
      Caption         =   "적용(&A)"
      Height          =   510
      Left            =   6045
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   7830
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FEF5F3&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   7365
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   7830
      Width           =   1320
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1125
      Left            =   540
      TabIndex        =   20
      Top             =   8055
      Visible         =   0   'False
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1984
      _Version        =   393217
      BackColor       =   15924219
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpBMReport.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboHD 
      Height          =   300
      Left            =   5070
      Style           =   2  '드롭다운 목록
      TabIndex        =   19
      Top             =   6840
      Width           =   3630
   End
   Begin VB.ComboBox cboPB 
      Height          =   300
      Left            =   5055
      Style           =   2  '드롭다운 목록
      TabIndex        =   18
      Top             =   180
      Width           =   3630
   End
   Begin VB.ComboBox cboCmt 
      Height          =   300
      Left            =   5070
      Style           =   2  '드롭다운 목록
      TabIndex        =   17
      Top             =   5220
      Width           =   3630
   End
   Begin VB.TextBox txtCnt 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   4485
      TabIndex        =   8
      Text            =   "500"
      Top             =   2400
      Width           =   780
   End
   Begin FPSpread.vaSpread tblData 
      Height          =   2265
      Left            =   480
      TabIndex        =   6
      Top             =   2820
      Width           =   8025
      _Version        =   196608
      _ExtentX        =   14155
      _ExtentY        =   3995
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   6
      ScrollBars      =   0
      ShadowColor     =   15265518
      ShadowDark      =   15265518
      SpreadDesigner  =   "frmTmpBMReport.frx":0AE6
   End
   Begin RichTextLib.RichTextBox rtfPb 
      Height          =   1215
      Left            =   435
      TabIndex        =   0
      Top             =   540
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   15924219
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpBMReport.frx":1481
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfCmt 
      Height          =   1080
      Left            =   435
      TabIndex        =   1
      Top             =   5565
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1905
      _Version        =   393217
      BackColor       =   15924219
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpBMReport.frx":169D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfHD 
      Height          =   570
      Left            =   450
      TabIndex        =   2
      Top             =   7185
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1005
      _Version        =   393217
      BackColor       =   15924219
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpBMReport.frx":18B9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MedControls1.LisLabel lblCmt 
      Height          =   315
      Left            =   450
      TabIndex        =   21
      Top             =   5235
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
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
      Caption         =   "Comment"
      Appearance      =   0
   End
   Begin VB.Label lblE 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   7890
      TabIndex        =   13
      Top             =   2430
      Width           =   615
   End
   Begin VB.Label lblDiv 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   7620
      TabIndex        =   12
      Top             =   2445
      Width           =   300
   End
   Begin VB.Label lblM 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   6945
      TabIndex        =   11
      Top             =   2445
      Width           =   645
   End
   Begin VB.Label lblME 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "G:E ratio"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   240
      Left            =   5775
      TabIndex        =   10
      Top             =   2445
      Width           =   1200
   End
   Begin VB.Label lblDiff1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   240
      Left            =   5235
      TabIndex        =   9
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label lblDiff 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "DIFFERENTIAL COUNT(Total count :            "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   510
      TabIndex        =   7
      Top             =   2430
      Width           =   4740
   End
   Begin VB.Label lblHD 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "< HEMATOLOGIC DIAGNOSIS >"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   6810
      Width           =   3795
   End
   Begin VB.Label lblBM 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "< BONE MARROW FINDINGS >"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   630
      TabIndex        =   4
      Top             =   1995
      Width           =   3720
   End
   Begin VB.Label lblPB 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "< PERIPHERAL BLOOD FINDINGS >"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   585
      TabIndex        =   3
      Top             =   195
      Width           =   4275
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFCF7&
      FillStyle       =   0  '단색
      Height          =   420
      Left            =   450
      Shape           =   4  '둥근 사각형
      Top             =   90
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFCF7&
      FillStyle       =   0  '단색
      Height          =   420
      Left            =   435
      Shape           =   4  '둥근 사각형
      Top             =   1905
      Width           =   4560
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFCF7&
      FillStyle       =   0  '단색
      Height          =   420
      Left            =   450
      Shape           =   4  '둥근 사각형
      Top             =   6720
      Width           =   4545
   End
End
Attribute VB_Name = "frmTmpBMReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DescClick(ByVal SelDesc As String)    'Event
Private objSql As New clsLISSqlStatement

Private Sub cboPB_Click()
    Dim RS As New Recordset
    Dim strSQL As String
    
    If Trim(cboPB.Text) = "" Then GoTo NoData
    
    strSQL = objSql.GetPBFindings(Trim(medGetP(cboPB.Text, 1, vbTab)))
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        rtfPb.Text = RS.Fields("text1").Value & ""
    End If
NoData:
    Set RS = Nothing
End Sub

Private Sub cboCmt_Click()
    Dim RS As New Recordset
    Dim strSQL As String
    
    If Trim(cboCmt.Text) = "" Then GoTo NoData
    
    strSQL = objSql.GetBMComment(Trim(medGetP(cboCmt.Text, 1, vbTab)))
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        rtfCmt.Text = RS.Fields("text1").Value & ""
    End If
NoData:
    Set RS = Nothing
End Sub

Private Sub cboHD_Click()
    Dim RS As New Recordset
    Dim strSQL As String
    
    If Trim(cboHD.Text) = "" Then GoTo NoData
    
    strSQL = objSql.GetHDiagnosis(Trim(medGetP(cboHD.Text, 1, vbTab)))
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        rtfHD.Text = RS.Fields("text1").Value & ""
    End If
NoData:
    Set RS = Nothing
End Sub

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdExit_Click()
    RaiseEvent DescClick("")
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strTmp As String
    Dim strText As String
    Dim ii As Integer
    Dim jj As Integer
    
    With tblData
        jj = 1
        For ii = 1 To .MaxRows
            
           
            .Row = ii
            .Col = jj:
    '                    If jj = 1 Then
                .Value = Format(.Value, "!" & String(19, "@"))
    '                    Else
    '                        .Value = Format(rs.Fields("field1").Value & "", String(20, "@"))
    '                    End If
            .Col = 2:
                .CellType = CellTypeEdit
                .Value = Format(Trim(.Value), String(5, "@"))
            .Col = 3:   .Value = Format("%", String(1, "@"))
            .Col = 4:
                .Value = Format(.Value, String(19, "@"))
            .Col = 5:
                .CellType = CellTypeEdit
                .Value = Format(Trim(.Value), String(5, "@"))
            .Col = 6:   .Value = Format("%", String(1, "@"))
            .Col = 7:
                .Value = Format(.Value, String(19, "@"))
            .Col = 8:
                .CellType = CellTypeEdit
                .Value = Format(Trim(.Value), String(5, "@"))
            .Col = 9:   .Value = Format("%", String(1, "@"))
                  
        Next
            
       
    End With
    
    
    With rtfText
        .Text = lblPB.Caption & vbNewLine & vbNewLine
        .Text = .Text & rtfPb.Text & vbNewLine & vbNewLine
        .Text = .Text & lblBM.Caption & vbNewLine & vbNewLine
        .Text = .Text & Trim(lblDiff.Caption) & Space(3) & Trim(txtCnt.Text) & Space(2) & Trim(lblDiff1.Caption) & Space(10) & Trim(lblME.Caption) & Space(2) & Trim(lblM.Caption) & Space(2) & Trim(lblDiv.Caption) & Space(2) & Trim(lblE.Caption) & vbNewLine & vbNewLine
        tblData.Row = 1: tblData.Row2 = tblData.MaxRows
        tblData.Col = 1: tblData.COL2 = tblData.MaxCols
        tblData.BlockMode = True
        strTmp = tblData.Clip
        tblData.BlockMode = False
        .Text = .Text & Replace(strTmp, vbTab, Space(1)) & vbNewLine & vbNewLine
'        Debug.Print tblData.Clip
        .Text = .Text & lblCmt.Caption & vbNewLine & vbNewLine
        .Text = .Text & rtfCmt.Text & vbNewLine & vbNewLine
        .Text = .Text & lblHD.Caption & vbNewLine & vbNewLine
        .Text = .Text & rtfHD.Text & vbNewLine
        
        
        .SelStart = .Find(lblPB.Caption, 0, , rtfWholeWord + rtfNoHighlight)
        .SelLength = Len(lblPB.Caption)
        .SelBold = 1
        .SelFontSize = 11
'        .SelUnderline = True
        .SelColor = vbBlack
        
        .SelStart = .Find(lblBM.Caption, 0, , rtfWholeWord + rtfNoHighlight)
        .SelLength = Len(lblBM.Caption)
        .SelBold = 1
        .SelFontSize = 11
'        .SelUnderline = True
        .SelColor = vbBlack
        
        .SelStart = .Find(lblCmt.Caption, 0, , rtfWholeWord + rtfNoHighlight)
        .SelLength = Len(lblCmt.Caption)
        .SelBold = 1
        .SelFontSize = 10
        .SelColor = vbBlack
        
        .SelStart = .Find(lblHD.Caption, 0, , rtfWholeWord + rtfNoHighlight)
        .SelLength = Len(lblHD.Caption)
        .SelBold = 1
        .SelFontSize = 11
'        .SelUnderline = True
        .SelColor = vbBlack
        
        
    End With
    
    strText = rtfText.TextRTF
'    Debug.Print strText
    RaiseEvent DescClick(strText)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim RS As New Recordset
    Dim strSQL As String
    Dim ii As Integer
    Dim jj As Integer
    
    Clear
    
    strSQL = objSql.GetBMTestCd
    
    RS.Open strSQL, dbconn
    With tblData
        
        If RS.RecordCount > 0 Then
            RS.MoveFirst
            ii = 1
            jj = 1
            Do Until RS.EOF
                    .Row = ii
                    .Col = jj:
                    If jj = 1 Then
                        .Value = Format(RS.Fields("field1").Value & "", "!" & String(19, "@"))
                    Else
                        .Value = Format(RS.Fields("field1").Value & "", String(20, "@"))
                    End If
'                    .Col = 2:   .Value = Format("0.0", String(5, "@"))
                    .Col = 3:   .Value = Format("%", String(1, "@"))
'                    .Col = 5:   .Value = Format("0.0", String(5, "@"))
                    
                    .Col = 6:   .Value = Format("%", String(1, "@"))
'                    .Col = 8:   .Value = Format("0.0", String(5, "@"))
                    .Col = 9:   .Value = Format("%", String(1, "@"))
                    jj = jj + 3
                    If jj > 7 Then
                        jj = 1
                        ii = ii + 1
                    End If
                    RS.MoveNext
            Loop
            
        End If
    End With
    LoadTemp
NoData:
    Set RS = Nothing
End Sub



Private Sub LoadTemp()
    Dim RS As New Recordset
    Dim strSQL As String
    
    strSQL = objSql.GetPBFindings
    
    RS.Open strSQL, dbconn
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        cboPB.Clear
        Do Until RS.EOF
            cboPB.AddItem RS.Fields("cdval1").Value & "" & vbTab & RS.Fields("field1").Value & ""
            RS.MoveNext
        Loop
    End If
    
    
    strSQL = objSql.GetBMComment
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open strSQL, dbconn
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        cboCmt.Clear
        Do Until RS.EOF
            cboCmt.AddItem RS.Fields("cdval1").Value & "" & vbTab & RS.Fields("field1").Value & ""
            RS.MoveNext
        Loop
    End If
    
    
    strSQL = objSql.GetHDiagnosis
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open strSQL, dbconn
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        cboHD.Clear
        Do Until RS.EOF
            cboHD.AddItem RS.Fields("cdval1").Value & "" & vbTab & RS.Fields("field1").Value & ""
            RS.MoveNext
        Loop
    End If
    
    Set RS = Nothing
End Sub

Private Sub Clear()
    rtfPb.Text = ""
    rtfCmt.Text = ""
    rtfHD.Text = ""
    rtfText.Text = ""
    txtCnt.Text = ""
    
    tableClear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub

Private Sub tblData_KeyDown(KeyCode As Integer, Shift As Integer)
    With tblData
        If KeyCode = vbKeyReturn Then
            If .ActiveRow = .MaxRows And .ActiveCol = 8 Then
                cboCmt.SetFocus
            ElseIf .ActiveCol <> 8 And .ActiveRow = .MaxRows Then
                .Row = 1: .Col = .ActiveCol + 3
                .Action = ActionActiveCell
            Else
                .Row = .ActiveRow + 1: .Col = .ActiveCol
                .Action = ActionActiveCell
            End If
             
        End If
    End With
End Sub


Private Sub tblData_LostFocus()
    Dim ii As Single
    Dim jj As Single
    ii = 0: jj = 0
    
    With tblData
        'm
        .Row = 1: .Col = 2: ii = .Value
        .Row = 2: .Col = 2: ii = ii + .Value
        .Row = 3: .Col = 2: ii = ii + .Value
        .Row = 4: .Col = 2: ii = ii + .Value
        .Row = 5: .Col = 2: ii = ii + .Value
        .Row = 6: .Col = 2: ii = ii + .Value
        .Row = 1: .Col = 5: ii = ii + .Value
        .Row = 2: .Col = 5: ii = ii + .Value
        'g
        .Row = 4: .Col = 8: ii = ii + .Value
        'e
        .Row = 3: .Col = 5: jj = .Value
        .Row = 4: .Col = 5: jj = jj + .Value
        .Row = 5: .Col = 5: jj = jj + .Value
        .Row = 6: .Col = 5: jj = jj + .Value
        
    End With
    If ii = 0 Or jj = 0 Then Exit Sub
    lblM.Caption = Round(ii / jj, 1)
    lblE.Caption = "1.0"
End Sub

Private Sub txtCnt_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCnt.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        With tblData
            .SetFocus
            .Row = 1: .Col = 2
            .Action = ActionActiveCell
        End With
    End If
End Sub

Private Sub tableClear()
    Dim ii As Integer
    
    With tblData
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = 2:   .Value = Format("0.0", String(5, "@"))
            .Col = 5:   .Value = Format("0.0", String(5, "@"))
            .Col = 8:   .Value = Format("0.0", String(5, "@"))
        Next
    End With
End Sub


