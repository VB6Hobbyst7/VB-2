VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form frmTmpResult 
   BackColor       =   &H00E8EEEE&
   Caption         =   "일반 & Text 결과 Templete"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "frmTmpResult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows 기본값
   Begin DRcontrol1.DrFrame fraResult 
      Height          =   3735
      Left            =   375
      TabIndex        =   26
      Top             =   5685
      Visible         =   0   'False
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   6588
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   15518662
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView lvwResult 
         Height          =   3600
         Left            =   60
         TabIndex        =   27
         Top             =   60
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   6350
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16775406
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Template Text"
            Object.Width           =   12435
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox rtfTmp 
      Height          =   7500
      Left            =   90
      TabIndex        =   25
      Top             =   435
      Visible         =   0   'False
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   13229
      _Version        =   393217
      BackColor       =   15924219
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpResult.frx":08CA
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E8EEEE&
      Height          =   1845
      Left            =   0
      TabIndex        =   15
      Top             =   -75
      Width           =   9060
      Begin VB.ComboBox cboTemplate 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   6180
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Top             =   135
         Width           =   2805
      End
      Begin VB.ComboBox cboTmp1 
         Height          =   300
         Left            =   4770
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   465
         Width           =   4215
      End
      Begin RichTextLib.RichTextBox rtfTmp1 
         Height          =   1005
         Left            =   75
         TabIndex        =   17
         Top             =   780
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   1773
         _Version        =   393217
         BackColor       =   15924219
         ScrollBars      =   3
         RightMargin     =   9000
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmTmpResult.frx":096F
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Template Code"
         ForeColor       =   &H00313D46&
         Height          =   180
         Left            =   4800
         TabIndex        =   20
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label lblTitle1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   165
         TabIndex        =   18
         Top             =   300
         Width           =   3960
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000001&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H80000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   105
         Top             =   180
         Width           =   4110
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFF9F7&
         BackStyle       =   1  '투명하지 않음
         Height          =   285
         Left            =   4770
         Shape           =   4  '둥근 사각형
         Top             =   150
         Width           =   1365
      End
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H00FEF5F3&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   5055
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   9495
      Width           =   1320
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00FEF5F3&
      Caption         =   "적용(&A)"
      Height          =   510
      Left            =   6390
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   9495
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FEF5F3&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   7725
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   9495
      Width           =   1320
   End
   Begin VB.Frame fraTest 
      BackColor       =   &H00E8EEEE&
      Height          =   3420
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   9060
      Begin FPSpread.vaSpread tblData 
         Height          =   2520
         Left            =   15
         TabIndex        =   9
         Top             =   855
         Width           =   9030
         _Version        =   196608
         _ExtentX        =   15928
         _ExtentY        =   4445
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
         GrayAreaBackColor=   15265518
         MaxCols         =   9
         MaxRows         =   6
         ShadowColor     =   15265518
         ShadowDark      =   15265518
         SpreadDesigner  =   "frmTmpResult.frx":0A14
      End
      Begin VB.Label lblTest 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   270
         Width           =   4095
      End
      Begin VB.Shape shpSubMenu 
         BackColor       =   &H80000001&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H80000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   75
         Top             =   180
         Width           =   4100
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E8EEEE&
      Height          =   2715
      Left            =   0
      TabIndex        =   2
      Top             =   5010
      Width           =   9060
      Begin VB.ComboBox cboTemplate 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6225
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   120
         Width           =   2805
      End
      Begin VB.ComboBox cboTmp2 
         Height          =   300
         Left            =   4860
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   450
         Width           =   4170
      End
      Begin RichTextLib.RichTextBox rtfTmp2 
         Height          =   1905
         Left            =   45
         TabIndex        =   13
         Top             =   765
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   3360
         _Version        =   393217
         BackColor       =   15924219
         ScrollBars      =   3
         RightMargin     =   9000
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmTmpResult.frx":1283
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Template Code"
         ForeColor       =   &H00313D46&
         Height          =   180
         Left            =   4875
         TabIndex        =   22
         Top             =   210
         Width           =   1320
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFF9F7&
         BackStyle       =   1  '투명하지 않음
         Height          =   285
         Left            =   4845
         Shape           =   4  '둥근 사각형
         Top             =   135
         Width           =   1365
      End
      Begin VB.Label lblTitle2 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   300
         Width           =   3915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000001&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H80000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   60
         Top             =   180
         Width           =   4100
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E8EEEE&
      Height          =   1845
      Left            =   0
      TabIndex        =   1
      Top             =   7635
      Width           =   9060
      Begin VB.ComboBox cboTemplate 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   6225
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   90
         Width           =   2790
      End
      Begin VB.ComboBox cboTmp3 
         Height          =   300
         Left            =   4815
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   450
         Width           =   4200
      End
      Begin RichTextLib.RichTextBox rtfTmp3 
         Height          =   1005
         Left            =   75
         TabIndex        =   11
         Top             =   765
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   1773
         _Version        =   393217
         BackColor       =   15924219
         ScrollBars      =   3
         RightMargin     =   9000
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmTmpResult.frx":1328
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Template Code"
         ForeColor       =   &H00313D46&
         Height          =   180
         Left            =   4845
         TabIndex        =   24
         Top             =   210
         Width           =   1440
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFF9F7&
         BackStyle       =   1  '투명하지 않음
         Height          =   285
         Left            =   4815
         Shape           =   4  '둥근 사각형
         Top             =   135
         Width           =   1365
      End
      Begin VB.Label lblTitle3 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00794444&
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   315
         Width           =   3915
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H80000000&
         BorderWidth     =   3
         FillColor       =   &H00EEEBED&
         FillStyle       =   0  '단색
         Height          =   495
         Left            =   75
         Top             =   195
         Width           =   4100
      End
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2175
      Left            =   525
      TabIndex        =   14
      Top             =   5025
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   15924219
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpResult.frx":13CD
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
End
Attribute VB_Name = "frmTmpResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DescClick(ByVal SelDesc As String)    'Event

Private objSql      As New clsLISSqlStatement
Private objETest    As New clsLISSpecialTest
Private mvarTestCd  As String
Private mvarFocus   As String

Public objcbo       As ComboBox

Public Sub LoadData(ByVal pTestCd As String)
    Dim RS          As New Recordset
    Dim strSQL      As String
    Dim strTitle    As String
    Dim ii          As Long
    Dim jj          As Long
    
    mvarTestCd = pTestCd
    
    medClearTable tblData
    rtfTmp1.Text = ""
    rtfTmp2.Text = ""
    rtfTmp3.Text = ""
    lblTest.Caption = ""
    strSQL = objSql.SqlLAB031CodeList(LC2_TempletTest, "cdval2,field1,field2,field3,field4,text1", mvarTestCd, , "ORDER BY cdval2")
    RS.Open strSQL, dbconn
    
    lblTest.Caption = ""
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        With tblData
            .MaxRows = medGetP(RS.Fields("field4").Value & "", 1, "*")
            .MaxCols = medGetP(RS.Fields("field4").Value & "", 2, "*") * 3
            If lblTest.Caption = "" Then lblTest.Caption = RS.Fields("text1").Value & ""
            ii = 1
            Do Until RS.EOF
                .Row = ii
                .RowHeight(ii) = 16.8
                For jj = 1 To .MaxCols
                    .Col = jj
                    If (jj Mod 3) = 1 Then
                        .Value = RS.Fields("field1").Value & ""
                        .ColWidth(jj) = 15.63
                    ElseIf (jj Mod 3) = 2 Then
                        .Value = RS.Fields("field2").Value & ""
                        .ColWidth(jj) = 8
                    Else
                        .Value = RS.Fields("field3").Value & ""
                        .ColWidth(jj) = 5
                        RS.MoveNext
                    End If
                Next jj
                ii = ii + 1
            Loop
        End With
    End If
    
    Set RS = Nothing
    
    Call LoadTemp
    Call LoadCombo
End Sub

Private Sub LoadCombo()
    Dim ii As Long
    
    If objcbo.ListCount = 0 Then Exit Sub
    
    For ii = 1 To objcbo.ListCount
        cboTemplate(0).AddItem objcbo.List(ii - 1)
        cboTemplate(1).AddItem objcbo.List(ii - 1)
        cboTemplate(2).AddItem objcbo.List(ii - 1)
    Next
End Sub

Private Sub LoadTemp()
    Dim RS As New Recordset
    Dim strSQL As String
    
    strSQL = objSql.SqlLAB031CodeList(LC2_TempletText1, "cdval2,field1,field2,text1", mvarTestCd, , "ORDER BY cdval2")
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        cboTmp1.Clear
        lblTitle1.Caption = ""
        Do Until RS.EOF
            cboTmp1.AddItem RS.Fields("cdval2").Value & "" & vbTab & RS.Fields("field1").Value & ""
            If lblTitle1.Caption = "" Then lblTitle1.Caption = RS.Fields("field2").Value & ""
            RS.MoveNext
        Loop
    End If
    
    strSQL = objSql.SqlLAB031CodeList(LC2_TempletText2, "cdval2,field1,field2,text1", mvarTestCd, , "ORDER BY cdval2")
    
    Set RS = Nothing
    Set RS = New Recordset
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        cboTmp2.Clear
        lblTitle2.Caption = ""
        Do Until RS.EOF
            cboTmp2.AddItem RS.Fields("cdval2").Value & "" & vbTab & RS.Fields("field1").Value & ""
            If lblTitle2.Caption = "" Then lblTitle2.Caption = RS.Fields("field2").Value & ""
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
    
    strSQL = objSql.SqlLAB031CodeList(LC2_TempletText3, "cdval2,field1,field2,text1", mvarTestCd, , "ORDER BY cdval2")
    
    Set RS = Nothing
    Set RS = New Recordset
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        cboTmp3.Clear
        lblTitle3.Caption = ""
        Do Until RS.EOF
            cboTmp3.AddItem RS.Fields("cdval2").Value & "" & vbTab & RS.Fields("field1").Value & ""
            If lblTitle3.Caption = "" Then lblTitle3.Caption = RS.Fields("field2").Value & ""
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub

Private Sub cboTemplate_Click(Index As Integer)
    Dim sTemp       As String
    Dim sRType      As String
    Dim sTCode      As String
    Dim i           As Long
    Dim iPos        As Long

    If cboTemplate(Index).ListIndex < 0 Then Exit Sub

    sTemp = cboTemplate(Index).List(cboTemplate(Index).ListIndex)
    sRType = medGetP(sTemp, 4, vbTab)
    sTCode = medGetP(sTemp, 1, vbTab)
    
    Select Case Index
        Case "0":
            With rtfTmp1
                .TextRTF = objETest.GetTemplateRst(sRType, sTCode)
                .SelStart = 0
                .SelLength = Len(.Text)
                .SelStart = 0
                .SelLength = 0
            End With
        Case "1":
            With rtfTmp2
                .TextRTF = objETest.GetTemplateRst(sRType, sTCode)
                .SelStart = 0
                .SelLength = Len(.Text)
                .SelStart = 0
                .SelLength = 0
            End With
        Case "2":
            With rtfTmp3
                .TextRTF = objETest.GetTemplateRst(sRType, sTCode)
                .SelStart = 0
                .SelLength = Len(.Text)
                .SelStart = 0
                .SelLength = 0
            End With
    End Select
End Sub

Private Sub cboTmp1_Click()
    Dim RS As New Recordset
    Dim strSQL As String

    If Trim(cboTmp1.Text) = "" Then GoTo NoData
    strSQL = objSql.SqlLAB031CodeList(LC2_TempletText1, "cdval2,field1,field2,text1", mvarTestCd, Trim(medGetP(cboTmp1.Text, 1, vbTab)), "ORDER BY cdval2")

    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        rtfTmp1.Text = RS.Fields("text1").Value & ""
    End If
NoData:
    Set RS = Nothing
End Sub

Private Sub cboTmp2_Click()
    Dim RS As New Recordset
    Dim strSQL As String

    If Trim(cboTmp2.Text) = "" Then GoTo NoData
    strSQL = objSql.SqlLAB031CodeList(LC2_TempletText2, "cdval2,field1,field2,text1", mvarTestCd, Trim(medGetP(cboTmp2.Text, 1, vbTab)), "ORDER BY cdval2")

    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        rtfTmp2.Text = RS.Fields("text1").Value & ""
    End If
NoData:
    Set RS = Nothing
End Sub

Private Sub cboTmp3_Click()
    Dim RS As New Recordset
    Dim strSQL As String

    If Trim(cboTmp3.Text) = "" Then GoTo NoData
    strSQL = objSql.SqlLAB031CodeList(LC2_TempletText3, "cdval2,field1,field2,text1", mvarTestCd, Trim(medGetP(cboTmp3.Text, 1, vbTab)), "ORDER BY cdval2")

    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        rtfTmp3.Text = RS.Fields("text1").Value & ""
    End If
NoData:
    Set RS = Nothing
End Sub

Private Sub cmdExit_Click()
    RaiseEvent DescClick("")
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strTmp      As String
    Dim strText     As String
    Dim strHeader   As String
    Dim strSpHeader As String
    Dim strNewLine  As String
    Dim strTitle1   As String
    Dim strTitle2   As String
    Dim strTitle3   As String
    Dim strTest     As String
    Dim ii As Long
    Dim jj As Long
    
    strHeader = "\rtf1\ansi\ansicpg949\deff0{\fonttbl{\f0\fnil\fcharset129 \'b1\'bc\'b8\'b2;}}" & _
                "\viewkind4\uc1\pard\b\lang1042\f0\fs18 "
    strNewLine = "\par"
    
    strSpHeader = "\rtf1\ansi\ansicpg949\deff0{\fonttbl{\f0\fnil\fcharset129 \'b5\'b8\'bf\'f2\'c3\'bc;}}" & _
                  "\viewkind4\uc1\pard\lang1042\f0\fs18 "
    
    With tblData
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = 1 To .MaxCols
                .Col = jj:
                If (jj Mod 3) = 1 Then
                    If jj = 1 Then
                        .Value = Format(.Value, "!" & String(16, "@"))
                    Else
                        .Value = Format(.Value, String(16, "@"))
                    End If
                ElseIf (jj Mod 3) = 2 Then
                    .Value = Format(Trim(.Value), String(9, "@"))
                Else
                    .Value = Format(Trim(.Value), String(5, "@"))
                End If
            Next jj
        Next ii
    End With
    
    If Trim(lblTest.Caption) <> "" And tblData.DataRowCnt > 0 Then
        tblData.Row = 1: tblData.Row2 = tblData.MaxRows
        tblData.Col = 1: tblData.COL2 = tblData.MaxCols
        tblData.BlockMode = True
        strTmp = tblData.Clip
        tblData.BlockMode = False
        strTmp = Replace(strTmp, vbTab, Space(1))
        strTmp = Replace(strTmp, vbNewLine, Space(1) & strNewLine & Space(1))
    End If
    
    If lblTitle1.Caption <> "" And rtfTmp1.Text <> "" Then
        strTitle1 = strHeader & lblTitle1.Caption & strNewLine & strNewLine
    Else
        strTitle1 = ""
    End If
    
    If lblTest.Caption <> "" And tblData.DataRowCnt > 0 Then
        strTest = strHeader & lblTest.Caption & strNewLine & strNewLine
    Else
        strTest = ""
    End If
    
    If lblTitle2.Caption <> "" And rtfTmp2.Text <> "" Then
        strTitle2 = strHeader & lblTitle2.Caption & strNewLine & strNewLine
    Else
        strTitle2 = ""
    End If
    
    If lblTitle3.Caption <> "" And rtfTmp3.Text <> "" Then
        strTitle3 = strHeader & lblTitle3.Caption & strNewLine & strNewLine
    Else
        strTitle3 = ""
    End If
    
    With rtfText
        .Text = ""
        .TextRTF = "{" & strTitle1 & Mid(rtfTmp1.TextRTF, 2, Len(rtfTmp1.TextRTF) - 4) & strNewLine & _
                   strTest & strSpHeader & strTmp & strNewLine & _
                   strTitle2 & Mid(rtfTmp2.TextRTF, 2, Len(rtfTmp2.TextRTF) - 4) & strNewLine & _
                   strTitle3 & Mid(rtfTmp3.TextRTF, 2, Len(rtfTmp3.TextRTF))
        
'        If Trim(lblTitle1.Caption) <> "" AND rtfTmp1.Text <> "" Then
'            .SelStart = .Find(lblTitle1.Caption, 0, , rtfWholeWord + rtfNoHighlight)
'            .SelLength = Len(lblTitle1.Caption)
'            .SelBold = 1
'            .SelFontSize = 9
'            .SelColor = vbBlack
'        End If
'
'        If Trim(lblTest.Caption) <> "" AND tblData.DataRowCnt > 0 Then
'            .SelStart = .Find(lblTest.Caption, 0, , rtfWholeWord + rtfNoHighlight)
'            .SelLength = Len(lblTest.Caption)
'            .SelBold = 1
'            .SelFontSize = 9
'            .SelColor = vbBlack
'        End If
'
'        If Trim(lblTitle2.Caption) <> "" AND rtfTmp2.Text <> "" Then
'            .SelStart = .Find(lblTitle2.Caption, 0, , rtfWholeWord + rtfNoHighlight)
'            .SelLength = Len(lblTitle2.Caption)
'            .SelBold = 1
'            .SelFontSize = 9
'            .SelColor = vbBlack
'        End If
'
'        If Trim(lblTitle3.Caption) <> "" AND rtfTmp3.Text <> "" Then
'            .SelStart = .Find(lblTitle3.Caption, 0, , rtfWholeWord + rtfNoHighlight)
'            .SelLength = Len(lblTitle3.Caption)
'            .SelBold = 1
'            .SelFontSize = 9
'            .SelColor = vbBlack
'        End If
    End With

    strText = rtfText.TextRTF
    RaiseEvent DescClick(strText)
    Unload Me
End Sub

Private Sub Form_Load()
    medInitLvwHead lvwResult, "결과 Temlate,코드", "7050,0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
    Set objETest = Nothing
End Sub

Private Sub Clear()
    Call LoadData(mvarTestCd)
End Sub

Private Sub rtfTmp_DblClick()
    Select Case mvarFocus
        Case "1": rtfTmp1.TextRTF = rtfTmp.TextRTF
        Case "2": rtfTmp2.TextRTF = rtfTmp.TextRTF
        Case "3": rtfTmp3.TextRTF = rtfTmp.TextRTF
    End Select
    
    rtfTmp.Visible = False
End Sub

Private Sub rtfTmp1_DblClick()
    rtfTmp.TextRTF = rtfTmp1.TextRTF
    rtfTmp.Visible = True
    mvarFocus = "1"
End Sub

Private Sub rtfTmp2_DblClick()
    rtfTmp.TextRTF = rtfTmp2.TextRTF
    rtfTmp.Visible = True
    mvarFocus = "2"
End Sub

Private Sub rtfTmp3_DblClick()
    rtfTmp.TextRTF = rtfTmp3.TextRTF
    rtfTmp.Visible = True
    mvarFocus = "3"
End Sub

Private Sub rtfTmp1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF2 Then Exit Sub
    'If cboTemplate.ListIndex < 0 Then Exit Sub

    Dim sTemp As String
    Dim sRType As String
    Dim sTCode As String
    Dim iSPos As Long, iEPos As Long
    Dim strKey As String
    
    sTemp = cboTemplate(0).List(cboTemplate(0).ListIndex)
    sRType = medGetP(sTemp, 4, vbTab)
    sTCode = medGetP(sTemp, 1, vbTab)
    
    With rtfTmp1
        iSPos = .Find("<#" & sTCode & "_VALUE", .SelStart)
        If iSPos < 0 Then
            iSPos = .Find("<#" & sTCode & "_VALUE", 0)
            If iSPos < 0 Then Exit Sub
        End If
        iEPos = .Find(">", iSPos)
        .SelStart = iSPos
        .SelLength = iEPos - iSPos + 1
        .SelProtected = False
        strKey = .SelText
    End With
    mvarFocus = "1"
    Call LoadRstTemplateList(strKey)
    fraResult.Visible = True
    fraResult.ZOrder 0
    lvwResult.SetFocus
End Sub

Private Sub rtfTmp2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF2 Then Exit Sub
    'If cboTemplate.ListIndex < 0 Then Exit Sub
    Dim sTemp As String
    Dim sRType As String
    Dim sTCode As String
    Dim iSPos As Long, iEPos As Long
    Dim strKey As String
    
    sTemp = cboTemplate(1).List(cboTemplate(1).ListIndex)
    sRType = medGetP(sTemp, 4, vbTab)
    sTCode = medGetP(sTemp, 1, vbTab)
    
    With rtfTmp2
        iSPos = .Find("<#" & sTCode & "_VALUE", .SelStart)
        If iSPos < 0 Then
            iSPos = .Find("<#" & sTCode & "_VALUE", 0)
            If iSPos < 0 Then Exit Sub
        End If

        iEPos = .Find(">", iSPos)
        .SelStart = iSPos
        .SelLength = iEPos - iSPos + 1
        .SelProtected = False
        strKey = .SelText
    End With
    mvarFocus = "2"
    Call LoadRstTemplateList(strKey)
    fraResult.Visible = True
    fraResult.ZOrder 0
    lvwResult.SetFocus
End Sub

Private Sub rtfTmp3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyF2 Then Exit Sub
    Dim sTemp As String
    Dim sRType As String
    Dim sTCode As String
    Dim iSPos As Long, iEPos As Long
    Dim strKey As String
    
    sTemp = cboTemplate(2).List(cboTemplate(2).ListIndex)
    sRType = medGetP(sTemp, 4, vbTab)
    sTCode = medGetP(sTemp, 1, vbTab)
   
    With rtfTmp3
        iSPos = .Find("<#" & sTCode & "_VALUE", .SelStart)
        If iSPos < 0 Then
            iSPos = .Find("<#" & sTCode & "_VALUE", 0)
            If iSPos < 0 Then Exit Sub
        End If
        iEPos = .Find(">", iSPos)
        .SelStart = iSPos
        .SelLength = iEPos - iSPos + 1
        .SelProtected = False
        strKey = .SelText
    End With
    mvarFocus = "3"
    Call LoadRstTemplateList(strKey)
    fraResult.Visible = True
    fraResult.ZOrder 0
    lvwResult.SetFocus
End Sub

Private Sub LoadRstTemplateList(ByVal strKey As String)
    Dim objRs As Recordset
    Dim objComSql As New clsLISSqlStatement
    Dim iTmx As ListItem
    
    lvwResult.ListItems.Clear
    Set objRs = New Recordset
    objRs.Open objComSql.SqlLAB031CodeList(LC2_SpeAddTemp, "*", strKey), dbconn
    
    With objRs
        While Not .EOF
            Set iTmx = lvwResult.ListItems.Add(, , "" & .Fields("text1").Value)
            iTmx.SubItems(1) = .Fields("cdval2").Value & ""
            .MoveNext
        Wend
    End With
    Set objRs = Nothing
    Set objComSql = Nothing
End Sub

Private Sub lvwResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        fraResult.Visible = False
    End If
End Sub

Private Sub lvwResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iTmx As ListItem
    
    Set iTmx = lvwResult.HitTest(X, Y)
    If iTmx Is Nothing Then Exit Sub
    
    Select Case mvarFocus
        Case "1":
                rtfTmp1.SelProtected = False
                rtfTmp1.SelColor = DCM_Black
                rtfTmp1.SelText = iTmx.Text
            
                rtfTmp1.SetFocus
                fraResult.Visible = False
                DoEvents
                
                rtfTmp1.SetFocus
        Case "2":
                rtfTmp2.SelProtected = False
                rtfTmp2.SelColor = DCM_Black
                rtfTmp2.SelText = iTmx.Text
            
                rtfTmp2.SetFocus
                fraResult.Visible = False
                DoEvents
                
                rtfTmp2.SetFocus

        Case "3":
                rtfTmp3.SelProtected = False
                rtfTmp3.SelColor = DCM_Black
                rtfTmp3.SelText = iTmx.Text
            
                rtfTmp3.SetFocus
                fraResult.Visible = False
                DoEvents
                
                rtfTmp3.SetFocus
    End Select
End Sub

Private Sub lvwResult_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraResult.Visible = False
    Select Case mvarFocus
        Case "1": rtfTmp1.SetFocus
        Case "2": rtfTmp2.SetFocus
        Case "3": rtfTmp3.SetFocus
    End Select
End Sub
