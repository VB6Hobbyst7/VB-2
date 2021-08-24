VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmTmpEIReport 
   BackColor       =   &H00E8EEEE&
   Caption         =   "ELECTROPHORESIS & IMMUNOELECTROPHORESIS REPORT"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "frmTmpEIReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9060
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H00FEF5F3&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   4680
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   5130
      Width           =   1320
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00FEF5F3&
      Caption         =   "적용(&A)"
      Height          =   510
      Left            =   6000
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   5130
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FEF5F3&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   7320
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   5130
      Width           =   1320
   End
   Begin VB.ComboBox cboCmt 
      Height          =   300
      Left            =   5025
      Style           =   2  '드롭다운 목록
      TabIndex        =   5
      Top             =   3555
      Visible         =   0   'False
      Width           =   3630
   End
   Begin FPSpread.vaSpread tblData 
      Height          =   2640
      Left            =   540
      TabIndex        =   1
      Top             =   780
      Width           =   8025
      _Version        =   196608
      _ExtentX        =   14155
      _ExtentY        =   4657
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
      MaxCols         =   10
      MaxRows         =   7
      ScrollBars      =   0
      ShadowColor     =   15265518
      ShadowDark      =   15265518
      SpreadDesigner  =   "frmTmpEIReport.frx":08CA
   End
   Begin RichTextLib.RichTextBox rtfCmt 
      Height          =   1080
      Left            =   390
      TabIndex        =   0
      Top             =   3900
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1905
      _Version        =   393217
      BackColor       =   15924219
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpEIReport.frx":166C
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
      Left            =   405
      TabIndex        =   6
      Top             =   3570
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
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1125
      Left            =   165
      TabIndex        =   8
      Top             =   5295
      Visible         =   0   'False
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1984
      _Version        =   393217
      BackColor       =   15924219
      ScrollBars      =   3
      RightMargin     =   9000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmTmpEIReport.frx":1711
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
   Begin VB.Label lblEI 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "< ELECTROPHORESIS AND IMMUNOELECTROPHORESIS REPORT >"
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
      Left            =   285
      TabIndex        =   7
      Top             =   330
      Width           =   8475
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFCF7&
      FillStyle       =   0  '단색
      Height          =   420
      Left            =   165
      Shape           =   4  '둥근 사각형
      Top             =   240
      Width           =   8730
   End
End
Attribute VB_Name = "frmTmpEIReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Public Event DescClick(ByVal SelDesc As String)    'Event

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strTmp As String
    Dim strText As String
    Dim strPreData As String
    Dim iLen   As Long
    Dim ii As Integer
    Dim jj As Integer

    With tblData
        jj = 1
        For ii = 1 To .MaxRows


            .Row = ii
            
            If ii = 1 Then
                .Col = jj:
                    .Value = Format(.Value, "!" & String(19, "@"))
                .Col = 2:
                    '.CellType = CellTypeEdit
                    .Value = Format(Trim(.Value), String(5, "@"))
                .Col = 3:   .Value = Format(.Value, String(1, "@"))
                .Col = 4:
                    '.CellType = CellTypeEdit
                    .Value = Format(.Value, String(5, "@"))
                .Col = 5:
                    .Value = Format(Trim(.Value), String(5, "@"))
                .Col = 6:
                    '.CellType = CellTypeEdit
                    .Value = Format(.Value, String(5, "@"))
                .Col = 7:
                    .Value = Format(.Value, String(1, "@"))
                .Col = 8:
                    '.CellType = CellTypeEdit
                    .Value = Format(Trim(.Value), String(1, "@"))
                .Col = 9:   .Value = Format(.Value, String(5, "@"))
                .Col = 10:
                    '.CellType = CellTypeEdit
                    .Value = Format(.Value, String(1, "@"))
            Else
                .Col = jj:
                    .Value = Format(.Value, "!" & String(19, "@"))
                .Col = 2:
                    .CellType = CellTypeEdit
                    .Value = Format(Trim(.Value), String(1, "@"))
                .Col = 3:   .Value = Format(.Value, String(1, "@"))
                .Col = 4:
                    .CellType = CellTypeEdit
                    .Value = Format(.Value, String(1, "@"))
                .Col = 5:
                    .Value = Format(Trim(.Value), String(1, "@"))
                .Col = 6:
                    .CellType = CellTypeEdit
                    .Value = Format(.Value, String(1, "@"))
                    strPreData = .Value
                .Col = 7:
                    If strPreData = "" Then
                        .Value = Format(.Value, String(12, "@"))
                    Else
                        iLen = (12 - Len(strPreData))
                        .Value = Format(.Value, String(iLen, "@"))
                    End If
                    strPreData = .Value
                .Col = 8:
                    .CellType = CellTypeEdit
                    If strPreData = "" Then
                        .Value = Format(.Value, String(15, "@"))
                    Else
                        iLen = (15 - Len(strPreData))
                        .Value = Format(.Value, String(iLen, "@"))
                    End If
                    strPreData = .Value
                .Col = 9:
                    If strPreData = "" Then
                        .Value = Format(.Value, String(12, "@"))
                    Else
                        iLen = (12 - Len(strPreData))
                        .Value = Format(.Value, String(iLen, "@"))
                    End If
                .Col = 10:
                    .CellType = CellTypeEdit
                    .Value = Format(.Value, String(1, "@"))

            End If
        Next


    End With


    With rtfText
        .Text = .Text & lblEI.Caption & vbNewLine & vbNewLine
        '.Text = .Text & Trim(lblDiff.Caption) & Space(3) & Trim(txtCnt.Text) & Space(2) & Trim(lblDiff1.Caption) & Space(10) & Trim(lblME.Caption) & Space(2) & Trim(lblM.Caption) & Space(2) & Trim(lblDiv.Caption) & Space(2) & Trim(lblE.Caption) & vbNewLine & vbNewLine
        tblData.Row = 1: tblData.Row2 = tblData.MaxRows
        tblData.Col = 1: tblData.COL2 = tblData.MaxCols
        tblData.BlockMode = True
        strTmp = tblData.Clip
        tblData.BlockMode = False
        .Text = .Text & Replace(strTmp, vbTab, Space(1)) & vbNewLine & vbNewLine
        Debug.Print tblData.Clip
        .Text = .Text & lblCmt.Caption & vbNewLine & vbNewLine
        .Text = .Text & rtfCmt.Text & vbNewLine & vbNewLine
        '.Text = .Text & lblHD.Caption & vbNewLine & vbNewLine
'        .Text = .Text & rtfHD.Text & vbNewLine


        .SelStart = .Find(lblEI.Caption, 0, , rtfWholeWord + rtfNoHighlight)
        .SelLength = Len(lblEI.Caption)
        .SelBold = 1
        .SelFontSize = 10
        .SelUnderline = False
        .SelColor = vbBlack

        .SelStart = .Find(lblCmt.Caption, 0, , rtfWholeWord + rtfNoHighlight)
        .SelLength = Len(lblCmt.Caption)
        .SelBold = 1
        .SelFontSize = False
        .SelColor = vbBlack

    End With

    strText = rtfText.TextRTF
'    Debug.Print strText
    RaiseEvent DescClick(strText)
    Unload Me
End Sub
