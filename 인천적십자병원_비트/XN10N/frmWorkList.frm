VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   " 워크리스트 조회"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15930
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtBarcode 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   9090
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   1140
      Width           =   3525
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   315
      Left            =   840
      TabIndex        =   15
      Top             =   1710
      Width           =   195
   End
   Begin VB.TextBox txtQuery 
      Height          =   645
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmWorkList.frx":000C
      Top             =   9810
      Visible         =   0   'False
      Width           =   15435
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   15930
      TabIndex        =   1
      Top             =   0
      Width           =   15930
      Begin VB.TextBox txtSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13410
         TabIndex        =   13
         Text            =   "1"
         Top             =   450
         Width           =   1125
      End
      Begin VB.CommandButton cmdSendClose 
         Caption         =   "전송후 닫기"
         Height          =   375
         Left            =   9600
         TabIndex        =   10
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         Height          =   375
         Left            =   10890
         TabIndex        =   8
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "전송"
         Height          =   375
         Left            =   8490
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         Height          =   375
         Left            =   7380
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   3540
         TabIndex        =   2
         Top             =   450
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   5430
         TabIndex        =   4
         Top             =   450
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   40457
      End
      Begin VB.Label lblQuery 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "쿼리보기"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   13470
         TabIndex        =   16
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Seq"
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
         Height          =   180
         Left            =   12810
         TabIndex        =   14
         Top             =   510
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   5250
         TabIndex        =   5
         Top             =   540
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회기간"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   2580
         TabIndex        =   3
         Top             =   540
         Width           =   780
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   2370
         Picture         =   "frmWorkList.frx":0012
         Top             =   510
         Width           =   150
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmWorkList.frx":03FC
         Top             =   0
         Width           =   12900
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   8115
      Left            =   300
      TabIndex        =   11
      Top             =   1650
      Width           =   15465
      _Version        =   393216
      _ExtentX        =   27279
      _ExtentY        =   14314
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   20
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   21
      OperationMode   =   2
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   14548991
      SpreadDesigner  =   "frmWorkList.frx":1B3F
      UserResize      =   2
      ScrollBarTrack  =   1
      ShowScrollTips  =   3
   End
   Begin VB.Label Label3 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검체번호"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7230
      TabIndex        =   17
      Top             =   1200
      Width           =   1755
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   ">> 2016-10-16 부터  2016-10-16 까지의 워크리스트 내역입니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   300
      TabIndex        =   9
      Top             =   1260
      Width           =   5895
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkAll_Click()
    Dim iRow As Long
    
    With spdWork
        If chkAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 1
            Next iRow
        ElseIf chkAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = 1
                
                .Value = 0
            Next iRow
        End If
    End With
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSearch_Click()
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"))
    
End Sub

Private Sub cmdSend_Click()
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdWork, intWRow, colBARCODE)
                For intORow = 1 To frmMain.spdOrder.MaxRows
                    frmMain.spdOrder.Row = intORow
                    frmMain.spdOrder.Col = colBARCODE
                    If strBarno = GetText(frmMain.spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next
                If blnSame = False Then
                    frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), frmMain.spdOrder.MaxRows, colSPECNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), frmMain.spdOrder.MaxRows, colCHECKBOX)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), frmMain.spdOrder.MaxRows, colHOSPDATE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), frmMain.spdOrder.MaxRows, colBARCODE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.MaxRows, colSEQNO)
                    'Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colRACKNO), frmMain.spdOrder.MaxRows, colRACKNO)
                    'Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPOSNO), frmMain.spdOrder.MaxRows, colPOSNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), frmMain.spdOrder.MaxRows, colCHARTNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), frmMain.spdOrder.MaxRows, colPID)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), frmMain.spdOrder.MaxRows, colINOUT)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), frmMain.spdOrder.MaxRows, colPNAME)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), frmMain.spdOrder.MaxRows, colPSEX)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), frmMain.spdOrder.MaxRows, colPAGE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), frmMain.spdOrder.MaxRows, colPJUMIN)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), frmMain.spdOrder.MaxRows, colOCNT)
                    
                    varItems = GetText(spdWork, intWRow, colITEMS)
                    varItems = Split(varItems, "/")
                    For intItems = 0 To UBound(varItems)
                        For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                            frmMain.spdOrder.Row = 0
                            frmMain.spdOrder.Col = intOCol
                            If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                                .Row = frmMain.spdOrder.MaxRows
                                Call SetText(frmMain.spdOrder, "◆", frmMain.spdOrder.MaxRows, intOCol)
                            End If
                        Next
                    Next
                    
                    frmMain.spdOrder.RowHeight(-1) = 12
                End If
            End If
        Next
    End With
    
End Sub

Private Sub cmdSendClose_Click()
    
    Call cmdSend_Click
    
    Call cmdClose_Click
    
End Sub

Private Sub Form_Load()
    
    Call CtlInitializing

End Sub


Private Sub CtlInitializing()
    
    spdWork.MaxRows = 0
    
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    lblStatus.Caption = ""
    txtQuery.Text = ""
    txtSeq.Text = "1"
    
    txtBarcode.Text = ""
    
End Sub


Private Sub lblQuery_DblClick()
    If txtQuery.Visible = True Then
        txtQuery.Visible = False
    Else
        txtQuery.Visible = True
    End If
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim varItems As Variant
    Dim intItems As Integer
    Dim intOCol  As Integer
    
    If Row = 0 Then Exit Sub
    
    frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
    
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colSPECNO), frmMain.spdOrder.MaxRows, colSPECNO)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colCHECKBOX), frmMain.spdOrder.MaxRows, colCHECKBOX)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colHOSPDATE), frmMain.spdOrder.MaxRows, colHOSPDATE)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colBARCODE), frmMain.spdOrder.MaxRows, colBARCODE)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colSEQNO), frmMain.spdOrder.MaxRows, colSEQNO)
    'Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colRACKNO), frmMain.spdOrder.MaxRows, colRACKNO)
    'Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colPOSNO), frmMain.spdOrder.MaxRows, colPOSNO)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colCHARTNO), frmMain.spdOrder.MaxRows, colCHARTNO)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colPID), frmMain.spdOrder.MaxRows, colPID)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colINOUT), frmMain.spdOrder.MaxRows, colINOUT)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colPNAME), frmMain.spdOrder.MaxRows, colPNAME)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colPSEX), frmMain.spdOrder.MaxRows, colPSEX)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colPAGE), frmMain.spdOrder.MaxRows, colPAGE)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colPJUMIN), frmMain.spdOrder.MaxRows, colPJUMIN)
    Call SetText(frmMain.spdOrder, GetText(spdWork, Row, colOCNT), frmMain.spdOrder.MaxRows, colOCNT)
    
    varItems = GetText(spdWork, Row, colITEMS)
    varItems = Split(varItems, "/")
    For intItems = 0 To UBound(varItems)
        For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
            frmMain.spdOrder.Row = 0
            frmMain.spdOrder.Col = intOCol
            If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                '.Row = frmMain.spdOrder.MaxRows
                Call SetText(frmMain.spdOrder, "◆", frmMain.spdOrder.MaxRows, intOCol)
            End If
        Next
    Next
    
    frmMain.spdOrder.RowHeight(-1) = 12
    
    Call spdWork.DeleteRows(Row, 1)
    spdWork.MaxRows = spdWork.MaxRows - 1

End Sub

Private Sub spdWork_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    Dim strSeq As String
    
    If KeyAscii = vbKeyReturn Then
        With spdWork
            If .ActiveCol = colSEQNO Then
                strSeq = GetText(spdWork, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "숫자만 입력이 가능합니다"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdWork, strSeq + 1, intRow, colSEQNO)
                Next
            End If
        End With
    End If
    
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And txtBarcode.Text <> "" Then
        Call GetWorkList_Barcode(Trim(txtBarcode.Text))
    End If
End Sub

Private Sub txtSeq_KeyPress(KeyAscii As Integer)
    Dim intRow      As Integer
    
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(txtSeq.Text) Then
            MsgBox "숫자만 입력이 가능합니다", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If
        
        With spdWork
            For intRow = .ActiveRow To .DataRowCnt
                .Row = intRow
                .Col = colSEQNO
                .Text = Val(txtSeq.Text)
                txtSeq.Text = txtSeq.Text + 1
            Next
        End With
    End If

End Sub
