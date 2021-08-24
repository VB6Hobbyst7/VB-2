VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   " 워크리스트 조회"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14805
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin FPSpread.vaSpread vasExcel 
      Height          =   1725
      Left            =   3720
      TabIndex        =   17
      Top             =   4650
      Visible         =   0   'False
      Width           =   5085
      _Version        =   393216
      _ExtentX        =   8969
      _ExtentY        =   3043
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":000C
   End
   Begin VB.CommandButton cmdMake 
      BackColor       =   &H00C0FFFF&
      Caption         =   "검사전송"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9720
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   1170
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Top             =   1710
      Width           =   195
   End
   Begin VB.TextBox txtQuery 
      Height          =   645
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmWorkList.frx":58D1
      Top             =   9810
      Visible         =   0   'False
      Width           =   14145
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
      ScaleWidth      =   14805
      TabIndex        =   0
      Top             =   0
      Width           =   14805
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
         TabIndex        =   12
         Text            =   "1"
         Top             =   450
         Width           =   885
      End
      Begin VB.CommandButton cmdSendClose 
         Caption         =   "전송후 닫기"
         Height          =   375
         Left            =   9600
         TabIndex        =   9
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         Height          =   375
         Left            =   10890
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "전송"
         Height          =   375
         Left            =   8490
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         Height          =   375
         Left            =   7380
         TabIndex        =   5
         Top             =   420
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   3540
         TabIndex        =   1
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
         Format          =   127336449
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   5430
         TabIndex        =   3
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
         Format          =   127336449
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
         Left            =   13440
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   4
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
         TabIndex        =   2
         Top             =   540
         Width           =   780
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   2370
         Picture         =   "frmWorkList.frx":58D7
         Top             =   510
         Width           =   150
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmWorkList.frx":5CC1
         Top             =   0
         Width           =   12900
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   8115
      Left            =   300
      TabIndex        =   10
      Top             =   1650
      Width           =   14145
      _Version        =   393216
      _ExtentX        =   24950
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
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   14548991
      SpreadDesigner  =   "frmWorkList.frx":7404
      UserResize      =   2
      ScrollBarTrack  =   1
      ShowScrollTips  =   3
   End
   Begin MSComDlg.CommonDialog ExcelFile 
      Left            =   13080
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   8
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

Private Sub cmdMake_Click()

    '-- 로컬 워크리스트 저장
'    Call SetLocalDB_WorkList

    Dim RS1         As ADODB.Recordset
    Dim sFileName   As String
    Dim iRow        As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim strResult   As String
    Dim varResult   As Variant
    Dim strChartNo  As String
    Dim strRegDate  As String
    Dim strItems    As String
    
    If spdWork.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        
        vasExcel.MaxRows = 0
        vasExcel.MaxCols = 20
        
        With spdWork
            For iRow = 1 To .MaxRows
                If iRow = 1 Then
                    j = 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colHOSPDATE)), 0, j:    j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colBARCODE)), 0, j:     j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colSEQNO)), 0, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colRACKNO)), 0, j:      j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colPOSNO)), 0, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colINOUT)), 0, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colCHARTNO)), 0, j:     j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colPID)), 0, j:         j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colPNAME)), 0, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colPSEX)), 0, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colPAGE)), 0, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colPJUMIN)), 0, j:      j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colKEY1)), 0, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colKEY2)), 0, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colOCNT)), 0, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colRCNT)), 0, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colSTATE)), 0, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, 0, colITEMS)), 0, j:       j = j + 1
                    SetText vasExcel, "검사코드", 0, j:       j = j + 1
                    
                    vasExcel.MaxCols = j - 1
                End If
                
                j = 1
                
                If Trim(GetText(spdWork, iRow, colCHECKBOX)) = "1" Then
                    vasExcel.MaxRows = vasExcel.MaxRows + 1
                    k = vasExcel.MaxRows
                    strRegDate = ""
                    strChartNo = ""
                    strItems = ""
                    
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colHOSPDATE)), k, j:    j = j + 1: strRegDate = Trim(GetText(spdWork, iRow, colHOSPDATE))
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colBARCODE)), k, j:     j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colSEQNO)), k, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colRACKNO)), k, j:      j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colPOSNO)), k, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colINOUT)), k, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colCHARTNO)), k, j:     j = j + 1: strChartNo = Trim(GetText(spdWork, iRow, colCHARTNO))
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colPID)), k, j:         j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colPNAME)), k, j:       j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colPSEX)), k, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colPAGE)), k, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colPJUMIN)), k, j:      j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colKEY1)), k, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colKEY2)), k, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colOCNT)), k, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colRCNT)), k, j:        j = j + 1
                    SetText vasExcel, Trim(GetText(spdWork, iRow, colITEMS)), k, j:       j = j + 1
                    
                    
                    SQL = ""
                    SQL = SQL & "SELECT DISTINCT L.LABODRCOD as ITEM"
                    SQL = SQL & ", L.LABODRSTP as SEQ " & vbCrLf
                    SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D, ME_MAN M" & vbCrLf
                    SQL = SQL & " WHERE L.LABCHTNUM = '" & strChartNo & "'" & vbCr
                    SQL = SQL & "   AND L.LABODRDTE = '" & strRegDate & "'" & vbCr
                    SQL = SQL & "   AND L.LABKEYNUM = D.DATKEYNUM " & vbCrLf                    '-- 테이블연결키값
                    SQL = SQL & "   AND L.LABATTEND = D.DATATTEND " & vbCrLf                    '-- 내원번호
                    SQL = SQL & "   AND L.LABATTEND = M.MANATTEND " & vbCrLf                    '-- 내원번호
                    SQL = SQL & "   AND L.LABCHTNUM = D.DATCHTNUM " & vbCrLf                    '-- 챠트번호
                    SQL = SQL & "   AND L.LABCHTNUM = M.MANCHTNUM " & vbCrLf                    '-- 챠트번호
                    SQL = SQL & "   AND L.LABODRDTE = D.DATODRDTE " & vbCrLf                    '-- 처방일자
                    SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")" & vbCrLf
                    SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL) " & vbCrLf    '-- 취소여부
                    SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)" & vbCrLf
                    SQL = SQL & "   AND L.LABENDDEP < '3' " & vbCrLf                            '-- 처리상태 (2:접수, 3:결과입력)
                    
                    '-- Record Count 가져옴
                    AdoCn.CursorLocation = adUseClient
                    Set RS1 = AdoCn.Execute(SQL, , 1)
                    If Not RS1.EOF = True And Not RS1.BOF = True Then
                        Do Until RS1.EOF
                            If strItems = "" Then
                                strItems = Trim(RS1.Fields("ITEM")) & ""
                            Else
                                strItems = strItems & "/" & Trim(RS1.Fields("ITEM"))
                            End If
                        Loop
                    End If
                    
                    SetText vasExcel, strItems, k, j:     j = j + 1
                End If
            Next iRow
            
            vasExcel.RowHeight(-1) = 15
            
        End With
        
        If vasExcel.DataRowCnt < 1 Then
            MsgBox "출력할 자료를 선택하세요", , "알 림"
            Exit Sub
        Else
            ExcelFile.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
            ExcelFile.ShowSave
            sFileName = ExcelFile.FileName
'            SaveExcel sFileName, vasExcel
        End If
    End If

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
    Dim strChartNo      As String
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
                    'If strChartNo = GetText(frmMain.spdOrder, intORow, colCHARTNO) Then
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
    
End Sub




Private Sub lblQuery_DblClick()
    If txtQuery.Visible = True Then
        txtQuery.Visible = False
    Else
        txtQuery.Visible = True
    End If
End Sub

Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SetSpreadSort(spdWork, 0)
    End If
    
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    Dim strBarno_Work   As String
    
    If Row = 0 Then Exit Sub
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colCHARTNO
    strBarno_Work = Trim(spdWork.Text)
    
    With frmMain.spdOrder
        blnSame = False
        For intORow = 1 To .MaxRows
            .Row = intORow
            .Col = colCHARTNO
            If strBarno_Work = Trim(.Text) Then
                blnSame = True
                Exit For
            End If
        Next
        
        If blnSame = False Then
            frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), frmMain.spdOrder.MaxRows, colSPECNO)
            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), frmMain.spdOrder.MaxRows, colCHECKBOX)
            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), frmMain.spdOrder.MaxRows, colHOSPDATE)
            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), frmMain.spdOrder.MaxRows, colBARCODE)
            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), frmMain.spdOrder.MaxRows, colSEQNO)
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
    
    End With
    
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

