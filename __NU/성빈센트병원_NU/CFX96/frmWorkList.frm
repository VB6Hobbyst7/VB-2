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
   ClientWidth     =   13110
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtRow 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   26
      Text            =   "12"
      Top             =   1350
      Width           =   465
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   1770
      Width           =   435
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   315
      Left            =   870
      TabIndex        =   13
      Top             =   1770
      Width           =   195
   End
   Begin VB.TextBox txtQuery 
      Height          =   645
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmWorkList.frx":000C
      Top             =   9810
      Visible         =   0   'False
      Width           =   12195
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   0
      ScaleHeight     =   1605
      ScaleWidth      =   13110
      TabIndex        =   0
      Top             =   0
      Width           =   13110
      Begin VB.CommandButton cmdPrint 
         Caption         =   "인쇄"
         Height          =   375
         Left            =   10170
         TabIndex        =   24
         Top             =   420
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   7350
         TabIndex        =   19
         Top             =   1020
         Width           =   3015
         Begin VB.OptionButton optCheck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "예비결과"
            Height          =   195
            Index           =   2
            Left            =   1890
            TabIndex        =   23
            Top             =   210
            Width           =   1065
         End
         Begin VB.OptionButton optCheck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "접수"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   22
            Top             =   210
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton optCheck 
            BackColor       =   &H00FFFFFF&
            Caption         =   "전체"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   21
            Top             =   210
            Width           =   675
         End
      End
      Begin VB.TextBox txtBarNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   10710
         TabIndex        =   16
         Text            =   "O55RT05B0"
         Top             =   1260
         Width           =   1875
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "장비전송"
         Height          =   375
         Left            =   8880
         TabIndex        =   15
         Top             =   420
         Width           =   1275
      End
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
         TabIndex        =   11
         Text            =   "1"
         Top             =   450
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdSendClose 
         Caption         =   "전송후 닫기"
         Height          =   375
         Left            =   10860
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         Height          =   375
         Left            =   11280
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "화면전송"
         Height          =   375
         Left            =   7770
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         Height          =   375
         Left            =   6660
         TabIndex        =   5
         Top             =   420
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   2850
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
         Format          =   126681089
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   4740
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
         Format          =   126681089
         CurrentDate     =   40457
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회구분"
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
         Left            =   6360
         TabIndex        =   20
         Top             =   1260
         Width           =   780
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
         Left            =   120
         TabIndex        =   18
         Top             =   1260
         Width           =   5895
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "바코드번호"
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
         Left            =   11190
         TabIndex        =   17
         Top             =   990
         Width           =   975
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
         TabIndex        =   14
         Top             =   150
         Visible         =   0   'False
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
         TabIndex        =   12
         Top             =   510
         Visible         =   0   'False
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
         Left            =   4560
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
         Left            =   1890
         TabIndex        =   2
         Top             =   540
         Width           =   780
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   1680
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
      Height          =   7875
      Left            =   330
      TabIndex        =   9
      Top             =   1740
      Width           =   12165
      _Version        =   393216
      _ExtentX        =   21458
      _ExtentY        =   13891
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
      SpreadDesigner  =   "frmWorkList.frx":1B3F
      UserResize      =   2
      ScrollBarTrack  =   1
      ShowScrollTips  =   3
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

Private Sub cmdIn_Click()
    Dim intRow As Integer
    
    intRow = txtRow.Text
    
    With spdWork
        .MaxRows = .MaxRows + 1
        'Call .InsertRows(1, 1)
        Call .InsertRows(intRow, 1)
        Call .SetText(colCHECKBOX, intRow, "1")
        
        Call .SetText(colCHARTNO, intRow, CStr(.MaxRows))
        Call .SetText(colBARCODE, intRow, CStr(.MaxRows))
        Call .SetText(colPNAME, intRow, "테스트")
        
        
    End With

End Sub

Private Sub cmdOrder_Click()
    Dim lngFIleNum  As Long
    Dim strCFXFile  As String
    
    Dim strBarno    As String
    Dim strPNM      As String
    Dim strPID      As String
    Dim iCnt        As Integer
    Dim varTmp      As Variant
    Dim ORDERPATH   As String
    Dim i           As Integer
    Dim J, K, M     As Integer
    
    With frmMain.CFXFile
        .CancelError = True
        .FileName = gComm.ORDPATH & "Nimbus.lis"
        If Len(Dir(.FileName)) Then
             Close #lngFIleNum
             Kill .FileName
        End If
        lngFIleNum = FreeFile
        
        Open .FileName For Append As #lngFIleNum

        strCFXFile = ""
        J = 1
        K = 1
        M = 1
        
        For iCnt = 1 To spdWork.MaxRows + 3
            If iCnt = 48 Then
                Exit For
            End If
            
            spdWork.GetText 1, iCnt, varTmp
            If GetText(spdWork, iCnt, colCHECKBOX) = "1" Then
                strBarno = GetText(spdWork, iCnt, colBARCODE)
                strPNM = GetText(spdWork, iCnt, colPNAME)
                strPID = GetText(spdWork, iCnt, colCHARTNO)
                
                If iCnt = 1 Then
                    strCFXFile = strCFXFile & "Row,Column,*Target Name,*Sample Name,Sample No,Patient Id" & vbNewLine
                End If
                strCFXFile = strCFXFile & Chr(64 + J) & "," & K & ",," & strPNM & "," & strPID & "" & vbNewLine
                strCFXFile = strCFXFile & Chr(64 + J) & "," & K + 6 & ",," & strPNM & "," & strPID & "" & vbNewLine
                J = J + 1
                If J = 9 Then
                    J = 1
                    K = K + 1
                    If K = 9 Then
                        K = 1
                    End If
                End If
                    
                Call SetText(spdWork, "", iCnt, colCHECKBOX)
            End If
            
            If iCnt >= spdWork.MaxRows Then
                Select Case M
                    Case 1, 5: strPNM = "NC"
                    Case 2, 6: strPNM = "PC1"
                    Case 3, 7: strPNM = "PC2"
                    Case 4, 8: strPNM = "PC3"
                End Select
                M = M + 1
                
                strCFXFile = strCFXFile & Chr(64 + J) & "," & K & ",," & strPNM & "," & "," & "" & vbNewLine
                strCFXFile = strCFXFile & Chr(64 + J) & "," & K + 6 & ",," & strPNM & "," & "," & "" & vbNewLine
                J = J + 1
                If J = 9 Then
                    J = 1
                    K = K + 1
                    If K = 9 Then
                        K = 1
                    End If
                End If
            End If
        Next
        
'        strCFXFile = strCFXFile & "A,6,,NC" & vbNewLine
'        strCFXFile = strCFXFile & "F,6,,PC1" & vbNewLine
'        strCFXFile = strCFXFile & "G,6,,PC2" & vbNewLine
'        strCFXFile = strCFXFile & "H,6,,PC3" & vbNewLine
'        strCFXFile = strCFXFile & "A,12,,NC" & vbNewLine
'        strCFXFile = strCFXFile & "F,12,,PC1" & vbNewLine
'        strCFXFile = strCFXFile & "G,12,,PC2" & vbNewLine
'        strCFXFile = strCFXFile & "H,12,,PC3" & vbNewLine
        
       
        If strCFXFile <> "" Then
            strCFXFile = Mid(strCFXFile, 1, Len(strCFXFile) - 2)
            Print #lngFIleNum, strCFXFile
            MsgBox "오더 파일 생성 완료", vbOKOnly + vbInformation, Me.Caption
        End If
        strCFXFile = ""
        Close #lngFIleNum
        
    End With
End Sub

Private Sub cmdPrint_Click()
    
    spdWork.PrintOrientation = PrintOrientationPortrait 'PrintOrientationLandscape '가로출력
    spdWork.Action = 13
    
End Sub

Private Sub cmdSearch_Click()
    Dim strState As String
    
    If optCheck(0).Value = True Then
        strState = "0,2"
    ElseIf optCheck(1).Value = True Then
        strState = "0"
    ElseIf optCheck(2).Value = True Then
        strState = "2"
    End If
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), strState)
    
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
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colRACKNO), frmMain.spdOrder.MaxRows, colRACKNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPOSNO), frmMain.spdOrder.MaxRows, colPOSNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), frmMain.spdOrder.MaxRows, colCHARTNO)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), frmMain.spdOrder.MaxRows, colPID)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), frmMain.spdOrder.MaxRows, colINOUT)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), frmMain.spdOrder.MaxRows, colPNAME)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), frmMain.spdOrder.MaxRows, colPSEX)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), frmMain.spdOrder.MaxRows, colPAGE)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), frmMain.spdOrder.MaxRows, colPJUMIN)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colKEY1), frmMain.spdOrder.MaxRows, colKEY1)
                    Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colKEY2), frmMain.spdOrder.MaxRows, colKEY2)
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
    
    txtBarNo.Text = ""
    
    txtRow.Text = 1
    
    'spdWork.MaxRows = 10
    
End Sub


Private Sub lblQuery_DblClick()
    If txtQuery.Visible = True Then
        txtQuery.Visible = False
    Else
        txtQuery.Visible = True
    End If
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

Private Sub txtBarNo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call GetSampleInfo(0, frmWorkList.spdWork, txtBarNo.Text)
    End If

End Sub
