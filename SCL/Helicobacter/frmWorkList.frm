VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "워크리스트"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   16080
   Icon            =   "frmWorkList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   16080
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraHidden 
      Caption         =   "Frame1"
      Height          =   5865
      Left            =   6510
      TabIndex        =   15
      Top             =   2070
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdSeq 
         Caption         =   "Seq 매치"
         Height          =   375
         Left            =   270
         TabIndex        =   28
         Top             =   3750
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox chkNC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NC"
         Height          =   285
         Left            =   1620
         TabIndex        =   27
         Top             =   3360
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chkPC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PC"
         Height          =   285
         Left            =   570
         TabIndex        =   26
         Top             =   3360
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtNCCnt 
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
         Height          =   300
         Left            =   2280
         TabIndex        =   25
         Text            =   "1"
         Top             =   3360
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtPCCnt 
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
         Height          =   300
         Left            =   1230
         TabIndex        =   24
         Text            =   "1"
         Top             =   3360
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtSeqNo 
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
         Left            =   1230
         TabIndex        =   22
         Text            =   "1"
         Top             =   2280
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '평면
         BackColor       =   &H00800000&
         Caption         =   "저장포함"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   2040
         TabIndex        =   21
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optResult 
         Appearance      =   0  '평면
         BackColor       =   &H00A5704B&
         Caption         =   "전체"
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
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   20
         Top             =   1230
         Width           =   765
      End
      Begin VB.OptionButton optResult 
         Appearance      =   0  '평면
         BackColor       =   &H00A5704B&
         Caption         =   "결과있음"
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
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   19
         Top             =   1230
         Width           =   1125
      End
      Begin VB.OptionButton optResult 
         Appearance      =   0  '평면
         BackColor       =   &H00A5704B&
         Caption         =   "결과없음"
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
         Height          =   255
         Index           =   2
         Left            =   2340
         TabIndex        =   18
         Top             =   1230
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.CommandButton cmdSendClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "전송/닫기"
         Height          =   375
         Left            =   1380
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   510
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "화면전송"
         Height          =   375
         Left            =   180
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   540
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSeqNo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   450
         TabIndex        =   23
         Top             =   2340
         Visible         =   0   'False
         Width           =   510
      End
   End
   Begin MSComDlg.CommonDialog CFXFile 
      Left            =   13740
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00A5704B&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   16080
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      Begin VB.TextBox txtToWN 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5850
         MaxLength       =   5
         TabIndex        =   30
         Top             =   210
         Width           =   645
      End
      Begin VB.TextBox txtFromWN 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4980
         MaxLength       =   5
         TabIndex        =   29
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton cmdRemove 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
         Caption         =   "-Test"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13320
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdTest 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Caption         =   "+Test"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12630
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   9240
         Style           =   1  '그래픽
         TabIndex        =   12
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdPC 
         BackColor       =   &H00C0C0FF&
         Caption         =   "+ PC"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14730
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdNC 
         Appearance      =   0  '평면
         BackColor       =   &H00FFC0C0&
         Caption         =   "+ NC"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14010
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdOrder 
         BackColor       =   &H00FFFFFF&
         Caption         =   "오더전송"
         Height          =   375
         Left            =   8250
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdWorkPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "워크출력"
         Height          =   375
         Left            =   10230
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "워크조회"
         Height          =   375
         Left            =   7260
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "닫기"
         Height          =   375
         Left            =   11220
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   180
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   148307969
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   3030
         TabIndex        =   2
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   148307969
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "W/N"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   4620
         TabIndex        =   32
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   0
         Left            =   5670
         TabIndex        =   31
         Top             =   270
         Width           =   165
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   495
         Left            =   6990
         Top             =   90
         Width           =   9345
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회기간 :"
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
         Left            =   360
         TabIndex        =   4
         Top             =   270
         Width           =   930
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   465
         Left            =   270
         Top             =   120
         Width           =   6375
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
         Left            =   2820
         TabIndex        =   3
         Top             =   270
         Width           =   150
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   8835
      Left            =   30
      TabIndex        =   5
      Top             =   750
      Width           =   16395
      _Version        =   393216
      _ExtentX        =   28919
      _ExtentY        =   15584
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      ColsFrozen      =   22
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   23
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmWorkList.frx":06C2
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intFrom As Integer
Private intTo   As Integer


Private Sub cmdClear_Click()
    spdWork.MaxRows = 0
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Public Sub cmdNC_Click()
        
    spdWork.MaxRows = spdWork.MaxRows + 1
    Call SetText(spdWork, "1", spdWork.MaxRows, colCHECKBOX)
    Call SetText(spdWork, "NC", spdWork.MaxRows, colPNAME)

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
    Dim J, k, M     As Integer
    Dim l           As Integer
    
    With CFXFile
        .CancelError = True
        '.Filename = "C:\Users\SG CFX96-IVD\Desktop\INTERFACE\IMPORT\" & "Nimbus.lis"
        .Filename = gHOSP.IMPPATH & "Nimbus.lis"
        
        If Len(Dir(.Filename)) Then
             Close #lngFIleNum
             Kill .Filename
        End If
        lngFIleNum = FreeFile
        
        Open .Filename For Append As #lngFIleNum

        strCFXFile = ""
        J = 1
        k = 1
        M = 1
        l = 0
        
        For iCnt = 1 To spdWork.MaxRows '+ 4
            If iCnt = 48 Then
                Exit For
            End If
            
            spdWork.GetText 1, iCnt, varTmp
            If GetText(spdWork, iCnt, colCHECKBOX) = "1" Then
                strBarno = GetText(spdWork, iCnt, colBARCODE)
                strPNM = GetText(spdWork, iCnt, colPNAME)
                strPID = GetText(spdWork, iCnt, colPID)
                
                If iCnt = 1 Then
                    strCFXFile = strCFXFile & "Row,Column,*Target Name,*Sample Name,Sample No,Patient Id" & vbNewLine
                End If
                strCFXFile = strCFXFile & Chr(64 + J) & "," & k & ",," & strPNM & "," & strBarno & "," & strPID & "" & vbNewLine
                'strCFXFile = strCFXFile & Chr(64 + J) & "," & k + 6 & ",," & strPNM & "," & strPID & "" & vbNewLine
                
                Call SetText(spdWork, Chr(64 + J), iCnt, colRACKNO)
                Call SetText(spdWork, k, iCnt, colPOSNO)
            
                J = J + 1
                If J = 9 Then
                    J = 1
                    k = k + 1
                    If k = 9 Then
                        k = 1
                    End If
                End If
                    
                Call SetText(spdWork, "", iCnt, colCHECKBOX)
            End If
            
'            If iCnt >= spdWork.MaxRows Then
'                M = M + 1
'                strPNM = "PC"
'                strCFXFile = strCFXFile & Chr(64 + J) & "," & k & ",," & strPNM & "," & "," & "" & vbNewLine
'                J = J + 1
'                If J = 9 Then
'                    J = 1
'                    k = k + 1
'                    If k = 9 Then
'                        k = 1
'                    End If
'                End If
'                strPNM = "NC"
'                strCFXFile = strCFXFile & Chr(64 + J) & "," & k & ",," & strPNM & "," & "," & "" & vbNewLine
'                J = J + 1
'                If J = 9 Then
'                    J = 1
'                    k = k + 1
'                    If k = 9 Then
'                        k = 1
'                    End If
'                End If
'            End If
        Next
        
        If strCFXFile <> "" Then
            strCFXFile = Mid(strCFXFile, 1, Len(strCFXFile) - 2)
            Print #lngFIleNum, strCFXFile
            MsgBox "오더 파일 생성 완료", vbOKOnly + vbInformation, Me.Caption
        End If
        
        strCFXFile = ""
        Close #lngFIleNum
        
    End With
End Sub

Public Sub cmdPC_Click()
    
    spdWork.MaxRows = spdWork.MaxRows + 1
    Call SetText(spdWork, "1", spdWork.MaxRows, colCHECKBOX)
    Call SetText(spdWork, "PC", spdWork.MaxRows, colPNAME)

End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    
    With spdWork
        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = colCHECKBOX
            If .Value = "1" Then
                Call DeleteRow(spdWork, i, i)
                .MaxRows = .MaxRows - 1
            End If
        Next
    End With

End Sub

Private Sub cmdSearch_Click()
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)
    
    spdWork.RowHeight(-1) = 15

End Sub

Private Sub cmdSend_Click()
    Dim i               As Integer
    Dim intRow          As Integer
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
                    intRow = frmMain.spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                
                        varItems = GetText(spdWork, intWRow, colITEMS)
                        varItems = Split(varItems, "/")
                        For intItems = 0 To UBound(varItems)
                            For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                                frmMain.spdOrder.Row = 0
                                frmMain.spdOrder.Col = intOCol
                                If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                                    .Row = frmMain.spdOrder.MaxRows
                                    Call SetText(frmMain.spdOrder, "◇", frmMain.spdOrder.MaxRows, intOCol)
'                                    GoTo RST
                                End If
                            Next
                        Next
                    Next
                    
                    frmMain.spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
    End With
    
End Sub

Private Sub cmdSendClose_Click()
    
    Call cmdSend_Click
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdSeq_Click()
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim strSeq          As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                For intORow = 1 To frmMain.spdOrder.MaxRows
                    If GetText(spdWork, intWRow, colSEQNO) = GetText(frmMain.spdOrder, intORow, colSEQNO) Then
                        
                        Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), intORow, colBARCODE)
                        DoEvents
                        If GetSampleInfo(intORow, frmMain.spdOrder) = -1 Then
                            'MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                        Else
                            '정보수정
                            SQL = ""
                            SQL = SQL & "UPDATE PATRESULT SET "
                            SQL = SQL & "  BARCODE       = '" & Trim(GetText(frmMain.spdOrder, intORow, colBARCODE)) & "'" & vbCr
                            SQL = SQL & " ,INOUT         = '" & Trim(GetText(frmMain.spdOrder, intORow, colINOUT)) & "'" & vbCr
                            SQL = SQL & " ,CHARTNO       = '" & Trim(GetText(frmMain.spdOrder, intORow, colCHARTNO)) & "'" & vbCr
                            SQL = SQL & " ,PID           = '" & Trim(GetText(frmMain.spdOrder, intORow, colPID)) & "'" & vbCr
                            SQL = SQL & " ,PNAME         = '" & Trim(GetText(frmMain.spdOrder, intORow, colPNAME)) & "'" & vbCr
                            SQL = SQL & " ,PSEX          = '" & Trim(GetText(frmMain.spdOrder, intORow, colPSEX)) & "'" & vbCr
                            SQL = SQL & " ,PAGE          = '" & Trim(GetText(frmMain.spdOrder, intORow, colPAGE)) & "'" & vbCr
''                            SQL = SQL & " ,PJUMIN        = '" & Trim(GetText(frmMain.spdOrder, intORow, colPJUMIN)) & "'" & vbCr
'                            SQL = SQL & " ,PANICVALUE    = '" & Trim(GetText(frmMain.spdOrder, intORow, colKEY1)) & "'" & vbCr
                            SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(frmMain.spdOrder, intORow, colEXAMDATE)) & "'" & vbCr
                            SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(frmMain.spdOrder, intORow, colSAVESEQ)) & vbCr
                            SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "' & vbCr"
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        Exit For
                    End If
                Next intORow
            End If
        Next intWRow
    End With
End Sub

Private Sub cmdTest_Click()
    
    spdWork.MaxRows = spdWork.MaxRows + 1
    Call SetText(spdWork, "1", spdWork.MaxRows, colCHECKBOX)
    Call SetText(spdWork, CStr(spdWork.MaxRows), spdWork.MaxRows, colBARCODE)
    Call SetText(spdWork, "TEST" & CStr(spdWork.MaxRows), spdWork.MaxRows, colPNAME)
    
End Sub

Private Sub cmdWorkPrint_Click()
    
    If spdWork.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        spdWork.PrintOrientation = PrintOrientationLandscape 'PrintOrientationPortrait
        spdWork.Action = 13
    End If
    

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    
    Call CtlInitializing

    '-- 컬럼보이기설정
    Call SetColumnView(spdWork)
    
    spdWork.ColWidth(spdWork.MaxCols) = 20
        
'    spdWork.MaxRows = 10
    
    
'    Dim i As Integer
'
'    For i = 1 To 10
'        Call SetText(spdWork, i, i, colBARCODE)
'        Call SetText(spdWork, i * 10, i, colITEMS)
'
'    Next
    '-- 검사명 보이기
'    Call SetExamCode(spdWork)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub CtlInitializing()
    
    spdWork.MaxRows = 0
    
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    txtSeqNo.Text = "1"
    
    txtFromWN = "1"
    txtToWN = "99999"
    
    '순번사용
    If gHOSP.RSTTYPE = "1" Then
        lblSeqNo.Visible = True
        txtSeqNo.Visible = True
    Else
        lblSeqNo.Visible = False
        txtSeqNo.Visible = False
    End If
    
End Sub

Private Sub Form_Resize()
    
    If Me.ScaleHeight = 0 Then Exit Sub

    spdWork.WIDTH = Me.ScaleWidth - 300
    spdWork.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300

    spdWork.ColWidth(colSTATE + 1) = 60 '(spdWork.Width / 40) * intColSum

End Sub

Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer

    If Row = 0 And Col <> colCHECKBOX Then
        Call SetSpreadSort(spdWork, 0)
        Exit Sub
    End If
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdWork, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "1", i, colCHECKBOX)
            Next
        End If
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
    End If
    
'    txtQuery.Visible = True
'    txtQuery.Text = GetText(spdWork, Row, colITEMS)
    
End Sub

'Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim i               As Integer
'    Dim intRow          As Integer
'    Dim intWRow         As Integer
'    Dim intORow         As Integer
'    Dim intWCol         As Integer
'    Dim intOCol         As Integer
'    Dim strBarno        As String
'    Dim blnSame         As Boolean
'    Dim varItems        As Variant
'    Dim intItems        As Integer
'
'    Dim strBarno_Work   As String
'
'    If Row = 0 Then Exit Sub
'    If Col <> colBARCODE Then
'        Exit Sub
'    End If
'
'    intWRow = Row
'    spdWork.Row = Row
'    spdWork.Col = colBARCODE
'    strBarno_Work = Trim(spdWork.Text)
'
'    With frmMain.spdOrder
'        blnSame = False
'        For intORow = 1 To .MaxRows
'            .Row = intORow
'            .Col = colBARCODE
'            If strBarno_Work = Trim(.Text) Then
'                blnSame = True
'                Exit For
'            End If
'        Next
'
'        If blnSame = False Then
'            frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
'            intRow = frmMain.spdOrder.MaxRows
'
'            For i = colCHECKBOX To colSTATE
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, i), intRow, i)
'
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), intRow, colSPECNO)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), intRow, colCHECKBOX)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), intRow, colHOSPDATE)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), intRow, colBARCODE)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), intRow, colSEQNO)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), intRow, colCHARTNO)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), intRow, colPID)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), intRow, colINOUT)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), intRow, colPNAME)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), intRow, colPSEX)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), intRow, colPAGE)
''    '            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), introw, colPJUMIN)
''                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), intRow, colOCNT)
'
'                varItems = GetText(spdWork, intWRow, colITEMS)
'                varItems = Split(varItems, "/")
'                For intItems = 0 To UBound(varItems)
'                    For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
'                        frmMain.spdOrder.Row = 0
'                        frmMain.spdOrder.Col = intOCol
'                        If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
'                            .Row = intRow
'                            Call SetText(frmMain.spdOrder, "◇", intRow, intOCol)
'                        End If
'                    Next
'                Next
'            Next
'
'            frmMain.spdOrder.RowHeight(-1) = 15
'        End If
'
'    End With
'
'End Sub

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

Private Sub spdWork_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    intFrom = NewRow

End Sub

Private Sub spdWork_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Integer

    With spdWork
        For i = 1 To .MaxRows
            Call SetText(spdWork, "0", i, colCHECKBOX)
        Next
        
        For i = .SelBlockRow To .SelBlockRow2
            Call SetText(spdWork, "1", i, colCHECKBOX)
        Next
    End With


End Sub
