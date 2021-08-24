VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmFlagCfg 
   BorderStyle     =   1  '단일 고정
   Caption         =   "FLAG 설정"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   Icon            =   "frmFlagCfg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6480
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdExit 
      Caption         =   "닫 기"
      Height          =   420
      Left            =   5010
      TabIndex        =   2
      Top             =   5895
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저 장"
      Height          =   420
      Left            =   3840
      TabIndex        =   1
      Top             =   5895
      Width           =   1125
   End
   Begin FPSpread.vaSpread spdFlag 
      Height          =   5550
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   6075
      _Version        =   393216
      _ExtentX        =   10716
      _ExtentY        =   9790
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmFlagCfg.frx":0E42
      TextTip         =   1
   End
   Begin VB.Menu mnuPopupA00 
      Caption         =   "PopupMenuA00"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupA 
         Caption         =   "전체선택"
         Index           =   1
      End
      Begin VB.Menu mnuPopupA 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopupA 
         Caption         =   "전체취소"
         Index           =   3
      End
   End
   Begin VB.Menu mnuPopupB00 
      Caption         =   "PopupMenuB00"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupB 
         Caption         =   "추가"
         Index           =   1
      End
      Begin VB.Menu mnuPopupB 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopupB 
         Caption         =   "삭제"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmFlagCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrRtn
    
    Dim objDB   As Object
    Dim vTmp
    Dim ii%, iSeq%
    Dim sRet$
    Dim sTSeq$, sTFlagCd$, sTFlagInfo$, sTDispCd$, sTUseYn$, sTRmk$
    
    sTSeq = "": sTFlagCd = "": sTFlagInfo = "": sTDispCd = "": sTUseYn = "": sTRmk = ""
    
    With spdFlag
        iSeq = 0
        For ii = 1 To .MaxRows
            Call .GetText(3, ii, vTmp)
            If Trim(vTmp) <> "" Then
                iSeq = iSeq + 1
            
                sTSeq = sTSeq & Format(iSeq, "000") & Chr(124)
                
                .Col = 2: .Row = ii
                If .Value = vbChecked Then
                    sTUseYn = sTUseYn & "Y" & Chr(124)
                Else
                    sTUseYn = sTUseYn & Chr(124)
                End If
                
                Call .GetText(3, ii, vTmp): sTFlagCd = sTFlagCd & Trim(vTmp) & Chr(124)
                Call .GetText(4, ii, vTmp): sTFlagInfo = sTFlagInfo & Trim(vTmp) & Chr(124)
                Call .GetText(5, ii, vTmp): sTDispCd = sTDispCd & Trim(vTmp) & Chr(124)
                Call .GetText(6, ii, vTmp): sTRmk = sTRmk & Trim(vTmp) & Chr(124)
            End If
        Next ii
    End With
    
    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    sRet = objDB.Add_IFFlagInfo(gsMachineCd, sTSeq, sTFlagCd, sTFlagInfo, sTDispCd, sTUseYn, sTRmk)
    
    Set objDB = Nothing
    
    If IsNumeric(sRet) = True Then
        MsgBox "등록시 오류발생...", vbCritical, Me.Caption
    Else
        MsgBox "정상적으로 등록되었습니다.", vbInformation, Me.Caption
        Unload Me
    End If
    
ErrRtn:
    If Err <> 0 Then
        Set objDB = Nothing
        MsgBox Err.Description, vbCritical, Me.Caption
    End If
End Sub


Private Sub Form_Load()

    spdFlag.MaxRows = 0

    Call DispIFFlagInfo
    
End Sub
Private Sub DispIFFlagInfo()
    On Error GoTo ErrRtn
    
    Dim ii%, iItemCnt%
    Dim sRetVal3$
    Dim objDB   As Object
    
    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    'flag info
    sRetVal3 = objDB.Get_IFflaginfo(gsMachineCd)
    
    If sRetVal3 <> "NONE" Then
        iItemCnt = GetByOneUserSymbol(sRetVal3, sRetVal3, Chr$(3))
        Call MakeIFFlagStruct(sRetVal3, iItemCnt)
    End If
    
    Set objDB = Nothing
    
    With spdFlag
        For ii = 1 To iItemCnt
            If Trim(gIFFlagInfo(ii).SEQ) = "" Then Exit For
            
            .MaxRows = .MaxRows + 1
            .RowHeight(.MaxRows) = 11
        
            Call .SetText(1, .MaxRows, Trim(gIFFlagInfo(ii).SEQ))
            If Trim(gIFFlagInfo(ii).USEYN) = "Y" Then
                Call .SetText(2, .MaxRows, vbChecked)
            End If
            Call .SetText(3, .MaxRows, Trim(gIFFlagInfo(ii).FLAGCD))
            Call .SetText(4, .MaxRows, Trim(gIFFlagInfo(ii).FLAGINFO))
            Call .SetText(5, .MaxRows, Trim(gIFFlagInfo(ii).DISPCD))
            Call .SetText(6, .MaxRows, Trim(gIFFlagInfo(ii).REMARK))
        Next ii
    End With
    
ErrRtn:
    If Err <> 0 Then
        Set objDB = Nothing
        MsgBox "DispIFFlagInfo Err - " & Err.Description, vbExclamation, Me.Caption
    End If
End Sub




Private Sub mnuPopupA_Click(Index As Integer)
    
    Dim ii%
    
    With spdFlag
        If Index = 2 Then Exit Sub
        
        For ii = 1 To .MaxRows
            If Index = 1 Then
                Call .SetText(2, ii, vbChecked)
            Else
                Call .SetText(2, ii, vbUnchecked)
            End If
        Next ii
    End With
    
End Sub

Private Sub mnuPopupB_Click(Index As Integer)
    
    Dim vTmp
    
    With spdFlag
        Select Case Index
            Case 1
                .MaxRows = .MaxRows + 1
                .RowHeight(.MaxRows) = 11
            
                If .MaxRows > 20 Then
                    .TopRow = .MaxRows - 19
                End If
            
            Case 3
                Call .GetText(3, .ActiveRow, vTmp)
                If Trim(vTmp) <> "" Then
                    If MsgBox("해당 항목을 화면에서 삭제하시겠습니까?" & vbCrLf _
                            & Trim(vTmp) & " [Row:" & Trim(.ActiveRow) & "]", vbQuestion + vbYesNo, "삭제확인") <> vbYes Then
                        Exit Sub
                    End If
                End If
                .Row = .ActiveRow
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
        End Select
    End With
    
End Sub


Private Sub spdFlag_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Dim sTmp$
    Dim vTmp
    
    If Row < 1 Or Col <> 3 Then Exit Sub
    
    If MsgBox("장비 FLAG를 입력 또는 수정하시겠습니까?", vbQuestion + vbYesNo, "FLAG 입력확인") <> vbYes Then Exit Sub
    
    With spdFlag
        Call .GetText(3, Row, vTmp)
        
        sTmp = InputBox("추가(수정)을 원하는 FLAG를 입력해 주십시요.", "FLAG", Trim(vTmp))
        
        If Trim(sTmp) <> "" Then
            Call .SetText(3, Row, sTmp)
        End If
    End With
    
End Sub

    
Private Sub spdFlag_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Col = 2 Then
        Call PopupMenu(mnuPopupA00)
    Else
        Call PopupMenu(mnuPopupB00)
    End If
    
End Sub
