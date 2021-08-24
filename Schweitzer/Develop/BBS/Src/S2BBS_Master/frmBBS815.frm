VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS815 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검체보관장소 마스터"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS815.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboCenter 
      Height          =   300
      Left            =   1320
      Style           =   2  '드롭다운 목록
      TabIndex        =   17
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   4080
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   7440
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   5400
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   7440
      Width           =   1260
   End
   Begin VB.TextBox txtRmk 
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   6120
      Width           =   5295
   End
   Begin VB.TextBox txtLegCd 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   285
      Left            =   3120
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "신규등록"
      Top             =   5340
      Width           =   990
   End
   Begin VB.TextBox txtSlotCnt 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7380
      TabIndex        =   7
      Top             =   5700
      Width           =   990
   End
   Begin VB.TextBox txtRowCnt 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   285
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   2
      Top             =   5700
      Width           =   990
   End
   Begin VB.TextBox txtColCnt 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   285
      Left            =   5220
      MaxLength       =   3
      TabIndex        =   3
      Top             =   5700
      Width           =   990
   End
   Begin FPSpread.vaSpread tblSearch 
      Height          =   4095
      Left            =   300
      TabIndex        =   0
      Top             =   660
      Width           =   10215
      _Version        =   196608
      _ExtentX        =   18018
      _ExtentY        =   7223
      _StockProps     =   64
      BackColorStyle  =   1
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
      GridShowVert    =   0   'False
      MaxCols         =   6
      MaxRows         =   10
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS815.frx":076A
      TextTip         =   4
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "센터"
      Height          =   180
      Left            =   840
      TabIndex        =   16
      Top             =   300
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Remark:"
      Height          =   180
      Left            =   2160
      TabIndex        =   15
      Top             =   6180
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Slot:"
      Height          =   180
      Left            =   6420
      TabIndex        =   13
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Rack :"
      Height          =   180
      Left            =   2310
      TabIndex        =   12
      Top             =   5400
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Rows:"
      Height          =   180
      Left            =   2310
      TabIndex        =   11
      Top             =   5760
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Cols:"
      Height          =   180
      Left            =   4260
      TabIndex        =   10
      Top             =   5760
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Total SLOT:"
      Height          =   180
      Left            =   8220
      TabIndex        =   9
      Top             =   300
      Width           =   1020
   End
   Begin VB.Label lblTotSlot 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9300
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   300
      TabIndex        =   14
      Top             =   4860
      Width           =   10215
   End
End
Attribute VB_Name = "frmBBS815"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form명   : frmBBS801
'|  2. 기  능   : 검체보관장소 마스터
'|  4. 작성일   : 2000.11.20
'|
'|  CopyRight(C) 2000 대련엠티에스
'+--------------------------------------------------------------------------------------+
Option Explicit
Private objSql As clsBBSMSTStatement
Private objPop As clsPopupMenu
Private Const MENU_DEL& = 1

Private Sub cboCenter_Click()
    Call Search
    If tblSearch.MaxRows > 0 Then Call TblDisplay(1)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim RS As Recordset
    Dim strTmp As VbMsgBoxResult
    Dim strTmp1 As VbMsgBoxResult
    Dim Centercd As String
    
    If txtLegCd = "" Or txtLegCd = "신규등록" Then
        MsgBox "Leg코드를 넣어 주세요..", vbInformation, Me.Caption
        txtLegCd.SetFocus
        Set RS = Nothing
        Set objSql = Nothing
        Exit Sub
    End If
    If txtColCnt <> "" And IsNumeric(txtColCnt) = False Then
        MsgBox "행수는 숫자만 넣어주세요..", vbInformation, Me.Caption
        txtColCnt.SetFocus
        Exit Sub
    ElseIf txtColCnt = "" Then
        MsgBox "행수를 넣어주세요..", vbInformation, Me.Caption
        txtColCnt.SetFocus
        Exit Sub
    End If
    If txtRowCnt <> "" And IsNumeric(txtRowCnt) = False Then
        MsgBox "열수는 숫자만 넣어주세요..", vbInformation, Me.Caption
        txtRowCnt.SetFocus
        Exit Sub
    ElseIf txtRowCnt = "" Then
        MsgBox "열수를 넣어주세요..", vbInformation, Me.Caption
        txtColCnt.SetFocus
        Exit Sub
    End If
    
    Centercd = medGetP(cboCenter.Text, 1, " ")
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DBConn
    Set RS = objSql.GetBBS003(Centercd, Trim(txtLegCd))
    
    If RS.EOF = False Then
        strTmp1 = MsgBox("수정하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
        If strTmp1 = vbCancel Then
            Clear
            Set RS = Nothing
            Set objSql = Nothing
            Clear
            Exit Sub
        Else '수정
            If objSql.InsertBBS003(Centercd, Trim(txtLegCd), Val(txtRowCnt), Val(txtColCnt), Trim(txtRmk), False) = True Then
                MsgBox "수정하였습니다.", vbInformation, Me.Caption
                Search
            End If
        End If
    Else
    '저장여부 확인...
        strTmp = MsgBox("저장하시겠습니까?", vbInformation + vbOKCancel, Me.Caption)
        If strTmp = vbCancel Then
            Clear
            Set RS = Nothing
            Set objSql = Nothing
            Exit Sub
        Else '저장
            If objSql.InsertBBS003(Centercd, Trim(txtLegCd), Val(txtRowCnt), Val(txtColCnt), Trim(txtRmk), True) = True Then
                MsgBox "저장성공하였습니다.", vbInformation, Me.Caption
                Search
            End If
        End If
    End If
    Set RS = Nothing
    Set objSql = Nothing
    Clear
End Sub

Private Sub Clear()
    txtLegCd.Text = ""
    txtRowCnt.Text = ""
    txtColCnt.Text = ""
    txtSlotCnt.Text = ""
    txtRmk.Text = ""
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_Load()
    Dim RS As Recordset
    Dim objcom003 As clsCom003
    
    
    '센터코드 셋팅
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter)
    Set objcom003 = Nothing
    cboCenter.ListIndex = medComboFind(cboCenter, ObjSysInfo.BuildingCd & Space(1) & ObjSysInfo.BuildingNm)
End Sub

Private Sub TblDisplay(ByVal Row As Long)
    '스프레드 내용을 가져오자..
    With tblSearch
        .Row = Row
        .Col = 1: txtLegCd = .Value
        .Col = 2: txtRowCnt = .Value
        .Col = 3: txtColCnt = .Value
        .Col = 4: txtSlotCnt = .Value
        .Col = 5: txtRmk = .Value
    End With
End Sub

Private Sub tblSearch_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    If NewRow < 0 Then Exit Sub
    Call TblDisplay(NewRow)
End Sub

Private Sub tblSearch_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim RS As Recordset
    Dim strRack As String
    Dim strCenterCd As String
    Dim strSql As String
    
    If ClickType = 1 Then
        strCenterCd = medGetP(cboCenter.Text, 1, " ")
        tblSearch.Row = Row
        tblSearch.Col = 1
        strRack = tblSearch.Text
        
        If strCenterCd = "" Then Exit Sub
        If strRack = "" Then Exit Sub
        
        Set objPop = New clsPopupMenu
        With objPop
            .AddMenu MENU_DEL, "보관장소 삭제"
            
            .PopupMenus Me.hWnd
            
            If .MenuID = MENU_DEL Then
                If MsgBox("보관장소를 삭제하시겠습니까?", vbYesNo) = vbYes Then
                    'BBS003 보관장소 마스터와 BBS206 검체보관 내역에 없는 경우에만 삭제 가능하도록
                    Set RS = New Recordset
                    
                    strSql = " select * from " & T_BBS206
                    strSql = strSql & " where " & DBW("centercd=", strCenterCd)
                    strSql = strSql & " and " & DBW("legcd=", strRack)
                    strSql = strSql & " and (spcyy is not null )"
                    strSql = strSql & " and (spcno is not null )"
                    
                    RS.Open strSql, DBConn
                    
                    If RS.EOF = False Then
                        MsgBox "이미 해당 보관장소에 보관중인 검체가 있습니다." & vbNewLine & _
                               "보관장소를 삭제하려면 검체를 먼저 폐기하거나 다른 보관장소에 이동 후에만 삭제가 가능합니다.", vbCritical
                        Set RS = Nothing
                        Set objPop = Nothing
                        Exit Sub
                    End If
                    
                    Set RS = Nothing
                    
                    On Error GoTo ExecuteErr
                    DBConn.BeginTrans
                    
                    'BBS003 삭제
                    strSql = " delete " & T_BBS003 & _
                            " where " & DBW("centercd=", strCenterCd) & _
                            " and " & DBW("legcd=", strRack)
                    DBConn.Execute strSql
                        
                    'BBS206 삭제
                    strSql = " delete " & T_BBS206 & _
                            " where " & DBW("centercd=", strCenterCd) & _
                            " and " & DBW("legcd=", strRack)
                    DBConn.Execute strSql
                    
                    DBConn.CommitTrans
                    
                    '스프레드 로우삭제
                    tblSearch.Row = Row
                    tblSearch.Action = ActionDeleteRow
                    
                    MsgBox "정상적으로 처리되었습니다.", vbExclamation
                    GoTo Skip
ExecuteErr:
                    DBConn.RollbackTrans
                    MsgBox "처리도중 오류가 발생하였습니다.", vbExclamation
Skip:
                End If
            End If
        End With
        
        Set objPop = Nothing
    End If
End Sub

Private Sub txtLegCd_GotFocus()
    If txtLegCd = "신규등록" Then
        txtLegCd = ""
    End If
    With txtLegCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtColCnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtLegCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtLegCd_LostFocus()
    Dim RS As Recordset
    
    If txtLegCd = "" Then
        txtRowCnt = ""
        txtColCnt = ""
        txtSlotCnt = ""
        txtRmk = ""
        Exit Sub
    End If
    
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DBConn
    Set RS = objSql.GetBBS003(Trim(txtLegCd))
    If RS.EOF = True Then
        Set RS = Nothing
        Set objSql = Nothing
        txtRowCnt = ""
        txtColCnt = ""
        txtSlotCnt = ""
        txtRmk = ""
        Exit Sub
    Else
        txtRowCnt = RS.Fields("rowcnt").Value & ""
        txtColCnt = RS.Fields("colcnt").Value & ""
        txtSlotCnt = RS.Fields("rowcnt").Value & "" * RS.Fields("colcnt").Value & ""
        txtRmk = RS.Fields("rmk").Value & ""
        Set RS = Nothing
        Set objSql = Nothing
    End If
End Sub

Private Sub txtRowCnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRowCnt_LostFocus()
    If txtColCnt = "" Then
        Exit Sub
    Else
        txtSlotCnt.Text = Val(txtRowCnt) * Val(txtColCnt)
    End If
End Sub

Private Sub txtColCnt_LostFocus()
    If txtRowCnt = "" Then
        Exit Sub
    Else
        txtSlotCnt.Text = Val(txtRowCnt) * Val(txtColCnt)
    End If
End Sub
Private Sub Search()
    Dim RS As Recordset
    Dim totslot As Long
    
    tblSearch.MaxRows = 0
    
    Set objSql = New clsBBSMSTStatement
'    objSql.setDbConn DBConn
    Set RS = objSql.GetBBS003(medGetP(cboCenter.Text, 1, " "))
    lblTotSlot = ""
    If RS.EOF = True And RS.BOF = True Then
        Set RS = Nothing
        Set objSql = Nothing
        Exit Sub
    Else '스프레드에 내용을 뿌리자...
        totslot = 0
        With tblSearch
            .MaxRows = 0
            Do Until RS.EOF = True
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1: .Value = Trim(RS.Fields("legcd"))
                .Col = 2: .Value = Trim(RS.Fields("rowcnt"))
                .Col = 3: .Value = Trim(RS.Fields("colcnt"))
                .Col = 4: .Value = Trim(RS.Fields("rowcnt")) * Trim(RS.Fields("colcnt"))
                .Col = 5: .Value = Trim(RS.Fields("rmk")) & ""
                .Col = 6: .Value = Trim(RS.Fields("centercd"))
                
                totslot = totslot + Trim(RS.Fields("rowcnt")) * Trim(RS.Fields("colcnt"))
                
                RS.MoveNext
            Loop
        End With
        lblTotSlot = totslot
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub SetCenter()
    Dim RS As Recordset
    
    Set RS = ReadCom003(BC2_CENTER)
    cboCenter.Clear
    If RS Is Nothing Then Exit Sub
    With RS
        cboCenter.AddItem .Fields("cdval1") & "-" & .Fields("field1")
        .MoveNext
    End With
    
    Set RS = Nothing
End Sub





