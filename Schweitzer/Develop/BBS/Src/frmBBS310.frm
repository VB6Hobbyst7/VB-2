VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS310 
   BackColor       =   &H00DBE6E6&
   Caption         =   "�����������"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   12585
   WindowState     =   2  '�ִ�ȭ
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  ���׹�ȣ"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1275
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   9930
      Begin VB.TextBox txtBldNo 
         Height          =   330
         Left            =   1695
         TabIndex        =   0
         Top             =   705
         Width           =   2595
      End
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ��Է�"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   705
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "���׹�ȣ"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '�׷���
      TabIndex        =   3
      Tag             =   "128"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   9555
      Style           =   1  '�׷���
      TabIndex        =   2
      Tag             =   "124"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�������(&S)"
      Height          =   510
      Left            =   8235
      Style           =   1  '�׷���
      TabIndex        =   1
      Tag             =   "124"
      Top             =   7575
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblBlood 
      Height          =   5145
      Left            =   2280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2325
      Width           =   9930
      _Version        =   196608
      _ExtentX        =   17515
      _ExtentY        =   9075
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   12
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS310.frx":0000
      StartingColNumber=   0
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   1995
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  ���� ����Ʈ"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmBBS310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcSEL = 1
    TcBLOODNO
    tcCOMPO
    tcPTID
    tcPTNM
    tcSTATUS
    tcBLDSRC
    tcBLDYY
    tcBLDNO
    tcCompocd
    tcSTSCD
    tcSPLITOUTFG
End Enum

Private Sub cmdClear_Click()
    Call ClearAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        ClearAll
    End If
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub ClearAll()
    txtBldNo = ""
    Clear
End Sub

Private Sub Clear()
    medClearTable tblBlood
End Sub

Private Sub Form_Load()
    ClearAll
End Sub

Private Function ChkBldNo(ByVal pBldNo As String, BldSrc As String, BldYY As String, BldNo As String) As Boolean
    If chkBar.value = 1 Then
        pBldNo = Mid(pBldNo, 1, Len(pBldNo) - 2)
        BldSrc = Mid(pBldNo, 1, 2)
        BldYY = Mid(pBldNo, 3, 2)
        BldNo = Mid(pBldNo, 5, 6)
    Else
        BldSrc = medGetP(pBldNo, 1, "-")
        BldYY = medGetP(pBldNo, 2, "-")
        BldNo = medGetP(pBldNo, 3, "-")
    End If

    If BldSrc = "" Or BldYY = "" Or BldNo = "" Then
        ChkBldNo = False
    Else
        ChkBldNo = True
    End If

End Function

Private Sub Query()
    Dim i As Long, r As Long
    Dim BldSrc As String
    Dim BldYY As String
    Dim BldNo As String
    Dim Rs      As Recordset
    Dim objSQL  As clsBBSSQLStatement
    
    If txtBldNo = "" Then Exit Sub
    
    Clear
    
    '���׹�ȣ�� �������� �˻��ϰ�, ��ȣ�� 3������ҷ� �и��Ѵ�.
    If ChkBldNo(txtBldNo, BldSrc, BldYY, BldNo) = False Then
        MsgBox "���׹�ȣ�� �������� �ʽ��ϴ�.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    Set objSQL = New clsBBSSQLStatement
    Set Rs = objSQL.GetBloodInfo(BldSrc, BldYY, BldNo)
    Set objSQL = Nothing
    
    If Rs Is Nothing Then Exit Sub
    
    With tblBlood
        If Rs.RecordCount < 1 Then
            MsgBox "������ �����ϴ�.", vbCritical, Me.Caption
        Else
            r = .DataRowCnt
            For i = 1 To Rs.RecordCount
                '���������� �͸� ��ȸ---------------------
                If Rs.Fields("reserved").value & "" = "1" Or Rs.Fields("pherefg").value & "" = "1" Then
                    r = r + 1
                    If r > .MaxRows Then .MaxRows = .MaxRows + 1
                    
                    .Row = r
                    .Col = TblColumn.tcSEL:         .value = 1
                    .Col = TblColumn.tcPTID:        .value = Rs.Fields("ptid").value & ""
                    .Col = TblColumn.tcPTNM:        .value = GetPtNm(Rs.Fields("ptid").value & "")
                    .Col = TblColumn.tcCOMPO:       .value = Rs.Fields("compocd").value & "" & " " & _
                                                             medGetP(Get_CompNm(Rs.Fields("compocd").value & ""), 1, COL_DIV)
                    .Col = TblColumn.tcBLDSRC:      .value = Rs.Fields("bldsrc").value & ""
                    .Col = TblColumn.tcBLDYY:       .value = Rs.Fields("bldyy").value & ""
                    .Col = TblColumn.tcBLDNO:       .value = Rs.Fields("bldno").value & ""
                    .Col = TblColumn.tcCompocd:     .value = Rs.Fields("compocd").value & ""
                    .Col = TblColumn.tcSTSCD:       .value = Rs.Fields("stscd").value & ""
                    .Col = TblColumn.tcSPLITOUTFG:  .value = Rs.Fields("splitoutfg").value & ""
                    
                    .Col = TblColumn.TcBLOODNO:     .value = Rs.Fields("bldsrc").value & "" & "-" & Rs.Fields("bldyy").value & "" & "-" & Rs.Fields("bldno").value & ""
                    
                    .Col = TblColumn.tcSTATUS:
                                                 If Rs.Fields("splitoutfg").value & "" = "1" Then
                                                    .value = "��ȹ���"
                                                 Else
                                                    Select Case Rs.Fields("stscd").value & ""
                                                        Case BBSBloodStatus.stsENTER
                                                            .value = "�԰�"
                                                        Case BBSBloodStatus.stsASSIGN:
                                                            .value = "Assign"
                                                        Case BBSBloodStatus.stsBAG:
                                                            .value = "ȸ��"
                                                        Case BBSBloodStatus.stsDELIVERY:
                                                            .value = "���"
                                                        Case BBSBloodStatus.stsENTER:
                                                            .value = ""
                                                        Case BBSBloodStatus.stsEXPIRE:
                                                            .value = "���"
                                                        Case BBSBloodStatus.stsRETURN:
                                                            .value = ""
                                                    End Select
                                                 End If
                End If
                Rs.MoveNext
            Next i
        End If
    End With
    Set Rs = Nothing
    
End Sub

Private Function Save() As Boolean
    '�������� ��Ҵ� revsered��� �ʵ�(flag)�� ���� �����Ѵ�.
    'ȯ��id�� �״�� �д�.
    'Assign������ ������ ������ Ǯ�� ���Ѵ�.
    Dim objBeginTrans As clsBeginTrans
    
    Dim BldSrc        As String
    Dim BldYY         As String
    Dim BldNo         As String
    Dim CompoCd       As String
    
    Dim SSQL    As String
    
    Dim i       As Long
    
    Set objBeginTrans = New clsBeginTrans
    
On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    
    With tblBlood

        For i = 1 To .DataRowCnt
            .Row = i
            '���õ� ���׸�--------------------------------------
            .Col = TblColumn.tcSEL
            If .value = 1 Then
                'Assign���� ������ ���׸�-----------------------
                .Col = TblColumn.tcSTSCD
                If .value <= BBSBloodStatus.stsRETURN Then
                    .Col = TblColumn.tcBLDSRC:  BldSrc = .value
                    .Col = TblColumn.tcBLDYY:   BldYY = .value
                    .Col = TblColumn.tcBLDNO:   BldNo = .value
                    .Col = TblColumn.tcCompocd: CompoCd = .value
                    
                    SSQL = objBeginTrans.GetSQL_CancelBloodResreved(BldSrc, BldYY, BldNo, CompoCd)
                    DBConn.Execute SSQL
                End If
            End If
        Next i
    
    End With
    
    DBConn.CommitTrans
    Save = True
    Set objBeginTrans = Nothing
    Exit Function
    
SAVE_ERROR:
    DBConn.RollbackTrans
    Save = False
    Set objBeginTrans = Nothing
    MsgBox Err.Description, vbExclamation
End Function
Private Sub txtBldNo_Change()
    Dim lngLen As Long
    
    If chkBar.value = 1 Then Exit Sub
    
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_GotFocus()
    With txtBldNo
        .SelStart = 0
        .SelLength = Len(.Text)
        .tag = .Text
    End With
End Sub


Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub txtBldNo_LostFocus()
    If chkBar.value <> 1 Then
        If Len(Trim(txtBldNo)) <= 6 Then Exit Sub
    End If
    If Trim(txtBldNo) = "" Then Exit Sub
    If txtBldNo.tag = txtBldNo Then Exit Sub
    
    Call Query
    With txtBldNo
        .SelStart = 0
        .SelLength = Len(.Text)
        .tag = .Text
    End With
End Sub
