VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS425 
   BackColor       =   &H00DBE6E6&
   Caption         =   "���� ���"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14535
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "��ȸ(&Q)"
      Height          =   510
      Left            =   6615
      Style           =   1  '�׷���
      TabIndex        =   7
      Tag             =   "124"
      Top             =   8130
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2306
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1470
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
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2306
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   405
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�������(&S)"
      Height          =   510
      Left            =   7999
      Style           =   1  '�׷���
      TabIndex        =   8
      Tag             =   "124"
      Top             =   8130
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   9379
      Style           =   1  '�׷���
      TabIndex        =   9
      Tag             =   "124"
      Top             =   8130
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   10759
      Style           =   1  '�׷���
      TabIndex        =   10
      Tag             =   "128"
      Top             =   8130
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   840
      Left            =   2306
      TabIndex        =   1
      Top             =   630
      Width           =   9930
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ�� �Է�(&B)"
         Height          =   315
         Left            =   3915
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Value           =   1  'Ȯ��
         Width           =   1755
      End
      Begin VB.TextBox txtBldNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1530
         MaxLength       =   12
         TabIndex        =   3
         Top             =   300
         Width           =   2160
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   330
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1050
         _ExtentX        =   1852
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
   Begin FPSpread.vaSpread tblBlood 
      Height          =   6120
      Left            =   2310
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1785
      Width           =   9930
      _Version        =   196608
      _ExtentX        =   17515
      _ExtentY        =   10795
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
      MaxCols         =   13
      MaxRows         =   20
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS425.frx":0000
      StartingColNumber=   0
      TextTip         =   4
   End
End
Attribute VB_Name = "frmBBS425"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RowHeight& = 12

'����, ���׹�ȣ, ��������, ����ȯ��, ������, ��������, ����, bldsrc, bldyy, bldno, compocd, stscd, splitoutfg
Private Enum TblColumn
    tcSEL = 1
    TcBLOODNO
    tcCompo
    tcPTID
    tcDONOR
    tcDONORACCDT
    tcSTATUS
    tcBLDSRC
    tcBLDYY
    tcBldNo
    tcCompoCd
    tcSTSCD
    tcSPLITOUTFG
    tcCHKDUP
End Enum

Private Sub cmdClear_Click()
    Call InitForm
    On Error Resume Next
    txtBldNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS425 = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim objPro As clsProgress
    Dim strSql As String
    Dim RS As Recordset
    Dim i As Long
    
    strSql = " SELECT * FROM " & T_BBS401
    strSql = strSql & " where bldsrc > ' '"
    strSql = strSql & " and bldyy > ' '"
    strSql = strSql & " and bldno > 0"
    strSql = strSql & " and (reserved='1' or pherefg='1')"
    
    Set RS = New Recordset
    RS.Open strSql, DBConn
    
    If RS.EOF Then
        MsgBox "������ ������ �����ϴ�.", vbExclamation
        Set RS = Nothing
        Exit Sub
    End If
    
    Call medClearTable(tblBlood)
    tblBlood.MaxRows = 20
    tblBlood.RowHeight(-1) = RowHeight
    
    Set objPro = New clsProgress
    With objPro
        .Container = Me
        .Left = LisLabel2.Left
        .Top = LisLabel2.Top
        .Width = LisLabel2.Width
        .Height = LisLabel2.Height
        .Max = RS.RecordCount
    End With
    
'����, ���׹�ȣ, ��������, ����ȯ��, ������, ��������, ����, bldsrc, bldyy, bldno, compocd, stscd, splitoutfg
    With tblBlood
        .ReDraw = False
        Do Until RS.EOF
            If .MaxRows < .DataRowCnt Then
                .MaxRows = .MaxRows + 1
            End If
            .Row = .DataRowCnt + 1
            i = i + 1
            objPro.value = i
            
            .Col = TblColumn.tcSEL: .value = 1
            .Col = TblColumn.TcBLOODNO: .value = RS.Fields("bldsrc").value & "" & "-" & RS.Fields("bldyy").value & "" & "-" & Format(RS.Fields("bldno").value & "", "000000")
            .Col = TblColumn.tcCompo: .value = RS.Fields("compocd").value & "" & " " & _
                                                     medGetP(Get_CompNm(RS.Fields("compocd").value & ""), 1, COL_DIV)
            .Col = TblColumn.tcPTID: .value = GetPtNm(RS.Fields("ptid").value & "") & "(" & RS.Fields("ptid").value & "" & ")"
            .Col = TblColumn.tcDONOR: .value = GetDonorNm(RS.Fields("donorid").value & "")
            .Col = TblColumn.tcDONORACCDT: .value = Format(RS.Fields("donoraccdt").value & "", "####-##-##")
            .Col = TblColumn.tcSTATUS:
                                         If RS.Fields("splitoutfg").value & "" = "1" Then
                                            .value = "��ȹ���"
                                         Else
                                            Select Case RS.Fields("stscd").value & ""
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
            
            .Col = TblColumn.tcBLDSRC:      .value = RS.Fields("bldsrc").value & ""
            .Col = TblColumn.tcBLDYY:       .value = RS.Fields("bldyy").value & ""
            .Col = TblColumn.tcBldNo:       .value = RS.Fields("bldno").value & ""
            .Col = TblColumn.tcCompoCd:     .value = RS.Fields("compocd").value & ""
            .Col = TblColumn.tcSTSCD:       .value = RS.Fields("stscd").value & ""
            .Col = TblColumn.tcSPLITOUTFG:  .value = RS.Fields("splitoutfg").value & ""
            
            RS.MoveNext
        Loop
        .ReDraw = True
    End With
    
    Set RS = Nothing
    Set objPro = Nothing
End Sub

Private Function GetDonorNm(ByVal vDonorid As String) As String
    Dim strSql As String
    Dim RS As Recordset
    
    strSql = " select * from " & T_BBS601 & " where " & DBW("donorid=", vDonorid)
    
    Set RS = New Recordset
    RS.Open strSql, DBConn
    If RS.EOF = False Then
        GetDonorNm = RS.Fields("donornm").value & ""
    End If
    
    Set RS = Nothing
End Function

Private Sub cmdSave_Click()
    Dim strBldSrc        As String
    Dim strBldYY         As String
    Dim strBldNo         As String
    Dim strCompoCd       As String
    Dim strSql    As String
    Dim i       As Long
    
    If MsgBox("������ �׸��� ������� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
On Error GoTo ErrTrap
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
                    .Col = TblColumn.tcBLDSRC:  strBldSrc = .value
                    .Col = TblColumn.tcBLDYY:   strBldYY = .value
                    .Col = TblColumn.tcBldNo:   strBldNo = .value
                    .Col = TblColumn.tcCompoCd: strCompoCd = .value
                    
                    strSql = " UPDATE " & T_BBS401 & " " & _
                            " SET " & DBW("reserved=", "0", 1) & DBW("pherefg=", "0") & _
                            " WHERE " & DBW("bldsrc=", strBldSrc) & _
                            " AND " & DBW("bldyy=", strBldYY) & _
                            " AND " & DBW("bldno=", strBldNo) & _
                            " AND " & DBW("compocd=", strCompoCd)
                    DBConn.Execute strSql
                End If
            End If
        Next i
    End With
    
    DBConn.CommitTrans
    MsgBox "���������� ó���Ǿ����ϴ�.", vbInformation
    Call InitForm
    
    Exit Sub
    
ErrTrap:
    DBConn.RollbackTrans
    MsgBox "���������� ó������ �ʾҽ��ϴ�.", vbExclamation
End Sub

Private Sub Form_Load()
    Call InitForm
End Sub

Private Sub InitForm()
    txtBldNo.Text = ""
    Call medClearTable(tblBlood)
    tblBlood.MaxRows = 20
    tblBlood.RowHeight(-1) = RowHeight
End Sub

Private Sub txtBldNo_Change()
    If chkBar.value = 1 Then Exit Sub
    Dim lngLen As Long
    
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
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtBldNo.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
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

Private Sub txtBldNo_Validate(Cancel As Boolean)
    Dim strBldNo As String
    Dim strBNum As String
    
    If txtBldNo.Text = "" Then Exit Sub
    
    strBldNo = GetBldNo
    
'    strBNum = Replace(strBldNo, "-", "")
    
    If CheckDup(strBldNo) Then
        Cancel = True
        MsgBox "�̹� �Էµ� ���׹�ȣ�Դϴ�.", vbExclamation
    Else
        If InsertTable = False Then
            Cancel = True
        End If
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function GetBldNo() As String
    '�Էµ� ���׹�ȣ�� ##-##-#������� ��ȯ�Ѵ�.
    If chkBar.value = 1 Then
        GetBldNo = Mid(txtBldNo.Text, 1, 2) & "-" & Mid(txtBldNo.Text, 3, 2) & "-" & Mid(txtBldNo.Text, 5, 6)
    Else
        GetBldNo = txtBldNo.Text
    End If
End Function

Private Function CheckDup(ByVal vBldNo As String) As Boolean
    Dim i As Long
    CheckDup = False
    For i = 1 To tblBlood.DataRowCnt
        tblBlood.Row = i
        tblBlood.Col = TblColumn.TcBLOODNO
        If vBldNo = tblBlood.value Then
            CheckDup = True
            Exit For
        End If
    Next
End Function

Private Function InsertTable() As Boolean
    Dim objPro As clsProgress
    Dim strSql As String
    Dim RS As Recordset
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldNo As String
    
    If chkBar.value = 1 Then
        strBldSrc = Mid(txtBldNo.Text, 1, 2)
        strBldYY = Mid(txtBldNo.Text, 3, 2)
        strBldNo = Format(Mid(txtBldNo.Text, 5, 6), "00000#")
    Else
        strBldSrc = medGetP(txtBldNo.Text, 1, "-")
        strBldYY = medGetP(txtBldNo.Text, 2, "-")
        strBldNo = Format(medGetP(txtBldNo.Text, 3, "-"), "######")
    End If
    
    strSql = " SELECT * FROM " & T_BBS401
    strSql = strSql & " where " & DBW("bldsrc=", strBldSrc)
    strSql = strSql & " and " & DBW("bldyy=", strBldYY)
    strSql = strSql & " and " & DBW("bldno=", strBldNo)
    strSql = strSql & " and (reserved='1' or pherefg='1')"
    
    Set RS = New Recordset
    RS.Open strSql, DBConn
    
    If RS.EOF Then
        MsgBox "������ ������ �����ϴ�.", vbExclamation
        Set RS = Nothing
        Exit Function
    End If
    
'    Call medClearTable(tblBlood)
'    tblBlood.MaxRows = 20
    tblBlood.RowHeight(-1) = RowHeight

'����, ���׹�ȣ, ��������, ����ȯ��, ������, ��������, ����, bldsrc, bldyy, bldno, compocd, stscd, splitoutfg
    With tblBlood
        .ReDraw = False
        Do Until RS.EOF
            If .MaxRows < .DataRowCnt Then
                .MaxRows = .MaxRows + 1
            End If
            .Row = .DataRowCnt + 1
            
            .Col = TblColumn.tcSEL: .value = 1
            .Col = TblColumn.TcBLOODNO: .value = RS.Fields("bldsrc").value & "" & "-" & RS.Fields("bldyy").value & "" & "-" & Format(RS.Fields("bldno").value & "", "000000")
            .Col = TblColumn.tcCompo: .value = RS.Fields("compocd").value & "" & " " & _
                                                     medGetP(Get_CompNm(RS.Fields("compocd").value & ""), 1, COL_DIV)
            .Col = TblColumn.tcPTID: .value = GetPtNm(RS.Fields("ptid").value & "") & "(" & RS.Fields("ptid").value & "" & ")"
            .Col = TblColumn.tcDONOR: .value = GetDonorNm(RS.Fields("donorid").value & "")
            .Col = TblColumn.tcDONORACCDT: .value = Format(RS.Fields("donoraccdt").value & "", "####-##-##")
            .Col = TblColumn.tcSTATUS:
                                         If RS.Fields("splitoutfg").value & "" = "1" Then
                                            .value = "��ȹ���"
                                         Else
                                            Select Case RS.Fields("stscd").value & ""
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
            
            .Col = TblColumn.tcBLDSRC:      .value = RS.Fields("bldsrc").value & ""
            .Col = TblColumn.tcBLDYY:       .value = RS.Fields("bldyy").value & ""
            .Col = TblColumn.tcBldNo:       .value = RS.Fields("bldno").value & ""
            .Col = TblColumn.tcCompoCd:     .value = RS.Fields("compocd").value & ""
            .Col = TblColumn.tcSTSCD:       .value = RS.Fields("stscd").value & ""
            .Col = TblColumn.tcSPLITOUTFG:  .value = RS.Fields("splitoutfg").value & ""
            
            RS.MoveNext
        Loop
        .ReDraw = True
    End With
    
    InsertTable = True
    
    Set RS = Nothing
End Function
