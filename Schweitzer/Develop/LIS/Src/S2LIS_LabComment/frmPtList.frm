VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPtList 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "ȯ�� ����Ʈ"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmPtList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOut 
      Caption         =   "�����"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   690
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FEDECD&
      Caption         =   "�����ۼ�"
      Height          =   390
      Left            =   0
      Style           =   1  '�׷���
      TabIndex        =   16
      Top             =   390
      Visible         =   0   'False
      Width           =   1095
   End
   Begin DRcontrol1.DrFrame fraHistory 
      Height          =   2475
      Left            =   2835
      TabIndex        =   12
      Top             =   1410
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   4366
      Title           =   "���� ���� ����"
      BackColor       =   16776191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFECFF&
         Caption         =   "���"
         Height          =   300
         Left            =   585
         Style           =   1  '�׷���
         TabIndex        =   15
         Top             =   2040
         Width           =   570
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFECFF&
         Caption         =   "��ȸ"
         Height          =   300
         Left            =   1170
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   2040
         Width           =   570
      End
      Begin VB.ListBox lstBedinDt 
         Appearance      =   0  '���
         Height          =   1470
         Left            =   165
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H00D8DEDA&
      Caption         =   "&History"
      Height          =   390
      Left            =   2835
      Style           =   1  '�׷���
      TabIndex        =   11
      Top             =   1005
      Width           =   930
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   2265
      Top             =   765
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FEDECD&
      Caption         =   "�����ۼ�"
      Height          =   390
      Left            =   990
      Style           =   1  '�׷���
      TabIndex        =   8
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton cmdResult 
      BackColor       =   &H00FEDECD&
      Caption         =   "�������"
      Height          =   390
      Left            =   45
      Style           =   1  '�׷���
      TabIndex        =   7
      Top             =   1005
      Width           =   930
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00FFFBF7&
      Caption         =   "��󺸱�"
      Height          =   300
      Left            =   3795
      Style           =   1  '�׷���
      TabIndex        =   5
      Tag             =   "1"
      Top             =   690
      Width           =   930
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   7605
      Left            =   45
      TabIndex        =   4
      Top             =   1425
      Width           =   4680
      _Version        =   196608
      _ExtentX        =   8255
      _ExtentY        =   13414
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      MaxRows         =   50
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmPtList.frx":058A
      TextTip         =   4
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00D8DEDA&
      Caption         =   "Re&fresh"
      Height          =   390
      Left            =   3795
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   1005
      Width           =   930
   End
   Begin MSComCtl2.MonthView mnvDate 
      Height          =   2220
      Left            =   225
      TabIndex        =   9
      Top             =   465
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   15658734
      Appearance      =   1
      StartOfWeek     =   24248321
      CurrentDate     =   36598
   End
   Begin VB.Label lblChangeDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�Կ��� : "
      Height          =   180
      Left            =   270
      MouseIcon       =   "frmPtList.frx":0CBE
      MousePointer    =   99  '����� ����
      TabIndex        =   10
      Tag             =   "20102"
      ToolTipText     =   "Click�Ͻø� �Կ����� ������ �� �ֽ��ϴ�."
      Top             =   255
      Width           =   720
   End
   Begin VB.Label lblRptCnt 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3600
      TabIndex        =   6
      Top             =   255
      Width           =   120
   End
   Begin VB.Label lblTotCnt 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2715
      TabIndex        =   2
      Top             =   255
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��       ����         �� �����"
      Height          =   180
      Left            =   2265
      TabIndex        =   1
      Top             =   255
      Width           =   2280
   End
   Begin VB.Label lblBedinDt 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1500
      TabIndex        =   0
      Top             =   255
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DBF2FD&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   480
      Left            =   75
      Shape           =   4  '�ձ� �簢��
      Top             =   105
      Width           =   4635
   End
End
Attribute VB_Name = "frmPtList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngOldRow As Long
Private lngOldColor As Long

Private Sub cmdAll_Click()
    Dim strTmp As String
    
    With tblPtList
        If lngOldRow > 0 Then
            .Row = lngOldRow
            .Col = 10
            If .Value = "Y" Then    '���� ����ڰ� �����ߴ� ȯ��
                .Col = 1
                strTmp = .Value 'ȯ��ID
                Call UnlockPtnt(strTmp, Format(lblBedinDt.Caption, CS_DateDbFormat))
            End If
        End If
        lngOldRow = -1
        frmMain.lblMsg1.Caption = ""
    End With
    If cmdAll.Tag = "1" Then  '��󺸱�
        cmdAll.Tag = "0"
        cmdAll.Caption = "��κ���"
    Else        '��κ���
        cmdAll.Tag = "1"
        cmdAll.Caption = "��󺸱�"
    End If
    Call Get_Data
End Sub

Private Sub cmdCancel_Click()
    fraHistory.Visible = False
End Sub

Private Sub cmdHistory_Click()
    Dim SqlStmt As String
    Dim Rs As Recordset
    
    If lngOldRow <= 0 Then Exit Sub
    tblPtList.Row = lngOldRow
    tblPtList.Col = 1
    
    SqlStmt = " select bedindt from " & T_LAB501 & " " & _
              " where  " & DBW("ptid = ", tblPtList.Value) & _
              " and    " & DBW("bedindt <> ", gBedInDT) & _
              " order by bedindt"
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    lstBedinDt.Clear
    While (Not Rs.EOF)
        lstBedinDt.AddItem Format(Trim("" & Rs.Fields("BedinDt").Value), CS_DateLongMask)
        Rs.MoveNext
    Wend
    
    If lstBedinDt.ListCount > 0 Then
        fraHistory.Visible = True
        fraHistory.ZOrder 0
    Else
        fraHistory.Visible = False
    End If
    
'    Rs.RsClose
    Set Rs = Nothing
    
End Sub

Private Sub cmdQuery_Click()
    
    If lstBedinDt.ListCount = 0 Then Exit Sub
    
    With frmReport
        .Show
        .ZOrder 0
        If lngOldRow <= 0 Then Exit Sub
        DoEvents
        tblPtList.Row = lngOldRow
        tblPtList.Col = 1
        If .ptid <> tblPtList.Value Or Not .QueryFg Then
            .ptid = tblPtList.Value
            .BedinDt = Format(lstBedinDt.Text, CS_DateDbFormat)
            Call .StartQuery
        End If
    End With
        
End Sub

Private Sub cmdReport_Click()
    With frmReport
        .Show
        .ZOrder 0
        If lngOldRow <= 0 Then Exit Sub
        DoEvents
        tblPtList.Row = lngOldRow
        tblPtList.Col = 1
        If .ptid <> tblPtList.Value Or Not .QueryFg Then
            .ptid = tblPtList.Value
            .BedinDt = Format(lblBedinDt.Caption, CS_DateDbFormat)
            Call .StartQuery
        End If
    End With
End Sub

Private Sub cmdResult_Click()
    With frmResultReview
        .Show
        .ZOrder 0
        If lngOldRow <= 0 Then Exit Sub
        DoEvents
        tblPtList.Row = lngOldRow
        tblPtList.Col = 1
        If .txtPtId.Text <> tblPtList.Value Or Not .QueryFg Then
            .txtPtId.Text = tblPtList.Value
            .BedinDt = Format(lblBedinDt.Caption, CS_DateDbFormat)
            Call .Call_PtId_LostFocus
        End If
    End With
End Sub

Private Sub Command1_Click()
    Dim strPtid As String
    Dim strDate As String
    Set objLabComments = New clsLabComments

    With objLabComments
        strPtid = gPatId
        strDate = sBedInDT
    End With

    Call f_set_501(strPtid, strDate)
    
    With frmReport
        .Show
        .ZOrder 0
        .ptid = strPtid
        .BedinDt = Format(strDate, CS_DateDbFormat)
        Call .StartQuery
    End With
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    
    lngOldRow = -1
    lngOldColor = 0
    lblBedinDt.Caption = Format(DateAdd("d", Val(objDoctor.Daycnt) * (-1), Now), CS_DateLongFormat)
    gBedInDT = Format(lblBedinDt.Caption, CS_DateDbFormat)
    
    Call Get_Data
    Call Command1_Click
End Sub

Private Sub cmdRefresh_Click()
    Dim strTmp As String
    
    With tblPtList
        If lngOldRow > 0 Then
            .Row = lngOldRow
            .Col = 10
            If .Value = "Y" Then    '���� ����ڰ� �����ߴ� ȯ��
                .Col = 1
                strTmp = .Value 'ȯ��ID
                Call UnlockPtnt(strTmp, Format(lblBedinDt.Caption, CS_DateDbFormat))
            End If
        End If
        lngOldRow = -1
        frmMain.lblMsg1.Caption = ""
    End With
    Call Get_Data
End Sub


Private Sub Get_Data()

    Dim i As Integer
    Dim SqlStmt As String
    Dim tmpRs As Recordset
    Dim Rs As New Recordset
    Dim strSql  As String
    
    Me.Caption = "ȯ�� ����Ʈ (����Ÿ �ε���..)"
    lblTotCnt.Caption = ""
    lblRptCnt.Caption = ""
    
    Screen.MousePointer = vbArrowHourglass
    
'    '** ���� �������� : �����Ǻ� ������ ���� �ʰ� �����ϴ� �ɷ� �� ����... By M.G.Choi 2006.09.05
'    SqlStmt = " Select a." & F_INPTID & " ptid, " & F_BEDINDT2("a") & " bedindt, a." & F_BEDINTM & " bedintm, a." & F_PTWARDID & " wardid, " & _
'              "        a." & F_PTROOMID & " roomid, a." & F_PTDIV & " ptdiv, c." & F_PTNM & " ptnm, b.empnm, d.field1 as PtDivNm, " & _
'              "        e.rptdt, e.rptid, e.donefg, e.prtfg " & _
'              " from   " & T_LAB015 & " b, " & T_HIS001 & " c, " & T_LAB032 & " d, " & T_LAB501 & " e, " & T_HIS002 & " a " & _
'              " where  a." & F_BEDINDT & " = to_date('" & gBedinDt & "', 'yyyymmdd')" & _
'              " and    a." & F_PTWARDID & " <> 'ER' " & _
'              " and    " & DBJ("e.ptid =* a." & F_INPTID) & _
'              " and    " & DBJ("e.bedindt =* trim(" & F_BEDINDT2("a") & ")") & _
'              " and    " & DBJ("b.empid =* e.rptid") & _
'              " and    c." & F_PTID & " = a." & F_INPTID & _
'              " and    " & DBJ(DBW("d.cdindex = ", LC3_PtDiv)) & _
'              " and    " & DBJ("d.cdval1 =* a." & F_PTDIV) & _
'              " and    a." & F_BEDOUTDT & " is null " & _
'              " order  by ptid"
    
    '** ���� ===================================================================================
    If chkOut.Value = 1 Then
        SqlStmt = "SELECT a.ptid ptid,                       "
        SqlStmt = SqlStmt + "       a.bedindt bedindt,                 "
        SqlStmt = SqlStmt + "       a.bedintm bedintm,                 "
        SqlStmt = SqlStmt + "       a.wardid wardid,                   "
        SqlStmt = SqlStmt + "       a.hosilid roomid,                  "
        SqlStmt = SqlStmt + "       c.pattype ptdiv,                   "
        SqlStmt = SqlStmt + "       c.patname ptnm,                    "
        SqlStmt = SqlStmt + "       b.empnm,                           "
        SqlStmt = SqlStmt + "       d.field1 AS PtDivNm,               "
        SqlStmt = SqlStmt + "       a.rptdt,                           "
        SqlStmt = SqlStmt + "       a.rptid,                           "
        SqlStmt = SqlStmt + "       a.donefg,                          "
        SqlStmt = SqlStmt + "       a.prtfg                            "
        SqlStmt = SqlStmt + "  FROM s2com006 b,                        "
        SqlStmt = SqlStmt + "       ORAA1.APPATBAT c,                  "
        SqlStmt = SqlStmt + "       s2lab032 d,                        "
        SqlStmt = SqlStmt + "       s2lab501 a                         "
        SqlStmt = SqlStmt + " WHERE a.bedindt = '" & gBedInDT & "'     "
        SqlStmt = SqlStmt + "       AND b.empid(+) = a.rptid           "
        SqlStmt = SqlStmt + "       AND c.patno = a.ptid               "
        SqlStmt = SqlStmt + "       AND d.cdindex(+) = 'C237'          "
        SqlStmt = SqlStmt + "       AND d.cdval1(+) = c.pattype        "
        SqlStmt = SqlStmt + "       AND                                "
        SqlStmt = SqlStmt + "       (                                  "
        SqlStmt = SqlStmt + "           a.bedoutdt IS NULL             "
        SqlStmt = SqlStmt + "           OR a.bedoutdt = '99999999'     "
        SqlStmt = SqlStmt + "       )                                  "
        SqlStmt = SqlStmt + "ORDER BY ptnm                             "
        SqlStmt = SqlStmt + "                                          "
    Else
        SqlStmt = " Select a." & F_INPTID & " ptid, " & F_BEDINDT2("a") & " bedindt, a." & F_BEDINTM & " bedintm, a." & F_PTWARDID & " wardid, " & _
                  "        a." & F_PTROOMID & " roomid, a." & F_PTDIV & " ptdiv, c." & F_PTNM & " ptnm, b.empnm, d.field1 as PtDivNm, " & _
                  "        e.rptdt, e.rptid, e.donefg, e.prtfg " & _
                  " from   " & T_LAB015 & " b, " & T_HIS001 & " c, " & T_LAB032 & " d, " & T_LAB501 & " e, " & T_HIS002 & " a " & _
                  " where  a." & F_BEDINDT & " = to_date('" & gBedInDT & "', 'yyyymmdd')" & _
                  " and    a." & F_PTWARDID & " <> 'ER' " & _
                  " and    " & DBJ("e.ptid =* a." & F_INPTID) & _
                  " and    " & DBJ("e.bedindt =* trim(" & F_BEDINDT2("a") & ")") & _
                  " and    " & DBJ("b.empid =* e.rptid") & _
                  " and    c." & F_PTID & " = a." & F_INPTID & _
                  " and    " & DBJ(DBW("d.cdindex = ", LC3_PtDiv)) & _
                  " and    " & DBJ("d.cdval1 =* a." & F_PTDIV) & _
                  " and    a." & F_BEDOUTDT & " is null " & _
                  " order  by ptid"
    End If
    '===========================================================================================
    
'    Debug.Print SqlStmt
'    Set tmpRs = OpenRecordSet(SqlStmt)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    With tblPtList
        .ReDraw = False
        .MaxRows = 0
        lblTotCnt.Caption = tmpRs.RecordCount
        For i = 1 To tmpRs.RecordCount
            If Trim("" & tmpRs.Fields("DoneFg").Value) <> "" Then
                If cmdAll.Tag = "0" Then    '��󺸱�
                    If Val("" & tmpRs.Fields("DoneFg").Value) > 1 Or Trim("" & tmpRs.Fields("RptId").Value) <> objDoctor.DoctId Then
                        GoTo Skip
                    End If
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    If objDoctor.DoctId <> Trim("" & tmpRs.Fields("RptId").Value) Then
                        .Col = -1
                        .ForeColor = &H808080   'ȸ��
                    End If
                End If
            Else
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            End If
            .Col = 1: .Value = Trim("" & tmpRs.Fields("Ptid").Value)
            .Col = 2: .Value = Trim("" & tmpRs.Fields("Ptnm").Value)
            .Col = 3: .Value = Trim("" & tmpRs.Fields("PtDivNm").Value)
            
'            Select Case PROJECT_HOSCD
'                Case "05":
'                    If tmpRs.Fields("Ptid").Value & "" = "01" Then .Value = .Value & "*"
'                Case Else
'            End Select
            
            .Col = 4: .Value = Trim("" & tmpRs.Fields("WardId").Value)
            .Col = 5: .Value = Trim("" & tmpRs.Fields("EmpNm").Value)
            .Col = 7: .Value = Trim("" & tmpRs.Fields("PrtFg").Value)
            .Col = 8: .Value = Trim("" & tmpRs.Fields("RptId").Value)
            .Col = 9: .Value = Trim("" & tmpRs.Fields("DoneFg").Value)
            If .Value = "0" Or .Value = "1" Then
                .Col = 5: .Value = "������"
            ElseIf .Value = "2" Then
                .Col = 6: .Value = "Y"  '���Ῡ��
            End If
            .Col = 11: .Value = Trim("" & tmpRs.Fields("RptDt").Value)
            
            '** ���� ----------------------------------------------------------
            'If .Value <> "" Then lblRptCnt.Caption = Val(lblRptCnt.Caption) + 1
            '------------------------------------------------------------------
            
            '** ���� �����Ǻ� Count �� �����Ѵ�.
            '   By M.G.Choi 2005.11.16
            If objDoctor.DoctId = Trim("" & tmpRs.Fields("RptId").Value) Then
                If .Value <> "" Then lblRptCnt.Caption = Val(lblRptCnt.Caption) + 1
            End If
                
Skip:
            tmpRs.MoveNext
        Next
        .RowHeight(-1) = 10.5
        .ReDraw = True
    End With
    
    
    
'    tmpRs.RsClose
    Set tmpRs = Nothing
    Set Rs = Nothing
    Me.Caption = "ȯ�� ����Ʈ"
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If gPtntId <> "" Then
        Call UnlockPtnt(gPtntId, gBedInDT)
        frmMain.lblMsg1.Caption = ""
    End If

End Sub

Private Sub lblChangeDate_Click()
    mnvDate.Value = lblBedinDt.Caption
    mnvDate.Visible = True
    mnvDate.ZOrder 0
    mnvDate.SetFocus
End Sub

Private Sub mnvDate_DateClick(ByVal DateClicked As Date)
    lblBedinDt.Caption = Format(mnvDate.Value, CS_DateLongFormat)
    gBedInDT = Format(lblBedinDt.Caption, CS_DateDbFormat)
    mnvDate.Visible = False
    Call cmdRefresh_Click
End Sub


Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)

    Static iSortOrder As Integer
    Dim tmpColNm As String
    Dim strTmp As String
    Dim blnCheck As Boolean
    
    With tblPtList
        
        If Row = 0 Then  'Sort...
            .Row = -1: .Col = -1
            .SortBy = SortByRow
            .SortKey(1) = Col
            If iSortOrder = SortKeyOrderAscending Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                iSortOrder = SortKeyOrderDescending
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                iSortOrder = SortKeyOrderAscending
            End If
            .Action = ActionSort
            Exit Sub
        End If
        
        If lngOldRow = Row Then Exit Sub
        If lngOldRow > 0 Then
            .Row = lngOldRow
            .Col = -1
            .ForeColor = lngOldColor
            '.BackColor = lngOldColor
            .Col = 10
            If .Value = "Y" Then    '���� ����ڰ� �����ߴ� ȯ��
                .Col = 1
                strTmp = .Value 'ȯ��ID
                Call UnlockPtnt(strTmp, Format(lblBedinDt.Caption, CS_DateDbFormat))
            End If
        End If
        
        .Row = Row
        .Col = Col
        lngOldRow = Row
        'lngOldColor = .BackColor
        lngOldColor = .ForeColor
        .Col = -1
        .ForeColor = DCM_LightRed
        
        .Col = 1
        strTmp = .Value 'ȯ��ID
        blnCheck = CheckStatus(strTmp, Format(lblBedinDt.Caption, CS_DateDbFormat), objDoctor.DoctId)
        If chkOut.Value = 1 Then
            If blnCheck Then
                .Col = 10: .Value = "Y"
                cmdReport.Enabled = True
                cmdResult.Enabled = True
                gPtntId = strTmp
                
                '-- ���� -------------
    '            Call cmdReport_Click
                '---------------------
                
                '-- ���� �ش� ������ �Ϸ� ���� Ȯ�� Check
                '   By M.G.Choi 2005.11.16
                If objDoctor.RptCount <= objDoctor.Ptcnt Then
                    Call cmdReport_Click
                Else
                    MsgBox "�Ϸ� ������ �ʰ� �Ǿ����ϴ�." & Chr(13) & "�������� ���� �Ͻ� �� �ֽ��ϴ�.", vbCritical, "���� �ʰ�"
                End If
                
    '            If Trim(lblRptCnt.Caption) <= objDoctor.RptCount Then
    '                Call cmdReport_Click
    '            Else
    '                MsgBox "�Ϸ� ������ �ʰ� �Ǿ����ϴ�." & Chr(13) & "�������� ���� �Ͻ� �� �ֽ��ϴ�.", vbCritical, "���� �ʰ�"
    '            End If
                '--------------------------------------------------------------------------------------
            End If
        Else
            If blnCheck Then
                .Col = 10: .Value = "Y"
                cmdReport.Enabled = True
                cmdResult.Enabled = True
                gPtntId = strTmp
                
                '-- ���� -------------
    '            Call cmdReport_Click
                '---------------------
                
                '-- ���� �ش� ������ �Ϸ� ���� Ȯ�� Check
                '   By M.G.Choi 2005.11.16
                If objDoctor.RptCount <= objDoctor.Ptcnt Then
                    Call cmdReport_Click
                Else
                    MsgBox "�Ϸ� ������ �ʰ� �Ǿ����ϴ�." & Chr(13) & "�������� ���� �Ͻ� �� �ֽ��ϴ�.", vbCritical, "���� �ʰ�"
                End If
                
    '            If Trim(lblRptCnt.Caption) <= objDoctor.RptCount Then
    '                Call cmdReport_Click
    '            Else
    '                MsgBox "�Ϸ� ������ �ʰ� �Ǿ����ϴ�." & Chr(13) & "�������� ���� �Ͻ� �� �ֽ��ϴ�.", vbCritical, "���� �ʰ�"
    '            End If
                '--------------------------------------------------------------------------------------
            
            Else
                .Col = 10: .Value = ""
                cmdReport.Enabled = False
                cmdResult.Enabled = False
                MsgBox "�̹� ����Ǿ��ų� ���� �������� ȯ���Դϴ�.", vbExclamation, "�޼���"
                gPtntId = ""
            End If
        End If
        .Col = 1
        frmMain.lblMsg1.Caption = "���� ���õ� ȯ�� : " & .Value
        .Col = 2
        frmMain.lblMsg1.Caption = frmMain.lblMsg1.Caption & "  " & .Value
        
    End With
    
End Sub

Private Sub f_set_501(ByVal strPtid As String, ByVal strBedinDt As String)
    Dim tmpColNm As String
    Dim strTmp As String
    Dim blnCheck As Boolean

    blnCheck = CheckStatus(strPtid, Format(strBedinDt, CS_DateDbFormat), objDoctor.DoctId)
    
    If blnCheck Then
        cmdReport.Enabled = True
        cmdResult.Enabled = True
        gPtntId = strPtid
        If objDoctor.RptCount <= objDoctor.Ptcnt Then

        Else
            MsgBox "�Ϸ� ������ �ʰ� �Ǿ����ϴ�." & Chr(13) & "�������� ���� �Ͻ� �� �ֽ��ϴ�.", vbCritical, "���� �ʰ�"
        End If
    Else
        cmdReport.Enabled = False
        cmdResult.Enabled = False
        MsgBox "�̹� ����Ǿ��ų� ���� �������� ȯ���Դϴ�.", vbExclamation, "�޼���"
        gPtntId = ""
    End If

End Sub

Private Sub tblPtList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    tblPtList.SetFocus
End Sub

Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
    If Col = 3 Then
        tblPtList.Row = Row: tblPtList.Col = 3
        MultiLine = 1
        TipText = vbCrLf & "  " & tblPtList.Value & "  " & vbCrLf
        TipWidth = 2000
        Call tblPtList.SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    Else
        ShowTip = False
    End If
    
End Sub

Private Sub Timer1_Timer()
    
    Static TimeCount As Long
    
    TimeCount = TimeCount + 1
    If TimeCount = 5 Then Call Get_Data: TimeCount = 0 '5�� ����
    
End Sub

Public Sub Call_Refresh()
    Call cmdRefresh_Click
End Sub
