VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkList 
   BorderStyle     =   1  '���� ����
   Caption         =   "WorkList"
   ClientHeight    =   9075
   ClientLeft      =   2490
   ClientTop       =   810
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   13035
   Begin VB.CommandButton cmdDel 
      Caption         =   "����(&D)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8730
      Style           =   1  '�׷���
      TabIndex        =   26
      Top             =   60
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Sample No]"
      Height          =   675
      Index           =   2
      Left            =   60
      TabIndex        =   22
      Top             =   780
      Width           =   4155
      Begin VB.TextBox txtSampleNo 
         Appearance      =   0  '���
         Height          =   330
         Left            =   1620
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtSampleNoCnt 
         Appearance      =   0  '���
         Height          =   330
         Left            =   3360
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ü��: "
         Height          =   180
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Top             =   300
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Sample No:"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "�ݱ�(&Q)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11940
      TabIndex        =   19
      Top             =   60
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Profile �ϰ�����]"
      Height          =   675
      Index           =   1
      Left            =   60
      TabIndex        =   17
      Top             =   1500
      Width           =   6435
      Begin VB.TextBox txtRowT 
         Appearance      =   0  '���
         Height          =   330
         Left            =   1620
         TabIndex        =   7
         Top             =   240
         Width           =   675
      End
      Begin VB.ComboBox cboProfile 
         Height          =   300
         Left            =   3300
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtRowF 
         Appearance      =   0  '���
         Height          =   330
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Profile:"
         Height          =   180
         Index           =   3
         Left            =   2520
         TabIndex        =   21
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Row:          ~"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[���ڵ��ȣ]"
      Height          =   675
      Index           =   0
      Left            =   4260
      TabIndex        =   14
      Top             =   780
      Width           =   8715
      Begin VB.TextBox txtBarcodeStartRow 
         Appearance      =   0  '���
         Height          =   330
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtBarcodeSeqF 
         Appearance      =   0  '���
         Height          =   330
         Left            =   5820
         TabIndex        =   4
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtBarcodeSeqT 
         Appearance      =   0  '���
         Height          =   330
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
      Begin MSComCtl2.DTPicker dtp�˻����� 
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   100335617
         CurrentDate     =   39954
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "���ڵ� Seq:          ~"
         Height          =   180
         Index           =   5
         Left            =   4740
         TabIndex        =   25
         Top             =   300
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Row:"
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   15
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10860
      TabIndex        =   13
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9780
      Style           =   1  '�׷���
      TabIndex        =   12
      Top             =   60
      Width           =   1035
   End
   Begin VB.CheckBox chkAll 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   600
      TabIndex        =   10
      Top             =   2280
      Width           =   195
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6795
      Left            =   60
      TabIndex        =   9
      Top             =   2220
      Width           =   12915
      _Version        =   393216
      _ExtentX        =   22781
      _ExtentY        =   11986
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   29
      MaxRows         =   100
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSear.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "WorkList"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   9
      Left            =   480
      TabIndex        =   18
      Top             =   180
      Width           =   1185
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   435
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '�ܻ�
      Height          =   435
      Index           =   1
      Left            =   3960
      Top             =   240
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   60
      X2              =   12900
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '����
      Caption         =   "����Ϸ� : ������, �̿Ϸ� : ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '�ܻ�
      Height          =   615
      Index           =   1
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   4035
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iIndex As Integer

Public glRow As Long
Public gOCnt As Integer
Public gCount As String

Private Sub cboProfile_Change()
'''    If Trim(cboProfile) = "" Then MsgBox "������ Profile�� �����Ͻʽÿ�!", vbCritical, "�������": cboProfile.SetFocus: Exit Sub
'''    If Trim(txtRowF) = "" Then MsgBox "Row(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
'''    If Trim(txtRowT) = "" Then MsgBox "Row(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
'''    If IsNumeric(txtRowF) = False Then MsgBox "Row(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
'''    If IsNumeric(txtRowT) = False Then MsgBox "Row(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
'''
'''    If Val(txtRowF) < 1 Then MsgBox "Row(From)�� 1 Row �̻��̾�� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
'''    If Val(txtRowT) < 1 Then MsgBox "Row(To)�� 1 Row �̻��̾�� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
'''
'''    If Val(txtRowF) > Val(txtRowT) Then
'''        MsgBox "Row ������ (��)�Է��Ͻʽÿ�", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
'''    End If
'''
'''    For intX = Val(txtRowF) To Val(txtRowT)
'''        If intX > vasList.MaxRows Then Exit For
'''        Call SET_CODE_SPREAD_COMBO_CELL_L(vasList, 4, intX, Trim(Left(cboProfile, 2)), 2)
'''        Call SET_PROFILE(CLng(intX))
'''    Next intX
'''
'''    vasList.SetFocus
End Sub

Private Sub cboProfile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtRowF) = "" Then MsgBox "Row(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        If Trim(txtRowT) = "" Then MsgBox "Row(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
        If IsNumeric(txtRowF) = False Then MsgBox "Row(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        If IsNumeric(txtRowT) = False Then MsgBox "Row(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
        
        If Val(txtRowF) < 1 Then MsgBox "Row(From)�� 1 Row �̻��̾�� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        If Val(txtRowT) < 1 Then MsgBox "Row(To)�� 1 Row �̻��̾�� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
        
        If Val(txtRowF) > Val(txtRowT) Then
            MsgBox "Row ������ (��)�Է��Ͻʽÿ�", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        End If
            
        For intX = Val(txtRowF) To Val(txtRowT)
            If intX > vasList.MaxRows Then Exit For
            Call SET_CODE_SPREAD_COMBO_CELL_L(vasList, 4, intX, Trim(Left(cboProfile, 2)), 2)
            
            If Trim(cboProfile) = "" Then
                vasList.ClearRange 5, intX, vasList.MaxCols, intX, -1
            Else
                Call SET_PROFILE(CLng(intX))
            End If
        Next intX
        
        txtRowF.SetFocus
    End If
End Sub

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 0
        Next iRow
    End If
End Sub

'Private Sub chkAllOrder_Click()
'    If chkAllOrder.Value = 1 Then
'        vasOrder.Row = -1
'        vasOrder.Col = 1
'        vasOrder.Value = 1
'    Else
'        vasOrder.Row = -1
'        vasOrder.Col = 1
'        vasOrder.Value = 0
'    End If
'End Sub

'Private Sub cmdCalendar_Click(Index As Integer)
'    iIndex = Index
'    If Index = 0 Then
'        monvCal.Left = 1680
'        monvCal.Top = 570
'        monvCal.Visible = True
'
'        monvCal.Value = dtpSDate.Text
'    ElseIf Index = 1 Then
'        monvCal.Left = 3600
'        monvCal.Top = 570
'        monvCal.Visible = True
'
'        monvCal.Value = dtpEDate.Text
'    End If
'    'monvCal.Visible = True
'End Sub

'Private Sub cmdClose_Click()
'    Dim sCnt As String
'    Dim iRow As Integer
'
'    Dim sExamCode As String
'    Dim sSubCode As String
'    Dim sExamName As String
'
'    Dim sEQUIPCODE As String
'    Dim sAge As String
'    Dim i As Integer
'
'    sCnt = ""
'
'    '2005/10/18 �̻���
'    'ó�����ڿ��� �˻����ڷ� ������
'    'txtDate = Format(txtDate.Text, "yyyymmdd")
'    txtDate = Format(frmInterface.txtToday.Text, "yyyymmdd")
'
'    For iRow = 1 To vasOrder.DataRowCnt
'        vasOrder.Row = iRow
'        vasOrder.Col = 1
'
'        If vasOrder.Value = 1 Then
'            sExamCode = Trim(GetText(vasOrder, iRow, 2))
''            sSubCode = Trim(GetText(vasOrder, iRow, 3))
'            sExamName = Trim(GetText(vasOrder, iRow, 3))
'
'            sEQUIPCODE = GetEquip_ExamCode(sExamCode)
''            txtDate = "20080311"
'            SQL = "delete from PAT_RES WHERE EXAMDATE = '" & txtDate & "' " & vbCrLf & _
'                  "and EQUIPNO = '" & gEquip & "' and pid = '" & Trim(txtPID.Text) & "' " & vbCrLf & _
'                " and examcode = '" & sExamCode & "'"
''            SQL = " Insert Into PAT_RES(EXAMDATE, EQUIPNO, BARCODE, EQUIPCODE,  " & vbCrLf & _
''                  " examcode, subcode, examname, pid, pname, psex, page, recedate, resdate, sendflag)  " & vbCrLf & _
''                  " Values ( '" & Trim(txtDate) & "', '" & gEquip & "', '" & Trim(txtPID.Text) & "' , '" & Trim(sEQUIPCODE) & "', " & vbCrLf & _
''                  " '" & sExamCode & "', '" & sSubCode & "', '" & sExamName & "', '" & Trim(txtPID.Text) & "', " & vbCrLf & _
''                  " '" & Trim(txtName.Text) & "', '" & Trim(txtSex.Text) & "', " & sAge & ", " & vbCrLf & _
''                  " '" & Trim(txtReqDate) & "', '', 'O') "
'
'            res = SendQuery(gLocal, SQL)
'
'            If res = -1 Then
'                SaveQuery SQL
'            End If
'        End If
'    Next iRow
'
'    sspOrder.Visible = False
'End Sub
'End Sub

Private Sub cmdClear_Click()
    If vasList.MaxRows > 0 Then vasList.MaxRows = 0
End Sub

Private Sub cmdDel_Click()
    Dim i As Integer
    Dim SEQNO As String
    For i = vasList.MaxRows To 1 Step -1
        If GET_CELL(vasList, 1, i) = "1" Then
            DeleteRow vasList, i, i
        End If
    Next i
'    SEQNO = GET_CELL(vasList, 2, 1)
'    For i = 2 To vasList.DataRowCnt
'            SEQNO = CInt(SEQNO) + 1
'            SET_CELL vasList, 2, i, Format(SEQNO, "0###")
'    Next i
    
    
End Sub

Private Sub cmdSave_Click()
    Dim strSaveYN   As String
    
    For intX = 1 To vasList.MaxRows
        If GET_CELL(vasList, 1, intX) = "1" Then strSaveYN = "Y": Exit For
    Next intX
    If strSaveYN <> "Y" Then MsgBox "������ �ڷᰡ �����ϴ�.", vbCritical, "���� �Ұ�": Exit Sub
    
    If MsgBox("���� �ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "���� ����") = vbCancel Then Exit Sub
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    With vasList
        For intX = 1 To .MaxRows
            If GET_CELL(vasList, 1, intX) = "1" Then
                '/���ڵ尡 �ٸ����¿��� Sample No�� �ٸ� �� �ִ�.
                '/�˻����� ���� Sample No�� �ߺ��ɼ� ����.
                gstrQuy = "SELECT * "
                gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
                gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPNO  = '" & gtypREG_INFO.EQUIPCD & "' "
'                gstrQuy = gstrQuy & vbCrLf & "   AND EXAMDATE = '" & Format(frmInterface.dtpExamDate, "yyyymmdd") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND EXAMDATE = '" & Format(dtp�˻�����, "yyyymmdd") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND SEQNO    = '" & Trim(GET_CELL(vasList, 2, intX)) & "' "
                If ReadSQL(gstrQuy, ADR) = False Then
                    Call CloseDB
                    Exit Sub
                End If
                
                If Not ADR Is Nothing Then
                    ADR.Close: Set ADR = Nothing
                    Call CloseDB
                    
'                    MsgBox "�˻�����(" & frmInterface.dtpExamDate.Value & ")�� ������ Sample No(" & Trim(GET_CELL(vasList, 2, intX)) & ")�� �����մϴ�." & vbCrLf & _
                    "Sample No�� ������ �ϱ�ٶ��ϴ�.", vbCritical, "Sample No �ߺ�����"
                                        
                    MsgBox "�˻�����(" & dtp�˻�����.Value & ")�� ������ Sample No(" & Trim(GET_CELL(vasList, 2, intX)) & ")�� �����մϴ�." & vbCrLf & _
                    "Sample No�� ������ �ϱ�ٶ��ϴ�.", vbCritical, "Sample No �ߺ�����"
                                        
                                        
                    Exit Sub
                End If
            End If
        Next intX
    End With

    ADC.BeginTrans

    With vasList
        For intX = 1 To .MaxRows
            If GET_CELL(vasList, 1, intX) = "1" Then
                gstrQuy = "DELETE FROM PAT_RES "
'                gstrQuy = gstrQuy & vbCrLf & " WHERE EXAMDATE = '" & Format(frmInterface.dtpExamDate, "yyyymmdd") & "' "
                gstrQuy = gstrQuy & vbCrLf & " WHERE EXAMDATE = '" & Format(dtp�˻�����, "yyyymmdd") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPNO  = '" & gtypREG_INFO.EQUIPCD & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND BARCODE  = '" & Trim(GET_CELL(vasList, 3, intX)) & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND SEQNO    = '" & Trim(GET_CELL(vasList, 2, intX)) & "' "
                If RunSQL(gstrQuy) = False Then
                    ADC.RollbackTrans
                    Call CloseDB
                    Call ErrQuery(gstrQuy, 0)
                    Exit Sub
                End If
                
                For intY = 5 To vasList.MaxCols
                    If GET_CELL(vasList, intY, intX) = "1" Then
                        gstrQuy = "INSERT INTO PAT_RES "
                        gstrQuy = gstrQuy & vbCrLf & " (EXAMDATE,   EQUIPNO,    BARCODE,    receno,     pid, "
                        gstrQuy = gstrQuy & vbCrLf & "  pname,      pjumin,     page,       psex,       resdate, "
                        gstrQuy = gstrQuy & vbCrLf & "  EQUIPCODE,  examcode,   examtype,   result,     sendflag, "
                        gstrQuy = gstrQuy & vbCrLf & "  examname,   refflag,    panicflag,  deltaflag,  unit, "
                        gstrQuy = gstrQuy & vbCrLf & "  refvalue,   panicvalue, SEQNO,      diskno,     posno) "
                        gstrQuy = gstrQuy & vbCrLf & " VALUES "
'                        gstrQuy = gstrQuy & vbCrLf & " ('" & Format(CDate(frmInterface.dtpExamDate), "yyyymmdd") & "', " '/EXAMDATE
                        gstrQuy = gstrQuy & vbCrLf & " ('" & Format(CDate(dtp�˻�����), "yyyymmdd") & "', " '/EXAMDATE
                        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(gtypREG_INFO.EQUIPCD) & "', " '/EQUIPNO
                        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(GET_CELL(vasList, 3, intX)) & "', " '/BARCODE
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/receno
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/pid
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/pname
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/pjumin
                        gstrQuy = gstrQuy & vbCrLf & "  0, " '/page
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/psex
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/resdate
                        gstrQuy = gstrQuy & vbCrLf & "  '" & Format(intY - 4, "00") & "', " '/EQUIPCODE
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/examcode
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/examtype
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/result
                        gstrQuy = gstrQuy & vbCrLf & "  '0', " '/sendflag(0.W/S,1.�������, 2.���ۿϷ�)
                        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(GET_CELL(vasList, intY, 0)) & "', " '/examname
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/refflag
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/panicflag
                        gstrQuy = gstrQuy & vbCrLf & "  '', " '/deltaflag
                        gstrQuy = gstrQuy & vbCrLf & "  '', "     '/unit
                        gstrQuy = gstrQuy & vbCrLf & "  '', "      '/refvalue
                        gstrQuy = gstrQuy & vbCrLf & "  '', "    '/panicvalue
                        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(GET_CELL(vasList, 2, intX)) & "', "      '/SEQNO
                        gstrQuy = gstrQuy & vbCrLf & "  '', "       '/diskno
                        gstrQuy = gstrQuy & vbCrLf & "  '') "       '/posno
                        If RunSQL(gstrQuy) = False Then
                            ADC.RollbackTrans
                            Call CloseDB
                            Call ErrQuery(gstrQuy, 0)
                            Exit Sub
                        End If
                    End If
                Next intY
            End If
        Next intX
    End With
    
    ADC.CommitTrans
    
    Call CloseDB
'    Unload Me
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub dtp�˻�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    'dtpSDate.SetFocus
    vasActiveCell vasList, 1, 2
    
    With vasList
        If .MaxRows > 0 Then .MaxRows = 0
        
        .Col = 4
        .Row = -1
        
        strTemp = " " + Chr$(9)
        cboProfile.AddItem " "
        Dim lngRow  As Long
        
        If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub
    
        gstrQuy = "SELECT EQ_PROFILECD, EQ_PROFILENM "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_PROFILE "
        gstrQuy = gstrQuy & vbCrLf & " GROUP BY EQ_PROFILECD, EQ_PROFILENM "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY EQ_PROFILECD "
        If ReadSQL(gstrQuy, ADR) = False Then
            Call CloseDB
            Exit Sub
        End If
        
        If Not ADR Is Nothing Then
            Do Until ADR.EOF
                cboProfile.AddItem TEXT_LSET(Trim(ADR!EQ_PROFILECD & ""), 2) & "." & Trim(ADR!EQ_PROFILENM & "")
                
                strTemp = strTemp & TEXT_LSET(Trim(ADR!EQ_PROFILECD & ""), 2) & "." & Trim(ADR!EQ_PROFILENM & "") + Chr$(9)
                
                ADR.MoveNext
            Loop
            ADR.Close: Set ADR = Nothing
        End If
        Call CloseDB
    
        .TypeComboBoxList = strTemp
    End With
End Sub


Private Sub Form_Load()
'    dtpSDate.Text = Format(Date, "yyyy-mm-dd")
'    dtpEDate.Text = dtpSDate.Text
    
    '2010.03.16 �̻��� - ���糯¥�� ���� �ȵ�
    dtp�˻�����.Value = Format(Date, "YYYY-MM-DD")
    
    ClearSpread vasList
'    ClearSpread vasListTemp
    
    chkAll.Value = 0
'    cmdSearch_Click
End Sub

Private Sub txtBarcodeSeqF_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtBarcodeSeqF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtBarcodeSeqF_LostFocus()
    If IsNumeric(txtBarcodeSeqF) = True Then txtBarcodeSeqF = Format(txtBarcodeSeqF, "0000")
End Sub

Private Sub txtBarcodeSeqT_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtBarcodeSeqT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtBarcodeStartRow) = "" Then MsgBox "Start Row�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtBarcodeStartRow.SetFocus: Exit Sub
        If IsNumeric(txtBarcodeStartRow) = False Then MsgBox "Start Row�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtBarcodeStartRow.SetFocus: Exit Sub
        If Val(txtBarcodeStartRow) < 1 Then MsgBox "Start Row�� 1 �̻��̾�� �մϴ�!", vbCritical, "�������": txtBarcodeStartRow.SetFocus: Exit Sub
        
        
        
        If Trim(txtBarcodeSeqF) = "" Then MsgBox "���ڵ� Seq(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtBarcodeSeqF.SetFocus: Exit Sub
        If Trim(txtBarcodeSeqT) = "" Then MsgBox "���ڵ� Seq(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtBarcodeSeqT.SetFocus: Exit Sub
        If IsNumeric(txtBarcodeSeqF) = False Then MsgBox "���ڵ� Seq(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtBarcodeSeqF.SetFocus: Exit Sub
        If IsNumeric(txtBarcodeSeqT) = False Then MsgBox "���ڵ� Seq(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtBarcodeSeqT.SetFocus: Exit Sub
        
        If Val(txtBarcodeSeqF) > Val(txtBarcodeSeqT) Then
            MsgBox "���ڵ� Seq ������ (��)�Է��Ͻʽÿ�", vbCritical, "�������": txtBarcodeSeqF.SetFocus: Exit Sub
        End If
        
        '/////// ���ڵ� seq ������ Maxrows ���� ////////20110623 ȿ�� �߰�
        txtSampleNoCnt = CStr(Val(txtBarcodeSeqT) - Val(txtBarcodeSeqF) + 1)
        Call txtSampleNoCnt_KeyDown(vbKeyReturn, 0)
        
        If Val(txtBarcodeStartRow) > vasList.MaxRows Then Exit Sub
        
        intY = Val(txtBarcodeStartRow)
        
        For intX = Val(txtBarcodeSeqF) To Val(txtBarcodeSeqT)
            
            If intY > vasList.MaxRows Then Exit Sub
            vasList.Row = intY
            vasList.Col = 3: vasList.Text = gID_Par.BARCID & Format(dtp�˻�����, "yyyymmdd") & Format(intX, "0000")
            
            intY = intY + 1
        Next intX
        txtSampleNo = CStr(Val(GET_CELL(vasList, 2, vasList.MaxRows)) + 1)
        txtSampleNo = Format(txtSampleNo, "0##0")
        txtBarcodeStartRow.SetFocus
    End If
End Sub

Private Sub txtBarcodeSeqT_LostFocus()
    If IsNumeric(txtBarcodeSeqT) = True Then txtBarcodeSeqT = Format(txtBarcodeSeqT, "0000")
End Sub

Private Sub txtBarcodeStartRow_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtBarcodeStartRow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRowF_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtRowF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRowT_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtRowT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtRowF) = "" Then MsgBox "Row(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        If Trim(txtRowT) = "" Then MsgBox "Row(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
        If IsNumeric(txtRowF) = False Then MsgBox "Row(From)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        If IsNumeric(txtRowT) = False Then MsgBox "Row(To)�� ���������� �Է��ؾ� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
        
        If Val(txtRowF) < 1 Then MsgBox "Row(From)�� 1 Row �̻��̾�� �մϴ�!", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        If Val(txtRowT) < 1 Then MsgBox "Row(To)�� 1 Row �̻��̾�� �մϴ�!", vbCritical, "�������": txtRowT.SetFocus: Exit Sub
        
        If Val(txtRowF) > Val(txtRowT) Then
            MsgBox "Row ������ (��)�Է��Ͻʽÿ�", vbCritical, "�������": txtRowF.SetFocus: Exit Sub
        End If
        
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSampleNo_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtSampleNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtSampleNo) = False Then MsgBox "Sample No�� 0000~9999������ ���ڷ� �Է��Ͻʽÿ�!", vbCritical, "���Է�": txtSampleNo.SetFocus: Exit Sub
        
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSampleNo_LostFocus()
    txtSampleNo = Format(txtSampleNo, "0000")
End Sub

Private Sub txtSampleNoCnt_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtSampleNoCnt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtSampleNo) = False Then MsgBox "Sample No�� 0000~9999������ ���ڷ� �Է��Ͻʽÿ�!", vbCritical, "���Է�": txtSampleNo.SetFocus: Exit Sub
        If IsNumeric(txtSampleNoCnt) = False Then MsgBox "��ü���� ���ڷ� �Է��Ͻʽÿ�!", vbCritical, "���Է�": txtSampleNoCnt.SetFocus: Exit Sub

        For intX = 0 To Val(txtSampleNoCnt) - 1
            vasList.MaxRows = vasList.MaxRows + 1: vasList.Row = vasList.MaxRows
            
            vasList.Col = 2: vasList.Text = Format(Val(txtSampleNo) + intX, "0000")
            '''vasList.Col = 3: vasList.Text = "09" & Format(dtp�˻�����, "yyyymmdd") & Format(Val(txtSampleNo) + intX, "0000")
            
        Next intX
        
        SendKeys "{TAB}"
    End If
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 0 Then
        vasList.Value = 1
        Else
        vasList.Value = 0
        End If
    Next i
End Sub

Private Sub vasList_Change(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    If Col <> 4 Then Exit Sub
    
    If Trim(Left(GET_CELL(vasList, 4, Row), 2)) = "" Then
        vasList.ClearRange 5, Row, vasList.MaxCols, Row, -1
    End If
    
    For i = 5 To vasList.MaxCols
        vasList.Row = Row
        vasList.Col = i
        vasList.Text = "0"
    Next i
    
    Call SET_PROFILE(Row)
End Sub

Public Sub SET_PROFILE(argRow As Long)
'    ---------------------------------- �ӽú��� ����
    Dim TESTName As String
    Dim i As Integer
'    -----------------------------------
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT A.EQUIPCODE, B.EXAMNAME "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_PROFILE A, EQUIPEXAM B "
    gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQ_PROFILECD = '" & Trim(Left(GET_CELL(vasList, 4, argRow), 2)) & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPCODE = B.EQUIPCODE "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EXAMCODE = B.EXAMCODE "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY EQ_PROFILECD "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If
    
    If Not ADR Is Nothing Then
        
        Do Until ADR.EOF
            TESTName = Trim(ADR!examname)
            For i = 5 To vasList.MaxCols
                If Trim(GET_CELL(vasList, i, 0)) = TESTName Then
                    'vasList.Col = Val(ADR!EQUIPCODE & "") + 4
                    vasList.Col = i
                    Exit For
                End If
            Next i
            vasList.Row = argRow
            vasList.Text = "1"
            
            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
    End If
    Call CloseDB
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Integer
    Dim chkVal As String
    
    If Row = 0 And Col >= 5 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = Col
            If iRow = 1 Then
                If vasList.Value = 1 Then
                    chkVal = 0
                Else
                    chkVal = 1
                End If
            End If
            vasList.Value = chkVal
        Next iRow
    End If
End Sub
