VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDCU001 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�������� ����"
   ClientHeight    =   7035
   ClientLeft      =   1155
   ClientTop       =   1845
   ClientWidth     =   8430
   ForeColor       =   &H00FF0000&
   HelpContextID   =   41001
   Icon            =   "frmDCU001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.OptionButton optWorkFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��������"
      Height          =   255
      Index           =   1
      Left            =   3195
      TabIndex        =   10
      Top             =   570
      Width           =   1110
   End
   Begin VB.OptionButton optWorkFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��������"
      Height          =   255
      Index           =   0
      Left            =   3195
      TabIndex        =   9
      Top             =   255
      Value           =   -1  'True
      Width           =   1110
   End
   Begin VB.TextBox txtText 
      ForeColor       =   &H00000000&
      Height          =   4035
      Left            =   360
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   3
      Top             =   2280
      Width           =   7770
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����(&X)"
      Height          =   450
      Left            =   6930
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   6435
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����(&S)"
      Height          =   450
      Left            =   4380
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   6435
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear(&C)"
      Height          =   450
      Left            =   5655
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   6435
      Width           =   1215
   End
   Begin VB.Frame fraDept 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�����μ�"
      ForeColor       =   &H00C14F3E&
      Height          =   675
      Left            =   4545
      TabIndex        =   7
      Top             =   180
      Width           =   3570
      Begin VB.CheckBox chkBBS 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��������"
         ForeColor       =   &H00C14F3E&
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   300
         Width           =   1095
      End
      Begin VB.CheckBox chkAPS 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ܺ���"
         ForeColor       =   &H00C14F3E&
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   300
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkLIS 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�ӻ󺴸�"
         ForeColor       =   &H00C14F3E&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   360
      ScaleHeight     =   1155
      ScaleWidth      =   7695
      TabIndex        =   11
      Top             =   1020
      Width           =   7755
      Begin VB.TextBox txtTitle 
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "���� �̸��� �־� �ּ���."
         Top             =   135
         Width           =   5985
      End
      Begin VB.TextBox txtUsers 
         Appearance      =   0  '���
         BackColor       =   &H00DBE6E6&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1275
         MaxLength       =   30
         TabIndex        =   13
         ToolTipText     =   "���� �̸��� �־� �ּ���."
         Top             =   675
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker dtpLimitDay 
         Height          =   315
         Left            =   4935
         TabIndex        =   12
         Top             =   675
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Format          =   72613889
         CurrentDate     =   36833
      End
      Begin VB.Label Label2 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���� :"
         Height          =   315
         Left            =   435
         TabIndex        =   17
         Top             =   195
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�۾��� :"
         Height          =   255
         Left            =   435
         TabIndex        =   16
         Top             =   735
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�Խ� ������ :"
         Height          =   315
         Left            =   3675
         TabIndex        =   15
         Top             =   735
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�����ϱ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EE772F&
      Height          =   390
      Left            =   1005
      TabIndex        =   8
      Top             =   345
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   540
      Picture         =   "frmDCU001.frx":030A
      Top             =   300
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00808080&
      FillColor       =   &H00F7FDF8&
      FillStyle       =   0  '�ܻ�
      Height          =   630
      Index           =   0
      Left            =   360
      Shape           =   4  '�ձ� �簢��
      Top             =   180
      Width           =   2355
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  '�ܻ�
      Height          =   630
      Index           =   1
      Left            =   420
      Shape           =   4  '�ձ� �簢��
      Top             =   240
      Width           =   2355
   End
End
Attribute VB_Name = "frmDCU001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form��   : frmDCU001
'|  2. ��  ��   : �������� ����
'|  3. �ۼ���   : �� ����
'|  4. �ۼ���   : 2000.11.6
'|
'|  CopyRight(C) 2002 Pomis
'+--------------------------------------------------------------------------------------+
Option Explicit

Private ObjSql As clsDCUSqlStmt
'Private objEmp As clsBasisData
Private mvarProjectId As String 'APS, BBS, LIS ���θ� �޾ƿ��� ����
Private mvarEmpId As String     '���� ID�� �޾ƿ��� ����

Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property

Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

Public Property Let EmpId(ByVal vData As String)
    mvarEmpId = vData
End Property

Public Property Get EmpId() As String
    EmpId = mvarEmpId
End Property

Private Sub cmdClear_Click()
    '������...
    txtTitle.Text = ""
    txtText.Text = ""
    txtTitle.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'-------------------------------'
'   2002-08-06 ������ : �̻��
'-------------------------------'
Private Sub cmdSave_Click()
'    Dim Rs As New Recordset
    Dim strTmp As VbMsgBoxResult
    Dim strTitle As String
    Dim strNote As String
    Dim aryTmp() As String
    Dim aryTmp1() As String
    Dim lngSeq As Long
    
    'üũ����.
    If txtTitle.Text = "" Then
        MsgBox "������ �Է��Ͽ� �ּ���.", vbInformation, Me.Caption
        txtTitle.SetFocus
        Exit Sub
    ElseIf Format(dtpLimitDay.Value, "yyyymmdd") < Format(Now, "yyyymmdd") Then
        MsgBox "���� ���� ��¥�� ���� �� �� �����ϴ�.", vbInformation, Me.Caption
        dtpLimitDay.SetFocus
        Exit Sub
    ElseIf chkLIS.Value = 0 And chkAPS.Value = 0 And chkBBS.Value = 0 Then
        MsgBox "�ϳ��� �μ��� ���� �ּ���.", vbInformation, Me.Caption
        chkLIS.SetFocus
        Exit Sub
    ElseIf Len(txtText.Text) > 2000 Then
        MsgBox "�������� ������ �ٿ� �ּ���.", vbInformation, Me.Caption
        txtText.SetFocus
        Exit Sub
    End If
    
    aryTmp = Split(txtTitle, "'")
    aryTmp1 = Split(txtText, "'")
    strTitle = Join(aryTmp, "''")
    strNote = Join(aryTmp1, "''")
    Set ObjSql = New clsDCUSqlStmt
    
    '�Ϸù�ȣ�� ���� �´�.
'    Set Rs = ObjSql.Get_COM011
'    If Rs.EOF = False And IsNull(Rs.Fields(0).Value) = True Then
'        lngSeq = 1
'    Else
'        lngSeq = Val("" & Rs.Fields(0).Value) + 1
'    End If
'    Set Rs = Nothing
    
    lngSeq = ObjSql.GetCOM011MaxSeq
    
    '���忩�� Ȯ��...
    strTmp = MsgBox("�����Ͻðڽ��ϱ�?", vbInformation + vbOKCancel, Me.Caption)
    If strTmp = vbCancel Then
        Set ObjSql = Nothing
        Exit Sub
    Else
        ' ��������, �������� ����

       If optWorkFg(0).Value = True Then
           ObjSql.WorkFg = "0"
       ElseIf optWorkFg(1).Value = True Then
           ObjSql.WorkFg = "1"
       End If

        '����
        If ObjSql.Insert_COM011(Trim(lngSeq), Format(Now, "yyyymmdd"), Format(dtpLimitDay.Value, "yyyymmdd"), _
                            Trim(strTitle), strNote, txtUsers.Text, chkLIS.Value, chkAPS.Value, chkBBS.Value) = True Then
            MsgBox "���� �����Ͽ����ϴ�.", vbInformation, Me.Caption
        End If
    End If
    Set ObjSql = Nothing
    txtTitle.Text = ""
    txtText.Text = ""
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtTitle.SetFocus
End Sub

'-------------------------------'
'   2002-08-06 ������ : �̻��
'-------------------------------'
Private Sub Form_Load()
    Dim strTmpEmpNM As String
    
    '��������

    optWorkFg(0).Visible = True
    optWorkFg(1).Visible = True
    
    'APS, BBS, LIS ���θ� �޾ƿ´�.
    Select Case mvarProjectId
        Case "APS": chkAPS.Value = 1
        Case "BBS": chkBBS.Value = 1
        Case "LIS": chkLIS.Value = 1
    End Select
    
    Set ObjSql = New clsDCUSqlStmt
'    Set objEmp = New clsBasisData
    
'    strTmpEmpNM = ObjSql.GetHIS005EmpNm(mvarEmpId)
    
'    If strTmpEmpNM = "" Then strTmpEmpNM = GetEmpNm(mvarEmpId)   ' ObjSql.GetCOM006EmpNm(mvarEmpId)
    
    txtUsers.Text = GetEmpNm(mvarEmpId) 'strTmpEmpNM
    
    Set ObjSql = Nothing
'    Set objEmp = Nothing
    
    '��¥�� ��������...
    dtpLimitDay.Value = Format(DateAdd("d", 7, Date), "yyyy�� mm�� dd��")
    dtpLimitDay.CustomFormat = "yyyy�� MM�� dd��"
    dtpLimitDay.Format = dtpCustom

End Sub

'-------------------------------'
'   2002-08-06 �ۼ��� : �̻��
'-------------------------------'
Private Sub optWorkFg_Click(Index As Integer)
    If Index = 0 Then
        Label1.Caption = "�����ϱ�"
        frmDCU001.Caption = "�������� ����"
        Label2.Caption = "���� :"
        fraDept.Enabled = True
        chkLIS.Enabled = True
        chkAPS.Enabled = True
        chkBBS.Enabled = True
    ElseIf Index = 1 Then
        Label1.Caption = "��������"
        frmDCU001.Caption = "�������� ����"
        Label2.Caption = "�μ� :"
        fraDept.Enabled = False
        chkLIS.Enabled = False
        chkAPS.Enabled = False
        chkBBS.Enabled = False
    End If
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtText.SetFocus
End Sub
