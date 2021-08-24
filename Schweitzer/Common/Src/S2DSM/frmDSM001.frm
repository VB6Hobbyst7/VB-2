VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDSM001 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�� ����"
   ClientHeight    =   3945
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7815
   ForeColor       =   &H00DD6131&
   Icon            =   "frmDSM001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Begin MSComctlLib.ListView lvwForm 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "������ ���콺�� Ŭ���ϼ���."
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   6906
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776191
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�μ�"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��"
         Object.Width           =   2541
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�� �̸�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�� ����"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "�б� ���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "���� ���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "����Ʈ ���"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "���(&I)  "
   End
   Begin VB.Menu mnuUpdate 
      Caption         =   "���� (&E)"
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "����(&D)"
   End
   Begin VB.Menu mnuForm 
      Caption         =   "��ü��(&A)"
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "&Refresh  "
   End
   Begin VB.Menu mnuExit 
      Caption         =   "����(&X)"
   End
   Begin VB.Menu mnuPop 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdate1 
         Caption         =   "�� ��"
      End
      Begin VB.Menu mnuDelete1 
         Caption         =   "�� ��"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh1 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmDSM001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form��   : frmDSM001
'|  2. ��  ��   : �� List ���,����
'|  3. �ۼ���   : �� ����
'|  4. �ۼ���   : 2000.10.23
'|
'|  CopyRight(C) 2000 ��ÿ�Ƽ����
'+--------------------------------------------------------------------------------------+
Option Explicit

Private objSQL As clsDSMSqlStmt
Private mvarProjectId As String 'APS, BBS, LIS ���θ� �޾ƿ��� ����
Private strPID As String 'APS, BBS, LIS �޴� ����

Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property

Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

Private Sub Form_Load()
    'APS, BBS, LIS ���θ� �޾ƿ´�.
    Select Case mvarProjectId
        Case "APS": strPID = "A"
        Case "BBS": strPID = "B"
        Case "LIS": strPID = "L"
    End Select
    '��ġ�� ����....
    Me.Top = Me.Height / 2 + 300
    Me.Left = Me.Width / 2 - 1500
    '����Ʈ �並 display
    Set objSQL = New clsDSMSqlStmt
    Call objSQL.ShowListView(lvwForm, strPID)
    Set objSQL = Nothing
End Sub

Private Sub lvwForm_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'ColumnHeader Ŭ�� ��Ʈ
    Static intOrder As Integer
    
    With lvwForm
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '������ ���콺�� ������ �˾� �޴� �����.
    If Button = 2 Then
        frmDSM001.PopupMenu mnuPop
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Item As ListItem
    Dim strDept As String
    Dim strDept1 As String
    Dim strForm As String
    Dim strTmp As VbMsgBoxResult
    
    With lvwForm
        strDept = .ListItems(.SelectedItem.Index).Text
        If strDept = "LIS" Then
            strDept1 = "L"
        ElseIf strDept = "APS" Then
            strDept1 = "A"
        Else
            strDept1 = "B"
        End If
        strForm = .ListItems(.SelectedItem.Index).SubItems(1)
    End With
    '���� ���� Ȯ��...
    strTmp = MsgBox("�� ���� " & strForm & "�� ������ �����մϴ�.", vbInformation + vbOKCancel, Me.Caption)
    If strTmp = vbCancel Then
        Exit Sub
    Else
        Set objSQL = New clsDSMSqlStmt
        '����
        If objSQL.Del_COM007(strDept1, strForm) = True Then
            '����Ʈ �信�� ����
            Call lvwItem_Remove(strDept, strForm)
            MsgBox "�����Ͽ����ϴ�.", vbInformation, Me.Caption
        End If
        Set objSQL = Nothing
    End If
End Sub

Private Sub mnuDelete1_Click()
    'POPUP�޴�....
    Dim Item As ListItem
    Dim strDept As String
    Dim strDept1 As String
    Dim strForm As String
    Dim strTmp As VbMsgBoxResult
    
    With lvwForm
        strDept = .ListItems(.SelectedItem.Index).Text
        If strDept = "LIS" Then
            strDept1 = "L"
        ElseIf strDept = "APS" Then
            strDept1 = "A"
        Else
            strDept1 = "B"
        End If
        strForm = .ListItems(.SelectedItem.Index).SubItems(1)
    End With
    '���� ���� Ȯ��...
    strTmp = MsgBox("�� ���� " & strForm & "�� ������ �����մϴ�.", vbInformation + vbOKCancel, Me.Caption)
    If strTmp = vbOK Then
        Set objSQL = New clsDSMSqlStmt
        '����
        If objSQL.Del_COM007(strDept1, strForm) = True Then
            '����Ʈ �信�� ����
            Call lvwItem_Remove(strDept, strForm)
            MsgBox "�����Ͽ����ϴ�.", vbInformation, Me.Caption
        End If
        Set objSQL = Nothing
    Else
        Exit Sub
    End If
End Sub

Private Sub mnuExit_Click()
    Set objSQL = Nothing
    Unload Me
End Sub

Private Sub mnuForm_Click()
    Set objSQL = New clsDSMSqlStmt
    '��ü APS, BBS, LIS�� ����Ʈ �信 display
    Call objSQL.ShowListView_all(lvwForm)
    Set objSQL = Nothing
End Sub

Private Sub mnuInsert_Click()
    'APS, BBS, LIS ���θ� �޾ƿ´�.
    Select Case mvarProjectId
        Case "APS": frmDSM001P.optDept(1).Value = True
        Case "BBS": frmDSM001P.optDept(2).Value = True
        Case "LIS": frmDSM001P.optDept(0).Value = True
    End Select
    '����� �� display
    frmDSM001P.Show 1
End Sub

Private Sub mnuRefresh_Click()
    'update�� ����Ÿ display
    Select Case mvarProjectId
        Case "APS": strPID = "A"
        Case "BBS": strPID = "B"
        Case "LIS": strPID = "L"
    End Select
    '����Ʈ �信 �Ѹ���...
    Set objSQL = New clsDSMSqlStmt
    Call objSQL.ShowListView(lvwForm, strPID)
    Set objSQL = Nothing
End Sub

Private Sub mnuRefresh1_Click()
    'update�� ����Ÿ display
    Select Case mvarProjectId
        Case "APS": strPID = "A"
        Case "BBS": strPID = "B"
        Case "LIS": strPID = "L"
    End Select
    '����Ʈ �信 �Ѹ���...
    Set objSQL = New clsDSMSqlStmt
    Call objSQL.ShowListView(lvwForm, strPID)
    Set objSQL = Nothing
End Sub

Private Sub mnuUpdate_Click()
    'POPUP �޴�......
    Dim Item As ListItem
    Dim strDept As String
    Dim strRead As String
    Dim strWrite As String
    Dim strPrint As String
    
    '���õ� ����Ʈ�� ������ �ű���....
    With lvwForm
        strDept = .ListItems(.SelectedItem.Index).Text
        If strDept = "LIS" Then
            frmDSM001P.optDept(0).Value = True
        ElseIf strDept = "APS" Then
            frmDSM001P.optDept(1).Value = True
        Else
            frmDSM001P.optDept(2).Value = True
        End If
        frmDSM001P.txtForm = .ListItems(.SelectedItem.Index).SubItems(1)
        frmDSM001P.txtNm = .ListItems(.SelectedItem.Index).SubItems(2)
        frmDSM001P.txtDesc = .ListItems(.SelectedItem.Index).SubItems(3)
        strRead = .ListItems(.SelectedItem.Index).SubItems(4)
        strRead = .ListItems(.SelectedItem.Index).SubItems(4)
        frmDSM001P.chkRead.Value = IIf(strRead = "����", 1, 0)
        strWrite = .ListItems(.SelectedItem.Index).SubItems(5)
        frmDSM001P.chkWrite.Value = IIf(strWrite = "����", 1, 0)
        strPrint = .ListItems(.SelectedItem.Index).SubItems(6)
        frmDSM001P.chkPrint.Value = IIf(strPrint = "����", 1, 0)
    End With
    '������ �ʿ���� ���� ����.
    frmDSM001P.txtForm.Enabled = False
    frmDSM001P.txtForm.BackColor = &HEEEEEE
    frmDSM001P.optDept(0).Enabled = False
    frmDSM001P.optDept(1).Enabled = False
    frmDSM001P.optDept(2).Enabled = False
    '���� �� display
    frmDSM001P.Show 1
End Sub

Private Sub mnuUpdate1_Click()
    Dim Item As ListItem
    Dim strDept As String
    Dim strRead As String
    Dim strWrite As String
    Dim strPrint As String
    
    '���õ� ����Ʈ�� ������ �ű���....
    With lvwForm
        strDept = .ListItems(.SelectedItem.Index).Text
        If strDept = "LIS" Then
            frmDSM001P.optDept(0).Value = True
        ElseIf strDept = "APS" Then
            frmDSM001P.optDept(1).Value = True
        Else
            frmDSM001P.optDept(2).Value = True
        End If
        frmDSM001P.txtForm = .ListItems(.SelectedItem.Index).SubItems(1)
        frmDSM001P.txtNm = .ListItems(.SelectedItem.Index).SubItems(2)
        frmDSM001P.txtDesc = .ListItems(.SelectedItem.Index).SubItems(3)
        strRead = .ListItems(.SelectedItem.Index).SubItems(4)
        strRead = .ListItems(.SelectedItem.Index).SubItems(4)
        frmDSM001P.chkRead.Value = IIf(strRead = "����", 1, 0)
        strWrite = .ListItems(.SelectedItem.Index).SubItems(5)
        frmDSM001P.chkWrite.Value = IIf(strWrite = "����", 1, 0)
        strPrint = .ListItems(.SelectedItem.Index).SubItems(6)
        frmDSM001P.chkPrint.Value = IIf(strPrint = "����", 1, 0)
    End With
    '������ �ʿ���� ���� ����.
    frmDSM001P.txtForm.Enabled = False
    frmDSM001P.txtForm.BackColor = &HEEEEEE
    frmDSM001P.optDept(0).Enabled = False
    frmDSM001P.optDept(1).Enabled = False
    frmDSM001P.optDept(2).Enabled = False
    '���� �� display
    frmDSM001P.Show 1
End Sub

Private Sub lvwItem_Remove(ByVal Dept As String, ByVal FID As String)
    Dim itmX As Object
    Dim strForm As String
    Dim i As Long
    
    '����Ʈ �信�� ��� ����
    For i = 1 To lvwForm.ListItems.Count
        Set itmX = lvwForm.ListItems(i)
        strForm = itmX.SubItems(1)
        If itmX.Text = Dept And strForm = FID Then
            lvwForm.ListItems.Remove (i)
            itmX.EnsureVisible
            Exit For
        End If
    Next i
End Sub
