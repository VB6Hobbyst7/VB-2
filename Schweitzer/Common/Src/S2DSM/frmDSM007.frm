VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDSM007 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '���� ����
   Caption         =   "�ǻ� ���� ���"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   Icon            =   "frmDSM007.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11895
   StartUpPosition =   1  '������ ���
   Begin VB.CheckBox chkFireFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "������� ����"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Top             =   6945
      Width           =   2460
   End
   Begin MSComctlLib.ListView lvwEmpInformation 
      Height          =   3225
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5689
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776191
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ǻ� ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�ǻ� �̸�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "������ڵ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���� ����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "��ȭ ��ȣ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "�޴� ��ȭ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "�� ��"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Seq"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EEEBED&
      Caption         =   "Clear(&C)"
      Height          =   405
      Left            =   7185
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   6810
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00EEEBED&
      Caption         =   "����(&S)"
      Height          =   405
      Left            =   8355
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   6825
      Width           =   1050
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EEEBED&
      Caption         =   "����(&X)"
      Height          =   405
      Left            =   10695
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   6825
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00EEEBED&
      Caption         =   "����(&D)"
      Height          =   405
      Left            =   9525
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   6825
      Width           =   1050
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   3225
      Left            =   120
      ScaleHeight     =   3165
      ScaleWidth      =   11595
      TabIndex        =   6
      Top             =   3480
      Width           =   11655
      Begin VB.TextBox txtDeptNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   270
         Left            =   1875
         MaxLength       =   4
         TabIndex        =   26
         Text            =   "�μ��ڵ�"
         Top             =   1170
         Width           =   1290
      End
      Begin VB.TextBox txtCellNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   9930
         MaxLength       =   20
         TabIndex        =   15
         Text            =   "�޴���ȭ"
         Top             =   465
         Width           =   1545
      End
      Begin VB.TextBox txtEmpLngNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   5385
         MaxLength       =   20
         TabIndex        =   14
         Text            =   "�����̸���"
         Top             =   465
         Width           =   1545
      End
      Begin VB.TextBox txtTelNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   9930
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "��ȭ��ȣ"
         Top             =   885
         Width           =   1545
      End
      Begin VB.TextBox txtRemark 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   525
         Left            =   1875
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmDSM007.frx":06EA
         Top             =   1695
         Width           =   9630
      End
      Begin VB.CommandButton cmdCodeHelp 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3195
         TabIndex        =   11
         Tag             =   "Dept"
         Top             =   870
         Width           =   240
      End
      Begin VB.TextBox txtDeptCd 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   270
         Left            =   1875
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "�μ��ڵ�"
         Top             =   885
         Width           =   1290
      End
      Begin VB.TextBox txtEmpID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   1875
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "�����Ƶ�"
         Top             =   465
         Width           =   1545
      End
      Begin VB.TextBox txtEntDt 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   5370
         MaxLength       =   20
         TabIndex        =   8
         Text            =   "��������"
         Top             =   885
         Width           =   1545
      End
      Begin VB.CheckBox chkFireDt 
         BackColor       =   &H00EEEBED&
         Height          =   180
         Left            =   9930
         TabIndex        =   7
         Top             =   1365
         Width           =   195
      End
      Begin MSComCtl2.DTPicker dtpFireDt 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   10140
         TabIndex        =   16
         Top             =   1305
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16776191
         Format          =   104464385
         CurrentDate     =   36819
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� ���� ���"
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
         Left            =   135
         TabIndex        =   25
         Top             =   90
         Width           =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   180
         X2              =   11445
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "��� ����        : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   8205
         TabIndex        =   24
         Tag             =   "105"
         Top             =   1365
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ȭ ��ȣ        : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   8205
         TabIndex        =   23
         Tag             =   "105"
         Top             =   525
         Width           =   1710
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "����� �ڵ�      : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Tag             =   "105"
         Top             =   945
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�ǻ� I D         : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   135
         TabIndex        =   21
         Tag             =   "105"
         Top             =   525
         Width           =   1710
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�ǻ� ����        :"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   3645
         TabIndex        =   20
         Tag             =   "105"
         Top             =   525
         Width           =   1620
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�޴� ��ȭ        : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   8205
         TabIndex        =   19
         Tag             =   "105"
         Top             =   945
         Width           =   1710
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "��     ��        : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   135
         TabIndex        =   18
         Tag             =   "105"
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label lblEntDt 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� ����        : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   3645
         TabIndex        =   17
         Tag             =   "105"
         Top             =   945
         Width           =   1710
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "����Ʈ"
      Visible         =   0   'False
      Begin VB.Menu mnuListDel 
         Caption         =   "����"
      End
      Begin VB.Menu mnuListExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "��"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "����"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "frmDSM007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Coding By Legends
'Coding Date 2k/10
'���� ������ ���

'--------------------------------------------------
'���콺�� ������ ������ ���ϵ��� �ϴ� API
'���ε忡�� ���콺�� �����ϰ� ����ε忡�� ����
'�����ο��� �޽��� �ڽ����� ���� ����� �ٽ� ����.

Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetWindowRect Lib "user32" _
        (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim r As RECT
Dim X As Long
Dim deskhWnd As Long
'���콺�� ������ ������ ���ϵ��� �ϴ� API
'--------------------------------------------------

Private objMySql As New clsDSMSqlStmt

Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private Sub chkFireDt_Click()
    dtpFireDt.Enabled = (chkFireDt = "1")
End Sub

Private Sub chkFireFg_Click()

    Call ClearText
    Call objMySql.ShowEmpListView(lvwEmpInformation, chkFireFg.Value)
    Call lvwEmpInformation_ColumnClick(lvwEmpInformation.ColumnHeaders.Item(1))

End Sub

Private Sub cmdClear_Click()
    Call ClearText
    txtEmpID.SetFocus
End Sub

Private Sub cmdCodeHelp_Click()
    
    Dim strSQL As String
    
    strSQL = objMySql.Query(1)
        '
    Set objMyList = New clsPopUpList
    
    Call LockForm
    objMyList.Connection = DBConn
    Call objMyList.LoadPopUp(strSQL) ', 5000, 9000)
    
    txtDeptCd.Text = objMyList.SelectedItems(0)
    txtDeptNm.Text = objMyList.SelectedItems(1)
    
    Call LockForm
    Set objMySql = Nothing
End Sub

Private Sub cmdDelete_Click()
    Dim strMsg As VbMsgBoxResult
    
    If txtEmpID = "" Then
        MsgBox "������ �ǻ縦 ����Ʈ���� �����ϼ���.", vbInformation, "����Ȯ��"
        Call LockForm
        Exit Sub
    End If
    strMsg = MsgBox("'" & txtEmpLngNm & "' �� ���������� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "��������")
    
    If strMsg = vbNo Then Exit Sub
    
    Call objMySql.DelCOM098(Trim(txtEmpID))
    
    Call objMySql.ShowDoctListView(lvwEmpInformation, chkFireFg.Value)
    Call lvwEmpInformation_ColumnClick(lvwEmpInformation.ColumnHeaders.Item(1))
    Call ClearText
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmDSM007 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim blnUpdateFg As Boolean         '������Ʈ üũ�÷���
    Dim strSex As String            '����
    Dim strDegree As String         '����
    Dim strShiftCd As String        '�ٹ�����
    Dim strMsg As VbMsgBoxResult    '�޽���
    Dim blnFireCheck As Boolean        '���üũ
    Dim strSQL As String
    
    If txtEmpID.Text = "" Then
        MsgBox "�ǻ� ���̵� �Է��ϼ���."
        Call LockForm
        txtEmpID.SetFocus
        Exit Sub
    End If
           
    strSQL = objMySql.Query(7) & " where empno = " & Trim(txtEmpID.Text)
    blnUpdateFg = objMySql.UpdateCheck(Trim(txtEmpID), , , strSQL)
    blnFireCheck = objMySql.FireCheck(chkFireDt)
          
    Call objMySql.SetCOM098(blnUpdateFg, Trim(txtEmpID), Trim(txtEmpLngNm), Trim(txtDeptCd), Trim(txtDeptNm), Now, Trim(txtTelNo), Trim(txtCellNo), Trim(txtRemark), blnFireCheck)
    
    Call ClearText
    Call objMySql.ShowDoctListView(lvwEmpInformation, chkFireFg.Value)
    Call lvwEmpInformation_ColumnClick(lvwEmpInformation.ColumnHeaders.Item(1))
    
End Sub

Private Sub Form_Load()
        
    Me.Show
'    Call LockForm
    Call ClearText
'    Call LockForm
    DoEvents
    
    Call objMySql.ShowDoctListView(lvwEmpInformation, chkFireFg.Value)
    Call lvwEmpInformation_ColumnClick(lvwEmpInformation.ColumnHeaders.Item(1))
    DoEvents
'    Call LockForm
End Sub

Private Sub ClearText()
    txtEmpID = ""
    txtEmpLngNm = ""
    txtDeptCd = ""
    txtDeptNm = ""

    dtpFireDt.Value = GetSystemDate
    dtpFireDt.Enabled = False


    txtTelNo = ""
    txtCellNo = ""
    txtRemark = ""
    lblEntDt.Visible = False
    txtEntDt.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���콺 ������ ��ư�� Ŭ��������
        If Button = 2 Then frmDSM002.PopupMenu mnuForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnLockForm
    Set frmDSM007 = Nothing
End Sub

Private Sub lvwEmpInformation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'��Ʈ
    Static intOrder As Integer
    
    With lvwEmpInformation
        If ColumnHeader.Index = 1 Then
            .SortKey = 7
        Else
            .SortKey = ColumnHeader.Index - 1
        End If
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        
        intOrder = (intOrder + 1) Mod 2
        
        If ColumnHeader.Index <> 11 Then
            .SortKey = 7
            .SortOrder = lvwAscending
            .Sorted = True
        End If
    End With

End Sub

Private Sub lvwEmpInformation_ItemClick(ByVal Item As MSComctlLib.ListItem)
'ȭ�鿡 �ѷ��ִ�~
    With lvwEmpInformation
        txtEmpID = .ListItems(.SelectedItem.Index).Text
        txtEmpLngNm = Item.SubItems(1)
        txtDeptCd = Item.SubItems(2)
        txtDeptNm = Item.SubItems(3)
        lblEntDt.Visible = True
        txtEntDt.Visible = True
        txtEntDt = Item.SubItems(4)
        txtTelNo = Item.SubItems(5)
        'txtCellNo = Item.SubItems(14)
        txtRemark = Item.SubItems(6)
    End With

End Sub

Private Sub lvwEmpInformation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���콺 ������ ��ư�� Ŭ��������
    If Button = 2 Then frmDSM002.PopupMenu mnuList
End Sub

Private Sub mnuClear_Click()
    Call cmdClear_Click
End Sub

Private Sub mnuDel_Click()
    Call cmdDelete_Click
End Sub

Private Sub mnuExit_Click()
    Call cmdExit_Click
End Sub

Private Sub mnuListDel_Click()
    Call cmdDelete_Click
End Sub

Private Sub mnuListExit_Click()
    Call cmdExit_Click
End Sub

Private Sub mnuSave_Click()
    Call cmdSave_Click
End Sub

Private Sub LockForm()
    ' This code confines the cursor to the inside of frmAPS208
    '--------------------------------------------------
'    X = GetWindowRect(Me.hwnd, r)  ' API puts window coords into RECT
'    X = ClipCursor(r)  ' Confine the cursor
    '--------------------------------------------------
End Sub

Private Sub UnLockForm()
    ' This code releases the cursor
    '--------------------------------------------------
'    deskhWnd = GetDesktopWindow()  ' API gets desktop's handle
'    X = GetWindowRect(deskhWnd, r)  ' API puts window coords into RECT
'    X = ClipCursor(r)  ' "Confine" the cursor to the entire screen.
    '--------------------------------------------------
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then frmDSM002.PopupMenu mnuForm
End Sub
