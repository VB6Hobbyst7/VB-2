VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmDSM005 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "����� ���"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmDSM005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00EBF3ED&
      Caption         =   "Ȯ ��(&O)"
      Height          =   510
      Left            =   1815
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   4110
      Width           =   1095
   End
   Begin VB.CommandButton cmdCacel 
      BackColor       =   &H00EBF3ED&
      Caption         =   "�� ��(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   3045
      Style           =   1  '�׷���
      TabIndex        =   7
      Top             =   4110
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   3840
      Left            =   150
      ScaleHeight     =   3780
      ScaleWidth      =   5565
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   5625
      Begin VB.TextBox txtEmpId 
         Height          =   330
         Left            =   1605
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   585
         Width           =   1215
      End
      Begin VB.TextBox txtLoginId 
         Height          =   345
         Left            =   1605
         TabIndex        =   0
         Top             =   165
         Width           =   3705
      End
      Begin VB.TextBox txtLoginPass 
         Height          =   345
         IMEMode         =   3  '��� ����
         Left            =   1605
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1470
         Width           =   3705
      End
      Begin VB.ComboBox cboGroupId 
         Height          =   300
         Left            =   1605
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   4
         Top             =   2355
         Width           =   3690
      End
      Begin VB.TextBox txtLoginDesc 
         Height          =   345
         Left            =   1605
         TabIndex        =   5
         Top             =   2760
         Width           =   3705
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00EBF3ED&
         Caption         =   "�� �׷� ���"
         Height          =   405
         Left            =   4005
         Style           =   1  '�׷���
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3210
         Width           =   1305
      End
      Begin VB.CommandButton cmdList 
         BackColor       =   &H00EBF3ED&
         Caption         =   "..."
         Height          =   330
         Left            =   2835
         Style           =   1  '�׷���
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   570
         Width           =   330
      End
      Begin VB.TextBox txtRePass 
         Height          =   345
         IMEMode         =   3  '��� ����
         Left            =   1605
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1905
         Width           =   3705
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   345
         Left            =   135
         TabIndex        =   11
         Top             =   150
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "LogIn     ID :"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   345
         Left            =   135
         TabIndex        =   12
         Top             =   585
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "��         �� :"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   345
         Left            =   135
         TabIndex        =   13
         Top             =   1020
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "��         �� :"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Left            =   135
         TabIndex        =   14
         Top             =   1455
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "���  ��ȣ :"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   345
         Left            =   135
         TabIndex        =   15
         Top             =   1890
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "��й�ȣȮ�� :"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   345
         Left            =   135
         TabIndex        =   16
         Top             =   2325
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "��         �� :"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   345
         Left            =   135
         TabIndex        =   17
         Top             =   2760
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "��         �� :"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEmpNm 
         Height          =   330
         Left            =   3210
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   570
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
         BackColor       =   15463405
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDept 
         Height          =   330
         Left            =   1605
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1020
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   582
         BackColor       =   15463405
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmDSM005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarEditFg As Boolean
Private mvarLogInID As String
Private mvarChangePwd As Boolean
'LoginId , EmpId, Dept, passwd, passwd, groupcd, Desc

Private mvarEmpID As String
Private mvarDeptCd As String
Private mvarLogInPass As String
Private mvarGroupID As String
Private mvarLoginDesc As String

Public Property Let EditFg(ByVal vData As Boolean)
    mvarEditFg = vData
End Property

Public Property Get EditFg() As Boolean
    EditFg = mvarEditFg
End Property

Public Property Let LoginId(ByVal vData As String)
    mvarLogInID = vData
End Property

Public Property Get LoginId() As String
    LoginId = mvarLogInID
End Property

Public Property Let EmpId(ByVal vData As String)
    mvarEmpID = vData
End Property

Public Property Get EmpId() As String
    EmpId = mvarEmpID
End Property

Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

Public Property Let LogInPass(ByVal vData As String)
    mvarLogInPass = vData
End Property

Public Property Get LogInPass() As String
    LogInPass = mvarLogInPass
End Property

Public Property Let GroupID(ByVal vData As String)
    mvarGroupID = vData
End Property

Public Property Get GroupID() As String
    GroupID = mvarGroupID
End Property

Public Property Let LoginDesc(ByVal vData As String)
    mvarLoginDesc = vData
End Property

Public Property Get LoginDesc() As String
    LoginDesc = mvarLoginDesc
End Property

'Password����
Public Property Let ChangePwd(ByVal vData As Boolean)
    mvarChangePwd = vData
End Property

Public Property Get ChangePwd() As Boolean
    ChangePwd = mvarChangePwd
End Property

Private Sub cboGroupId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdCacel_Click()
    Unload Me
End Sub

Private Sub cmdList_Click()
'    frmDSM005P.Show vbModal, Me
    Dim objEmpList As clsPopUpList
    
    Set objEmpList = New clsPopUpList
    
    With objEmpList
        .Connection = DBConn
        .ColumnHeaderText = "���̵�;����ڸ�;�μ�"
        .ColumnHeaderWidth = "1440;1440;0"
        .FormWidth = 3330
        .SortColumn = 2
        .SqlStmt = " select empid, empnm, deptcd from " & T_COM006 & " order by empid, empnm "
                
        .LoadPopUp
        
        txtEmpId.Text = .SelectedItems(0)
        lblEmpNm.Caption = .SelectedItems(1)
        lblDept.Caption = .SelectedItems(2)
    End With
    
    Set objEmpList = Nothing
End Sub

Private Sub cmdNew_Click()
'    frmDSM003.Show vbModal, Me
    frmDSM003_N.Show vbModal, Me
End Sub

Private Sub cmdOK_Click()
    If CheckValidation = False Then Exit Sub
    
    Call SaveUser
End Sub

Private Function SaveUser() As Boolean
'����� ������ �����ϰ� �ٽ� �ν�Ʈ
    Dim strSQL As String
        
    On Error GoTo ErrTrap
    DBConn.BeginTrans
    strSQL = " delete " & T_COM010 & _
           " where " & DBW("loginid=", txtLoginId.Text)
    DBConn.Execute strSQL
                   
    strSQL = " insert into " & T_COM010 & " (loginid, loginpass,empid, logindesc, groupid) values (" & _
           DBV("loginid", txtLoginId.Text, 1) & DBV("loginpass", txtLoginPass.Text, 1) & _
           DBV("empid", txtEmpId.Text, 1) & DBV("logindesc", txtLoginDesc.Text, 1) & DBV("groupid", medGetP(cboGroupId.Text, 1, COL_DIV)) & ")"
    DBConn.Execute strSQL
    
    DBConn.CommitTrans
    
    MsgBox "���������� ó���Ǿ����ϴ�.", vbInformation
    Unload Me
    
    Exit Function
ErrTrap:
    DBConn.RollbackTrans
    MsgBox "ó������ ������ �߻��Ͽ����ϴ�.", vbExclamation
End Function

Private Function CheckValidation() As Boolean
    CheckValidation = False
    If txtLoginId.Text = "" Then
        MsgBox "�α��� ���̵� �Է��Ͻʽÿ�.", vbExclamation
        txtLoginId.SetFocus
        Exit Function
    End If
    
    If txtEmpId.Text = "" Then
        MsgBox "���� ���̵� �Է��Ͻʽÿ�.", vbExclamation
        txtEmpId.SetFocus
        Exit Function
    End If
    
    If txtLoginPass.Text = "" Then
        MsgBox "��й�ȣ�� �Է��Ͻʽÿ�.", vbExclamation
        txtLoginPass.SetFocus
        Exit Function
    End If
    
    If txtRePass.Text = "" Then
        MsgBox "��й�ȣ�� Ȯ���� �Է��Ͻʽÿ�.", vbExclamation
        txtRePass.SetFocus
        Exit Function
    End If
    
    If cboGroupId.ListIndex < 0 Then
        MsgBox "����� �׷��� �����Ͻʽÿ�.", vbExclamation
        cboGroupId.SetFocus
        Exit Function
    End If
    
    '�����ڴ� ��� �׷쿡 ���Ե� �� �ְ�
    'Supervisior�� �Ŵ��� ���� �׷��� ������ ���ְ�
    'manager�� end user�׷��� ������ �� �ֵ��� �Ѵ�.
    'D,S,M,E
'ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Or ObjMyUser.IsSupervisor
    If ObjMyUser.IsManager Then
        If (medGetP(cboGroupId.Text, 3, COL_DIV) = "D") Or (medGetP(cboGroupId.Text, 3, COL_DIV) = "S") Then
            MsgBox "������ �׷��� ����� ������ ������մϴ�.", vbExclamation
            cboGroupId.SetFocus
            Exit Function
        End If
    ElseIf ObjMyUser.IsSupervisor Then
        If medGetP(cboGroupId.Text, 3, COL_DIV) = "D" Then
            MsgBox "������ �׷��� ����� ������ ������մϴ�.", vbExclamation
            cboGroupId.SetFocus
            Exit Function
        End If
    End If
    
    CheckValidation = True
End Function

Private Sub Form_Load()
    Call InitForm
    Call LoadGroup
    
    If mvarEditFg Then
        txtLoginId.Text = mvarLogInID
        txtEmpId.Text = mvarEmpID
        lblEmpNm.Caption = GetEmpNm(mvarEmpID)
        lblDept.Caption = GetDeptNm(mvarDeptCd)
        txtLoginPass.Text = mvarLogInPass
        txtRePass.Text = mvarLogInPass
        cboGroupId.ListIndex = medComboFind(cboGroupId, mvarGroupID)
        txtLoginDesc.Text = mvarLoginDesc
        txtLoginId.Enabled = False
        txtEmpId.Enabled = False
        cmdList.Enabled = False
    Else
        txtLoginId.Enabled = True
        txtEmpId.Enabled = True
        cmdList.Enabled = True
    End If
    
    If mvarChangePwd Then
        '����� �⺻���� ��ȸ
        txtLoginId.Text = mvarLogInID
        Call CheckExist
        txtLoginId.Enabled = False
        txtEmpId.Enabled = False
        cmdList.Enabled = False
    End If
    
    '�����ڸ��� ��й�ȣ�� Ȯ���� �� �ִ�.
    'Append By legends 2003/09/29
    
    If ObjMyUser.IsDeveloper Then
        txtLoginPass.PasswordChar = ""
        txtRePass.PasswordChar = ""
    Else
        txtLoginPass.PasswordChar = "*"
        txtRePass.PasswordChar = "*"
    End If
End Sub

Private Sub InitForm()
    txtLoginId.Text = ""
    txtEmpId.Text = ""
    lblEmpNm.Caption = ""
    lblDept.Caption = ""
    txtLoginPass.Text = ""
    txtRePass.Text = ""
    cboGroupId.clear
    txtLoginDesc.Text = ""
End Sub

Private Sub LoadGroup()
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_COM008 & " order by groupid, groupnm "
    
    Set Rs = New Recordset
    
    Rs.Open strSQL, DBConn
    
    cboGroupId.clear
    Do Until Rs.EOF
        cboGroupId.AddItem Rs.Fields("groupid").Value & "" & COL_DIV & Rs.Fields("groupnm").Value & "" & Space(100) & COL_DIV & Rs.Fields("userfg").Value & ""
            
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
End Sub

'Private Sub UserSet()
'    With clsRef
'         txtID = .LoginId
'         LislblEmpID(0).Caption = .EmpId
'         LislblEmpID(1).Caption = .EmpNm
'         txtPass = .LoginPass
'         txtRePass = .LoginPass
'         .ComBo_List cboGroup, 1
'         .Set_ComboList cboGroup, .GroupNm
'         cboGroupId.ListIndex = cboGroup.ListIndex
'         txtDesc = .LoginDesc
'    End With
'End Sub

'Private Sub Make_Combo()
'    With clsRef
'         .ComBo_List cboGroupId, 0
'         .ComBo_List cboGroup, 1
'    End With
'End Sub

'Private Function User_Insert() As Boolean
'    If txtID = "" Then
'       MsgBox "�ʼ� �Է��Դϴ�.", vbExclamation, "�������"
'       txtID.SetFocus
'       User_Insert = False
'       Exit Function
'    Else
'       User_Insert = True
'    End If
'
'    clsRef.COM010_Insert bstrLogInID, bstrLogInPass, bstrEmpID, _
'                         bstrLogInDesc, bstrGroupID
'
'End Function

'Private Function User_UpDate() As Boolean
'    If txtID = "" Then
'       MsgBox "�ʼ� �Է��Դϴ�.", vbExclamation, "��������"
'       txtID.SetFocus
'       User_UpDate = False
'       Exit Function
'    Else
'       User_UpDate = True
'    End If
'
'    clsRef.COM010_UpDate bstrLogInID, bstrLogInPass, bstrEmpID, _
'                         bstrLogInDesc, bstrGroupID
'End Function

Private Sub txtEmpId_Change()
    If lblEmpNm.Caption <> "" Then
        lblEmpNm.Caption = ""
    End If
End Sub

Private Sub txtEmpId_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtEmpId_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtEmpId.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEmpId_Validate(Cancel As Boolean)
    Dim strEmpNm As String
    
    If txtEmpId.Text = "" Then Exit Sub
    
    strEmpNm = GetEmpNm(txtEmpId.Text)
    
    If strEmpNm = "" Then
        Cancel = True
        MsgBox "��ϵ��� ���� ������Դϴ�.", vbExclamation
    Else
        lblEmpNm.Caption = strEmpNm
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtLoginDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtLoginId_Change()
    If txtEmpId.Text <> "" Then
        txtEmpId.Text = ""
        lblEmpNm.Caption = ""
        lblDept.Caption = ""
        txtLoginPass.Text = ""
        txtRePass.Text = ""
        cboGroupId.ListIndex = -1
        txtLoginDesc.Text = ""
    End If
End Sub

Private Sub txtLoginId_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtLoginId_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtLoginId.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtLoginId_Validate(Cancel As Boolean)
'������ ���ǰ� �ִ��� ���� üũ�ϰ� �ִ� ��쿡�� �ش� ������ ǥ�����ش�.
    If txtLoginId.Text = "" Then
        Cancel = True
        MsgBox "�α��� ���̵� �Է��Ͻʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    Call CheckExist
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function CheckExist() As Boolean
    Dim Rs As Recordset
    Dim strSQL As String
    
    CheckExist = False
    
    strSQL = " select b.loginid, a.empid, a.empnm, a.deptcd, b.loginpass, b.groupid, b.logindesc " & _
            " from " & T_COM006 & " a, " & T_COM010 & " b " & _
            " where a.empId = b.empId " & _
            " and " & DBW("b.loginid=", txtLoginId.Text)
            
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF = False Then
        CheckExist = True
        
        txtEmpId.Text = Rs.Fields("empid").Value & ""
        lblEmpNm.Caption = Rs.Fields("empnm").Value & ""
        lblDept.Caption = Rs.Fields("deptcd").Value & ""
        txtLoginPass.Text = Rs.Fields("loginpass").Value & ""
        txtRePass.Text = Rs.Fields("loginpass").Value & ""
        cboGroupId.ListIndex = medComboFind(cboGroupId, Rs.Fields("groupid").Value & "")
        txtLoginDesc.Text = Rs.Fields("logindesc").Value & ""
    End If
End Function

Private Sub txtLoginPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRePass_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRePass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRePass_Validate(Cancel As Boolean)
    If txtLoginPass.Text <> "" And txtRePass.Text <> "" Then
        If txtLoginPass.Text <> txtRePass.Text Then
            Cancel = True
            MsgBox "��й�ȣ�� ���� �ٸ��ϴ�.", vbExclamation
        End If
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub
