VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDCU002 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "��������"
   ClientHeight    =   6180
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   7845
   Icon            =   "frmDCU002.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ȯ��(&O)"
      Default         =   -1  'True
      Height          =   435
      Left            =   6315
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   5625
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���� ��������(&N)"
      Height          =   435
      Left            =   4590
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   5625
      Width           =   1695
   End
   Begin VB.CheckBox chkLoadAtStartup 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���� �� ǥ�� ����(&S)"
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   5700
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5340
      Left            =   240
      Picture         =   "frmDCU002.frx":08CA
      ScaleHeight     =   5280
      ScaleWidth      =   7245
      TabIndex        =   0
      Top             =   225
      Width           =   7305
      Begin VB.OptionButton optWorkFg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������"
         Height          =   255
         Index           =   0
         Left            =   4815
         TabIndex        =   9
         Top             =   135
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.OptionButton optWorkFg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   8
         Top             =   135
         Visible         =   0   'False
         Width           =   1125
      End
      Begin RichTextLib.RichTextBox Rtxt 
         Height          =   4440
         Left            =   165
         TabIndex        =   6
         Top             =   810
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   7832
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmDCU002.frx":0BD4
      End
      Begin VB.TextBox txtText 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4440
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   1
         Top             =   810
         Width           =   7005
      End
      Begin VB.Label lblUsers 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   540
         TabIndex        =   7
         Top             =   495
         Width           =   945
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�˸��ϴ�.."
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
         Left            =   540
         TabIndex        =   2
         Top             =   180
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmDCU002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form��   : frmDCU002
'|  2. ��  ��   : �������� �˸�
'|  3. �ۼ���   : �� ����
'|  4. �ۼ���   : 2000.11.6
'|
'|  CopyRight(C) 2002 Pomis
'+--------------------------------------------------------------------------------------+
Option Explicit

Private ObjSql As clsDCUSqlStmt
Private mvarProjectId As String 'APS, BBS, LIS ���θ� �޾ƿ��� ����
Private mvarTradeMark As String '
Private strPID As String        'APS, BBS, LIS �޴� ����
Private mvarAppName As String
Private mvarDataExists As Boolean
Private MesgKey As New Collection   '�޸𸮿� �ִ� �����ͺ��̽�
Private CurrentTip As Long          '���� ǥ�õ� �÷����� �ε���

Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property

Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

Public Property Let TradeMark(ByVal vData As String)
    mvarTradeMark = vData
End Property

Public Property Get TradeMark() As String
    TradeMark = mvarTradeMark
End Property

Public Property Get DataExists() As Boolean
    DataExists = mvarDataExists
End Property

Private Sub cmdNext_Click()
    Dim Rs As Recordset
    Dim strListString As String
    Dim strDt As String
    Dim strSeq As String
    Dim strValue As String
    Dim strTmp   As String
    Dim lngLoc   As Long
    Dim Deli     As String
    
    Deli = "YMHGR"
    CurrentTip = CurrentTip + 1
    
    '�ǳ��̸� ó������...�����ϴ�.
    If MesgKey.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    '�޸𸮿� ����� ������ ���� ����.
    strListString = MesgKey.Item(CurrentTip)
    strDt = medGetP(strListString, 1, vbTab)
    strSeq = medGetP(strListString, 2, vbTab)
    
    '����� ������ ���� ����.
    Set ObjSql = New clsDCUSqlStmt
    
    Set Rs = New Recordset
    
    On Error GoTo ErrTrap
    
    Rs.Open ObjSql.GetSQLCOM011(strSeq, strDt), DBConn
    
    '����� ���� üũ...
    If Rs.EOF Then
        Set Rs = Nothing
        Set ObjSql = Nothing
        Exit Sub
    End If
    
    'ȭ�鿡 ��Ÿ����.
    lblTitle.Caption = "" & Rs.Fields("Title").Value
    lblUsers.Caption = "�� �� �� ��  : " & Rs.Fields("users").Value & "" & "   �� �� �� �� : " & Format(Rs.Fields("inputday").Value & "", "####��##��##��")

    txtText.Text = "�� �� �� ��  : " & Rs.Fields("users").Value & "" & vbNewLine & _
                   "�� �� �� �� : " & Format(Rs.Fields("inputday").Value & "", "####��##��##��") & vbNewLine & vbNewLine & vbNewLine & _
                                    Rs.Fields("note").Value & ""
    
ErrTrap:
    Set Rs = Nothing
    Set ObjSql = Nothing
End Sub

Private Sub cmdOK_Click()
    Set MesgKey = Nothing
    Unload Me
    Set frmDCU002 = Nothing
End Sub

'-------------------------------'
'   2002-08-06 ������ : �̻��
'-------------------------------'
Private Sub Form_Activate()
    Dim Rs As Recordset
    Dim strMsg As String
    Dim i As Long
    Dim strValue As String
    Dim strTmp   As String
    Dim lngLoc   As Long
    Dim Deli     As String
    
    Deli = "YMHGR"
    
    Me.Visible = False
    
    '�������� ������ ��������...
    Set ObjSql = New clsDCUSqlStmt
    ObjSql.WorkFg = "0"
    
    Set Rs = New Recordset
    
    On Error GoTo ErrTrap
    
    Rs.Open ObjSql.GetSQLCOM011ByDeptFg(mvarProjectId), DBConn
    
    '�������� üũ.
    If Rs.EOF Then
        Set Rs = Nothing
        Set ObjSql = Nothing
        Unload Me
        Set frmDCU002 = Nothing
        Exit Sub
    ElseIf Rs.RecordCount = 1 Then
        Me.Visible = True
        cmdNext.Visible = False
        lblTitle.Caption = "" & Rs.Fields("Title").Value
        lblUsers.Caption = "�� �� �� ��  : " & Rs.Fields("users").Value & "   �� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��")

        txtText.Text = "�� �� �� ��  : " & Rs.Fields("users").Value & vbNewLine & _
                       "�� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��") & vbNewLine & vbNewLine & vbNewLine & _
                       Rs.Fields("note").Value
    Else
        Me.Visible = True
        lblTitle.Caption = "" & Rs.Fields("Title").Value
        lblUsers.Caption = "�� �� �� ��  : " & Rs.Fields("users").Value & "   �� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��")

        txtText.Text = "�� �� �� ��  : " & Rs.Fields("users").Value & vbNewLine & _
                       "�� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��") & vbNewLine & vbNewLine & vbNewLine & _
                       Rs.Fields("note").Value
                       
        Set MesgKey = New Collection
        cmdNext.Visible = True
        For i = 1 To Rs.RecordCount
            MesgKey.Add "" & Rs.Fields("inputday").Value & vbTab & Rs.Fields("seq").Value & "", "Key" & CStr(i)
            Rs.MoveNext
        Next
    End If
    Set Rs = Nothing
    Set ObjSql = Nothing
    
    ' ��������, �������� ǥ��

'    If ObjMyUser.IsDeveloper Or ObjMyUser.IsManager Then
'        optWorkFg(0).Visible = True
'        optWorkFg(1).Visible = True
'    End If
    
    Exit Sub
ErrTrap:
    Set Rs = Nothing
    Set ObjSql = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Dim strAppName As String
    

    Rtxt.Visible = False: txtText.Visible = True: Rtxt.ZOrder 0
    
    strAppName = mvarTradeMark & " " & mvarProjectId
    chkLoadAtStartup.Value = (GetSetting(strAppName, "Options", "ShowAtStart", 0) + 1) Mod 2

    CurrentTip = 1
    
    Call medAlwaysOn(frmDCU002, 1)
     
    Me.Visible = False
    
'2001/08/28 Modify By Legends
End Sub

Public Sub ChkMsgExist()
    Dim Rs As Recordset
    
    '�������� ������ ��������...
    Set ObjSql = New clsDCUSqlStmt
    ObjSql.WorkFg = "0"
    
    Set Rs = New Recordset
    
    On Error GoTo ErrTrap
    
    Rs.Open ObjSql.GetSQLCOM011ByDeptFg(mvarProjectId), DBConn
    
    '�������� üũ.
    If Rs.EOF Then
        mvarDataExists = False
    Else
        mvarDataExists = True
    End If
    
ErrTrap:
    Set Rs = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' ���� �� �� ���� ǥ�� ���θ� �����մϴ�.
    Dim strAppName As String
    
    strAppName = mvarTradeMark & " " & mvarProjectId
    SaveSetting strAppName, "Options", "ShowAtStart", (chkLoadAtStartup.Value + 1) Mod 2
    
    Set MesgKey = Nothing
    Set ObjSql = Nothing
End Sub

'-------------------------------'
'   2002-08-06 �ۼ��� : �̻��
'-------------------------------'
Private Sub optWorkFg_Click(Index As Integer)
    Dim Rs  As Recordset
    Dim i   As Long
    
    Set ObjSql = New clsDCUSqlStmt
    Set Rs = New Recordset
    
On Error GoTo Errors
    If Index = 0 Then
        ObjSql.WorkFg = "0"
        Rs.Open ObjSql.GetSQLCOM011ByDeptFg(mvarProjectId), DBConn
        frmDCU002.Caption = "��������"
    ElseIf Index = 1 Then
        ObjSql.WorkFg = "1"
        Rs.Open ObjSql.GetSQLCOM011ByDeptFg(mvarProjectId), DBConn
        frmDCU002.Caption = "��������"
    End If
    
    If Rs.RecordCount = 1 Then
        Me.Visible = True
        cmdNext.Visible = False
        lblTitle.Caption = "" & Rs.Fields("Title").Value
        lblUsers.Caption = "�� �� �� ��  : " & Rs.Fields("users").Value & "   �� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��")

        txtText.Text = "�� �� �� ��  : " & Rs.Fields("users").Value & vbNewLine & _
                       "�� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��") & vbNewLine & vbNewLine & vbNewLine & _
                       Rs.Fields("note").Value

    Else
        Me.Visible = True
        lblTitle.Caption = "" & Rs.Fields("Title").Value
        lblUsers.Caption = "�� �� �� ��  : " & Rs.Fields("users").Value & "   �� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��")
        Set MesgKey = New Collection
        cmdNext.Visible = True
        
        txtText.Text = "�� �� �� ��  : " & Rs.Fields("users").Value & vbNewLine & _
               "�� �� �� �� : " & Format(Rs.Fields("inputday").Value, "####��##��##��") & vbNewLine & vbNewLine & vbNewLine & _
               Rs.Fields("note").Value

        For i = 1 To Rs.RecordCount
            MesgKey.Add "" & Rs.Fields("inputday").Value & vbTab & Rs.Fields("seq").Value, "Key" & CStr(i)
            Rs.MoveNext
        Next
    End If

    Set Rs = Nothing
    Set ObjSql = Nothing
    Exit Sub
    
Errors:
    Set ObjSql = Nothing
    Set Rs = Nothing
    MsgBox Err.Description, vbCritical, "����"
End Sub
