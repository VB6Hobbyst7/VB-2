VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmDSM003 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�׷� ���"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmDSM003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�׷� ����"
      Height          =   1575
      Left            =   128
      TabIndex        =   12
      Top             =   120
      Width           =   5955
      Begin VB.ComboBox cboGroupID 
         Height          =   300
         ItemData        =   "frmDSM003.frx":030A
         Left            =   2160
         List            =   "frmDSM003.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "�׷� ID�� �Է��ϼ���."
         Top             =   240
         Width           =   3555
      End
      Begin VB.TextBox txtGroupNm 
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   2175
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "�׷��"
         Top             =   675
         Width           =   3555
      End
      Begin VB.TextBox txtGroupDesc 
         Appearance      =   0  '���
         BackColor       =   &H00FFFBFF&
         Height          =   300
         Left            =   2175
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "�׷켳��"
         Top             =   1095
         Width           =   3555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��  ��  ��       : "
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
         Left            =   180
         TabIndex        =   17
         Tag             =   "105"
         Top             =   735
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�׷�   ID        : "
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
         Left            =   180
         TabIndex        =   16
         Tag             =   "105"
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�׷�  ����       : "
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
         Left            =   180
         TabIndex        =   15
         Tag             =   "105"
         Top             =   1155
         Width           =   1710
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���� �ο�"
      Height          =   680
      Left            =   128
      TabIndex        =   13
      Top             =   1815
      Width           =   5955
      Begin VB.OptionButton optUserFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&End User"
         Height          =   180
         Index           =   3
         Left            =   4365
         TabIndex        =   18
         Tag             =   "E"
         Top             =   300
         Width           =   1200
      End
      Begin VB.OptionButton optUserFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Su&pervisor"
         Height          =   180
         Index           =   2
         Left            =   3030
         TabIndex        =   5
         Tag             =   "S"
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optUserFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "De&veloper"
         Height          =   180
         Index           =   1
         Left            =   1695
         TabIndex        =   4
         Tag             =   "D"
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optUserFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Manager"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Tag             =   "M"
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EEEBED&
      Caption         =   "�ݱ�(&X)"
      Height          =   405
      Left            =   4710
      Style           =   1  '�׷���
      TabIndex        =   11
      Top             =   6180
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00EEEBED&
      Caption         =   "����(&S)"
      Height          =   405
      Left            =   3540
      Style           =   1  '�׷���
      TabIndex        =   10
      Top             =   6180
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���� Product"
      Height          =   680
      Left            =   128
      TabIndex        =   14
      Top             =   2610
      Width           =   5955
      Begin VB.CheckBox chkDeptFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� ����"
         Height          =   255
         Index           =   1
         Left            =   2070
         TabIndex        =   7
         Tag             =   "BBS"
         Top             =   255
         Width           =   1215
      End
      Begin VB.CheckBox chkDeptFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� ����"
         Height          =   255
         Index           =   0
         Left            =   495
         TabIndex        =   6
         Tag             =   "APS"
         Top             =   255
         Width           =   1215
      End
      Begin VB.CheckBox chkDeptFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�ӻ� ����"
         Height          =   255
         Index           =   2
         Left            =   3645
         TabIndex        =   8
         Tag             =   "LIS"
         Top             =   255
         Width           =   1215
      End
   End
   Begin FPSpread.vaSpread tblGroupInformation 
      Height          =   2640
      Left            =   120
      TabIndex        =   9
      Top             =   3405
      Width           =   5940
      _Version        =   196608
      _ExtentX        =   10478
      _ExtentY        =   4657
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   7
      MaxRows         =   50
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14411494
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmDSM003.frx":030E
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "����(&D)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�ݱ�(&X)"
      End
   End
End
Attribute VB_Name = "frmDSM003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Expire By legends
'2003/09/29
'�� ���� ���̻� ������� ����

'Coding By Legends
'Coding Date 2k/10
'�׷���

Private objMySql As New clsDSMSqlStmt
Private lngToggle(1 To 3) As Long   '�������� Ŭ���� ����� ���Ѻ���

Private mvarProjectId As String 'APS, BBS, LIS ���θ� �޾ƿ��� ����

Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property
Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

Private Sub cboGroupID_Change()
    If Trim(txtGroupNm.Text) = "" Then Exit Sub
    Call ClearText  '�� �ʱ�ȭ
End Sub

Private Sub cboGroupID_Click()
    If Trim(cboGroupID.Text) = "" Then Exit Sub
    Call ShowFormValue  '���� �׷�Ƶ� �ش�Ǵ� ���� �����ش�.
End Sub

Private Sub cboGroupID_KeyDown(KeyCode As Integer, Shift As Integer)
'�޺��� ����Ʈ �߰�
    Dim i As Long
    
    If Trim(cboGroupID.Text) = "" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboGroupID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboGroupID_LostFocus()
    Dim i As Long
    
    If Trim(cboGroupID.Text) = "" Then Exit Sub
    
    With cboGroupID
        For i = 0 To .ListCount
            If .List(i) = .Text Then
                Call ShowFormValue  '���� �׷�Ƶ� �ش�Ǵ� ���� �����ش�.
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub chkDeptFg_Click(Index As Integer)

    Call ShowFormSpread(chkDeptFg(Index).Tag, chkDeptFg(Index).Value)

End Sub

Private Sub ShowFormSpread(ByVal pDeptFg As String, ByVal pShowFg As Long)
    
    Dim i As Long
    
    If pShowFg = 0 Then
        With tblGroupInformation
            For i = .MaxRows To 1 Step -1
                .Row = i
                .Col = 1
                If .Value = pDeptFg Then
                    .Action = ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                End If
            Next
        End With
    Else
        Call objMySql.ShowSpread(tblGroupInformation, pDeptFg)
    End If
End Sub

Private Sub cmdExit_Click()
    Set objMySql = Nothing
    Unload Me
End Sub

Private Sub cmdSave_Click()
'��� ����

    Dim i As Long
    Dim strDeptFg As String     '�μ�����
    Dim strFormId As String     '�� �Ƶ�
    Dim strReadFg As String     '�б����
    Dim strWriteFg As String    '�������
    Dim strPrintFg As String    '��±���
    Dim strMsg As VbMsgBoxResult    '�޽��� ��� ����
    Dim blnCOM008UpdateCheck As Boolean
    Dim blnCOM009UpdateCheck As Boolean
    Dim strUserFg As String
    Dim strSQL As String
    
    If cboGroupID = "" Then
        MsgBox "�׷�ID�� �Է��ϰų� �����ϼ���.", vbInformation, "����Ȯ��"
        cboGroupID.SetFocus
        Exit Sub
    End If
    
    If txtGroupNm = "" Then
        MsgBox "�׷���� �Է��ϼ���.", vbInformation, "����Ȯ��"
        txtGroupNm.SetFocus
        Exit Sub
    End If
        
    If optUserFg(0).Value = "0" And optUserFg(1).Value = "0" And optUserFg(2).Value = "0" And optUserFg(3).Value = "0" Then
        MsgBox "���õ� �׷��� ������ �����Ͽ� �ֽʽÿ�", vbInformation, "����Ȯ��"
        Exit Sub
    End If
        
    If optUserFg(0).Value = "1" Then strUserFg = "M"
    If optUserFg(1).Value = "1" Then strUserFg = "D"
    If optUserFg(2).Value = "1" Then strUserFg = "S"
    If optUserFg(3).Value = "1" Then strUserFg = "E"
            
    strSQL = objMySql.Query(5) & " where groupid = '" & cboGroupID.Text & "'"
            
    strMsg = MsgBox("�׷� ID�� '" & cboGroupID & "' �� �׷쿡 ����մϴ�." & vbNewLine & _
                    "�½��ϱ�?", vbQuestion + vbYesNo, "����Ȯ��")
                        
    If strMsg = vbYes Then
        '������Ʈ üũ
        blnCOM008UpdateCheck = objMySql.UpdateCheck(cboGroupID, , , strSQL)
        'COM008�� ����
        Call objMySql.SetGroupDB("0", blnCOM008UpdateCheck, , cboGroupID, Trim(txtGroupNm), Trim(txtGroupDesc), _
                                strUserFg, chkDeptFg(0).Value, chkDeptFg(1).Value, chkDeptFg(2).Value)
        
        'COM009 ����
        Call objMySql.DelCOM89("1", cboGroupID)
        
        With tblGroupInformation
            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 1
                Select Case Trim(.Text)
                    Case "APS": strDeptFg = "A"
                    Case "BBS": strDeptFg = "B"
                    Case "LIS": strDeptFg = "L"
                End Select
                .Col = 6: strFormId = .Text
'                .Col = 3: strReadFg = IIf(.Text = "1", "1", "0")
                .Col = 3: strReadFg = .Value
                If Not .CellType = CellTypeCheckBox Then strReadFg = "X"
'                .Col = 4: strWriteFg = IIf(.Text = "1", "1", "0")
                .Col = 4: strWriteFg = .Value
                If Not .CellType = CellTypeCheckBox Then strWriteFg = "X"
'                .Col = 5: strPrintFg = IIf(.Text = "1", "1", "0")
                .Col = 5: strPrintFg = .Value
                If Not .CellType = CellTypeCheckBox Then strPrintFg = "X"
                
                '������Ʈ üũ
'                blnCOM009UpdateCheck = objMySql.UpdateCheck(cboGroupID, strDeptFg, Trim(strFormID), _
                                                           "SELECT * FROM COM009")
               
                'COM009�� ����
'                Call objMySql.SetGroupDB("1", , blnCOM009UpdateCheck, cboGroupID, , , , , , , _
                                         strDeptFg, Trim(strFormID), strReadFg, strWriteFg, strPrintFg)
                Call objMySql.SetGroupDB("1", , , cboGroupID, , , , , , , _
                                         strDeptFg, Trim(strFormId), strReadFg, strWriteFg, strPrintFg)
            Next
        End With
        
'        If objMySql.SetGroupDB Then
            MsgBox "����Ǿ����ϴ�.", vbInformation, "����Ȯ��"
'        End If
        
        Call Form_Load
        cboGroupID.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Call GetGroupId
End Sub

Private Sub Form_Load()
    Call InitializeForm
End Sub

Private Sub InitializeForm()
    Call ClearText  '�� �ʱ�ȭ
    
    lngToggle(1) = 0: lngToggle(2) = 0: lngToggle(3) = 0 '�������� Ŭ���� ���Ǵ� ���� �ʱ�ȭ
End Sub

Private Sub GetGroupId()
'�׷� �Ƶ� ���´�.
    Dim Rs As New Recordset
    Dim objSQL As clsDSMSqlStmt
    
    Set objSQL = New clsDSMSqlStmt
    
    Rs.Open objMySql.Query(5), dbconn
    
    If Rs.EOF = False Then
        With cboGroupID 'frmDSM003.cboGroupID
            .clear
            Do Until Rs.EOF
                .AddItem "" & Rs.Fields("GroupID").Value
                Rs.MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
    Set objSQL = Nothing
End Sub

Private Sub ShowFormValue()
    
    Call ShowText  '�ؽ�Ʈ ���� �����ش�.

    Call ShowSpreadValue(cboGroupID.Text)    '�������忡 ���� �����ش�.
End Sub

Public Sub ShowSpreadValue(ByVal pText As String)
'��� �ִ� ���� �������忡 �����ش�.

    Dim Rs As New Recordset
    Dim objSQL As clsDSMSqlStmt
    Dim strSQL As String
    Dim strDeptFg1 As String, strDeptFg2 As String
    Dim strFrmId1 As String, strFrmId2 As String
    Dim i As Long
    
    
    Set objSQL = New clsDSMSqlStmt
    
    strSQL = objSQL.Query(6) & Trim(pText) & "'"
    
    Rs.Open strSQL, dbconn
    
    If Rs.EOF Then
        Exit Sub
    End If
    
    With tblGroupInformation   'frmDSM003.tblGroupInformation
        .ReDraw = False
        
        .Col = 3: .Col2 = 5
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Value = ""
        .BlockMode = False
        
        Do Until Rs.EOF
            strFrmId1 = "" & Rs.Fields("FormID").Value
            Select Case "" & Rs.Fields("DeptFg").Value
                Case "L": strDeptFg1 = "LIS"
                Case "B": strDeptFg1 = "BBS"
                Case "A": strDeptFg1 = "APS"
            End Select
            
            For i = 1 To .MaxRows
                .Row = i
                .Col = 1: strDeptFg2 = .Text
                .Col = 6: strFrmId2 = .Text
                
                If (strFrmId1 = strFrmId2) And (strDeptFg1 = strDeptFg2) Then
            
                    .Row = i
                    .Col = 3:
                    If "" & Rs.Fields("ReadFg").Value = "X" Then
                        .CellType = CellTypeStaticText
                    Else
                        .Value = "" & Rs.Fields("ReadFg").Value
                    End If
                
                    .Col = 4:
                    If "" & Rs.Fields("WriteFg").Value = "X" Then
                        .CellType = CellTypeStaticText
                    Else
                        .Value = "" & Rs.Fields("WriteFg").Value
                    End If
                
                    .Col = 5:
                    If "" & Rs.Fields("PrintFg").Value = "X" Then
                        .CellType = CellTypeStaticText
                    Else
                        .Value = "" & Rs.Fields("PrintFg").Value
                    End If
            
                End If
            Next
            
            Rs.MoveNext
        Loop
        .ReDraw = True
    End With
    
    Set objSQL = Nothing
End Sub

Private Sub ShowText()
'�ؽ�Ʈ ���� �����ش�.
    
    If objMySql.ShowTextValue(Trim(cboGroupID.Text)) Then
        With objMySql
            txtGroupNm = .GroupNm
            txtGroupDesc = .GroupDesc
            chkDeptFg(0).Value = .APSFg
            chkDeptFg(1).Value = .BBSFg
            chkDeptFg(2).Value = .LISFg
            Select Case .UserFg
                Case "M": optUserFg(0).Value = "1"
                Case "D": optUserFg(1).Value = "1"
                Case "S": optUserFg(2).Value = "1"
                Case "E": optUserFg(3).Value = "1"
            End Select
        End With
    End If
    
End Sub

Private Sub ClearText()
    Dim i As Long
    
    txtGroupNm = ""
    txtGroupDesc = ""
    
    For i = 0 To 2
        optUserFg(i).Value = "0"
        chkDeptFg(i).Value = "0"
    Next
    optUserFg(3).Value = "0"
    
    tblGroupInformation.MaxRows = 0
    
    Select Case mvarProjectId   '�μ�����
        Case "APS": chkDeptFg(0).Value = 1: 'chkDeptFg(0).Enabled = False
        Case "BBS": chkDeptFg(1).Value = 1: 'chkDeptFg(1).Enabled = False
        Case "LIS": chkDeptFg(2).Value = 1: 'chkDeptFg(2).Enabled = False
    End Select
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���콺 ������ ��ư Ŭ���� �˾��޴��� �����ش�.
    If Button = 2 Then
        frmDSM003.PopupMenu mnuEdit
    End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���콺 ������ ��ư Ŭ���� �˾��޴��� �����ش�.
    If Button = 2 Then
        frmDSM003.PopupMenu mnuEdit
    End If
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���콺 ������ ��ư Ŭ���� �˾��޴��� �����ش�.
    If Button = 2 Then
        frmDSM003.PopupMenu mnuEdit
    End If
End Sub

Private Sub Frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���콺 ������ ��ư Ŭ���� �˾��޴��� �����ش�.
    If Button = 2 Then
        frmDSM003.PopupMenu mnuEdit
    End If
End Sub

Private Sub mnuDelete_Click()
    Call DeleteCOM008_9
End Sub

Private Sub DeleteCOM008_9()
'COM008, COM009 ����

    Dim strMsg As VbMsgBoxResult
    
    If cboGroupID = "" Then
        MsgBox "�׷�ID�� �Է��ϰų� �����ϼ���.", vbInformation, "����Ȯ��"
        cboGroupID.SetFocus
        Exit Sub
    End If
    
    strMsg = MsgBox("�׷� ID�� '" & cboGroupID & "' �� �׷��� �����մϴ�." & vbNewLine & _
                    "�½��ϱ�?", vbQuestion + vbYesNo, "����Ȯ��")
    
    If strMsg = vbYes Then
        If objMySql.DelCOM89("0", Trim(cboGroupID.Text)) Then
            MsgBox "�����Ǿ����ϴ�.", vbInformation, "����Ȯ��"
        End If
        
        Call Form_Load
        cboGroupID.SetFocus
    End If
End Sub
Private Sub mnuExit_Click()
    Set objMySql = Nothing
    Unload Me
End Sub

Private Sub mnuSave_Click()
    Call cmdSave_Click
End Sub

Private Sub tblGroupInformation_Click(ByVal Col As Long, ByVal Row As Long)
'���������� ����� Ŭ������ ��� ��ü ����, ���� �۾��� ����
    Dim i As Long
    
    If Col < 3 Then Exit Sub
    If Row > 0 Then Exit Sub
    
    lngToggle(Col - 2) = (lngToggle(Col - 2) + 1) Mod 2
    
    With tblGroupInformation
        .Col = Col
        For i = 1 To .MaxRows
            .Row = i
            If .CellType = CellTypeCheckBox Then
                .Value = lngToggle(Col - 2)
            End If
        Next
    End With
End Sub

Private Sub tblGroupInformation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���콺 ������ ��ư Ŭ���� �˾��޴��� �����ش�.
    If Button = 2 Then
        frmDSM003.PopupMenu mnuEdit
    End If
End Sub

Private Sub txtGroupDesc_GotFocus()
    With txtGroupDesc
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtGroupNm_GotFocus()
    With txtGroupNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtGroupNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtGroupNm.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
