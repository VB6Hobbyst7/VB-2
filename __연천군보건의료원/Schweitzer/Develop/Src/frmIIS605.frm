VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIIS605 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '���� ���� â
   Caption         =   "WorkArea ����"
   ClientHeight    =   8925
   ClientLeft      =   4080
   ClientTop       =   285
   ClientWidth     =   11175
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&D)"
      Height          =   495
      Left            =   7470
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&S)"
      Height          =   495
      Left            =   6255
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������(&C)"
      Height          =   495
      Left            =   8685
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   3585
      TabIndex        =   8
      Top             =   -30
      Width           =   7545
      Begin VB.TextBox txtWaKorNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   20
         TabIndex        =   2
         Top             =   2910
         Width           =   3750
      End
      Begin VB.TextBox txtWaNum 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   1
         TabIndex        =   3
         Top             =   3690
         Width           =   3750
      End
      Begin VB.TextBox txtWaEngNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   20
         TabIndex        =   1
         Top             =   2115
         Width           =   3750
      End
      Begin VB.TextBox txtWaCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   2
         TabIndex        =   0
         Top             =   780
         Width           =   2505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� WorkArea��(����)"
         Height          =   180
         Left            =   255
         TabIndex        =   13
         Top             =   1800
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� WorkArea��(�ѱ�)"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   2595
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ������ȣ �ο�����"
         Height          =   180
         Left            =   255
         TabIndex        =   11
         Top             =   3375
         Width           =   1740
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   7500
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� WorkArea�ڵ�"
         Height          =   180
         Left            =   255
         TabIndex        =   10
         Top             =   480
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   9900
      Style           =   1  '�׷���
      TabIndex        =   7
      Top             =   8205
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwWaList 
      Height          =   7665
      Left            =   45
      TabIndex        =   9
      Top             =   450
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   13520
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ڵ��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�ڵ��"
         Object.Width           =   4251
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Workarea��(�ѱ�)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "������ȣ �ο�����"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblName 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "WorkArea ����Ʈ"
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
      Left            =   1035
      TabIndex        =   14
      Top             =   165
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '�ܻ�
      Height          =   375
      Left            =   60
      Top             =   60
      Width           =   3495
   End
End
Attribute VB_Name = "frmIIS605"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIIS605.frm
'   �ۼ���  : �̻��
'   ��  ��  : WorkArea ������
'   �ۼ���  : 2004-01-12
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mWorkarea As clsIISWorkarea     'Workare Ŭ����

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With
    
    Set mWorkarea = New clsIISWorkarea
    Call CtlClear
    Me.Show
    DoEvents

    Me.MousePointer = vbHourglass
    Call GetWAList
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS605").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mWorkarea = Nothing
    Set frmIIS605 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX        As ListItem
    Dim strWaCd     As String           'Workarea�ڵ�
    Dim strWAEngNm  As String           'Workarea��(����)
    Dim strWAKorNm  As String           'Workarea��(�ѱ�)
    Dim strWaNum    As String           '������ȣ �ο�����

    If CheckCode = False Then Exit Sub
    
    strWaCd = Trim(txtWaCd.Text)
    strWAEngNm = Trim(txtWaEngNm.Text)
    strWAKorNm = Trim(txtWaKorNm.Text)
    strWaNum = Trim(txtWaNum.Text)

    Me.MousePointer = vbHourglass

    '## Workarea�߰�
    Set itmX = lvwWaList.FindItem(strWaCd, lvwText)
    With mWorkarea
        .WaCd = strWaCd
        .WaEngNm = strWAEngNm
        .WaKorNm = strWAKorNm
        .WaNumRule = strWaNum
        
        If itmX Is Nothing Then
            If .AddWorkarea Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� ����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        Else
            If .ModifyWorkarea Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        End If
    End With
    
    Call CtlClear
    Call GetWAList
    
    With lvwWaList
        Set itmX = .FindItem(strWaCd, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
        End If
        Set itmX = Nothing
    End With
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim itmX        As ListItem
    Dim strWaCd     As String           'Workarea�ڵ�
    Dim intTemp     As Integer

    If txtWaCd.Text = "" Then
        MsgBox "Workarea �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtWaCd.SetFocus
        Exit Sub
    End If
    
    strWaCd = Trim(txtWaCd.Text)
    
    intTemp = MsgBox("���� �����ұ��?", vbYesNo + vbQuestion, "Ȯ��")
    If intTemp = vbNo Then Exit Sub

    Me.MousePointer = vbHourglass

    '## Workarea����
    Set itmX = lvwWaList.FindItem(strWaCd, lvwText)
    If itmX Is Nothing Then Exit Sub
    Set itmX = Nothing
    
    With mWorkarea
        .WaCd = strWaCd
        If .DelWorkarea Then
            mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
        End If
    End With
    Call CtlClear
    Call GetWAList
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lvwWaList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwWaList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwWaList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call CtlClear
    txtWaCd.Text = Item.Text
    txtWaEngNm.Text = Item.SubItems(1)
    txtWaKorNm.Text = Item.SubItems(2)
    txtWaNum.Text = Item.SubItems(3)
End Sub

Private Sub txtWaCd_GotFocus()
    With txtWaCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWaCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtWaCd_KeyPress(KeyAscii As Integer)
    '## �ҹ��ڰ� �ԷµǸ� �빮�ڷ� ����
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtWaCd_LostFocus()
    Dim itmX    As ListItem
    Dim strWaCd As String       'Workarea�ڵ�
    
    strWaCd = Trim(txtWaCd.Text)
    If strWaCd = "" Then Exit Sub
    
    With lvwWaList
        Set itmX = .FindItem(strWaCd, lvwText)
        If itmX Is Nothing Then
            Call CtlClear
            txtWaCd.Text = strWaCd
        Else
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwWaList_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub

Private Sub txtWaEngNm_GotFocus()
    With txtWaEngNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWaEngNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtWaKorNm_GotFocus()
    With txtWaKorNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWaKorNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtWaNum_GotFocus()
    With txtWaNum
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWaNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : WorkArea����Ʈ ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetWAList()
    Dim Rs      As ADODB.Recordset
    Dim itmX    As ListItem

On Error GoTo Errors
    With lvwWaList
        .ListItems.Clear
        
        Set Rs = mWorkarea.GetWorkarea
        If Rs.BOF Or Rs.EOF Then GoTo EndLine
        
        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , Rs.Fields("WACD").Value)
            itmX.SubItems(1) = Rs.Fields("WAENGNM").Value
            itmX.SubItems(2) = Rs.Fields("WAKORNM").Value
            itmX.SubItems(3) = Rs.Fields("WANUMRULE").Value
            
            Rs.MoveNext
        Loop
        
        If .ListItems.Count > 37 Then
            .ColumnHeaders(2).Width = 2210
        Else
            .ColumnHeaders(2).Width = 2410
        End If
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Set itmX = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "clsIISWorkarea", "GetWAList", Err.Description, GetSysDate
    MsgBox Error.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �Է�, ������ �ʿ��� ������ ��ȿ�� Check
'   ��ȯ : True(��ȿ), False(��ȿ)
'-----------------------------------------------------------------------------'
Private Function CheckCode() As Boolean
    '## Workarea �ڵ�
    If txtWaCd.Text = "" Then
        MsgBox "Workarea �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtWaCd.SetFocus
        Exit Function
    End If
    
    '## Workarea��(����)
    If txtWaEngNm.Text = "" Then
        MsgBox "Workarea��(����)�� �Է��ϼ���.", vbInformation, "����"
        txtWaEngNm.SetFocus
        Exit Function
    End If
    
    '## Workarea��(�ѱ�)
    If txtWaKorNm.Text = "" Then
        MsgBox "Workarea��(�ѱ�)�� �Է��ϼ���.", vbInformation, "����"
        txtWaKorNm.SetFocus
        Exit Function
    End If
    
    CheckCode = True
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtWaCd.Text = ""
    txtWaEngNm.Text = ""
    txtWaKorNm.Text = ""
    txtWaNum.Text = ""
End Sub
