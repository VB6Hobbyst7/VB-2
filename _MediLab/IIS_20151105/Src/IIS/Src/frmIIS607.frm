VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIIS607 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�׷�ó�� ����"
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
   Begin VB.CommandButton cmdAllDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��λ���(&A)"
      Height          =   495
      Left            =   6255
      Style           =   1  '�׷���
      TabIndex        =   4
      ToolTipText     =   "��ǥ�׸� �ش��ϴ� ���׸��� ��λ����մϴ�."
      Top             =   8205
      Width           =   1215
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&S)"
      Height          =   495
      Left            =   5040
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   8205
      Width           =   1215
   End
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   3585
      TabIndex        =   8
      Top             =   -30
      Width           =   7545
      Begin VB.TextBox txtChild 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2685
         Width           =   2160
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   1
         Left            =   2415
         Picture         =   "frmIIS607.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   10
         Top             =   2670
         Width           =   405
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   0
         Left            =   2415
         Picture         =   "frmIIS607.frx":0E42
         Style           =   1  '�׷���
         TabIndex        =   9
         Top             =   780
         Width           =   405
      End
      Begin VB.TextBox txtSeq 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1575
         Width           =   2160
      End
      Begin VB.TextBox txtParent 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   0
         Top             =   795
         Width           =   2160
      End
      Begin MedControls1.LisLabel lblParentNm 
         Height          =   345
         Left            =   2925
         TabIndex        =   17
         Top             =   780
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   609
         BackColor       =   16252919
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblChildNm 
         Height          =   345
         Left            =   2925
         TabIndex        =   18
         Top             =   2670
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   609
         BackColor       =   16252919
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ���׸� �ڵ�"
         Height          =   180
         Left            =   255
         TabIndex        =   19
         Top             =   2370
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� SEQ"
         Height          =   180
         Left            =   255
         TabIndex        =   14
         Top             =   1260
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   7500
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ��ǥ�׸� �ڵ�"
         Height          =   180
         Left            =   255
         TabIndex        =   13
         Top             =   480
         Width           =   1380
      End
   End
   Begin MSComctlLib.ListView lvwParent 
      Height          =   2820
      Left            =   45
      TabIndex        =   11
      Top             =   450
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   4974
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�˻��ڵ�"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�˻��"
         Object.Width           =   4251
      EndProperty
   End
   Begin MSComctlLib.ListView lvwChild 
      Height          =   4440
      Left            =   45
      TabIndex        =   12
      Top             =   3690
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   7832
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SEQ"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�˻��ڵ�"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�˻��"
         Object.Width           =   3263
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���׸� �ڵ�"
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
      Left            =   1200
      TabIndex        =   16
      Top             =   3405
      Width           =   1275
   End
   Begin VB.Label lblName 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ǥ�׸� �ڵ�"
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
      Left            =   1200
      TabIndex        =   15
      Top             =   165
      Width           =   1275
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
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '�ܻ�
      Height          =   375
      Left            =   60
      Top             =   3300
      Width           =   3495
   End
End
Attribute VB_Name = "frmIIS607"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIIS607.frm (�츮LIS�� �����Ҷ� ���)
'   �ۼ���  : �̻��
'   ��  ��  : �׷��׸� ������
'   �ۼ���  : 2004-02-20
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mPanel As clsIISPanel   '�׷��׸� Ŭ����
Private WithEvents mCode1 As clsIISCodeList     '�ڵ帮��Ʈ Ŭ����
Attribute mCode1.VB_VarHelpID = -1
Private WithEvents mCode2 As clsIISCodeList     '�ڵ帮��Ʈ Ŭ����
Attribute mCode2.VB_VarHelpID = -1

Private mTestCd As String       '�˻��ڵ�

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With

    Set mPanel = New clsIISPanel
    Call CtlClear
    Me.Show
    DoEvents

    Me.MousePointer = vbHourglass
    
    '## �׷��׸� ����Ʈ ��ȸ
    Call GetParentList
    
    '## �˻��ڵ� �����Ϳ��� ���� ǥ���ϴ� ���
    If mTestCd <> "" Then
        txtParent.Text = mTestCd
        Call txtParent_LostFocus
        txtSeq.SetFocus
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS607").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mPanel = Nothing
    Set frmIIS607 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX        As ListItem
    Dim strParentCd As String   '��ǥ�׸�
    Dim strSeq      As String   'SEQ
    Dim strChildCd  As String   '���׸�

    '## �Էµ� �ڵ��� ��ȿ�� Check
    If CheckCode = False Then Exit Sub
    
    strParentCd = Trim(txtParent.Text)
    strSeq = Format$(txtSeq.Text, "00")
    strChildCd = Trim(txtChild.Text)

    '## �����ϴ� ��ǥ�׸�+SEQ�̸� Update, �������� ������ Insert
    Me.MousePointer = vbHourglass
    
    Set itmX = lvwChild.FindItem(strSeq, lvwText)
    With mPanel
        .ParentCd = strParentCd
        .Seq = strSeq
        .ChildCd = strChildCd
        
        If itmX Is Nothing Then
            If .AddPanel Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� ����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        Else
            If .ModifyPanel Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        End If
    End With
    Call CtlClear
    Call GetParentList

    Set itmX = lvwParent.FindItem(strParentCd, lvwText)
    If Not (itmX Is Nothing) Then
        lvwParent.ListItems(itmX.Index).Selected = True
        lvwParent.ListItems(itmX.Index).EnsureVisible
        Call lvwParent_ItemClick(itmX)
    End If
    Set itmX = Nothing
    txtSeq.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdAllDelete_Click()
    Dim strParentCd As String   '���ڵ�
    Dim intTemp     As Integer

    If txtParent.Text = "" Then
        MsgBox "��ǥ�׸� �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtParent.SetFocus
        Exit Sub
    End If
    
    strParentCd = Trim(txtParent.Text)
    intTemp = MsgBox("��ǥ�׸� ���Ե� ��� ���׸��� �����˴ϴ�. ���� �����ұ��?", vbYesNo + vbQuestion, "Ȯ��")
    If intTemp = vbNo Then Exit Sub

    '## ��ǥ�׸� ���Ե� ��� ���׸� ����
    Me.MousePointer = vbHourglass
    
    With mPanel
        .ParentCd = strParentCd
        If .DelPanelAll Then
            mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
        End If
    End With
    Call CtlClear
    Call GetParentList
    txtParent.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim itmX        As ListItem
    Dim strParentCd As String   '��ǥ�׸�
    Dim strSeq      As String   'SEQ
    Dim intTemp     As Integer

    If txtParent.Text = "" Then
        MsgBox "��ǥ�׸� �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtParent.SetFocus
        Exit Sub
    End If
    
    If txtSeq.Text = "" Then
        MsgBox "SEQ�� �Է��ϼ���.", vbInformation, "����"
        txtSeq.SetFocus
        Exit Sub
    End If
    
    strParentCd = Trim(txtParent.Text)
    strSeq = Trim(txtSeq.Text)
    
    intTemp = MsgBox("���� �����ұ��?", vbYesNo + vbQuestion, "Ȯ��")
    If intTemp = vbNo Then Exit Sub

    Me.MousePointer = vbHourglass
    
    With mPanel
        .ParentCd = strParentCd
        .Seq = strSeq
        If .DelPanel Then
            mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
        End If
    End With
    Call CtlClear
    Call GetParentList
    
    Set itmX = lvwParent.FindItem(strParentCd, lvwText)
    If Not (itmX Is Nothing) Then
        lvwParent.ListItems(itmX.Index).Selected = True
        lvwParent.ListItems(itmX.Index).EnsureVisible
        Call lvwParent_ItemClick(itmX)
    End If
    Set itmX = Nothing
    txtSeq.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub txtParent_GotFocus()
    With txtParent
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    '## �ҹ��ڰ� �ԷµǸ� �빮�ڷ� ����
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtParent_LostFocus()
    Dim itmX        As ListItem
    Dim strParentCd As String       '�׷��ڵ�
    Dim strParentNm As String       '�׷��ڵ��
    
    '## 1.�Էµ� �˻��ڵ尡 lvwParent�� �����ϴ� ��� �ش��ڵ��� ������ ǥ���ϰ�
    '## 2.�������� ������ �ڵ��� �׷��ڵ忩�θ� �ľ��Ͽ� �׷��ڵ尡 �ƴϸ� ����޽���,
    '   �׷��ڵ��̸� �����Է��Ҽ� �ֵ��� �Ѵ�.
    strParentCd = Trim(txtParent.Text)
    If strParentCd = "" Then Exit Sub
    lblParentNm.Caption = "":   txtSeq.Text = ""
    txtChild.Text = "":         lblChildNm.Caption = ""
    lvwChild.ListItems.Clear
    
    Set itmX = lvwParent.FindItem(strParentCd, lvwText)
    If itmX Is Nothing Then
        '## �Էµ� �ڵ尡 �������� �ʴ� ���
        strParentNm = mPanel.GetPanelNm(strParentCd)
        If strParentNm = "" Then
            MsgBox "�Է��� �ڵ�� ��ǥ�׸� �ڵ尡 �ƴմϴ�.", vbInformation, "����"
            With txtParent
                .SetFocus
                .Text = ""
            End With
        Else
            lblParentNm.Caption = strParentNm
        End If
    Else
        '## �Էµ� �ڵ尡 �����ϴ� ���
        With lvwParent
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwParent_ItemClick(itmX)
        End With
    End If
    Set itmX = Nothing
End Sub

Private Sub txtSeq_GotFocus()
    With txtSeq
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSeq_LostFocus()
    Dim itmX        As ListItem
    Dim strSeq      As String       'SEQ
    
    '## �Էµ� SEQ�� �����ϸ� ������ ǥ�� ������ �����Է��Ҽ� �ֵ��� �Ѵ�.
    strSeq = Format$(Trim(txtSeq.Text), "00")
    If strSeq = "" Then Exit Sub
    
    Set itmX = lvwChild.FindItem(strSeq, lvwText)
    
    txtChild.Text = ""
    lblChildNm.Caption = ""
    If Not (itmX Is Nothing) Then
        With lvwChild
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwChild_ItemClick(itmX)
        End With
    Else
        txtSeq.Text = strSeq
    End If
    Set itmX = Nothing
End Sub

Private Sub txtChild_GotFocus()
    With txtChild
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtChild_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtChild_KeyPress(KeyAscii As Integer)
    '## �ҹ��ڰ� �ԷµǸ� �빮�ڷ� ����
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtChild_LostFocus()
    Dim itmX        As ListItem
    Dim strChildCd  As String       '���׸� �ڵ�
    Dim strChildNm  As String       '���׸� �ڵ��
    
    '## �Էµ� �ڵ尡 �����׸�, �󼼸��ڵ����� �ľ��ؼ� �ƴϸ� ����޽��� ���
    strChildCd = Trim(txtChild.Text)
    If strChildCd = "" Then Exit Sub
    
    strChildNm = mPanel.GetChildNm(strChildCd)
    If strChildNm = "" Then
        MsgBox "�Է��� �ڵ�� ���׸� �ڵ尡 �ƴմϴ�.", vbInformation, "����"
        With txtChild
            .SetFocus
            .Text = ""
        End With
    Else
        lblChildNm.Caption = strChildNm
    End If
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    Select Case Index
        Case 0
            Set mCode1 = New clsIISCodeList
            With mCode1
                .Caption = "��ǥ�׸� ����Ʈ"
                .HeaderCd = "�˻��ڵ�"
                .HeaderCdNm = "�˻��"
                .CodeListBySql mPanel.GetPanelListBySql
            End With
            Set mCode1 = Nothing
        Case 1
            Set mCode2 = New clsIISCodeList
            With mCode2
                .Caption = "���׸� ����Ʈ"
                .HeaderCd = "�˻��ڵ�"
                .HeaderCdNm = "�˻��"
                .CodeListBySql mPanel.GetChildListBySql
            End With
            Set mCode2 = Nothing
    End Select
End Sub

Private Sub lvwParent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwParent
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
    
    Dim Col As MSComctlLib.ColumnHeader
    
End Sub

Private Sub lvwParent_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## Textbox�� �ڵ�, �ڵ�� ǥ��
    Call CtlClear
    txtParent.Text = Item.Text
    lblParentNm.Caption = Item.SubItems(1)
    
    '## ���õ� ��ǥ�׸��� ���׸� ����Ʈ ǥ��
    Call GetChildList
End Sub

Private Sub lvwChild_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwChild
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwChild_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## Textbox�� �ڵ��, �ڵ�� ǥ��
    txtSeq.Text = Item.Text
    txtChild.Text = Item.SubItems(1)
    lblChildNm.Caption = Item.SubItems(2)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �׷��׸� lvwParent�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetParentList()
    Dim Rs          As ADODB.Recordset
    Dim itmX        As ListItem
    
On Error GoTo Errors
    Set Rs = mPanel.GetPanelList
    If Not (Rs.BOF Or Rs.EOF) Then
        With lvwParent
            .ListItems.Clear
            lvwChild.ListItems.Clear
            
            Do Until Rs.EOF
                Set itmX = .ListItems.Add(, , Rs.Fields("TESTCD").Value)
                itmX.SubItems(1) = Rs.Fields("TESTNM").Value
                Rs.MoveNext
            Loop
            
            If .ListItems.Count > 12 Then
                .ColumnHeaders(2).Width = 2210
            Else
                .ColumnHeaders(2).Width = 2410
            End If
        End With
    End If
    Rs.Close
    Set Rs = Nothing
    Set itmX = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "frmIIS607", "GetParentList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���׸��� lvwParent�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetChildList()
    Dim Rs          As ADODB.Recordset
    Dim itmX        As ListItem
    
On Error GoTo Errors
    Set Rs = mPanel.GetPanelChildList(Trim(txtParent.Text))
    If Not (Rs.BOF Or Rs.EOF) Then
        With lvwChild
            .ListItems.Clear
            
            Do Until Rs.EOF
                Set itmX = .ListItems.Add(, , Rs.Fields("SEQ").Value)
                itmX.SubItems(1) = Rs.Fields("TESTCD").Value
                itmX.SubItems(2) = Rs.Fields("TESTNM").Value & ""
                Rs.MoveNext
            Loop
            
            If .ListItems.Count > 21 Then
                .ColumnHeaders(3).Width = 1600
            Else
                .ColumnHeaders(3).Width = 1850
            End If
        End With
    End If
    Rs.Close
    Set Rs = Nothing
    Set itmX = Nothing
    Exit Sub

Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "frmIIS607", "GetChildList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �Էµ� �׷��ڵ�, ���ڵ��� ��ȿ�� Check
'   ��ȯ : True(��ȿ), False(��ȿ)
'-----------------------------------------------------------------------------'
Private Function CheckCode() As Boolean
                           
    If txtParent.Text = "" Then
        MsgBox "��ǥ�׸� �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtParent.SetFocus
        Exit Function
    End If
    
    If txtSeq.Text = "" Then
        MsgBox "SEQ�� �Է��ϼ���.", vbInformation, "����"
        txtSeq.SetFocus
        Exit Function
    End If
    
    If txtChild.Text = "" Then
        MsgBox "���׸� �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtChild.SetFocus
        Exit Function
    End If
    
    CheckCode = True
End Function

'-----------------------------------------------------------------------------'
'   ��� : CodeList���� �̺�Ʈ ó��1
'-----------------------------------------------------------------------------'
Private Sub mCode1_SelectedItem(ByRef pSelItem As String)
    Dim itmX As ListItem
    
    txtParent.Text = mGetP(pSelItem, 1, DIV)
    lblParentNm.Caption = mGetP(pSelItem, 2, DIV)
    
    With lvwParent
        Set itmX = .FindItem(txtParent.Text, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwParent_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : CodeList���� �̺�Ʈ ó��2
'-----------------------------------------------------------------------------'
Private Sub mCode2_SelectedItem(ByRef pSelItem As String)
    txtChild.Text = mGetP(pSelItem, 1, DIV)
    lblChildNm.Caption = mGetP(pSelItem, 2, DIV)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtParent.Text = "":        txtChild.Text = ""
    txtSeq.Text = "":           lblParentNm.Caption = ""
    lblChildNm.Caption = ""
End Sub