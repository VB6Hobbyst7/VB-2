VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIIS616 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '���� ���� â
   Caption         =   "������� ���� ����"
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
      TabIndex        =   3
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&S)"
      Height          =   495
      Left            =   6255
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������(&C)"
      Height          =   495
      Left            =   8685
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   3585
      TabIndex        =   7
      Top             =   -30
      Width           =   7545
      Begin VB.TextBox txtContent 
         BackColor       =   &H00F7FFF7&
         Height          =   5790
         Left            =   255
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   1
         Top             =   2115
         Width           =   7035
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   20
         TabIndex        =   0
         Top             =   780
         Width           =   2505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ������� ����"
         Height          =   180
         Left            =   255
         TabIndex        =   9
         Top             =   1800
         Width           =   1380
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
         Caption         =   "�� ������� ���� �ڵ�"
         Height          =   180
         Left            =   255
         TabIndex        =   8
         Top             =   480
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   9900
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   8205
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwCodeList 
      Height          =   7665
      Left            =   45
      TabIndex        =   6
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "������� ���� �ڵ�"
         Object.Width           =   6050
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������� ����"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblName 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������� ���� ����Ʈ"
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
      Left            =   870
      TabIndex        =   10
      Top             =   165
      Width           =   1935
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
Attribute VB_Name = "frmIIS616"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIIS616.frm
'   �ۼ���  : �̻��
'   ��  ��  : ������� ���� ����
'   �ۼ���  : 2004-03-09
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mTemplate As clsIISTemplate     '���ø� ������ Ŭ����

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With
    
    Set mTemplate = New clsIISTemplate
    Call CtlClear
    Me.Show
    DoEvents
    
    Me.MousePointer = vbHourglass
    Call GetCodeList
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS616").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTemplate = Nothing
    Set frmIIS616 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX        As ListItem
    Dim strCode     As String       '������� ���� �ڵ�
    Dim strContent  As String       '������� ����

    '## ������� ���� �ڵ�
    If txtCode.Text = "" Then
        MsgBox "������� ���� �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtCode.SetFocus
        Exit Sub
    End If

    '## ������� ����
    If txtContent.Text = "" Then
        MsgBox "������� ������ �Է��ϼ���.", vbInformation, "����"
        txtContent.SetFocus
        Exit Sub
    End If

    strCode = Trim(txtCode.Text)
    strContent = Trim(txtContent.Text)
    
    Me.MousePointer = vbHourglass

    Set itmX = lvwCodeList.FindItem(strCode, lvwText)
    With mTemplate
        .CdIndex = CMDYRSNCD
        .Code = strCode
        .Content = strContent

        '# ������� ���� �ڵ尡 ������ Update, ������ Insert
        If itmX Is Nothing Then
            If .AddTemplate Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� ����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        Else
            If .ModifyTemplate Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        End If
    End With

    Call CtlClear
    Call GetCodeList

    With lvwCodeList
        Set itmX = .FindItem(strCode, lvwText)
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
    Dim strCode     As String       '������� ���� �ڵ�
    Dim intTemp     As Integer

    '## ������� ���� �ڵ�
    If txtCode.Text = "" Then
        MsgBox "������� ���� �ڵ带 �Է��ϼ���.", vbInformation, "����"
        txtCode.SetFocus
        Exit Sub
    End If

    strCode = Trim(txtCode.Text)

    intTemp = MsgBox("���� �����ұ��?", vbYesNo + vbQuestion, "Ȯ��")
    If intTemp = vbNo Then Exit Sub

    Set itmX = lvwCodeList.FindItem(strCode, lvwText)
    If itmX Is Nothing Then
        MsgBox "�������� �ʴ� ������� ���� �ڵ� �Դϴ�.", vbInformation, "����"
        With txtCode
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    Set itmX = Nothing

    Me.MousePointer = vbHourglass

    With mTemplate
        .CdIndex = CMDYRSNCD
        .Code = strCode
        
        If .DelTemplate Then
            mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
        End If
    End With

    Call CtlClear
    Call GetCodeList

    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lvwCodeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwCodeList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwCodeList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call CtlClear
    
    txtCode.Text = Item.Text
    txtContent.Text = Item.SubItems(1)
End Sub

Private Sub txtCode_GotFocus()
    With txtCode
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCode_LostFocus()
    Dim itmX    As ListItem
    Dim strCode As String       '������� ���� �ڵ�
    
    strCode = Trim(txtCode.Text)
    If strCode = "" Then Exit Sub
    txtContent.Text = ""
    
    '## �����ϴ� �ڵ��̸� ��Ŀ���̵�, ����ǥ��
    With lvwCodeList
        Set itmX = .FindItem(strCode, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwCodeList_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub

Private Sub txtContent_LostFocus()
    Dim strContent As String    '������� ����
    
    strContent = Trim(txtContent.Text)
    If strContent = "" Then Exit Sub
    
    '## 4000�� ���ϸ� �Է� �ǵ��� �Ѵ�.
    If Len(strContent) >= 4000 Then
        MsgBox "������� ������ 4000�� ���Ϸ� �Է��ؾ� �մϴ�.", vbInformation, "����"
        txtContent.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ڵ� ����Ʈ�� lvwCodeList�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetCodeList()
    Dim Rs   As ADODB.Recordset
    Dim itmX As ListItem
    
On Error GoTo Errors
    With lvwCodeList
        .ListItems.Clear
        
        Set Rs = mTemplate.GetTemplates(CMDYRSNCD)
        If Rs.BOF Or Rs.EOF Then GoTo EndLine
        
        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , Rs.Fields("CODE").Value)
            itmX.SubItems(1) = Rs.Fields("CONTENT").Value & ""
            
            Rs.MoveNext
        Loop
        Set itmX = Nothing
        
        If .ListItems.Count > 37 Then
            .ColumnHeaders(1).Width = 3190
        Else
            .ColumnHeaders(1).Width = 3430
        End If
    End With

EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "frmIIS616", "GetCodeList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtCode.Text = ""
    txtContent.Text = ""
End Sub

