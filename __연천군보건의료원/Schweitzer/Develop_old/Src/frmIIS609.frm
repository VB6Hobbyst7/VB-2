VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIIS609 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '���� ���� â
   Caption         =   "�˻���� ����"
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
      TabIndex        =   14
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&S)"
      Height          =   495
      Left            =   6255
      Style           =   1  '�׷���
      TabIndex        =   13
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������(&C)"
      Height          =   495
      Left            =   8685
      Style           =   1  '�׷���
      TabIndex        =   15
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   3585
      TabIndex        =   19
      Top             =   -30
      Width           =   7545
      Begin VB.CheckBox chkUseFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "������"
         Height          =   270
         Left            =   5250
         TabIndex        =   17
         Top             =   420
         Width           =   1230
      End
      Begin VB.Frame fraTemp 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�µ�����"
         Height          =   1440
         Left            =   255
         TabIndex        =   30
         Top             =   6525
         Width           =   7035
         Begin VB.TextBox txtTempHigh 
            BackColor       =   &H00F7FFF7&
            Height          =   330
            Left            =   4590
            MaxLength       =   5
            TabIndex        =   12
            Top             =   870
            Width           =   1125
         End
         Begin VB.TextBox txtTempLow 
            BackColor       =   &H00F7FFF7&
            Height          =   330
            Left            =   2445
            MaxLength       =   5
            TabIndex        =   11
            Top             =   870
            Width           =   1125
         End
         Begin VB.PictureBox picTempScale 
            BackColor       =   &H00F7FFF7&
            Height          =   360
            Left            =   1965
            ScaleHeight     =   300
            ScaleWidth      =   3690
            TabIndex        =   33
            Top             =   360
            Width           =   3750
            Begin VB.OptionButton optTempScale 
               BackColor       =   &H00F7FFF7&
               Caption         =   "����(C)"
               Height          =   300
               Index           =   0
               Left            =   75
               TabIndex        =   9
               Tag             =   "35136"
               Top             =   15
               Value           =   -1  'True
               Width           =   1230
            End
            Begin VB.OptionButton optTempScale 
               BackColor       =   &H00F7FFF7&
               Caption         =   "ȭ��(F)"
               Height          =   300
               Index           =   1
               Left            =   1740
               TabIndex        =   10
               Tag             =   "35135"
               Top             =   15
               Width           =   1230
            End
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "High"
            Height          =   180
            Left            =   4095
            TabIndex        =   35
            Top             =   930
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "Low"
            Height          =   180
            Left            =   1965
            TabIndex        =   34
            Top             =   930
            Width           =   360
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "�� �����µ�"
            Height          =   180
            Left            =   315
            TabIndex        =   32
            Top             =   930
            Width           =   960
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "�� �µ�����"
            Height          =   180
            Left            =   315
            TabIndex        =   31
            Top             =   420
            Width           =   960
         End
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00F7FFF7&
         Height          =   1920
         Left            =   2220
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   4560
         Width           =   5070
      End
      Begin VB.ComboBox cboVandCd 
         BackColor       =   &H00F7FFF7&
         Height          =   300
         Left            =   2220
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   7
         Top             =   4020
         Width           =   2220
      End
      Begin VB.ComboBox cboLocation 
         BackColor       =   &H00F7FFF7&
         Height          =   300
         Left            =   2220
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   6
         Top             =   3480
         Width           =   2220
      End
      Begin MSComCtl2.DTPicker dtpPurchDt 
         Height          =   300
         Left            =   2220
         TabIndex        =   5
         Top             =   2925
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16252919
         CalendarTitleBackColor=   16252919
         Format          =   133890049
         CurrentDate     =   38048
      End
      Begin VB.PictureBox picEqpDiv 
         BackColor       =   &H00F7FFF7&
         Height          =   360
         Left            =   2220
         ScaleHeight     =   300
         ScaleWidth      =   3690
         TabIndex        =   25
         Top             =   1830
         Width           =   3750
         Begin VB.OptionButton optEqpDiv 
            BackColor       =   &H00F7FFF7&
            Caption         =   "�µ��������"
            Height          =   300
            Index           =   1
            Left            =   1740
            TabIndex        =   3
            Tag             =   "35135"
            Top             =   15
            Width           =   1545
         End
         Begin VB.OptionButton optEqpDiv 
            BackColor       =   &H00F7FFF7&
            Caption         =   "�Ϲ����"
            Height          =   300
            Index           =   0
            Left            =   75
            TabIndex        =   2
            Tag             =   "35136"
            Top             =   15
            Value           =   -1  'True
            Width           =   1560
         End
      End
      Begin VB.TextBox txtModelNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2370
         Width           =   3750
      End
      Begin VB.TextBox txtEqpNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1290
         Width           =   3750
      End
      Begin VB.TextBox txtEqpCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   2220
         MaxLength       =   8
         TabIndex        =   0
         Top             =   390
         Width           =   2505
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ���"
         Height          =   180
         Left            =   255
         TabIndex        =   29
         Top             =   4590
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ����ȸ��"
         Height          =   180
         Left            =   255
         TabIndex        =   28
         Top             =   4055
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� �ǹ�����"
         Height          =   180
         Left            =   255
         TabIndex        =   27
         Top             =   3520
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ��������"
         Height          =   180
         Left            =   255
         TabIndex        =   26
         Top             =   2985
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ����"
         Height          =   180
         Left            =   255
         TabIndex        =   23
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ��񱸺�"
         Height          =   180
         Left            =   255
         TabIndex        =   22
         Top             =   1915
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� �𵨸�(Serial-No)"
         Height          =   180
         Left            =   255
         TabIndex        =   21
         Top             =   2450
         Width           =   1740
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   7500
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ����ڵ�"
         Height          =   180
         Left            =   255
         TabIndex        =   20
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   9900
      Style           =   1  '�׷���
      TabIndex        =   16
      Top             =   8205
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwEqpList 
      Height          =   7665
      Left            =   45
      TabIndex        =   18
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����ڵ�"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   4251
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��񱸺�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�𵨸�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "��������"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ǹ�����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "����ȸ��"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "���"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "�µ�����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Low"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "High"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "�������"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblName 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻���� ����Ʈ"
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
      Left            =   1110
      TabIndex        =   24
      Top             =   165
      Width           =   1455
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
Attribute VB_Name = "frmIIS609"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIIS609.frm (�츮LIS�� �����Ҷ� ���)
'   �ۼ���  :
'   ��  ��  : �˻���� ������
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mEqp As clsIISEqpMaster     '����� ������

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight
        
        '   - ������� �ػ󵵰� ���ص� �׻� ���� ScaleHeight�� �µ��� ����
        .Width = mdiIISMain.ScaleWidth - 4030
    End With

    Call CtlClear
    Set mEqp = New clsIISEqpMaster
    Me.Show
    DoEvents
    
    '## �ǹ�����, ��ü����, ��񸮽�Ʈ �ε�
    Me.MousePointer = vbHourglass
    Call GetLocations
    Call GetVandCds
    Call GetEqps
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS609").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mEqp = Nothing
    Set frmIIS609 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX         As ListItem
    Dim strEqpCd     As String          '����ڵ�
    Dim strTempScale As String          '�µ�����
    
    '## ��ȿ�� Check
    If CheckCode = False Then Exit Sub
    
    '## �Ϲ������ ��� �µ������� Null�� ����
    If fraTemp.Visible = False Then
        strTempScale = ""
    Else
        strTempScale = IIf(optTempScale(0).Value = True, "C", "F")
    End If
    
    Me.MousePointer = vbHourglass

    strEqpCd = Trim(txtEqpCd.Text)
    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)
    
    With mEqp
        .EqpCd = strEqpCd
        .EqpNm = Trim(txtEqpNm.Text)
        .EqpDiv = IIf(optEqpDiv(0).Value = True, "E", "C")
        .ModelNm = Trim(txtModelNm.Text)
        .PurchDt = Format$(dtpPurchDt.Value, "YYYYMMDD")
        .LocationCd = mGetP(cboLocation.Text, 2, DIV)
        .VandCd = mGetP(cboVandCd.Text, 2, DIV)
        .Remark = Trim(txtRemark.Text)
        .TempScale = strTempScale
        .TempLow = CSng(txtTempLow.Text)
        .TempHigh = CSng(txtTempHigh.Text)
        .InUseFg = IIf(chkUseFg.Value = 0, 1, 0)
        
        '## �����ϴ� ����ڵ��̸� Update �������� ������ Insert
        If itmX Is Nothing Then
            If .AddEqp Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� ����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        Else
            If .ModifyEqp Then
                mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
            End If
        End If
    End With
    Call CtlClear
    Call GetEqps
    
    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)
    If Not (itmX Is Nothing) Then
        lvwEqpList.ListItems(itmX.Index).Selected = True
        lvwEqpList.ListItems(itmX.Index).EnsureVisible
    End If
    Set itmX = Nothing
    txtEqpCd.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim itmX        As ListItem
    Dim strEqpCd    As String          '����ڵ�
    Dim intTemp     As Integer

    strEqpCd = Trim(txtEqpCd.Text)
    If strEqpCd = "" Then
        MsgBox "����ڵ带 �Է��ϼ���.", vbInformation, "����"
        Exit Sub
    End If
    
    intTemp = MsgBox("���� �����ұ��?", vbYesNo + vbQuestion, "Ȯ��")
    If intTemp = vbNo Then Exit Sub

    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)
    If itmX Is Nothing Then
        MsgBox "�������� �ʴ� ����ڵ� �Դϴ�.", vbInformation, "����"
        Exit Sub
    End If
    Set itmX = Nothing

    Me.MousePointer = vbHourglass

    With mEqp
        .EqpCd = strEqpCd
        If .DelEqp Then
            mdiIISMain.sbrStatus.Panels(2).Text = "���������� �����Ǿ����ϴ�."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "�����߿� ������ �߻��߽��ϴ�."
        End If
    End With
    Call CtlClear
    Call GetEqps
    txtEqpCd.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    txtEqpCd.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lvwEqpList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwEqpList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwEqpList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## ����ڵ忡 ���� �������� ǥ��
    Call CtlClear
    
    txtEqpCd.Text = Item.Text
    txtEqpNm.Text = Item.SubItems(1)
    optEqpDiv(IIf(Item.SubItems(2) = "E", 0, 1)).Value = True
    txtModelNm.Text = Item.SubItems(3)
    dtpPurchDt.Value = Format$(Item.SubItems(4), "####-##-##")
    cboLocation.ListIndex = mFindComboX(cboLocation, Item.SubItems(5))
    cboVandCd.ListIndex = mFindComboX(cboVandCd, Item.SubItems(6))
    txtRemark = Item.SubItems(7)
    optTempScale(IIf(Item.SubItems(8) = "C", 0, 1)).Value = True
    txtTempLow.Text = Item.SubItems(9)
    txtTempHigh.Text = Item.SubItems(10)
    chkUseFg.Value = IIf(Item.SubItems(11) = "0", 1, 0)
End Sub

Private Sub txtEqpCd_GotFocus()
    With txtEqpCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEqpCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEqpCd_KeyPress(KeyAscii As Integer)
    '## �ҹ��ڰ� �ԷµǸ� �빮�ڷ� ����
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtEqpCd_LostFocus()
    Dim itmX     As ListItem
    Dim strEqpCd As String      '����ڵ�
    
    strEqpCd = Trim(txtEqpCd.Text)
    If strEqpCd = "" Then Exit Sub
    Call CtlClear
    txtEqpCd.Text = strEqpCd
    
    '## �����ϴ� ����ڵ��̸� ��Ŀ���̵�, ����ǥ��
    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)
    If Not (itmX Is Nothing) Then
        lvwEqpList.ListItems(itmX.Index).Selected = True
        lvwEqpList.ListItems(itmX.Index).EnsureVisible
        Call lvwEqpList_ItemClick(itmX)
    End If
    Set itmX = Nothing
End Sub

Private Sub txtEqpNm_GotFocus()
    With txtEqpNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEqpNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtModelNm_GotFocus()
    With txtModelNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtModelNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRemark_GotFocus()
    With txtRemark
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTempLow_GotFocus()
    With txtTempLow
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTempLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTempLow_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack _
        And KeyAscii <> vbKeyDecimal And KeyAscii <> vbKeyDelete _
        And KeyAscii <> vbKeySubtract And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTempLow_Validate(Cancel As Boolean)
    Dim sngTemp As Single
    
    If txtTempLow.Text = "" Then
        MsgBox "�����µ��� �Է����ּ���.", vbInformation, "����"
        Cancel = True
        Exit Sub
    End If
    
    txtTempLow.Text = Format$(txtTempLow.Text, "0.0")
    sngTemp = CSng(txtTempLow.Text)
    If sngTemp < -99.9 Or sngTemp > 99.9 Then
        MsgBox "�����µ��� �ּ�:-99.9, �ִ�:99.9 ������ ���̾�� �մϴ�.", vbInformation, "����"
        Cancel = True
    End If
End Sub

Private Sub txtTempHigh_GotFocus()
    With txtTempHigh
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTempHigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTempHigh_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack _
        And KeyAscii <> vbKeyDecimal And KeyAscii <> vbKeyDelete _
        And KeyAscii <> vbKeySubtract And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTempHigh_Validate(Cancel As Boolean)
    Dim sngTemp As Single
    
    If txtTempHigh.Text = "" Then
        MsgBox "�����µ��� �Է����ּ���.", vbInformation, "����"
        Cancel = True
        Exit Sub
    End If
    
    txtTempHigh.Text = Format$(txtTempHigh.Text, "0.0")
    sngTemp = CSng(txtTempHigh.Text)
    If sngTemp < -99.9 Or sngTemp > 99.9 Then
        MsgBox "�����µ��� �ּ�:-99.9, �ִ�:99.9 ������ ���̾�� �մϴ�.", vbInformation, "����"
        Cancel = True
    End If
End Sub

Private Sub optEqpDiv_Click(Index As Integer)
    If Index = 1 Then
        fraTemp.Visible = True
    Else
        fraTemp.Visible = False
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ǹ������� cboLocation�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetLocations()
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    With cboLocation
        .Clear
        
        Set Rs = mEqp.GetLocations
        If Rs.BOF Or Rs.EOF Then GoTo EndLine
        
        Do Until Rs.EOF
            .AddItem Rs.Fields("LOCNM").Value & Space(100) & DIV & Rs.Fields("LOCCD").Value
            Rs.MoveNext
        Loop
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "frmIIS609", "GetLocations", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��ü������ cboVandCd�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetVandCds()
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    With cboVandCd
        .Clear
        
        Set Rs = mEqp.GetVands
        If Rs.BOF Or Rs.EOF Then GoTo EndLine
        
        Do Until Rs.EOF
            .AddItem Rs.Fields("VANDNM").Value & Space(100) & DIV & Rs.Fields("VANDCD").Value
            Rs.MoveNext
        Loop
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "frmIIS609", "GetVandCds", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��񸮽�Ʈ�� lvwEqpList�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetEqps()
    Dim Rs      As ADODB.Recordset
    Dim itmX    As ListItem
    
'On Error GoTo Errors
    With lvwEqpList
        .ListItems.Clear
        
        Set Rs = mEqp.GetEqps
        If Rs.BOF Or Rs.EOF Then GoTo EndLine
        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , Rs.Fields("EQPCD").Value)
            itmX.SubItems(1) = Rs.Fields("EQPNM").Value
            itmX.SubItems(2) = Rs.Fields("EQPDIV").Value
            itmX.SubItems(3) = Rs.Fields("MODELNM").Value & ""
            itmX.SubItems(4) = Rs.Fields("PURCHDT").Value & ""
            itmX.SubItems(5) = Rs.Fields("LOCATION").Value & ""
            itmX.SubItems(6) = Rs.Fields("VANDCD").Value & ""
            itmX.SubItems(7) = Rs.Fields("REMARK").Value & ""
            itmX.SubItems(8) = Rs.Fields("TEMPSCALE").Value & ""
            itmX.SubItems(9) = CStr(Rs.Fields("TEMPLOW").Value & "")
            itmX.SubItems(10) = CStr(Rs.Fields("TEMPHIGH").Value & "")
            itmX.SubItems(11) = Rs.Fields("INUSEFG").Value & ""
            
            Rs.MoveNext
        Loop
        Set itmX = Nothing
        
        
        If .ListItems.Count > 37 Then
            .ColumnHeaders(2).Width = 2210
        Else
            .ColumnHeaders(2).Width = 2410
        End If
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "frmIIS609", "GetEqps", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �Է�, ������ �ʿ��� ������ ��ȿ�� Check
'   ��ȯ : True(��ȿ), False(��ȿ)
'-----------------------------------------------------------------------------'
Private Function CheckCode() As Boolean
    If txtEqpCd.Text = "" Then
        MsgBox "����ڵ带 �Է��ϼ���.", vbInformation, "����"
        CheckCode = False
        Exit Function
    End If
    
    If txtEqpNm.Text = "" Then
        MsgBox "������ �Է��ϼ���.", vbInformation, "����"
        CheckCode = False
        Exit Function
    End If
    
    If cboLocation.ListIndex = -1 Then
        MsgBox "�ǹ������� �Է��ϼ���.", vbInformation, "����"
        CheckCode = False
        Exit Function
    End If
    
    If txtTempLow = "" Then txtTempLow.Text = "0"
    If txtTempHigh = "" Then txtTempHigh.Text = "0"
    
    CheckCode = True
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtEqpCd.Text = "":         txtEqpNm.Text = ""
    optEqpDiv(0).Value = True:  optTempScale(0).Value = True
    txtModelNm.Text = "":       txtRemark.Text = ""
    cboLocation.ListIndex = -1: cboVandCd.ListIndex = -1
    txtTempLow.Text = "":       txtTempHigh.Text = ""
    dtpPurchDt.Value = Now:     fraTemp.Visible = False
    chkUseFg.Value = 0
End Sub
