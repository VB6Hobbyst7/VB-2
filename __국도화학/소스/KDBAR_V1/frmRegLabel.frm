VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRegLabel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���������"
   ClientHeight    =   12015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20850
   BeginProperty Font 
      Name            =   "���� ���"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12015
   ScaleWidth      =   20850
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   19425
      Begin VB.ComboBox cboComp 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   5
         Top             =   360
         Width           =   3795
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ȸ"
         Height          =   375
         Left            =   5220
         Style           =   1  '�׷���
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ȭ������"
         Height          =   375
         Left            =   6330
         Style           =   1  '�׷���
         TabIndex        =   3
         ToolTipText     =   "����ȭ���� ��� ����ϴ�"
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtComp 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   585
      End
      Begin FPSpread.vaSpread spdRegProd 
         Height          =   4395
         Left            =   8580
         TabIndex        =   37
         Top             =   150
         Visible         =   0   'False
         Width           =   8775
         _Version        =   393216
         _ExtentX        =   15478
         _ExtentY        =   7752
         _StockProps     =   64
         ColsFrozen      =   8
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   21
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmRegLabel.frx":0000
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '����
         Caption         =   "�� ���� "
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   330
         TabIndex        =   6
         Top             =   390
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   90
      TabIndex        =   0
      Top             =   1050
      Width           =   19395
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
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
         Height          =   9375
         Left            =   5220
         TabIndex        =   7
         Top             =   300
         Width           =   14025
         Begin VB.TextBox txtCompNm 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6420
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   36
            Text            =   "ȭ�������"
            Top             =   780
            Width           =   2055
         End
         Begin VB.ComboBox cboProdCd 
            Height          =   375
            ItemData        =   "frmRegLabel.frx":1112
            Left            =   2670
            List            =   "frmRegLabel.frx":1114
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   35
            Top             =   360
            Width           =   3735
         End
         Begin VB.CommandButton cmdRemove 
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            Caption         =   "(-) �׸�����"
            Height          =   405
            Left            =   1830
            Style           =   1  '�׷���
            TabIndex        =   33
            Top             =   3900
            Width           =   1395
         End
         Begin VB.CommandButton cmdAdd 
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Caption         =   "(+) �׸��߰�"
            Height          =   405
            Left            =   420
            Style           =   1  '�׷���
            TabIndex        =   32
            Top             =   3900
            Width           =   1395
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
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
            Height          =   795
            Left            =   10410
            TabIndex        =   28
            Top             =   8550
            Width           =   3525
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00E0E0E0&
               Caption         =   "�ݱ�"
               BeginProperty Font 
                  Name            =   "���� ���"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   2340
               Style           =   1  '�׷���
               TabIndex        =   31
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdOK 
               Appearance      =   0  '���
               BackColor       =   &H00E0E0E0&
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "���� ���"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   60
               Style           =   1  '�׷���
               TabIndex        =   30
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelete 
               BackColor       =   &H00E0E0E0&
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "���� ���"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1200
               Style           =   1  '�׷���
               TabIndex        =   29
               Top             =   150
               Width           =   1095
            End
         End
         Begin VB.ComboBox Combo3 
            Height          =   375
            ItemData        =   "frmRegLabel.frx":1116
            Left            =   2670
            List            =   "frmRegLabel.frx":1118
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   26
            Top             =   2880
            Width           =   3735
         End
         Begin VB.ComboBox cboPrtSide 
            Height          =   375
            ItemData        =   "frmRegLabel.frx":111A
            Left            =   2670
            List            =   "frmRegLabel.frx":111C
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   25
            Top             =   2460
            Width           =   3735
         End
         Begin VB.TextBox txtLabelPrtDefaultNo 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   2
            TabIndex        =   22
            Text            =   "0"
            Top             =   2040
            Width           =   3720
         End
         Begin VB.TextBox txtLabelPrtNo 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   2
            TabIndex        =   20
            Text            =   "0"
            Top             =   1620
            Width           =   3720
         End
         Begin VB.ComboBox cboLabelType 
            Height          =   375
            ItemData        =   "frmRegLabel.frx":111E
            Left            =   2670
            List            =   "frmRegLabel.frx":1120
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   18
            Top             =   1200
            Width           =   3735
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6420
            MaxLength       =   5
            TabIndex        =   11
            Text            =   "ȭ�������"
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtCompCd 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "ȭ�������"
            Top             =   780
            Width           =   3705
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10230
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   9
            Text            =   "ȭ�������"
            Top             =   360
            Width           =   1140
         End
         Begin VB.CheckBox chkUsedYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���"
            Height          =   255
            Left            =   2730
            TabIndex        =   8
            Top             =   3390
            Width           =   795
         End
         Begin FPSpread.vaSpread spdRegLabelDetail 
            Height          =   4125
            Left            =   420
            TabIndex        =   27
            Top             =   4350
            Width           =   13425
            _Version        =   393216
            _ExtentX        =   23680
            _ExtentY        =   7276
            _StockProps     =   64
            ColsFrozen      =   8
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "���� ���"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   15921919
            GridShowVert    =   0   'False
            MaxCols         =   15
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            ScrollBarShowMax=   0   'False
            ShadowColor     =   16775150
            SpreadDesigner  =   "frmRegLabel.frx":1122
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "PPBox Reel �����ڵ�"
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   16
            Left            =   420
            TabIndex        =   24
            Top             =   2880
            Width           =   2190
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "PP Box Reel���"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   420
            TabIndex        =   23
            Top             =   2460
            Width           =   2200
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "�ڽ��� Reel �⺻����"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   420
            TabIndex        =   21
            Top             =   2040
            Width           =   2200
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "����¼���"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   420
            TabIndex        =   19
            Top             =   1620
            Width           =   2200
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "��ǰ����"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   420
            TabIndex        =   17
            Top             =   1200
            Width           =   2200
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "��ǰ����"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   8700
            TabIndex        =   16
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "��뿩��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   420
            TabIndex        =   15
            Top             =   3330
            Width           =   2200
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   420
            TabIndex        =   14
            Top             =   780
            Width           =   2200
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '���� ����
            Caption         =   "��ǰ��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   420
            TabIndex        =   13
            Top             =   360
            Width           =   2200
         End
         Begin VB.Label lblWorkDate 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   "����(M)"
            BeginProperty Font 
               Name            =   "���� ���"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   11400
            TabIndex        =   12
            Top             =   420
            Width           =   975
         End
      End
      Begin FPSpread.vaSpread spdRegLabel 
         Height          =   9255
         Left            =   210
         TabIndex        =   34
         Top             =   390
         Width           =   4935
         _Version        =   393216
         _ExtentX        =   8705
         _ExtentY        =   16325
         _StockProps     =   64
         ColsFrozen      =   8
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   19
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmRegLabel.frx":1D61
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmRegLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmRegLabel.frm
'   �ۼ���  : ������
'   ��  ��  : �� ���
'   �ۼ���  : 2020-02-07
'   ��  ��  : 1.0.0
'   ��  ��  : ����ȭ��
'-----------------------------------------------------------------------------'


Private Sub cboComp_Click()

    txtProdCd.Text = ""
    txtComp.Text = mGetP(cboComp.Text, 2, "|")

    Call GetProdList("", cboComp.Text)

'    Call GetProdList_CodeName("", txtComp.Text)

End Sub


Private Sub cboProdCd_Click()
    
    txtProdCd.Text = Trim(mGetP(cboProdCd.Text, 2, "|"))
    
    Call GetComp_CodeName(txtProdCd.Text)
    
End Sub

'-- �׸��߰�
Private Sub cmdAdd_Click()
    Dim intRow      As Integer
    Dim intNum      As Integer
    Dim intMaxNum   As Integer
    
    intMaxNum = 0
    With spdRegLabelDetail
        For intRow = 1 To .MaxRows
            intNum = GetText(spdRegLabelDetail, intRow, 1)
            If intMaxNum < intNum Then
                intMaxNum = intNum
            End If
        Next
    End With
    intMaxNum = intMaxNum + 1
    
    spdRegLabelDetail.MaxRows = spdRegLabelDetail.MaxRows + 1
        
    'ITEM_NO
    Call SetText(spdRegLabelDetail, intMaxNum, spdRegLabelDetail.MaxRows, 1)
    '����
    Call SetText(spdRegLabelDetail, CStr(spdRegLabelDetail.MaxRows), spdRegLabelDetail.MaxRows, 3)
    '�׸����
    Call SetText(spdRegLabelDetail, "��", spdRegLabelDetail.MaxRows, 5)
    Call SetText(spdRegLabelDetail, "0", spdRegLabelDetail.MaxRows, 6)
    Call SetText(spdRegLabelDetail, "0", spdRegLabelDetail.MaxRows, 7)
    Call SetText(spdRegLabelDetail, "1", spdRegLabelDetail.MaxRows, 9)
    'ȸ��
    Call SetText(spdRegLabelDetail, "0", spdRegLabelDetail.MaxRows, 10)
    '��뱸��
    Call SetText(spdRegLabelDetail, "1", spdRegLabelDetail.MaxRows, 11)

End Sub

Private Sub cmdClear_Click()
        
    spdRegLabel.MaxRows = 0
    spdRegLabelDetail.MaxRows = 0
    
    '-- 1 Line
    txtProdCd.Text = ""
    'txtProdNm.Text = ""
    txtCompCd.Text = ""
    txtProdLen.Text = ""
    cboLabelType.Clear
    cboLabelType.AddItem "Reel"
    cboLabelType.AddItem "PP Box"
    cboLabelType.AddItem "ICE Box"
    cboLabelType.ListIndex = 0
    
    cboPrtSide.Clear
    cboPrtSide.AddItem "�ƴϿ�"
    cboPrtSide.AddItem "��"
    cboPrtSide.ListIndex = 0
    
    '-- 2 Line
    txtLabelPrtNo.Text = "0"
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()

'    Call SetLabel

End Sub

Private Sub cmdOK_Click()

    Call SetLabel

End Sub

Private Sub SetLabel()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intItemNo   As Integer
    
    '�ʼ��Է� üũ
    If txtProdCd.Text = "" Then
        MsgBox "��ǰ�ڵ带 �Է��ϼ���", vbOKOnly + vbCritical, Me.Caption
        txtProdCd.SetFocus
        Exit Sub
    End If
        
    If txtCompCd.Text = "" Then
        MsgBox "���縦 �����ϼ���", vbOKOnly + vbCritical, Me.Caption
        txtCompCd.SetFocus
        Exit Sub
    End If
        
    If txtProdLen.Text = "" Then
        MsgBox "��ǰ���̸� �Է��ϼ���", vbOKOnly + vbCritical, Me.Caption
        txtProdLen.SetFocus
        Exit Sub
    End If
        
'    If txtLabelPrtNo.Text = "" Then
'        MsgBox "����¼����� �Է��ϼ���", vbOKOnly + vbCritical, Me.Caption
'        txtLabelPrtNo.SetFocus
'        Exit Sub
'    End If
        
        
    '-- ���
    gLblInfo.PRODCD = txtProdCd.Text                'Key
    gLblInfo.COMPCD = txtCompCd.Text                'Key
    gLblInfo.LBLTYPE = Mid(cboLabelType.Text, 1, 1) 'Key
    
    gLblInfo.LBLPRTNO = txtLabelPrtNo.Text
    gLblInfo.LBLPRTDEFAULTNO = txtLabelPrtDefaultNo.Text    '�ڽ��� ���⺻������ ��������
    gLblInfo.LBLPRTSIDE = IIf(cboPrtSide.Text = "�ƴϿ�", "N", "Y")
    gLblInfo.LBLBARSIDE1 = ""
    gLblInfo.LBLBARSIDE2 = ""
    gLblInfo.LBLBARSIDE3 = ""
    gLblInfo.LBLBARSIDE4 = ""
    gLblInfo.PRODMAXTOT = "0"

    With spdRegLabelDetail
        gLblMst.PRODCD = txtProdCd.Text                 'Key
        gLblMst.COMPCD = txtCompCd.Text                 'Key
        gLblMst.LBLTYPE = Mid(cboLabelType.Text, 1, 1)  'Key
        ReDim gLblMst.LBLITEM_NO(.MaxRows) As String    'Key
        
        ReDim gLblMst.LBLITEM_SEQ(.MaxRows) As String
        ReDim gLblMst.LBLITEM_NAME(.MaxRows) As String
        ReDim gLblMst.LBLITEM_MEMO(.MaxRows) As String
        ReDim gLblMst.LBLITEM_NMPRT(.MaxRows) As String
        ReDim gLblMst.LBLITEM_BARGU(.MaxRows) As String
        ReDim gLblMst.LBLITEM_BARCD(.MaxRows) As String
        ReDim gLblMst.LBLITEM_X(.MaxRows) As String
        ReDim gLblMst.LBLITEM_Y(.MaxRows) As String
        ReDim gLblMst.LBLITEM_FONT(.MaxRows) As String
        ReDim gLblMst.LBLITEM_ROT(.MaxRows) As String
        ReDim gLblMst.YN(.MaxRows) As String
        
        For intRow = 1 To .DataRowCnt
            gLblMst.LBLITEM_NO(intRow) = GetText(spdRegLabelDetail, intRow, 1)
            gLblMst.LBLITEM_SEQ(intRow) = GetText(spdRegLabelDetail, intRow, 3)
            gLblMst.LBLITEM_NAME(intRow) = GetText(spdRegLabelDetail, intRow, 2)
            gLblMst.LBLITEM_MEMO(intRow) = GetText(spdRegLabelDetail, intRow, 4)
            gLblMst.LBLITEM_NMPRT(intRow) = IIf(GetText(spdRegLabelDetail, intRow, 5) = "��", "Y", "N")
            gLblMst.LBLITEM_X(intRow) = GetText(spdRegLabelDetail, intRow, 6)
            gLblMst.LBLITEM_Y(intRow) = GetText(spdRegLabelDetail, intRow, 7)
            '���ڵ屸�� :
            ' >> ���ڵ��ϰ�� ���ڵ� ����   : "1" : 1D , "2" : 2D
            ' >> �ƴ�    ��� ������ : . /
            If gLblMst.LBLITEM_NAME(intRow) = "���ڵ�" Then
                gLblMst.LBLITEM_BARGU(intRow) = Mid(GetText(spdRegLabelDetail, intRow, 8), 1, 1)
            Else
                gLblMst.LBLITEM_BARGU(intRow) = GetText(spdRegLabelDetail, intRow, 8)
            End If
            
            gLblMst.LBLITEM_BARCD(intRow) = "code128"
            gLblMst.LBLITEM_FONT(intRow) = GetText(spdRegLabelDetail, intRow, 9)
            gLblMst.LBLITEM_ROT(intRow) = GetText(spdRegLabelDetail, intRow, 10)
            gLblMst.YN(intRow) = IIf(GetText(spdRegLabelDetail, intRow, 11) = "1", "Y", "N")
        Next
    End With
    
    If chkUsedYN.Value = "1" Then
        gLblInfo.YN = "Y"
    Else
        gLblInfo.YN = "N"
    End If
                
    '-- Insert / Update ã�ƿ���
    Set AdoRs = Get_LabelList(txtProdCd.Text, txtCompCd.Text, Mid(cboLabelType.Text, 1, 1))
        
    '-- ����
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Label("IN") Then
            '�󼼳��� ����
            For intRow = 1 To spdRegLabelDetail.DataRowCnt
                intItemNo = GetText(spdRegLabelDetail, intRow, 1)
                Call SetLabel_Master(intItemNo)
            Next
            
            'Call CtlInitializing
            'Call GetProdList
        End If
    Else
        'UPDATE
        If Set_Label("UP") Then
            '�󼼳��� ����
            For intRow = 1 To spdRegLabelDetail.DataRowCnt
                intItemNo = GetText(spdRegLabelDetail, intRow, 1)
                Call SetLabel_Master(intItemNo)
            Next
            
            'Call CtlInitializing
            'Call GetProdList
        End If
    End If
    
    
End Sub

Private Sub SetLabel_Master(ByVal pItemNo As Integer, Optional pDelFlag As String)
'    Dim intRow      As Integer
'    Dim intCol      As Integer
    
    '-- Insert / Update ã�ƿ���
    Set AdoRs = Get_LabelMasterList(txtProdCd.Text, txtCompCd.Text, gLblInfo.LBLTYPE, pItemNo)
        
    '-- ����
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Label_Master("IN", pItemNo) Then
'            Call CtlInitializing
'            Call GetProdList
        End If
    Else
        If pDelFlag = "DEL" Then
            'DELETE
            If Set_Label_Master("DEL", pItemNo) Then
    '            Call CtlInitializing
    '            Call GetProdList
            End If
        Else
            'UPDATE
            If Set_Label_Master("UP", pItemNo) Then
    '            Call CtlInitializing
    '            Call GetProdList
            End If
        End If
    End If
    
End Sub
      
      
Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
    
    If pAdoRS Is Nothing Then
        '��ϵ� ���� ����
    Else
        cboProdCd.Clear
        Do Until pAdoRS.EOF
            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop
                            
        If pAdoRS.RecordCount > 0 Then
            cboProdCd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub
    
'ȭ�鿡�� �Ⱥ���(Hiddenó��)
Private Sub GetProdList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    
    Dim strCompNm       As String
    
    Set AdoRs = Get_ProdList(pProdCd, pCompCd)
    
    If AdoRs Is Nothing Then
        '��ϵ� ���� ����
    Else
        Do Until AdoRs.EOF
            With spdRegProd
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegProd, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 1)
                Call SetText(spdRegProd, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 2)
                strCompNm = GetCompList_Name(AdoRs.Fields("COMP_CD").Value & "")
                Call SetText(spdRegProd, strCompNm, .MaxRows, 3)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 4)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 5)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_MATERIAL_CD").Value & "", .MaxRows, 6)
                Call SetText(spdRegProd, AdoRs.Fields("EXPIR_MONTH").Value & "", .MaxRows, 7)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_STOR_TEMP").Value & "", .MaxRows, 8)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_SIZE").Value & "", .MaxRows, 9)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_CHIMEI_PN").Value & "", .MaxRows, 10)
                Call SetText(spdRegProd, AdoRs.Fields("VENDER_CD").Value & "", .MaxRows, 11)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_LINE_FA").Value & "", .MaxRows, 12)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_SLIT_FA").Value & "", .MaxRows, 13)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_CONTROL_YN").Value & "", .MaxRows, 14)
                Call SetText(spdRegProd, AdoRs.Fields("PROD_PCN_NO").Value & "", .MaxRows, 15)
                Call SetText(spdRegProd, AdoRs.Fields("ITEM_BARCODE").Value & "", .MaxRows, 16)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegLabel, "1", .MaxRows, 17)
                Else
                    Call SetText(spdRegLabel, "0", .MaxRows, 17)
                End If
                Call SetText(spdRegProd, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 18)
                Call SetText(spdRegProd, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 19)
                Call SetText(spdRegProd, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 20)
                Call SetText(spdRegProd, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 21)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
    spdRegProd.RowHeight(0) = 12
    spdRegProd.RowHeight(-1) = 12

End Sub

' �� ����Ʈ ������
Private Sub GetLabelList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLabelType As String)
    
    Dim strLabelType    As String
    
    Set AdoRs = Get_LabelList(pProdCd, pCompCd, pLabelType)
    
'        Call SetText(spdRegLabel, "��ǰ�ڵ�", 0, 1):            .ColWidth(1) = 0
'        Call SetText(spdRegLabel, "��ǰ��", 0, 2):              .ColWidth(1) = 10
'        Call SetText(spdRegLabel, "��ǰ����", 0, 3):            .ColWidth(4) = 8
'        Call SetText(spdRegLabel, "��ǰŸ��", 0, 4):            .ColWidth(5) = 8
'        Call SetText(spdRegLabel, "�����ڵ�", 0, 5):          .ColWidth(2) = 0
'        Call SetText(spdRegLabel, "�����", 0, 6):            .ColWidth(3) = 8
'        Call SetText(spdRegLabel, "��¸ż�", 0, 7):            .ColWidth(6) = 8
'        Call SetText(spdRegLabel, "�ڽ��縱�⺻����", 0, 8):    .ColWidth(7) = 10
'        Call SetText(spdRegLabel, "������±���", 0, 9):        .ColWidth(8) = 10
'        Call SetText(spdRegLabel, "�����1Ÿ��", 0, 10):         .ColWidth(9) = 10
'        Call SetText(spdRegLabel, "�����2Ÿ��", 0, 11):        .ColWidth(10) = 8
'        Call SetText(spdRegLabel, "����3Ÿ��", 0, 12):          .ColWidth(11) = 8
'        Call SetText(spdRegLabel, "����4Ÿ��", 0, 13):          .ColWidth(12) = 8
'        Call SetText(spdRegLabel, "�����ִ����", 0, 14):       .ColWidth(13) = 10
'        Call SetText(spdRegLabel, "��뿩��", 0, 15):           .ColWidth(14) = 10
'        Call SetText(spdRegLabel, "�Է���", 0, 16):             .ColWidth(15) = 10
'        Call SetText(spdRegLabel, "�Է��Ͻ�", 0, 17):           .ColWidth(16) = 10
'        Call SetText(spdRegLabel, "������", 0, 18):             .ColWidth(17) = 10
'        Call SetText(spdRegLabel, "�����Ͻ�", 0, 19):           .ColWidth(18) = 10

    If AdoRs Is Nothing Then
        '��ϵ� ���� ����
    Else
        Do Until AdoRs.EOF
            With spdRegLabel
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 1)
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 2)
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 3)
                strLabelType = AdoRs.Fields("PROD_LABEL_TYPE").Value & ""
                Select Case UCase(strLabelType)
                    Case "R": Call SetText(spdRegLabel, "Reel", .MaxRows, 4)
                    Case "P": Call SetText(spdRegLabel, "PP Box", .MaxRows, 4)
                    Case "I": Call SetText(spdRegLabel, "ICE Box", .MaxRows, 4)
                End Select
                Call SetText(spdRegLabel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 5)
                'strCompNm = GetCompList_Name(AdoRs.Fields("COMP_CD").Value & "")
                'Call SetText(spdRegLabel, strCompNm, .MaxRows, 3)
                Call SetText(spdRegLabel, AdoRs.Fields("COMP_NAME").Value & "", .MaxRows, 6)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_NO").Value & "", .MaxRows, 7)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_DEFAULT_NO").Value & "", .MaxRows, 8)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_SIDE").Value & "", .MaxRows, 9)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE01_TYPE").Value & "", .MaxRows, 10)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE02_TYPE").Value & "", .MaxRows, 11)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE03_TYPE").Value & "", .MaxRows, 12)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE04_TYPE").Value & "", .MaxRows, 13)
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_MAX_TOT").Value & "", .MaxRows, 14)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegLabel, "1", .MaxRows, 15)
                Else
                    Call SetText(spdRegLabel, "0", .MaxRows, 15)
                End If
                Call SetText(spdRegLabel, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdRegLabel, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                Call SetText(spdRegLabel, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdRegLabel, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub

' �� ����Ʈ ������
Private Sub GetLabelMaster(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLabelType As String)
    
    Dim strLabelType    As String
    Dim strItemName     As String
    Dim strBarGu        As String
    Dim strBarType      As String
    
    Set AdoRs = Get_LabelMaster(pProdCd, pCompCd, pLabelType)
            
    gLblMst.PRODCD = pProdCd
    gLblMst.COMPCD = pCompCd
    gLblMst.LBLTYPE = pLabelType
    
    If AdoRs Is Nothing Then
        '��ϵ� ���� ����
    Else
        Do Until AdoRs.EOF
            With spdRegLabelDetail
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_NO").Value & "", .MaxRows, 1)
                strItemName = AdoRs.Fields("LABEL_ITEM_NAME").Value & ""
                Call SetText(spdRegLabelDetail, strItemName, .MaxRows, 2)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_SEQ").Value & "", .MaxRows, 3)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_MEMO").Value & "", .MaxRows, 4)
                Call SetText(spdRegLabelDetail, IIf(AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y", "��", "�ƴϿ�"), .MaxRows, 5)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "", .MaxRows, 6)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "", .MaxRows, 7)
                
                strBarGu = AdoRs.Fields("LABEL_ITEM_BAR_GU").Value & ""
                If strItemName = "���ڵ�" Then
                    If strBarGu = "1" Then
                        strBarType = "1D ���ڵ�"
                    ElseIf strBarGu = "2" Then
                        strBarType = "2D ���ڵ�"
                    Else
                        strBarType = strBarGu
                    End If
                Else
                    strBarType = strBarGu
                End If
                
                If strItemName = "���ڵ�" Then
                    .Row = .MaxRows
                    .Col = 8
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "2D ���ڵ�"
                    .TypeComboBoxString = "1D ���ڵ�"
                    .Text = "1D ���ڵ�"
                Else
                    .Row = .MaxRows
                    .Col = 8
                    .CellType = CellTypeEdit
                    .TypeMaxEditLen = 1
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                    
                    Call SetText(spdRegLabelDetail, strBarType, .MaxRows, 8)
                End If
                
                
                
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_FONTSIZE").Value & "", .MaxRows, 9)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_ROT").Value & "", .MaxRows, 10)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegLabelDetail, "1", .MaxRows, 11)
                Else
                    Call SetText(spdRegLabelDetail, "0", .MaxRows, 11)
                End If
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub

Private Sub cmdRemove_Click()
    Dim intItemNo   As Integer
    
    intItemNo = GetText(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, 1)
    
    If MsgBox(GetText(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, 2) & " �׸��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbYes Then
        Call SetLabel_Master(intItemNo, "DEL")
    
        If spdRegLabelDetail.MaxRows > 0 Then
            Call DeleteRow(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, spdRegLabelDetail.ActiveRow)
            spdRegLabelDetail.MaxRows = spdRegLabelDetail.MaxRows - 1
        End If
    End If
    
End Sub

Private Sub cmdSearch_Click()

    If txtComp.Text = "" Then
        Exit Sub
    End If
    
    Call cmdClear_Click
    
    Call GetProdList("", txtComp.Text)
    
    Call GetProdList_CodeName("", txtComp.Text)
    
    Call GetLabelList("", txtComp.Text)
    
    
'''    txtProdCd.Text = ""
'''    txtComp.Text = mGetP(cboComp.Text, 2, "|")
'''
'''    Call GetProdList("", cboComp.Text)
'''
'''    Call GetProdList_CodeName("", txtComp.Text)
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    '���� ����Ʈ ��������
    Call GetCompList_CodeName
    
    '��ǰ ����Ʈ ��������
'    Call GetProdList_CodeName("", txtComp.Text)
    
End Sub


Private Function GetCompList_Name(Optional ByVal pCompCd As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_Name(pCompCd)

    If pAdoRS Is Nothing Then
        '��ϵ� ���� ����
    Else
        Do Until pAdoRS.EOF
            GetCompList_Name = pAdoRS.Fields("COMP_NAME").Value & ""

            pAdoRS.MoveNext
        Loop

    End If

    pAdoRS.Close

End Function

'-- ��� ���縮��Ʈ ��������
Private Sub GetCompList_CodeName()
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_CodeName
    
    If pAdoRS Is Nothing Then
        '��ϵ� ���� ����
    Else
        cboComp.Clear
        
        cboComp.AddItem "��ü" & Space(30) & "|" & "��ü"
        
        Do Until pAdoRS.EOF
            'cboComp.AddItem pAdoRS.Fields("COMP_NAME").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
            cboComp.AddItem pAdoRS.Fields("COMP_NAME").Value & Space(15 - Len(pAdoRS.Fields("COMP_NAME").Value)) & pAdoRS.Fields("COMP_LINE").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
            
            pAdoRS.MoveNext
        Loop
        
        If pAdoRS.RecordCount > 0 Then
            cboComp.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub

'-- ��ǰ���������� �ش� ���� ��������
Private Sub GetComp_CodeName(ByVal pProdCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_Comp_CodeName(pProdCd)
    
    If pAdoRS Is Nothing Then
        '��ϵ� ���� ����
    Else
        txtCompCd.Text = ""
        txtCompNm.Text = ""
        txtProdLen.Text = ""
        
        Do Until pAdoRS.EOF
            txtCompNm.Text = pAdoRS.Fields("COMP_NAME").Value & ""
            txtCompCd.Text = pAdoRS.Fields("COMP_CD").Value & ""
            txtProdLen.Text = pAdoRS.Fields("PROD_LENGTH").Value & ""
            
            pAdoRS.MoveNext
        Loop
        
    End If
    
    pAdoRS.Close
    
End Sub


'-- ��Ʈ���ʱ�ȭ
Private Sub CtlInitializing()
    
    With spdRegProd
        Call SetText(spdRegProd, "��ǰ�ڵ�", 0, 1):         .ColWidth(1) = 0
        Call SetText(spdRegProd, "�����ڵ�", 0, 2):       .ColWidth(2) = 0
        Call SetText(spdRegProd, "�����", 0, 3):         .ColWidth(3) = 8
        Call SetText(spdRegProd, "��ǰ��", 0, 4):           .ColWidth(4) = 11
        Call SetText(spdRegProd, "��ǰ����", 0, 5):         .ColWidth(5) = 10
        Call SetText(spdRegProd, "�����ڵ�", 0, 6):         .ColWidth(6) = 8
        Call SetText(spdRegProd, "��ȿ�Ⱓ", 0, 7):         .ColWidth(7) = 10
        Call SetText(spdRegProd, "�����µ�", 0, 8):         .ColWidth(8) = 10
        Call SetText(spdRegProd, "������", 0, 9):           .ColWidth(9) = 8
        Call SetText(spdRegProd, "CHIMEI�ڵ�", 0, 10):      .ColWidth(10) = 8
        Call SetText(spdRegProd, "VENDOR�ڵ�", 0, 11):      .ColWidth(11) = 8
        Call SetText(spdRegProd, "�������ΰ���", 0, 12):    .ColWidth(12) = 10
        Call SetText(spdRegProd, "SLITTING����", 0, 13):    .ColWidth(13) = 10
        Call SetText(spdRegProd, "��������Ż����", 0, 14):  .ColWidth(14) = 8
        Call SetText(spdRegProd, "PCN����", 0, 15):         .ColWidth(15) = 8
        Call SetText(spdRegProd, "���ڵ�", 0, 16):          .ColWidth(16) = 8
        Call SetText(spdRegProd, "��뿩��", 0, 17):        .ColWidth(17) = 10
        Call SetText(spdRegProd, "�Է���", 0, 18):          .ColWidth(18) = 10
        Call SetText(spdRegProd, "�Է��Ͻ�", 0, 19):        .ColWidth(19) = 10
        Call SetText(spdRegProd, "������", 0, 20):          .ColWidth(20) = 10
        Call SetText(spdRegProd, "�����Ͻ�", 0, 21):        .ColWidth(21) = 10
    
        .MaxRows = 0
    End With
    
    With spdRegLabel
        Call SetText(spdRegLabel, "��ǰ�ڵ�", 0, 1):            .ColWidth(1) = 0
        Call SetText(spdRegLabel, "��ǰ��", 0, 2):              .ColWidth(2) = 12
        Call SetText(spdRegLabel, "��ǰ����", 0, 3):            .ColWidth(3) = 8
        Call SetText(spdRegLabel, "��ǰŸ��", 0, 4):            .ColWidth(4) = 8
        Call SetText(spdRegLabel, "�����ڵ�", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdRegLabel, "�����", 0, 6):            .ColWidth(6) = 9
        Call SetText(spdRegLabel, "��¸ż�", 0, 7):            .ColWidth(7) = 8
        Call SetText(spdRegLabel, "�ڽ��縱�⺻����", 0, 8):    .ColWidth(8) = 10
        Call SetText(spdRegLabel, "������±���", 0, 9):        .ColWidth(9) = 10
        Call SetText(spdRegLabel, "�����1Ÿ��", 0, 10):        .ColWidth(10) = 10
        Call SetText(spdRegLabel, "�����2Ÿ��", 0, 11):        .ColWidth(11) = 8
        Call SetText(spdRegLabel, "����3Ÿ��", 0, 12):          .ColWidth(12) = 8
        Call SetText(spdRegLabel, "����4Ÿ��", 0, 13):          .ColWidth(13) = 8
        Call SetText(spdRegLabel, "�����ִ����", 0, 14):       .ColWidth(14) = 10
        Call SetText(spdRegLabel, "��뿩��", 0, 15):           .ColWidth(15) = 10
        Call SetText(spdRegLabel, "�Է���", 0, 16):             .ColWidth(16) = 10
        Call SetText(spdRegLabel, "�Է��Ͻ�", 0, 17):           .ColWidth(17) = 10
        Call SetText(spdRegLabel, "������", 0, 18):             .ColWidth(18) = 10
        Call SetText(spdRegLabel, "�����Ͻ�", 0, 19):           .ColWidth(19) = 10
    
        .MaxRows = 0
    End With
    
    '-- 1 Line
    txtProdCd.Text = ""
'    txtProdNm.Text = ""
    txtCompCd.Text = ""
    txtCompNm.Text = ""
    txtProdLen.Text = ""
    cboLabelType.Clear
    cboLabelType.AddItem "Reel"
    cboLabelType.AddItem "PP Box"
    cboLabelType.AddItem "ICE Box"
    cboLabelType.ListIndex = 0
    
    cboPrtSide.Clear
    cboPrtSide.AddItem "�ƴϿ�"
    cboPrtSide.AddItem "��"
    cboPrtSide.ListIndex = 0
    
    '-- 2 Line
    txtLabelPrtNo.Text = "0"

    '-- 3 Line
    With spdRegLabelDetail
        Call SetText(spdRegLabelDetail, "No", 0, 1):          .ColWidth(1) = 5
        Call SetText(spdRegLabelDetail, "�׸�", 0, 2):        .ColWidth(2) = 10
        Call SetText(spdRegLabelDetail, "����", 0, 3):        .ColWidth(3) = 6
        Call SetText(spdRegLabelDetail, "����", 0, 4):        .ColWidth(4) = 20
        Call SetText(spdRegLabelDetail, "�׸����", 0, 5):    .ColWidth(5) = 10
        Call SetText(spdRegLabelDetail, "X��ǥ", 0, 6):       .ColWidth(6) = 8
        Call SetText(spdRegLabelDetail, "Y��ǥ", 0, 7):       .ColWidth(7) = 8
        Call SetText(spdRegLabelDetail, "������", 0, 8):      .ColWidth(8) = 8
        Call SetText(spdRegLabelDetail, "��Ʈũ��", 0, 9):    .ColWidth(9) = 8
        Call SetText(spdRegLabelDetail, "ȸ��", 0, 10):        .ColWidth(10) = 8
        Call SetText(spdRegLabelDetail, "��뿩��", 0, 11):   .ColWidth(11) = 12
        Call SetText(spdRegLabelDetail, "�Է���", 0, 12):     .ColWidth(12) = 0
        Call SetText(spdRegLabelDetail, "�Է��Ͻ�", 0, 13):   .ColWidth(13) = 0
        Call SetText(spdRegLabelDetail, "������", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdRegLabelDetail, "�����Ͻ�", 0, 15):   .ColWidth(15) = 0
    
        .MaxRows = 0
    End With
    
    chkUsedYN.Value = "1"
    If gKUKDO.USERGRD = "1" Then
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    
    gSORT = 0

End Sub


Private Sub Label9_DblClick()

    If spdRegProd.Visible = True Then
        spdRegProd.Visible = False
    Else
        spdRegProd.Visible = True
    End If

End Sub

Private Sub spdRegLabel_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i           As Integer
    Dim strPrtSide  As String
    
    If Row = 0 Then
        Call SetSpreadSort(spdRegLabel)
        Exit Sub
    End If
        
    For i = 0 To cboProdCd.ListCount
        If Trim(mGetP(cboProdCd.List(i), 2, "|")) = GetText(spdRegLabel, Row, 1) Then
            cboProdCd.ListIndex = i
            Exit For
        End If
    Next
    
    '�޺��ڽ����� ó���ǹǷ� ���ʿ�
    'txtProdCd.Text = GetText(spdRegLabel, Row, 1)
    
    txtProdLen.Text = GetText(spdRegLabel, Row, 3)
    For i = 0 To cboLabelType.ListCount
        If cboLabelType.List(i) = GetText(spdRegLabel, Row, 4) Then
            cboLabelType.ListIndex = i
            Exit For
        End If
    Next
    
    txtLabelPrtNo.Text = GetText(spdRegLabel, Row, 7)
    txtLabelPrtDefaultNo.Text = GetText(spdRegLabel, Row, 8)
    
    strPrtSide = GetText(spdRegLabel, Row, 9)
    If strPrtSide = "Y" Then
        strPrtSide = "��"
    Else
        strPrtSide = "�ƴϿ�"
    End If
    For i = 0 To cboPrtSide.ListCount
        If cboPrtSide.List(i) = strPrtSide Then
            cboPrtSide.ListIndex = i
            Exit For
        End If
    Next
    
    If GetText(spdRegLabel, Row, 15) = "1" Then
        chkUsedYN.Value = "1"
    Else
        chkUsedYN.Value = "0"
    End If

    spdRegLabelDetail.MaxRows = 0
    
    Call GetLabelMaster(txtProdCd.Text, txtCompCd.Text, Mid(cboLabelType.Text, 1, 1))

    
End Sub



Private Sub spdRegLabelDetail_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    With spdRegLabelDetail
        If Col = 2 Then
            If GetText(spdRegLabelDetail, Row, Col) = "���ڵ�" Then
                .Row = Row
                .Col = 8
                .CellType = CellTypeComboBox
                
                .TypeComboBoxString = "2D ���ڵ�"
                .TypeComboBoxString = "1D ���ڵ�"
                .Text = "1D ���ڵ�"
            Else
            
            End If
        End If
    End With
End Sub
