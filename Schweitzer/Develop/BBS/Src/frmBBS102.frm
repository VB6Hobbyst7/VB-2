VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBBS102 
   BackColor       =   &H00DBE6E6&
   Caption         =   "����ó�����"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   Icon            =   "frmBBS102.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14580
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdOrderView 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ó�溰��ȸ(&C)"
      Height          =   510
      Left            =   7380
      Style           =   1  '�׷���
      TabIndex        =   48
      Top             =   8580
      Width           =   1500
   End
   Begin VB.CheckBox chkAutoPrint 
      BackColor       =   &H00800000&
      Caption         =   "�����Ƿ� �ڵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8280
      TabIndex        =   47
      Top             =   1530
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "���(&P)"
      Enabled         =   0   'False
      Height          =   510
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   43
      Tag             =   "15101"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   42
      Tag             =   "124"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   41
      Tag             =   "128"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&O)"
      Height          =   510
      Left            =   9180
      Style           =   1  '�׷���
      TabIndex        =   40
      Tag             =   "15101"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.Frame fraStore 
      BorderStyle     =   0  '����
      Height          =   2535
      Left            =   10260
      TabIndex        =   33
      Top             =   2580
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox lstLeg 
         Height          =   1680
         Left            =   60
         TabIndex        =   36
         Top             =   420
         Width           =   1215
      End
      Begin VB.ListBox lstRow 
         Height          =   1680
         Left            =   1320
         TabIndex        =   35
         Top             =   420
         Width           =   1215
      End
      Begin VB.ListBox lstCol 
         Height          =   1680
         Left            =   2580
         TabIndex        =   34
         Top             =   420
         Width           =   1215
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   315
         Left            =   60
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   60
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         BackColor       =   8388608
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "�������"
         Appearance      =   0
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1320
         TabIndex        =   39
         Top             =   2220
         Width           =   570
      End
      Begin VB.Label lblCancel 
         AutoSize        =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2100
         TabIndex        =   38
         Top             =   2220
         Width           =   705
      End
   End
   Begin VB.ComboBox cboLeg 
      Height          =   300
      ItemData        =   "frmBBS102.frx":000C
      Left            =   13125
      List            =   "frmBBS102.frx":000E
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   4
      Top             =   1455
      Width           =   990
   End
   Begin VB.CheckBox chkSPos 
      BackColor       =   &H00800000&
      Caption         =   "������� �ڵ��ο�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   10350
      TabIndex        =   3
      Top             =   1545
      Width           =   2055
   End
   Begin VB.ComboBox cboBuilding 
      Height          =   300
      Left            =   11265
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   1
      Top             =   75
      Width           =   3210
   End
   Begin VB.CommandButton cmdRePrint 
      BackColor       =   &H00F4F0F2&
      Height          =   315
      Left            =   14145
      Picture         =   "frmBBS102.frx":0010
      Style           =   1  '�׷���
      TabIndex        =   0
      ToolTipText     =   "�����ǥ�� ������մϴ�."
      Top             =   1455
      Width           =   315
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   255
      Index           =   1
      Left            =   12435
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1500
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   450
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Rack"
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   900
      Top             =   8430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   6690
      Left            =   75
      TabIndex        =   44
      Top             =   1785
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   11800
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
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
      MaxCols         =   43
      MaxRows         =   25
      OperationMode   =   1
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS102.frx":0542
      TextTip         =   4
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   315
      Left            =   75
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1455
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  ó�� ����Ʈ"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  ��ȸ ����"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1140
      Left            =   75
      TabIndex        =   6
      Top             =   315
      Width           =   14400
      Begin VB.CheckBox chkDc 
         BackColor       =   &H00DBE6E6&
         Caption         =   "DC����"
         Height          =   240
         Left            =   11820
         TabIndex        =   27
         Top             =   720
         Value           =   1  'Ȯ��
         Width           =   930
      End
      Begin VB.ComboBox cboInOut 
         Height          =   300
         ItemData        =   "frmBBS102.frx":15F9
         Left            =   4545
         List            =   "frmBBS102.frx":1606
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   16
         Top             =   270
         Width           =   990
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "��ȸ(&Q)"
         Height          =   510
         Left            =   12945
         Style           =   1  '�׷���
         TabIndex        =   15
         Tag             =   "15101"
         Top             =   390
         Width           =   1320
      End
      Begin VB.TextBox txtWardId 
         Appearance      =   0  '���
         Height          =   300
         Left            =   5550
         TabIndex        =   14
         Text            =   "7123456"
         Top             =   270
         Width           =   1110
      End
      Begin VB.CommandButton cmdWardId 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Left            =   6675
         Style           =   1  '�׷���
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   360
      End
      Begin VB.TextBox txtPtId 
         Appearance      =   0  '���
         Height          =   315
         Left            =   9615
         TabIndex        =   12
         Text            =   "7123456"
         Top             =   270
         Width           =   1155
      End
      Begin VB.CommandButton cmdPtId 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   330
         Left            =   10800
         Style           =   1  '�׷���
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   360
      End
      Begin VB.CheckBox chkStat 
         BackColor       =   &H00DBE6E6&
         Caption         =   "����ó�游"
         Height          =   240
         Left            =   10560
         TabIndex        =   10
         Top             =   720
         Width           =   1230
      End
      Begin VB.ComboBox cboOrd 
         Height          =   300
         ItemData        =   "frmBBS102.frx":161C
         Left            =   1200
         List            =   "frmBBS102.frx":1626
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   9
         Top             =   660
         Width           =   3150
      End
      Begin VB.CheckBox chkTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ü"
         Height          =   240
         Left            =   4560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   705
         Width           =   855
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   8535
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "ȯ��ID"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   330
         Left            =   1185
         TabIndex        =   17
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   84017155
         CurrentDate     =   36838
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   330
         Left            =   2910
         TabIndex        =   18
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   84017155
         CurrentDate     =   36838
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   7080
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   11190
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   270
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   285
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�� �� �� ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   675
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "����������"
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   495
         Left            =   5580
         TabIndex        =   21
         Top             =   540
         Width           =   4935
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "�Ϸ�"
            Height          =   255
            Index           =   4
            Left            =   3900
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   180
            Width           =   735
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "�˻���"
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   180
            Width           =   855
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "����"
            Height          =   255
            Index           =   2
            Left            =   1860
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Value           =   1  'Ȯ��
            Width           =   675
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "ä��"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   180
            Value           =   1  'Ȯ��
            Width           =   675
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "ó��"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   675
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "~"
         Height          =   180
         Left            =   2715
         TabIndex        =   32
         Top             =   345
         Width           =   135
      End
      Begin VB.Label lblAge 
         Height          =   195
         Left            =   11505
         TabIndex        =   31
         Top             =   180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblSex 
         Height          =   240
         Left            =   10725
         TabIndex        =   30
         Top             =   180
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�� ���콺������ ��ư�� ����Ͻø� ��ü�߰���û �� �˻���Һ��� ����� ��� ����."
      ForeColor       =   &H00854F3F&
      Height          =   180
      Left            =   75
      TabIndex        =   46
      Top             =   8775
      Width           =   6900
   End
End
Attribute VB_Name = "frmBBS102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcSEL = 1
    tcPTID
    tcPTNM
    tcABO
    tcORDNM
    
    tcORDDT
    tcUNITQTY
    TcMESG
    tcSTATnm
    tcDCNM
    
    tcSTSNM
    tcWARD
    tcROOM
    tcDEPT
    tcSPCNO
    
    tcSTORE
    tcACCNO
    tcCENTERNM
    tcBUSSDIV
    tcORDDTDB
    
    tcORDNO
    tcORDSEQ
    tcSTATFG
    tcDCFG
    tcBedInDT
    
    tCLegRowCol
    tcCENTERCD
    tcNOACCSSS
    tcPHERESIS
    tcSTSCD
    
    tcREASON
    tcDISEASE
    tcDISEASE2
    tcDISEASE3
    tcDISEASE4
    
    tcTime
    tcORDDIV
    tcDUPCHK
    tcREQDT
    tcDOCT
    
    tcTRANSDT
    tcACCDTTM
End Enum


Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1
Private WithEvents objPtInfo    As frmPtInfo
Attribute objPtInfo.VB_VarHelpID = -1
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1

Private Const MENU_ADD& = 1
Private Const MENU_SEP$ = 2
Private Const MENU_XM& = 3

Private Const RowHeight& = 12

Private aryLeg()
Private aryRow()
Private aryCol()
Private SortTF As Boolean

'Private Sub cboDateDiv_Click()
'    tblPtList.MaxRows = 0
'End Sub

Private Sub cboInOut_Click()
    If cboInOut.ListIndex = 0 Then
        txtWardId = ""
        lblWardNm.Caption = ""
        txtWardId.Enabled = False
        cmdWardId.Enabled = False
        
        txtWardId.BackColor = Me.BackColor
    Else
        txtWardId = ""
        lblWardNm.Caption = ""
        txtWardId.Enabled = True
        cmdWardId.Enabled = True
        
        txtWardId.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub cboInOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboOrd_Click()
    tblPtList.MaxRows = 0
End Sub

Private Sub chkTot_Click()
    chkQue(0).value = chkTot.value
    chkQue(1).value = chkTot.value
    chkQue(2).value = chkTot.value
    chkQue(3).value = chkTot.value
    chkQue(4).value = chkTot.value
End Sub

Private Sub cmdClear_Click()
    Call ClearAll
    dtpFrDt.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS102 = Nothing
End Sub

Private Sub cmdOrderView_Click()
    Dim i As Integer
    Dim pFrmName As String
    If Len(txtPtId.Text) < 2 Then GoTo End2Stop

    pFrmName = "frm401ResultView"
    
    medMain.lblSubMenu.Caption = "ó������ȸ" 'medGetP(Button.Tag, 1, "(")
    
    frmLisReview.ButtonKey = "LIS155A" 'Button.Key
    frmLisReview.Ptid = txtPtId.Text
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    Exit Sub

PermissionDenied:
   
'    blnFormShow = False
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"
End2Stop:
End Sub

Private Sub cmdPrint_Click()
    Me.MousePointer = 11
'    Call PrintTransReport
'    Call PrintIntionlize
'    Call PrintHeader_Trans("ȫ�浿", "EM", "0010313", "M", "Dise", "A+", "Trans", "IM", "��ö��", "�ӻ�")
    Call PrintOrderList
    Me.MousePointer = 0
End Sub

Private Sub cmdPtId_Click()
    objPtInfo.Show vbModal
End Sub

Private Sub cmdQuery_Click()
    cmdQuery.tag = "1"
    lblTitle.Caption = " ó�� ����Ʈ"

    If cboInOut.ListIndex = 1 Then
        If txtWardId = "" Then
            MsgBox "������ �����Ͻʽÿ�.", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    If cboInOut.ListIndex = 2 Then
        If txtWardId = "" Then
            MsgBox "������� �����Ͻʽÿ�.", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Me.MousePointer = 11
    
    Call Query
    
'    Call SpreadCellBorder(tblPtList)

    Me.MousePointer = 0
    
    If tblPtList.MaxRows > 0 Then
        cmdPrint.Enabled = True
        tblPtList.SetFocus
    Else
        cmdPrint.Enabled = False
        MsgBox "�ش��ڷᰡ �����ϴ�", vbInformation, Me.Caption
        If cboInOut.ListIndex = 0 Then
        Else
            txtWardId.SetFocus
        End If
    End If
    '2001-11-30�߰�
    cmdCollect.Enabled = True

End Sub

Private Sub cmdRePrint_Click()
    Dim i As Long
    Dim strPtnm As String
    Dim StrWARD As String
    Dim strPtid As String
    Dim strDiease As String
    Dim strABO As String
    Dim strTrans As String
    Dim strDoct As String
    Dim strDept As String
    Dim strSexAge As String
    
    If tblPtList.MaxRows <= 0 Then
        MsgBox "���� ó�泻�� ��ȸ�� �� ����ϼ���.", vbExclamation
        Exit Sub
    End If
    
    '���� �̻��� status �� ��쿡�� ����� ����
    
    tblPtList.Col = TblColumn.tcSTSNM
    tblPtList.Row = tblPtList.ActiveRow
    
    If tblPtList.value = "" Then Exit Sub
    If tblPtList.Row < 1 Then Exit Sub
        
    If tblPtList.value = STS_NM_ORDER Or tblPtList.value = STS_NM_COLLECT Then 'ó��,ä��
        MsgBox "����� ����� �ƴմϴ�. �����̻��� ������ ��쿡�� ������� �� �ֽ��ϴ�.", vbExclamation
        Exit Sub
    End If
    
'    Call PrintDeliveryList(True)
    Call PrintTransList(CStr(tblPtList.ActiveRow))
End Sub

Private Sub cmdWardId_Click()
    
    Set objListPop = New clsPopUpList
    With objListPop
        txtWardId.Text = "": lblWardNm.Caption = ""
        .Connection = DBConn
        .Delimiter = ";"
        Select Case cboInOut.ListIndex
            Case 1
                .FormCaption = "���� ��ȸ": .ColumnHeaderText = "�ڵ�;�ڵ��"
                .LoadPopUp GetSQLWardList
            Case 2
                .FormCaption = "�������ȸ": .ColumnHeaderText = "�ڵ�;�ڵ��"
                .LoadPopUp GetSQLDeptList
        End Select
        
        If .SelectedString <> "" Then
            If txtWardId <> .SelectedItems(0) Then
                tblPtList.MaxRows = 0
            End If
            txtWardId.Text = .SelectedItems(0)
            lblWardNm.Caption = .SelectedItems(1)
            dtpFrDt.SetFocus
        Else
            txtWardId.SetFocus
        End If
    End With
    Set objListPop = Nothing
    
End Sub

Private Sub dtpFrDt_Change()
    tblPtList.MaxRows = 0
End Sub

Private Sub dtpFrDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpToDt_Change()
    tblPtList.MaxRows = 0
End Sub

Private Sub dtpToDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim objAccess   As clsBBSAccess
    Dim objBBSsql   As clsGetSqlStatement
    Dim RS          As Recordset
    Dim Rsord       As Recordset
    Dim ii          As Long
    
    Set objPtInfo = New frmPtInfo
    Set objAccess = New clsBBSAccess
    Set objBBSsql = New clsGetSqlStatement
    Set Rsord = objBBSsql.Get_CompoRecordSet
    
    chkQue(0).Caption = STS_NM_ORDER
    chkQue(1).Caption = STS_NM_COLLECT
    chkQue(2).Caption = STS_NM_ACCESS
    chkQue(3).Caption = STS_NM_INPROGRESS
    chkQue(4).Caption = STS_NM_DONE
    
    With objAccess
        Set RS = New Recordset
        
        RS.Open .Get_LegPos(ObjSysInfo.BuildingCd), DBConn
        
        If RS.EOF = False Then
            cboLeg.Clear
            cboLeg.AddItem ""
            Do Until RS.EOF = True
                cboLeg.AddItem RS.Fields("legcd").value & ""
                RS.MoveNext
            Loop
        End If
        If cboLeg.ListCount <> 0 Then cboLeg.ListIndex = 0
        
    End With
    
    '�˻��׸�
    With Rsord
        cboOrd.Clear
        cboOrd.AddItem "��ü��������"
        For ii = 1 To .RecordCount
             cboOrd.AddItem .Fields("compocd").value & "" & Space(2) & .Fields("abbrnm").value & ""
            .MoveNext
        Next ii
    End With
    
    '�ǹ������� ����� ��� �ǹ�����Ʈ �ε�
    If ObjSysInfo.UseBuildingInfo Then
        cboBuilding.Visible = True
        Call LoadBuilding
    Else
        cboBuilding.Visible = False
    End If
    
    dtpFrDt = DateAdd("d", -3, GetSystemDate)
    dtpToDt = GetSystemDate
    
    cboInOut.ListIndex = 0
    chkStat.value = False
    Call ClearAll
    Me.Show
    
    Set RS = Nothing
    Set Rsord = Nothing
    Set objAccess = Nothing
    Set objBBSsql = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objPtInfo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub lblApply_Click()
    Dim LegCd   As String
    Dim RowNo   As String
    Dim ColNo   As String
    Dim store   As String
    
    Dim i       As Long
    Dim Row     As Long
    Dim spcno   As String
    
    If lstLeg.ListIndex < 0 Then
        Exit Sub
    ElseIf lstLeg.ListIndex > 0 Then
        If lstRow.ListIndex < 0 Then Exit Sub
        If lstCol.ListIndex < 0 Then Exit Sub
    End If
    
    If lstLeg.ListIndex = 0 Then
        LegCd = ""
        RowNo = ""
        ColNo = ""
        store = ""
    Else
        LegCd = lstLeg.Text
        RowNo = lstRow.Text
        ColNo = lstCol.Text
        store = LegCd & "(" & RowNo & "," & ColNo & ")"
    End If
    
    '----------�� ������Ҹ� �ٸ� ��ü��ȣ�� �����س�����?
    If store <> "" Then
        With tblPtList
            .Row = Row
            .Col = TblColumn.tcSPCNO: spcno = .value
            
            For i = 1 To .MaxRows
                .Row = i
                .Col = TblColumn.tcSPCNO
                If spcno <> .value Then
                    .Col = TblColumn.tcSTORE
                    If store = .value Then
                        MsgBox "�̹� �������̰ų� ����������� ����Դϴ�.", vbCritical, Me.Caption
                        Exit Sub
                    End If
                End If
            Next i
        End With
    End If
    
    '----------�ݿ�(���� ��ü��ȣ�̸� ������ҵ� ����)
    Row = Val(fraStore.tag)
    
    With tblPtList
        .Row = Row
        .Col = TblColumn.tcSTORE:     .value = store
                                      .ForeColor = vbBlue
        .Col = TblColumn.tCLegRowCol: .value = LegCd & ";" & RowNo & ";" & ColNo
        
        .Col = TblColumn.tcSPCNO:     spcno = .value
        
        For i = 1 To .MaxRows
            If i <> Row Then
                .Row = i
                .Col = TblColumn.tcSPCNO
                If .value = spcno Then
                    '���� ��ü��ȣ��. ����......
                    .Col = TblColumn.tcSTORE:     .value = store
                                                  .ForeColor = vbBlue
                    .Col = TblColumn.tCLegRowCol: .value = LegCd & ";" & RowNo & ";" & ColNo
                End If
            End If
        Next i
    End With
    
    fraStore.Visible = False
End Sub

Private Sub lblCancel_Click()
    fraStore.Visible = False
End Sub

Private Sub lstLeg_Click()
    Dim i       As Long
    Dim LegCd   As String
    Dim objXM   As clsCrossMatching
    Dim DrRS    As Recordset
    
    lstRow.Clear
    lstCol.Clear
    
    If lstLeg.ListIndex = 0 Then Exit Sub
    
    LegCd = lstLeg.Text
    
    Set objXM = New clsCrossMatching
    
    Set DrRS = New Recordset
    DrRS.Open objXM.Get_Row(LegCd, ObjSysInfo.BuildingCd), DBConn
    
    With DrRS
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                lstRow.AddItem .Fields("rowno").value & ""
                .MoveNext
            Next i
        End If
    End With
    Set DrRS = Nothing
    
    Set DrRS = New Recordset
    DrRS.Open objXM.Get_Col(LegCd, ObjSysInfo.BuildingCd), DBConn
    With DrRS
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                lstCol.AddItem .Fields("colno").value & ""
                .MoveNext
            Next i
        End If
    End With
    Set DrRS = Nothing
    
    Set objXM = Nothing
End Sub

'Private Sub mnuAddSpc_Click()
'
'    With tblPtList
'        .Row = .ActiveRow
'        .Col = TblColumn.tcACCNO
'        frmBBS204.txtAccNo = .value
'        frmBBS204.Show
'    End With
'End Sub

'Private Sub mnuMoveLoc_Click()
''2001-11-29 �߰�
'    Dim objBg       As clsBeginTrans
'    Dim Resp        As VbMsgBoxResult
'    Dim strSpcNo    As String
'    Dim strSQL      As String
'
'
'    Resp = MsgBox("�ش� ȯ���� �˻縦 " & ObjSysInfo.BuildingNm & " �˻�ǿ��� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "�˻���Һ���")
'    If Resp = vbNo Then Exit Sub
'
'    tblPtList.Col = TblColumn.tcSPCNO
'    strSpcNo = tblPtList.value
'
'    Set objBg = New clsBeginTrans
'    strSQL = objBg.Change_Location(medGetP(strSpcNo, 1, "-"), medGetP(strSpcNo, 2, "-"), _
'                                         ObjSysInfo.BuildingCd)
'    Set objBg = Nothing
'
'On Error GoTo Err_Trap
'    DBConn.BeginTrans
'
'    DBConn.Execute strSQL
'
'    DBConn.CommitTrans
'
'    Call Query
'
'    Exit Sub
'
'Err_Trap:
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbCritical, "����"
'End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_ADD
            With tblPtList
                .Row = .ActiveRow
                .Col = TblColumn.tcACCNO
                frmBBS204.txtAccNo = .value
                frmBBS204.Show
            End With
        Case MENU_XM
            With tblPtList
                .Row = .ActiveRow
                .Col = TblColumn.tcACCNO
                DoEvents
                frmBBS201.Show
                frmBBS201.txtSpcNO.Text = Mid(.value, 3)
                frmBBS201.CallByExtForm
            End With
    End Select
        
End Sub

Private Sub objPtInfo_Click(ByVal isSELECT As Boolean, ByVal ptInfo As S2BBS_Library.clsPtInformation)
    txtPtId.Text = "": lblPtNm.Caption = ""
    On Error Resume Next
    If txtPtId.Text <> ptInfo.Ptid Then tblPtList.MaxRows = 0
    txtPtId.Text = ptInfo.Ptid
    lblPtNm.Caption = ptInfo.ptnm

End Sub

Private Function CanSelect(ByVal Col As Long, ByVal Row As Long) As Boolean
    
    Dim objSql   As clsQueryOrder
    Dim CenterCd As String
    Dim noaccess As String
    Dim pheresis As String
    Dim sel      As String
    Dim spcno    As String
    Dim KeepOur  As Long
    Dim i        As Long
    
    '�߰��� ������ �Ұ����� ���̴�.....
    CanSelect = False
    
    With tblPtList
        '��ü��ȣ�� �ִ� �͸� ���
        '������ȣ�� ���� ��(ó�������)�� ���
        '������Ұ� ���� ��(��ü������)�� ���
        'D/Có���� ����
        '��ü�����ð� ������ �����͸� ���
        'irradiation ó���� �ƴ� ó�游 ���
        
        .Row = Row
        
        '�ǹ��ڵ尡 �ٸ��� �����Ҽ� ����.
        .Col = TblColumn.tcCENTERCD: CenterCd = .value
        If CenterCd <> ObjSysInfo.BuildingCd Then Exit Function
        
        'D/C�߻��� ó�濡 ���ؼ��� �����Ҽ� ����.
        .Col = TblColumn.tcDCFG
        If .value = "1" Then Exit Function
        
        '��ü��ȣ�� ������ �����Ҽ� ����.
        .Col = TblColumn.tcSPCNO
        If .value = "" Then Exit Function
        
        '������ȣ�� ������ �����Ҽ� ����.
        .Col = TblColumn.tcACCNO
        If .value <> "" Then Exit Function
        
        '���°� ó���ΰ��� �����Ҽ� ����.
        .Col = TblColumn.tcSTSNM
        If .value = STS_NM_ORDER Then Exit Function '"ó��"
        
        '72�ð��� ���� ��ü�� �����Ҽ� ����.
'        .Col = TblColumn.tcTime
'        If Val(.value) > KeepOur Then Exit Function
        
        'IRRAdiation ó���� �����Ҽ� ����.
        .Col = TblColumn.tcORDDIV
        If .value = "Z" Then Exit Function
    End With

    CanSelect = True
End Function
Private Sub SPreadSort(ByVal Col As Integer)
    With tblPtList
        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = Col
        If SortTF = True Then
            .SortKeyOrder(1) = SortKeyOrderAscending
            SortTF = False
        Else
            SortTF = True
            .SortKeyOrder(1) = SortKeyOrderDescending
        End If
        .Col = 1:  .COL2 = .MaxCols
        .Row = 1:  .Row2 = .MaxRows
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
    End With
End Sub
Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)
    Static BfRow    As Long
    Dim clrBackOdd  As Long
    Dim clrForeOdd  As Long
    Dim clrBackEven As Long
    Dim clrForeEven As Long
    
    Dim CenterCd    As String
    Dim noaccess    As String
    Dim pheresis    As String
    Dim sel         As String
    Dim spcno       As String
    Dim i           As Long
    
    If Row < 1 Then
        Call SPreadSort(Col)
        Exit Sub
    End If
    If Row > tblPtList.MaxRows Then Exit Sub
    If fraStore.Visible = True Then Exit Sub
        
    With tblPtList
    
        Call .GetOddEvenRowColor(clrBackOdd, clrForeOdd, clrBackEven, clrForeEven)
        
        If BfRow <> Row Then
            .Row = BfRow: .Row2 = BfRow
            .Col = 1: .COL2 = .MaxCols
            .BlockMode = True
            If (BfRow Mod 2) = 0 Then
                .BackColor = clrBackEven
            Else
                .BackColor = clrBackOdd
            End If
            .BlockMode = False
        End If
        
        .Row = Row: .Row2 = Row
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .BackColor = .SelBackColor
        .BlockMode = False
        
        BfRow = Row
    End With
    
    
    With tblPtList
        Select Case Col
            Case TblColumn.tcSTORE
                If chkSPos.value = 1 Then Exit Sub
                .Row = Row
                .Col = TblColumn.tcNOACCSSS: noaccess = .value
                .Col = TblColumn.tcCENTERCD: CenterCd = .value
                
                '-------------------���� ��ü������ �ȵ� �͸� ó��.
                If noaccess = "0" Then Exit Sub
                '---------------------�츮 ���Ϳ��� ó���� �� ����.
                If CenterCd <> ObjSysInfo.BuildingCd Then Exit Sub
                
                fraStore.tag = Row
                fraStore.Visible = True
            Case TblColumn.tcSEL
                .Col = Col
                .Row = Row
                If .CellType <> CellTypeCheckBox Then Exit Sub
                
                If CanSelect(Col, Row) = False Then
                    .Col = Col
                    .Row = Row
                    .value = 0
                    Exit Sub
                End If
                
                'pheresis ó���ϰ��� ó���ѰǴ�üũ�� �����ϴ�.....
                .Row = Row
                .Col = TblColumn.tcSPCNO: spcno = .value
                .Col = TblColumn.tcSEL:   sel = .value
                .value = IIf(sel = 1, 0, 1)
'                If pheresis <> "1" Then

                For i = 1 To .MaxRows
                    If i <> Row Then
                        .Row = i
                        .Col = TblColumn.tcORDDIV
                        'irradiationó���ΰ��� �����ϱ� ���ؼ�......
                        If .value = C_WORKAREA Then
                            .Col = TblColumn.tcSPCNO
                            '���� ä����ȣ�� ������...
                            If spcno = .value Then
                                '������ȣ�� ""(�����ʵȰŸ�)....
                                .Col = TblColumn.tcACCNO
                                If .value = "" Then
                                    '����ó���� ������ ���׿� ���ؼ���....
                                    .Col = Col
                                    If .CellType = CellTypeCheckBox Then
                                        .Col = TblColumn.tcSEL
                                        .value = IIf(sel = 1, 0, 1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
'                End If
        End Select
    End With
End Sub

Private Sub tblPtList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Then Exit Sub
    With tblPtList
        .Row = Row
        .Col = TblColumn.tcACCNO
        If .value = "" Then Exit Sub
        .Col = TblColumn.tcSTSNM
        If .value = STS_NM_DONE Or .value = STS_NM_END Then Exit Sub '"�Ϸ�","����"
        .Action = ActionActiveCell
        
        Set objPop = New clsPopupMenu
        With objPop
            .AddMenu MENU_ADD, "��ü�߰���û"
            .AddMenu MENU_SEP, "-"
            .AddMenu MENU_XM, "XM ������"
            
            .PopupMenus Me.hwnd
        End With
        Set objPop = Nothing
'
'
'        Set mnuPopup = frmControl.mnuPopup
'        Set mnuAddSpc = frmControl.mnuSub
'        mnuAddSpc.Caption = "��ü�߰���û"
'        PopupMenu mnuPopup
'        Set mnuPopup = Nothing
'        Set mnuAddSpc = Nothing
    End With
End Sub
Private Function GetTestInformation(ByVal sPtid As String) As String
    Dim objSql As clsCrossMatching
    Dim RS     As Recordset
    Dim strTmp As String
    Dim SSQL   As String
    Dim ii     As Integer
    
    Set objSql = New clsCrossMatching
    SSQL = objSql.TestResultXM(sPtid)
    If SSQL <> "" Then
    Set RS = New Recordset
    RS.Open SSQL, DBConn
        If Not RS.EOF Then
             Do Until RS.EOF
                 strTmp = strTmp & RS.Fields("workarea").value & "" & "-" & _
                          RS.Fields("accdt").value & "" & "-" & _
                          RS.Fields("accseq").value & "" & _
                          "    " & RS.Fields("abbrnm10").value & "" & " : " & _
                          RS.Fields("rstcd").value & "" & vbNewLine & "       "
                RS.MoveNext
            Loop
        End If
        Set RS = Nothing
    End If
    
    If strTmp <> "" Then
        strTmp = "  �� ���ð˻� �� " & vbNewLine & "       " & strTmp
        GetTestInformation = strTmp
    End If
    
    Set objSql = Nothing
End Function
Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim objQuery    As clsQueryOrder
    Dim objDisease  As clsDisease
    Dim RS          As Recordset
'    Dim blnComplete As Boolean
    Dim intord      As Integer
    
    Dim strAccNo     As String  '������ȣ
    Dim strSpcNo     As String  '��ü��ȣ
    Dim strStore     As String  '�������
    Dim StrWARD      As String  '����
    Dim strDept      As String  '�����
    Dim strReason    As String  '��������
    Dim strDisea1    As String  '���ܸ�
    Dim strDisea2    As String  '���ܸ�2
    Dim strDisea3    As String  '���ܸ�3
    Dim strDisea4    As String  '���ܸ�4
    Dim coldttm      As String  '����ð��� ������ �������� ����
    Dim strTime      As String
    Dim strDiseaDisp As String
    Dim strReqDt     As String
    Dim strAccdttm As String
    Dim strMesg      As String
    
    Dim strAccDt    As String
    Dim strAccSeq   As String
    
    'IRRADIATIONó���ΰ��..
    Dim strPtid      As String
    Dim strOrdDt     As String
    Dim strOrdNo     As String
    Dim strROrd      As String
    
    Dim i            As Long
    Dim strtip       As String
    Dim sICSStr         As String
    Dim strTmp          As String
    
    
    Dim blnCompleted As Boolean '�ϷῩ��
    Dim blnAccomplished As Boolean '���Ῡ��
    
    If Row < 1 Then Exit Sub
    
    
    Set objQuery = New clsQueryOrder
    Set objDisease = New clsDisease
    
    With tblPtList
        Call .SetTextTipAppearance("����ü", 9, False, False, &HFFFFC0, vbBlack)
        .Row = Row
        .Col = TblColumn.tcPTID:        strPtid = .value
        .Col = TblColumn.tcACCNO:       strAccNo = .value
        .Col = TblColumn.tcSPCNO:       strSpcNo = .value
        .Col = TblColumn.tcSTORE:       strStore = .value
        .Col = TblColumn.tcWARD:        StrWARD = .value
        .Col = TblColumn.tcDEPT:        strDept = .value
        .Col = TblColumn.tcREQDT:       strReqDt = .Text
        .Col = TblColumn.tcACCDTTM: strAccdttm = .Text
        .Col = TblColumn.TcMESG:        strMesg = .value
        .Col = TblColumn.tcORDDT:       strOrdDt = Replace(.value, "-", "")
        .Col = TblColumn.tcORDNO:       strOrdNo = .value
        '���ܸ��� ���Ѵ�.
        objDisease.Clear
        objDisease.Ptid = strPtid
        objDisease.OrdDt = strOrdDt
        objDisease.ordno = strOrdNo
        
        If objDisease.GetDisease Then
            i = 0
            Do
                If objDisease.EOF Then Exit Do

                If objDisease.DiseaseCd <> "" Then
                    i = i + 1
                    Select Case i
                        Case 1: strDisea1 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 2: strDisea2 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 3: strDisea3 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 4: strDisea4 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                    End Select
                End If
                objDisease.MoveNext
            Loop
        End If
        
        strDiseaDisp = strDisea1
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea2
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea3
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea4
                                               
        '��������
        strReason = objQuery.GetTransReason(strPtid, strOrdDt, strOrdNo): If strReason = "" Then strReason = "(����)"
        
        '----------------------------
        '��ü��� �ð��� ���ϱ����ؼ�
        '----------------------------
        If strSpcNo <> "-" Then
            Set RS = New Recordset
            RS.Open objQuery.Get_spcTime(medGetP(strSpcNo, 1, "-"), medGetP(strSpcNo, 2, "-")), DBConn
            If Not RS.EOF Then
                If Len(RS.Fields("coltm").value & "") = 4 Then
                    coldttm = RS.Fields("coltm").value & "" & "00"
                    coldttm = Format(RS.Fields("coldt").value & "", "0###-##-##") & " " & Format(coldttm, "0#:##:##")
                Else
                    coldttm = Format(RS.Fields("coldt").value & "", "0###-##-##") & " " & Format(RS.Fields("coltm").value & "", "0#:##:##")
                End If
                strTime = DateDiff("h", coldttm, GetSystemDate) & "�ð�"
            End If
            Set RS = Nothing
        End If
        
        .Col = TblColumn.tcORDDIV
        '-----------------------------------------------
        'irradiation ó���ΰ�� �˻����� ó�浵 �����ش�
        '-----------------------------------------------
        If .value = "Z" Then
            
            Set RS = objQuery.GetRelationOrder(strPtid, strOrdDt)
            If Not RS.EOF Then
                With RS
                    Do Until RS.EOF
                        intord = intord + 1
                    '�˻����ΰ�....
                        If .Fields("stscd").value & "" = "3" Then
                            Call CheckCompleted(.Fields("accdt").value & "", .Fields("accseq").value & "", .Fields("unitqty").value & "", _
                                                blnCompleted, blnAccomplished)
'                            blnComplete = CompleteOrderChk(.Fields("accdt").value & "", _
'                                                           .Fields("accseq").value & "", _
'                                                           .Fields("unitqty").value & "")
                            If intord <= 1 Then
                                If blnCompleted = False Then
                                    strROrd = strROrd & "  ����ó�� : " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_INPROGRESS & vbNewLine '�˻���
                                Else
                                    If blnAccomplished Then
                                        strROrd = strROrd & "  ����ó�� : " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_END & vbNewLine '����"
                                    Else
                                        strROrd = strROrd & "  ����ó�� : " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_DONE & vbNewLine '�Ϸ�"
                                    End If
                                End If
                            Else
                                If blnCompleted = False Then
                                    strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_INPROGRESS & vbNewLine '�˻���"
                                Else
                                    If blnAccomplished Then
                                        strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_END & vbNewLine '����"
                                    Else
                                        strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_DONE & vbNewLine '�Ϸ�"
                                    End If
                                End If
                            End If
                            
                        Else
                            If intord <= 1 Then
                                Select Case .Fields("stscd").value & ""
                                    Case "0": strROrd = strROrd & "  ����ó�� : " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_ORDER & vbNewLine 'ó��"
                                    Case "1": strROrd = strROrd & "  ����ó�� : " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_COLLECT & vbNewLine 'ä��"
                                    Case "2": strROrd = strROrd & "  ����ó�� : " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_ACCESS & vbNewLine '����"
                                End Select
                            Else
                                Select Case .Fields("stscd").value & ""
                                    Case "0": strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_ORDER & vbNewLine 'ó��"
                                    Case "1": strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_COLLECT & vbNewLine 'ä��"
                                    Case "2": strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(����:" & .Fields("unitqty").value & "" & ") �� " & STS_NM_ACCESS & vbNewLine '����"
                                End Select
                            End If
                        End If
                        .MoveNext
                    Loop
                End With
            End If
        End If
        
        sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
        
        strtip = "  ������ȣ : [" & strAccNo & "], ��ü��ȣ : [" & strSpcNo & "], ������� : [" & strStore & "]" & vbNewLine & "  ����ð� : " & strTime & vbNewLine & _
                 "  ����/��  : " & StrWARD & "/" & strDept '& vbNewLine & _
                 "  �������� : " & strREASON & vbNewLine & _
                 "  �����Ͻ� : " & strReqDt & vbNewLine & _
                 "  ó���� : " & strMesg & vbNewLine & _
                 strDiseaDisp
        
        If strReason <> "" Then strtip = strtip & vbNewLine & "  �������� : " & strReason
        If strReqDt <> "" Then strtip = strtip & vbNewLine & "  �����Ͻ� : " & strReqDt
        If strAccdttm <> "" Then strtip = strtip & vbNewLine & "  �����Ͻ� : " & strAccdttm
        If strMesg <> "" Then strtip = strtip & vbNewLine & "  ó���� : " & strMesg
        If sICSStr <> "" Then strtip = strtip & vbNewLine & " �������� : " & sICSStr
        
        If strDiseaDisp <> "" Then strtip = strtip & vbNewLine & "  �� �� �� : " & strDiseaDisp
        
        If strROrd <> "" Then strtip = strtip & vbNewLine & Mid(strROrd, 1, Len(strROrd) - 1)
        strtip = strtip & vbNewLine & objQuery.GetAccWorkLoad(strAccNo)
        
        '** �߰� X-Match �󼼰�� By M.G.Choi 2007.11.14
        strtip = strtip & vbNewLine & DetailRst(medGetP(strAccNo, 1, "-"), medGetP(strAccNo, 2, "-"))
        
        strTmp = GetTestInformation(strPtid)
        If strTmp <> "" Then
            strtip = strtip & vbNewLine & strTmp
        End If
        
        TipWidth = 6500 '6350
        MultiLine = 1
        TipText = vbNewLine & strtip & vbNewLine
        ShowTip = True
    End With
    
    Set RS = Nothing
    Set objQuery = Nothing
    Set objDisease = Nothing
    
End Sub

Private Function DetailRst(ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim strSQL      As String
    Dim RS          As New ADODB.Recordset
    Dim strTmp      As String
    Dim strS1       As String
    Dim strS2       As String
    Dim strS3       As String
    Dim strS4       As String
    
    strSQL = " select step1, step2, step3, step4 from " & T_BBS302 & _
             "  where workarea = 'B' " & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq)
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        strS1 = "saline" & IIf(RS.Fields("step1").value & "" = "1", "(O)", "(X)")
        strS2 = "bovine" & IIf(RS.Fields("step2").value & "" = "1", "(O)", "(X)")
        strS3 = "37'C" & IIf(RS.Fields("step3").value & "" = "1", "(O)", "(X)")
        strS4 = "coombs" & IIf(RS.Fields("step4").value & "" = "1", "(O)", "(X)")
        
        strTmp = "  X-match : " & strS1 & "," & strS2 & "," & strS3 & "," & strS4
    End If
    
    RS.Close
    Set RS = Nothing
    
    DetailRst = strTmp
    
End Function

Private Sub txtPtId_GotFocus()
    txtPtId.tag = txtPtId
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = cmdClear.name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = cmdExit.name Then Exit Sub
    
    If txtPtId.tag = txtPtId Then Exit Sub
    If SearchPTINFO = False Then
        txtPtId.SetFocus
    Else
        txtPtId.tag = txtPtId.Text
    End If

End Sub

Private Function SearchPTINFO() As Boolean
    SearchPTINFO = Search_PtInfo
    tblPtList.MaxRows = 0
End Function

Private Sub txtWardId_GotFocus()
    txtWardId.tag = txtWardId
    txtWardId.SelStart = 0
    txtWardId.SelLength = Len(txtWardId)
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If SearchWard = True Then
            txtWardId.tag = txtWardId
            SendKeys "{TAB}"
        Else
            txtWardId.SelStart = 0
            txtWardId.SelLength = Len(txtWardId)
        End If
    End If
End Sub

Private Sub txtWardId_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = cmdClear.name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = cmdExit.name Then Exit Sub
    
    If txtWardId.tag = txtWardId Then Exit Sub
    If SearchWard = False Then txtWardId.SetFocus
End Sub

Private Function SearchWard() As Boolean

    SearchWard = Search_Ward
    
    tblPtList.MaxRows = 0
End Function

Private Sub ClearAll()
    Call ICSPatientMark
    txtWardId = ""
    lblWardNm.Caption = ""
    txtPtId = ""
    lblPtNm.Caption = ""
    tblPtList.MaxRows = 0
    chkSPos.value = 1
    cboOrd.ListIndex = 0
End Sub

Private Function Search_PtInfo() As Boolean
    Dim objPtInfo As clsPtInformation
    Dim DrRS      As Recordset
    Dim ii        As Long
    Dim strLng    As String
    
    If txtPtId = "" Then
        lblPtNm.Caption = ""
        Search_PtInfo = True
    Else
        For ii = 1 To Val(BBS_PTID_LENGTH) - 1
            strLng = strLng & "0"
        Next ii
        

        If Len(Trim(txtPtId.Text)) <> BBS_PTID_LENGTH Then
            txtPtId.Text = Format(txtPtId.Text, strLng & "#")
        End If
        
        '��������
        Call ICSPatientMark(txtPtId.Text, enICSNum.BBS_ALL)
        
        Set objPtInfo = New clsPtInformation
        Set DrRS = New Recordset
        DrRS.Open objPtInfo.Get_Ptid(txtPtId), DBConn
        
        If DrRS.EOF = False Then
            With objPtInfo
                .BedPt_Chk txtPtId.Text, Format(GetSystemDate, PRESENTDATE_FORMAT)
                If .PtDiv = "BED" Then
                    'txtPtId = .ptid
                    lblPtNm.Caption = .ptnm
                    lblSex = .Sex
                    lblAge = .Age
                Else
                    'txtPtId = .ptid
                    lblPtNm.Caption = .ptnm
                    lblSex = .Sex
                    lblAge = .Age
                End If
            End With
            Search_PtInfo = True
        Else
            MsgBox "�ش�Ǵ� ȯ�ڰ� �����ϴ�. Ȯ���� ��ȸ�ϼ���.", vbInformation + vbOKOnly, Me.Caption
            txtPtId = ""
            lblPtNm.Caption = ""
            Search_PtInfo = False
        End If
        Set DrRS = Nothing
        Set objPtInfo = Nothing
    End If
End Function

Private Function Search_Ward() As Boolean
    If txtWardId = "" Then
        lblWardNm.Caption = ""
        Search_Ward = True
    Else
        txtWardId.Text = UCase(txtWardId.Text)
        lblWardNm.Caption = GetWardNm(txtWardId.Text)
        If lblWardNm.Caption = "" Then
            MsgBox "�ش�Ǵ� �ڷᰡ �����ϴ�. Ȯ���� �Է��ϼ���.", vbInformation + vbOKOnly, "�����Է�"
            lblWardNm.Caption = ""
            Search_Ward = False
        End If
    End If
End Function

Private Sub CheckCompleted(ByVal vAccdt As String, ByVal vAccseq As String, ByVal vUnitqty As Long, _
                           ByRef pCompleted As Boolean, ByRef pAccomplished As Boolean)
'2005/05/31 modify by legends
'�ϷῩ�ο� ���Ῡ�θ� ���ϱ� ���� ��ƾ
'�Ϸ� : ó�� ���� ��ŭ �غ�Ǿ� �ִ� ���
'���� : ó�� ���� ��ū ���� ���(��ȯ�ϸ� ���ƴ����� ����)

    Dim objXM As clsCrossMatching
    Dim A_Cnt As Long   'Assign����
    Dim C_Cnt As Long   'Assign Cancel ����
    Dim O_Cnt As Long   '������
    Dim R_Cnt As Long   '��ȯ����
    Dim X_Cnt As Long   '������
    Dim T_Cnt As Long   '��Assign ����
    Dim M_Cnt As Long   '�� ���� ����

    'pCompleted : Assign�� �Ϸ�Ǿ����� ����
    'pAccomplished : ��� �Ϸ�Ǿ����� ����

    'CompleteOrderChk=True�̸� �ϰ�ó��
    'CompleteOrderChk=�̿ϰ�ó��
    Set objXM = New clsCrossMatching
    
    pCompleted = False
    pAccomplished = False
    
    If vAccdt <> "" Then
        With objXM
            .Assign_Cnt vAccdt, Val(vAccseq)
            A_Cnt = .AssignCnt
            C_Cnt = .CancelCnt
            O_Cnt = .OutCnt
            R_Cnt = .RetCnt
            X_Cnt = .ExpCnt
        End With
        Set objXM = Nothing
        
        '������� ������� ó�������, Assign ������ ���Ѵ�.
        '��Assign ����=Assign����-Assign��� ����
        
        T_Cnt = A_Cnt - C_Cnt '���� Assign�� �� ��� Assign�Ǿ����� �Ϸ�
        M_Cnt = O_Cnt - (R_Cnt + X_Cnt) '���� ����-(��ȯ�� ����+���� ����)'���� ���
        
        '���� �ϳ��� ���ϰ� ����θ� �ߴٰ� ��� ����� ����ϸ� �������·� �ѹ�...
        
        '��� ����ߴٰ� ���Ǿ��� ��� ����� ǥ��(��ȯ�� ��� ����)
        'ó��=���=��� �� ��� ����� ǥ��
        
        'vUnitqty : ó�����
        'ó�������ŭ Assign�� �Ǿ����� �Ϸ�, �ƴϸ� �˻���
        If vUnitqty <= T_Cnt Then 'vUnitqty = T_Cnt
            If O_Cnt >= 1 Then '��� �׼��� �ѹ��̶� �� ���
                If M_Cnt >= 1 Then '���� ��� �Ѱ� �̻��� ���
                    pCompleted = True
                End If
            Else '��� �ϳ��� �ȵ� ���
                pCompleted = True
            End If
        Else
            pCompleted = False
        End If
        
'        If vUnitqty <= T_Cnt Then
'            pCompleted = True
'        End If
        
        If vUnitqty = M_Cnt Then
            pAccomplished = True
        End If
        
        '�Ʒ� ������ �߰��Ǿ���.2005/10/24
        If vUnitqty = O_Cnt And O_Cnt = X_Cnt Then
            pCompleted = True
            pAccomplished = True
        End If
    End If
    Set objXM = Nothing
End Sub

'Private Function CompleteOrderChk(ByVal accdt As String, ByVal accseq As String, ByVal unitqty As Long) As Boolean
'    Dim objXM As clsCrossMatching
'    Dim A_Cnt As Long   'Assign����
'    Dim C_Cnt As Long   'Assign Cancel ����
'    Dim O_Cnt As Long   '������
'    Dim R_Cnt As Long   '��ȯ����
'    Dim X_Cnt As Long   '������
'    Dim T_Cnt As Long   '��Assign ����
'
'
'    'CompleteOrderChk=True�̸� �ϰ�ó��
'    'CompleteOrderChk=�̿ϰ�ó��
'    Set objXM = New clsCrossMatching
'    CompleteOrderChk = False
'    If accdt <> "" Then
'
'        With objXM
'            .Assign_Cnt accdt, Val(accseq)
'            A_Cnt = .AssignCnt
'            C_Cnt = .CancelCnt
'            O_Cnt = .OutCnt
'            R_Cnt = .RetCnt
'            X_Cnt = .ExpCnt
'        End With
'        Set objXM = Nothing
'
'        '������� ������� ó�������, Assign ������ ���Ѵ�.
'        '��Assign ����=Assign����-Assign��� ����
'
'        T_Cnt = A_Cnt - C_Cnt
'       ' T_Cnt = A_Cnt - C_Cnt - R_Cnt - X_Cnt
'
'        If unitqty <= T_Cnt Then
'            CompleteOrderChk = True
'        End If
'    End If
'    Set objXM = Nothing
'End Function

'Private Function CheckAccomplished(ByVal vAccdt As String, ByVal vAccseq As String, ByVal vUnitqty As Long) As Boolean
''2005/05/31 Append by legends
''�ϰ� ���� üũ
''ó������� ���� �������� ���� ��� �ϰ� ó��
''��� �� �� ��ȯ���� ���� ��� ����.
'
'    Dim strSql As String
'    Dim Rs As Recordset
'
'    strSql = " select count(*) as cnt from " & T_BBS402
'    strSql = strSql & " where " & DBW("workarea=", "B")
'    strSql = strSql & " and " & DBW("accdt=", vAccdt)
'    strSql = strSql & " and " & DBW("accseq=", vAccseq)
'    strSql = strSql & " and (retfg<>'1' or retfg is not null)"
'
'    Set Rs = New Recordset
'    Rs.Open strSql, DBConn, , , adCmdText
'
'    If Rs.EOF Or Rs.BOF Then
'        CheckAccomplished = False
'    Else
'        CheckAccomplished = True
'    End If
'
'    Set Rs = Nothing
'End Function

Private Function IRR_DUPchk(ByVal Ptid As String, ByVal OrdDt As String) As Boolean
    Dim ii      As Integer
    Dim strTmp  As String
    
    strTmp = Ptid & COL_DIV & OrdDt
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcDUPCHK
            If .value = strTmp Then
                IRR_DUPchk = True
                Exit Function
            End If
        Next
    End With
End Function

Private Function GetABO(ByVal Ptid As String) As String
'������,���ۿ�,��������,���ڵ�,���� ��ȸ�Ѵ�.
    Dim ObjABO As clsABO
    
    Set ObjABO = New clsABO
    With ObjABO
        .Ptid = Ptid
        If .GetABO = True Then
            GetABO = .ABO & .Rh
        Else
            GetABO = ""
        End If
    End With
    Set ObjABO = Nothing
    
End Function

Private Sub Query()
    Dim i           As Long
    Dim j           As Long
    
    Dim RS          As Recordset
    Dim QueryOrder  As clsQueryOrder
    
    Dim accno       As String
    Dim reason      As String
    Dim status      As String
    Dim spcno       As String
    Dim storeleg    As String
    Dim storerow    As String
    Dim storecol    As String
    Dim center      As String
    
    Dim inout       As String
    Dim MaxRowCnt   As Long
    Dim TestDiv     As String
'    Dim blnComplete As Boolean
    
    Dim objPrgBar   As clsProgress
    Dim objDisease  As clsDisease

    
    '���ٰ� ���������̸� ���ڸ� ���߱� ���Ѻ�����
    Dim bkPtId      As String
    Dim bkReason    As String
    Dim bkReqDt     As String
    Dim bkOrdDt     As String
    Dim bkRoomid    As String
    Dim bkWard      As String
    Dim bkDept      As String
    
    Dim strDc       As String
    
    Dim blnCompleted As Boolean
    Dim blnAccomplished As Boolean
    
    tblPtList.MaxRows = 0
    
    Call Save_LegRowCol
    
    Set QueryOrder = New clsQueryOrder
    
    
    If cboOrd.ListIndex <> 0 Then TestDiv = medGetP(cboOrd.Text, 1, " ")
    '-----------
    '���º� ��ȸ
    '-----------
    If chkTot.value Then
        '�̿ϰḸ
        QueryOrder.stscd = "'0','1','2','3'"
        If TRANS_REQUIRE_USED = True Then QueryOrder.stscd = "'0','1','2','3','4'"
    Else
        'If chkAccess.value Then
            'ó��
            If chkQue(0).value Then QueryOrder.stscd = "'0'"
            'ä��
            If chkQue(1).value Then
                If QueryOrder.stscd <> "" Then
                    QueryOrder.stscd = QueryOrder.stscd & ",'1'"
                Else
                    QueryOrder.stscd = "'1'"
                End If
            End If
            '����
            If chkQue(2).value Then
                If TRANS_REQUIRE_USED Then
                    If QueryOrder.stscd <> "" Then
                        QueryOrder.stscd = QueryOrder.stscd & ",'2','3'"
                    Else
                        QueryOrder.stscd = "'2','3'"
                    End If
                Else
                    If QueryOrder.stscd <> "" Then
                        QueryOrder.stscd = QueryOrder.stscd & ",'2'"
                    Else
                        QueryOrder.stscd = "'2'"
                    End If
                End If
            End If
            '�˻���
            If chkQue(3).value Then
                If QueryOrder.stscd <> "" Then
                    If TRANS_REQUIRE_USED Then
                        QueryOrder.stscd = QueryOrder.stscd & ",'3','4'"
                    Else
                        QueryOrder.stscd = QueryOrder.stscd & ",'3'"
                    End If
                Else
                    If TRANS_REQUIRE_USED Then
                        QueryOrder.stscd = "'3','4'"
                    Else
                        QueryOrder.stscd = "'3'"
                    End If
                End If
            End If
            '�ϰ�
            If chkQue(4).value Then
                If chkQue(3).value = False Then
                    If QueryOrder.stscd <> "" Then
                        If TRANS_REQUIRE_USED Then
                            QueryOrder.stscd = QueryOrder.stscd & ",'3','4'"
                        Else
                            QueryOrder.stscd = QueryOrder.stscd & ",'3'"
                        End If
                    Else
                        If TRANS_REQUIRE_USED Then
                            QueryOrder.stscd = "'3','4'"
                        Else
                            QueryOrder.stscd = "'3'"
                        End If
                    End If
                End If
            End If
    End If
    
    Select Case cboInOut.ListIndex
        Case 0: inout = ""
        Case 1: inout = "2"
        Case 2: inout = "1"
    End Select
    If chkDc.value = "1" Then strDc = "1"
    
    Set RS = QueryOrder.QueryOrder(Format(dtpFrDt, PRESENTDATE_FORMAT), Format(dtpToDt, PRESENTDATE_FORMAT), chkStat.value, txtPtId.Text, inout, strDc, txtWardId, TestDiv)
    
    If RS Is Nothing Then
        Set RS = Nothing
        Set QueryOrder = Nothing
        Exit Sub
    End If
    
    
    Set objPrgBar = New clsProgress
    objPrgBar.Container = medMain.stsBar
    
    objPrgBar.Min = 1
    objPrgBar.Max = RS.RecordCount
    
    
    With tblPtList
        bkPtId = ""
        .ReDraw = False
        For i = 1 To RS.RecordCount
        
            objPrgBar.value = i
            
            '�ǹ������� ������ �´�.(��ü������ ����)
            Call QueryOrder.GetSpcNoAndStore(RS.Fields("ptid").value & "", spcno, storeleg, storerow, storecol, center)
            
            '2001-11-23 �߰� :
            '�ǹ������� ����� ���, �׸��� (��ü)�� �ƴҰ�� �ش� �ǹ��� ����Ÿ�� �����ش�.
            '�ǹ��ڵ尡 Ʋ���� �ǳʶڴ�.
            If center = "" Then center = ObjSysInfo.BuildingCd & vbTab & ObjSysInfo.BuildingNm
            

            If ObjSysInfo.UseBuildingInfo = 1 And cboBuilding.ListIndex <> 0 Then
                If medGetP(center, 1, vbTab) <> medGetP(cboBuilding.Text, 1, " ") Then: GoTo Skip
            End If
'            'X-Matching�� �ӻ󺴸��˻��׸񸶽��Ϳ����� �����ϱ⿡ ���������� ������Ѵ�.
'            If (RS.Fields("workarea").value & "") <> "B" And (RS.Fields("workarea").value & "") <> "" Then GoTo SKIP
            
'            blnComplete = CompleteOrderChk(Rs.Fields("accdt").value & "", Rs.Fields("accseq").value & "", Rs.Fields("unitqty").value & "")
            Call CheckCompleted(RS.Fields("accdt").value & "", RS.Fields("accseq").value & "", RS.Fields("unitqty").value & "", _
                                blnCompleted, blnAccomplished)
            '�˻��� or �Ϸ� ��ư ���õǾ�������....
            If chkQue(3).value Or chkQue(4).value Then
                '�Ϸ��ư�� ���õǾ�������.....
                If chkQue(4).value And chkQue(3).value = 0 Then
                    If RS.Fields("orddiv").value & "" = "Z" Then GoTo Skip1
                    'ó��,ä��,���� ��ȸ��....
                    If RS.Fields("stscd").value & "" = "0" Or RS.Fields("stscd").value & "" = "1" Or RS.Fields("stscd").value & "" = "2" Then GoTo Skip1
                    '�˻����� ó���� skip
                    If blnCompleted = False Then GoTo Skip
                    '�˻��߹�ư�� ���õǾ�������...
                ElseIf chkQue(3).value And chkQue(4).value = 0 Then
                    'ó���� �Ϸ� �Ǿ������� skip......
                    If blnCompleted = True Then GoTo Skip
                    If .MaxRows >= 0 And RS.Fields("orddiv").value & "" = "Z" Then
                        If IRR_DUPchk(RS.Fields("ptid").value & "", RS.Fields("orddt").value & "") = False Then GoTo Skip
                    End If
                    '���ǿ� ó��/ä��/������ ���õǾ�������...
                    If RS.Fields("stscd").value & "" = "0" Or RS.Fields("stscd").value & "" = "1" Or RS.Fields("stscd").value & "" = "2" Then GoTo Skip1
                End If
            End If
Skip1:
            MaxRowCnt = MaxRowCnt + 1
            .MaxRows = MaxRowCnt: .RowHeight(-1) = RowHeight
            .Row = MaxRowCnt
            accno = Trim(RS.Fields("accdt").value & "") & "-" & Val(Trim(RS.Fields("accseq").value & ""))
            If accno = "-0" Then accno = "" 'accno = "������"
            
            .Col = TblColumn.tcACCNO:      .value = accno
            .Col = TblColumn.tcPTID:       .value = RS.Fields("ptid").value & ""
            .Col = TblColumn.tcPTNM:       .value = GetPtNm(RS.Fields("ptid").value & "")
            .Col = TblColumn.tcORDNM:      .value = RS.Fields("testnm").value & ""
            .Col = TblColumn.tcORDDT:      .value = Format(RS.Fields("orddt").value & "", "####-##-##")
            .Col = TblColumn.tcUNITQTY:    .value = RS.Fields("unitqty").value & ""
            .Col = TblColumn.tcREASON:     .value = Trim(Trim0(reason))
            .Col = TblColumn.tcREQDT:      .value = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
          
            '2001-11-30�߰�
            '�����ǥ�� ����ǻ�/�ֱټ����� ����ϱ�����
            .Col = TblColumn.tcDOCT:       .value = RS.Fields("orddoct").value & ""
            .Col = TblColumn.tcWARD:       .value = RS.Fields("wardid").value & ""
            .Col = TblColumn.tcROOM:       .value = RS.Fields("hosilid").value & ""
            .Col = TblColumn.tcDEPT:       .value = RS.Fields("deptcd").value & ""
            .Col = TblColumn.tcBUSSDIV:    .value = RS.Fields("bussdiv").value & ""
            .Col = TblColumn.tcORDDTDB:    .value = RS.Fields("orddt").value & ""
            .Col = TblColumn.tcORDNO:      .value = Val(RS.Fields("ordno").value & "")
            .Col = TblColumn.tcORDSEQ:     .value = Val(RS.Fields("ordseq").value & "")
            .Col = TblColumn.tcSTATFG:     .value = RS.Fields("statfg").value & ""
            .Col = TblColumn.tcSTATnm:     .value = IIf(RS.Fields("statfg").value & "" = "1", "Y", ""): .ForeColor = vbRed: .FontBold = True
            .Col = TblColumn.tcBedInDT:    .value = RS.Fields("bedindt").value & ""
            .Col = TblColumn.tcDCFG:       .value = RS.Fields("dcfg").value & ""
            .Col = TblColumn.tcDCNM:       .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", ""): .ForeColor = vbBlue: .FontBold = True
            .Col = TblColumn.tcPHERESIS:   .value = RS.Fields("testdiv").value & ""
            .Col = TblColumn.tcSTSCD:      .value = RS.Fields("stscd").value & ""
            .Col = TblColumn.tcSTSNM
                                            If TRANS_REQUIRE_USED Then
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER: .ForeColor = DCM_Gray '"ó��"
                                                         Case "1": .value = STS_NM_COLLECT '"ä��"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"����"
                                                         Case "3": .value = STS_NM_REQUEST: .ForeColor = DCM_Red '"��û"
                                                                   '����ߴ� ��� ��ȯ�ϰų� ������ߴ� ��� ����� ����ϸ� �˻������� ǥ��...
                                                                   
                                                                   .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_REQUEST) '"����","�Ϸ�","��û"
                                                                   
                                                                   If .value = STS_NM_DONE Then .ForeColor = IIf(blnCompleted, &H8000&, DCM_Red) '"�Ϸ�"
'                                                                   If .value = STS_NM_DONE Then .ForeColor = DCM_Red '"�Ϸ�"
                                                                   If .value = STS_NM_END Then .ForeColor = DCM_Blue '"����"
                                                         Case "4": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS) '"����","�Ϸ�","�˻���"
                                                                   .ForeColor = IIf(blnCompleted, &H8000&, DCM_Brown)
                                                         Case Else: .value = ""
                                                    End Select
                                            Else
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER '"ó��"
                                                         Case "1": .value = STS_NM_COLLECT: .ForeColor = DCM_LightRed '"ä��"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"����"
                                                         Case "3": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS): .ForeColor = DCM_Brown '"����","�Ϸ�","�˻���"
                                                                   If .value = STS_NM_DONE Then .ForeColor = DCM_Red '"�Ϸ�"
                                                                   If .value = STS_NM_END Then .ForeColor = DCM_Blue '"����"
                                                         Case Else: .value = ""
                                                    End Select
                                            End If
                                            
            .Col = TblColumn.TcMESG: .value = RS.Fields("mesg").value & ""
            

            '--------------------------------------------------------------------------------------
            .Col = TblColumn.tcCENTERNM:    .value = medGetP(center, 2, vbTab)
            .Col = TblColumn.tcCENTERCD:    .value = medGetP(center, 1, vbTab)
            
            '�ٸ����Ϳ��ִ� ��üǥ��.
            If medGetP(center, 2, vbTab) <> ObjSysInfo.BuildingNm Then .Col = TblColumn.tcSTORE:   .value = medGetP(center, 1, vbTab)
            'Workareaǥ��
            .Col = TblColumn.tcORDDIV:      .value = RS.Fields("orddiv").value & ""
            '�������ǥ��
            If .value = C_WORKAREA Then
                If storerow = "0" Then storerow = ""
                If storecol = "0" Then storecol = ""
                
                .Col = TblColumn.tCLegRowCol:   .value = storeleg & ";" & storerow & ";" & storecol
                .Col = TblColumn.tcSPCNO:       .value = spcno
                
                If spcno = "" Then
                    .Col = TblColumn.tcSTORE:   .value = "" '.value = "��ä��"
                Else
                    If storeleg = "" Then
                        .Col = TblColumn.tcSTORE:    .value = ""
                        .Col = TblColumn.tcNOACCSSS: .value = "1"
                    Else
                        .Col = TblColumn.tcSTORE:    .value = storeleg & "(" & storerow & "," & storecol & ")"
                        .Col = TblColumn.tcNOACCSSS: .value = "0"
                    End If
                End If
            End If
            
            
            .Col = TblColumn.tcDUPCHK: .value = RS.Fields("ptid").value & "" & COL_DIV & RS.Fields("orddt").value & ""
            .Col = TblColumn.tcTRANSDT: '.value = QueryOrder.GetLatestTrandDt(RS.Fields("ptid").value & "")
            .Col = TblColumn.tcACCDTTM: .value = IIf(RS.Fields("rcvdt").value & "" = "", "", Format(RS.Fields("rcvdt").value & "", "0###-##-##") & " " & Format(RS.Fields("rcvtm").value & "", "0#:##:##"))
            
            
            '���ܸ��� ���Ѵ�.
            Set objDisease = Nothing
            Set objDisease = New clsDisease
            With objDisease
                .Clear
                .Ptid = RS.Fields("ptid").value & ""
                .OrdDt = RS.Fields("orddt").value & ""
                .ordno = RS.Fields("ordno").value & ""
            End With
            
            If objDisease.GetDisease = False Then
                .Col = TblColumn.tcDISEASE: .value = ""
                .Col = TblColumn.tcDISEASE2: .value = ""
                .Col = TblColumn.tcDISEASE3: .value = ""
                .Col = TblColumn.tcDISEASE4: .value = ""
            Else
                j = 0
                Do
                    If objDisease.EOF Then Exit Do
                    
                    If objDisease.DiseaseCd <> "" Then
                        j = j + 1
                        Select Case j
                            Case 1: .Col = TblColumn.tcDISEASE
                            Case 2: .Col = TblColumn.tcDISEASE2
                            Case 3: .Col = TblColumn.tcDISEASE3
                            Case 4: .Col = TblColumn.tcDISEASE4
                        End Select
                        .value = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                    End If
                    objDisease.MoveNext
                Loop
            End If
            Set objDisease = Nothing
            
            '-------------------------
            '�ߺ��Ǵ� ���� �Ⱥ��̰�...
            '-------------------------
            
            If bkPtId <> RS.Fields("ptid").value & "" Then
                bkPtId = RS.Fields("ptid").value & ""
                bkReason = reason
                bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                bkRoomid = RS.Fields("hosilid").value & ""
                bkWard = RS.Fields("wardid").value & ""
                bkDept = RS.Fields("deptcd").value & ""
                
            Else
                .Row = i - 1
                .Col = TblColumn.tcWARD: bkWard = .value
                .Col = TblColumn.tcDEPT: bkDept = .value
                
                .Row = i
                .Col = TblColumn.tcPTID: .ForeColor = .BackColor
                .Col = TblColumn.tcPTNM: .ForeColor = .BackColor
                If bkReason = reason Then
                    If reason <> "(����)" Then .Col = TblColumn.tcREASON: .ForeColor = .BackColor
                Else
                    bkReason = reason
                End If
                If bkWard = RS.Fields("wardid").value & "" Then
                    .Col = TblColumn.tcWARD: .ForeColor = .BackColor
                End If
                If bkDept = RS.Fields("deptcd").value & "" Then
                    .Col = TblColumn.tcDEPT: .ForeColor = .BackColor
                End If
                
                If bkRoomid = RS.Fields("hosilid").value & "" Then
                    .Col = TblColumn.tcROOM: .ForeColor = .BackColor
                Else
                    bkRoomid = RS.Fields("hosilid").value & ""
                End If
'                If bkReqDt = Format(RS.Fields("reqdt").value, "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value, 1, 4), "00:00") Then
'                    .Col = TblColumn.tcREQDT: .ForeColor = .BackColor
'                Else
'                    bkReqDt = Format(RS.Fields("reqdt").value, "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value, 1, 4), "00:00")
'                End If
                If bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##") Then
                    .Col = TblColumn.tcORDDT: .ForeColor = .BackColor
                Else
                    bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                End If
            End If
            
            'Irradiation ó���� ��� �۲��� ��������� ǥ�����ش�.
            .Row = .Row: .Row2 = .Row
            .Col = 1: .COL2 = .MaxCols
            .BlockMode = True
            If RS.Fields("irradfg").value & "" = "1" Then
                .FontBold = True
            Else
                If RS.Fields("statfg").value & "" = "1" Then
                    .Col = TblColumn.tcSTATFG
                    .FontBold = True
                Else
                    .FontBold = False
                End If
            End If
            .BlockMode = False
            
            '2007-06-29 �߰� (�������)
            .Col = 43
            .value = GetOUTDT(RS.Fields("ptid").value & "", RS.Fields("orddt").value & "")
            
Skip:

            RS.MoveNext
        Next i
'        .ReDraw = True
        '�������� �ϰ������� ������´�.
        Set objPrgBar = Nothing
        If .DataRowCnt > 0 Then Call GetBatchABO
    End With
    

    Set QueryOrder = Nothing
End Sub

Private Function GetOUTDT(ByVal pPtId As String, ByVal pOrdDt As String) As String
    Dim RS      As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error Resume Next
    
    strSQL = " select nvl(dschdate,to_char(sysdate,'yyyymmdd')) dschdate " & _
             "   from " & T_HIS002 & _
             "  where patno = " & DBS(pPtId) & _
             "    and nvl(dschdate,to_char(sysdate,'yyyymmdd')) >= " & DBS(pOrdDt)
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        GetOUTDT = RS.Fields("dschdate").value & ""
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

Private Sub GetBatchABO()
    Dim ObjABO      As clsABO
    Dim objPrgBar   As clsProgress
    Dim QueryOrder  As clsQueryOrder
    Dim ii          As Integer
    Dim tmpptid     As String
    Dim sPtid       As String
    Dim sORDDT      As String
    Dim sLastDt     As String
    
    Set ObjABO = New clsABO
    Set objPrgBar = New clsProgress
    Set QueryOrder = New clsQueryOrder
    
    objPrgBar.Container = medMain.stsBar

    With tblPtList
        objPrgBar.Max = .DataRowCnt
        .ReDraw = False
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn.tcPTID
            If tmpptid <> Trim(.value) Then
                sLastDt = ""
                sPtid = .value
                '���������ϱ�
                ObjABO.Ptid = sPtid
                If ObjABO.GetABO = False Then
                    .Col = TblColumn.tcABO:  .value = ""
                Else
                    .Col = TblColumn.tcABO:  .value = ObjABO.ABO & ObjABO.Rh
                End If
                sLastDt = QueryOrder.GetLatestTrandDt(sPtid)
                .Col = TblColumn.tcTRANSDT:  .value = sLastDt
            Else
                .Col = TblColumn.tcABO:      .value = ObjABO.ABO & ObjABO.Rh
                .Col = TblColumn.tcTRANSDT:  .value = sLastDt
            End If
            .Col = TblColumn.tcPTID: tmpptid = Trim(.value)
            If CanSelect(1, ii) Then
                .Row = ii
                .Col = TblColumn.tcSEL
                .CellType = CellTypeCheckBox
                .TypeCheckCenter = True
            Else
                .Row = ii
                .Col = TblColumn.tcSEL
                .CellType = CellTypeStaticText
                .Col = TblColumn.tcSTSNM
                If .value = STS_NM_DONE Or .value = STS_NM_END Then '"�Ϸ�","����"
                    .Col = TblColumn.tcSEL
                    .Text = "��"
                    .ForeColor = vbRed
                End If
            End If
            
            objPrgBar.value = ii: objPrgBar.Message = tmpptid & " �� �������� �˻����Դϴ�."
        Next
        .ReDraw = True
    End With
    
    Set ObjABO = Nothing
    Set QueryOrder = Nothing
    Set objPrgBar = Nothing
End Sub
Private Sub Save_LegRowCol()
'������� ������ �ڵ��� �ƴҰ�� ������Ҹ� �Է¹޾ƾ� �ϹǷ�
'���� ��ư Ŭ�������� �����Ͽ�
'�迭�� ��Ƴ��.
    Dim objXM   As New clsCrossMatching
    Dim DrRS    As New Recordset
    Dim strTmp  As String
    Dim ii      As Integer
    
    lstLeg.Clear
    lstLeg.AddItem "(����)"
    
    DrRS.Open objXM.Get_Leg(ObjSysInfo.BuildingCd), DBConn
    With DrRS
        For ii = 1 To .RecordCount
            lstLeg.AddItem .Fields("legcd").value & ""
            .MoveNext
        Next ii
    End With
    Set DrRS = Nothing
    Set objXM = Nothing
End Sub
Private Function SaveCheckNotAuto() As Boolean
'��������� �Է��� �Ǿ����� üũ�Ѵ�.
    Dim SavePos    As String
    Dim SaveTF     As String
    Dim DcFg       As String
    Dim strRowCol  As String
    Dim strCol     As String
    Dim ii As Integer
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                .Col = TblColumn.tcSTORE: SavePos = .value
                If SavePos <> "" Then
                    SaveCheckNotAuto = True
                Else
                    SaveCheckNotAuto = False
                    Exit Function
                End If
            End If
        Next
    End With
End Function
Private Function Save_Check() As Boolean
    Dim lngColCnt   As Long
    Dim ii          As Long
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                lngColCnt = lngColCnt + 1
                Exit For
            End If
        Next
    End With
    
    '��������� �Է¿��θ� Ȯ���Ѵ�.
    If chkSPos.value = 0 Then
        If SaveCheckNotAuto = False Then
            MsgBox "������Ұ� �����Ǿ����ϴ�." & vbNewLine & "Ȯ���Ͻ��� �����ϼ���.", vbInformation + vbOKOnly, Me.Caption
            Exit Function
        End If
    Else
        If cboLeg.ListIndex < 1 Then
            MsgBox "������� �ڵ� �ο��� ��� Rack�� �ݵ�� �����ϼž� �մϴ�.", vbInformation + vbOKOnly, "������� Rack����"
            Exit Function
        End If
    End If
    
    If lngColCnt = 0 Then
        '�����ϰ��� �ϴ� �Ǽ��� ���Ѵ�
        MsgBox "��������׸��� �����ϴ�.", vbCritical + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If Collect_Cnt = False Then Exit Function
    
    Save_Check = True

End Function

Private Sub cmdCollect_Click()
    Dim objNumbers     As clsBBSNumbers
    Dim objBg          As clsBeginTrans
    Dim RS             As Recordset
    Dim strColDt       As String
    Dim strColTm       As String
    Dim strAccDt       As String
    Dim lngAccNo       As Long
    Dim ii             As Integer
    
'    ������ ���� ������
    Dim strCenterCd As String
    Dim strPtid     As String
    Dim strOrdDt    As String
    Dim strPtnm As String
    Dim strSexAge As String
    Dim StrWARD As String
    Dim strDiease As String
    Dim strABO As String
    Dim strTrans As String
    Dim strDoct As String
    Dim strDept As String
    Dim strTmp      As String
    Dim strSpcYYR   As String
    Dim strFullSpc  As String
    Dim strLeg      As String
    Dim pheresis    As String
    Dim store_cnt   As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngSpcNoR   As Long
    Dim lngOrdseq   As Long
    Dim lngOrdNo    As Long
    Dim blnSave     As Boolean
    
    Dim SSQL        As String
    Dim strRow As String
    
    If Save_Check = False Then Exit Sub
    
    Set objBg = New clsBeginTrans
    
    Me.MousePointer = 11
    strCenterCd = ObjSysInfo.BuildingCd         '�����ڵ�
    strColDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strColTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
    Set objNumbers = New clsBBSNumbers
    With objNumbers
        strAccDt = .Get_AccdtFormat
        lngAccNo = Val(.Get_AccDT_Seq(strAccDt))
    End With
    
On Error GoTo Save_Spc_Error

    DBConn.BeginTrans
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            
            If Val(.value) = 1 Then
                strRow = strRow & ii & COL_DIV
                
                .Col = TblColumn.tcDCFG
                
                .Col = TblColumn.tcPTID:     strPtid = .value
                .Col = TblColumn.tcSPCNO:    strSpcYYR = Mid(.value, 1, 2)
                                             lngSpcNoR = Val(Mid(.value, 4))
                                             strFullSpc = strSpcYYR & CStr(lngSpcNoR)
                .Col = TblColumn.tcPHERESIS: pheresis = IIf(.value = "1", "1", "0")
                
                .Col = TblColumn.tcORDDT:    strOrdDt = Mid(.value, 1, 4) & Mid(.value, 6, 2) & Mid(.value, 9, 2)
                .Col = TblColumn.tcORDNO:    lngOrdNo = Val(.value)
                .Col = TblColumn.tcORDSEQ:   lngOrdseq = Val(.value)
                
                SSQL = objBg.Set_UpdateL101(strPtid, strOrdDt, CStr(lngOrdNo))
                DBConn.Execute SSQL
                
                SSQL = objBg.Set_UpdateL102(strPtid, strOrdDt, CStr(lngOrdNo), CStr(lngOrdseq), strAccDt, CStr(lngAccNo))
                DBConn.Execute SSQL
                
                
                SSQL = objBg.Set_BBS202_Insert(strAccDt, lngAccNo, strPtid, strOrdDt, CStr(lngOrdNo), CStr(lngOrdseq), ObjMyUser.EmpId, pheresis)
                DBConn.Execute SSQL
                
                'OCS ���� Acting Check
'                If OCSActingCheck(strPtid, strOrdDt, CStr(lngOrdNo), CStr(lngOrdseq)) = False Then GoTo Save_Spc_Error
                
               '��ü��ȣ�� �ִ� �� ó�� ������ ���� ���
               '���������� �ƴѰ��� ��ü �ش��ڷ�� �������� �ʴ´�.
               '�̹� ��ü�� �����Ǿ��ִ� ���� ��ü������Ҹ� update ������ �ʴ´�.
               
                .Col = TblColumn.tcACCNO
                If .value = "" And strFullSpc <> "" Then
                    If strTmp <> strPtid Then
                        If chkSPos.value = 0 Then    '����������� ����
                            Set RS = objBg.SavePositionRs(strCenterCd, strSpcYYR, CStr(lngSpcNoR))
                            If Not RS.EOF Then
                                strLeg = RS.Fields("legcd").value & ""
                                lngRow = Val(RS.Fields("rowno").value & "")
                                lngCol = Val(RS.Fields("colno").value & "")
                            Else
                                .Col = TblColumn.tcSTORE
                                strLeg = Mid(.value, 1, 1)
                                lngRow = Val(medGetP(medGetP(.value, 1, ","), 2, "("))
                                lngCol = Val(medGetP(medGetP(.value, 2, ","), 1, ")"))
                                SSQL = objBg.Set_UpdateB206(strCenterCd, strLeg, lngRow, lngCol, strSpcYYR, CStr(lngSpcNoR))
                                DBConn.Execute SSQL
                            End If
                            SSQL = objBg.Set_UpdateB201(strFullSpc, ObjMyUser.EmpId, strLeg, lngRow, lngCol)
                            DBConn.Execute SSQL
                            
                        Else                        '������� �ڵ�����
                            Set RS = objBg.SavePositionRs(strCenterCd, strSpcYYR, CStr(lngSpcNoR))
                            If Not RS.EOF Then
                                strLeg = RS.Fields("legcd").value & ""
                                lngRow = Val(RS.Fields("rowno").value & "")
                                lngCol = Val(RS.Fields("colno").value & "")
                            Else
                                store_cnt = store_cnt + 1
                                strLeg = aryLeg(store_cnt - 1)
                                lngRow = aryRow(store_cnt - 1)
                                lngCol = aryCol(store_cnt - 1)
    
                                SSQL = objBg.Set_UpdateB201(strFullSpc, ObjMyUser.EmpId, strLeg, lngRow, lngCol)
                                DBConn.Execute SSQL
                            End If
                            SSQL = objBg.Set_UpdateB206(strCenterCd, strLeg, lngRow, lngCol, strSpcYYR, CStr(lngSpcNoR))
                            DBConn.Execute SSQL
                            
                        End If
                        Set RS = Nothing
                    End If
                    
                    strTmp = strPtid
                End If
                
                '��ȸ�� �ӵ������� ���ؼ� ������ �ʿ䵥���͸� �����Ѵ�.
                Dim objCollect  As clsBBSCollection
                Dim SQLTmp      As String
                
                Set objCollect = New clsBBSCollection
                SQLTmp = objCollect.Set_AccUnitSQL_203(strPtid, strAccDt, CStr(lngAccNo))
                SSQL = medGetP(SQLTmp, 1, COL_DIV)
                DBConn.Execute SSQL
                SSQL = medGetP(SQLTmp, 2, COL_DIV)
                DBConn.Execute SSQL
                Set objCollect = Nothing
                lngAccNo = lngAccNo + 1
                blnSave = True
            End If
            
        Next ii
    End With
    
    If blnSave = True Then
        SSQL = objNumbers.Set_NumbersCom099(BN_ACC_NO, strAccDt, lngAccNo - 1)
        DBConn.Execute SSQL
    End If
    
    DBConn.CommitTrans
    Call Query
    
    Me.MousePointer = 0
    MsgBox "�����Ǿ����ϴ�.", vbInformation, "����"
    
    If blnSave And (chkAutoPrint.value = 1) Then '���������� ó���� ��쿡 �����ǥ�� ������ش�.
        DoEvents
        Call PrintTransList(strRow)
    End If
    
    Set objBg = Nothing
    Set objNumbers = Nothing
    Exit Sub
    
Save_Spc_Error:
    
    DBConn.RollbackTrans
    Me.MousePointer = 0
    MsgBox "���������� ó������ �ʾҽ��ϴ�.", vbInformation, "��������"
    Set objBg = Nothing
    Set objNumbers = Nothing
End Sub


'Private Function OCSActingCheck(ByVal strPtid As String, ByVal strOrdDt As String, _
'                                ByVal strOrdNo As String, ByVal strOrdSeq As String) As Boolean
'    Dim Rs          As Recordset
'    Dim SqlStmt     As String
'    Dim strOcsOrdNo As String
'    Dim strBussdiv  As String
'
'On Error GoTo Errors
'
'    '������ OCS ���� Table �� Acting_Check�� ���ش�.
'
'    SqlStmt = " SELECT a.ocsordno,b.bussdiv " & _
'              " FROM " & T_LAB101 & " b," & T_LAB102 & " a" & _
'              " WHERE " & DBW("a.ptid =", strPtid) & _
'              " AND " & DBW("a.orddt=", strOrdDt) & _
'              " AND " & DBW("a.ordno=", strOrdNo) & _
'              " AND " & DBW("a.ordseq=", strOrdSeq) & _
'              " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno"
'    Set Rs = New Recordset
'    Rs.Open SqlStmt, DBConn
'
'    If Not Rs.EOF Then
'        strOcsOrdNo = Val(Trim(Rs.Fields("ocsordno").value & ""))
'        strBussdiv = Trim(Rs.Fields("bussdiv").value & "")
'        '������ ipd_order_dmc,ipd_order_update_dmc ������Ʈ
'        '�ܷ��� opd_order_dmc ������Ʈ
'        If strBussdiv = enBussDiv.BussDiv_InPatient Then
'            SqlStmt = " UPDATE med_ocs.ipd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'            DBConn.Execute SqlStmt
'            SqlStmt = " UPDATE med_ocs.ipd_order_update_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'            DBConn.Execute SqlStmt
'        Else
'            SqlStmt = " UPDATE med_ocs.opd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'            DBConn.Execute SqlStmt
'        End If
'    End If
'
'    Set Rs = Nothing
'    OCSActingCheck = True
'    Exit Function
'
'Errors:
'    Set Rs = Nothing
'    OCSActingCheck = False
'End Function


Private Function Collect_Cnt() As Boolean
    Dim objSpec     As clsSpecManagement
    Dim strTmp      As String
    Dim strCollect  As String        '��������...
    Dim strGather   As String         'ä������...
    Dim store_cnt   As Integer
    Dim lngColCnt   As Integer
    Dim ii          As Integer
    
    Set objSpec = New clsSpecManagement

    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                lngColCnt = lngColCnt + 1
                .Col = TblColumn.tcPTID
                If .value <> strTmp Then
                    store_cnt = store_cnt + 1
                End If
                strTmp = .value
            End If
        Next
    End With
    If chkSPos.value = 1 Then
        If lngColCnt <> 0 Then
            With objSpec
                If .Save_Spc_Search(store_cnt, ObjSysInfo.BuildingCd, cboLeg.Text) Then
                    ReDim aryLeg(store_cnt)
                    ReDim aryRow(store_cnt)
                    ReDim aryCol(store_cnt)
                    For ii = 1 To store_cnt
                        aryLeg(ii - 1) = .Leg(ii)
                        aryRow(ii - 1) = .Row(ii)
                        aryCol(ii - 1) = .Col(ii)
                    Next
                    Collect_Cnt = True
                Else
                    Collect_Cnt = False
                End If
            End With
        End If
    End If
    Set objSpec = Nothing

End Function

Private Sub PrintOrderList()
'�������.....ũ����Ż
    Dim strPtid As String, strPtnm As String, strABO As String, strOrdDt As String, STRUNIT As String, strReqDt As String
    Dim strStat As String, STRDCFG As String, STRSTS As String, strSpcNo As String, strSave As String, STRBUILD As String
    Dim StrWARD As String, strDept As String, StrACC As String, strOrdNm As String, STRREAN As String, STRDISEA As String
    Dim strTmp  As String
    
    Dim strRfile   As String
    Dim strRptPath As String
    Dim intFNum    As Integer
    Dim ii         As Integer
    
    Dim sDupChk    As String
    Dim sICSStr    As String

    If tblPtList.MaxRows = 0 Then Exit Sub
    Me.MousePointer = 11
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            STRDISEA = ""
            
            .Col = TblColumn.tcPTID:    strPtid = .value
            
            If sDupChk <> strPtid Then
                sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
                .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
            Else
                .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
            End If
            
            sICSStr = ""
            sDupChk = strPtid
            .Col = TblColumn.tcABO:     strABO = Trim(.value)
            .Col = TblColumn.tcORDNM:   strOrdNm = Trim(.value)
            .Col = TblColumn.tcORDDT:   strOrdDt = Trim(.value)
            .Col = TblColumn.tcUNITQTY: STRUNIT = Trim(.value)
            .Col = TblColumn.tcREASON:  STRREAN = Trim(.value)
            .Col = TblColumn.tcREQDT:    strReqDt = Trim(.value)
            .Col = TblColumn.tcSTATnm:   strStat = Trim(.value)
            .Col = TblColumn.tcDCNM:    STRDCFG = Trim(.value)
            
            .Col = TblColumn.tcSTSNM:   STRSTS = Trim(.value)
            
            .Col = TblColumn.tcDISEASE: STRDISEA = Trim(.value)
            
            If STRDISEA <> "" Then
                .Col = TblColumn.tcDISEASE2
                If .value <> "" Then
                    STRDISEA = STRDISEA & "," & Trim(.value)
                Else
                    STRDISEA = STRDISEA
                End If
                .Col = TblColumn.tcDISEASE3
                If .value <> "" Then
                    STRDISEA = STRDISEA & "," & Trim(.value)
                Else
                    STRDISEA = STRDISEA
                End If
                .Col = TblColumn.tcDISEASE4
                If .value <> "" Then
                    STRDISEA = STRDISEA & "," & Trim(.value)
                Else
                    STRDISEA = STRDISEA
                End If
            End If
                        
            .Col = TblColumn.tcSPCNO:    strSpcNo = Trim(.value)
            .Col = TblColumn.tcSTORE:    strSave = Trim(.value)
            .Col = TblColumn.tcACCNO:    StrACC = Trim(.value)
            
            .Col = TblColumn.tcCENTERNM: STRBUILD = Trim(.value)
            .Col = TblColumn.tcWARD:     StrWARD = Trim(.value)
            .Col = TblColumn.tcDEPT:     strDept = Trim(.value)
            strTmp = strTmp & strPtid & vbTab & strPtnm & vbTab & strABO & vbTab & strOrdDt & vbTab & STRUNIT & vbTab & strReqDt & vbTab & _
                     strStat & vbTab & STRDCFG & vbTab & STRSTS & vbTab & strSpcNo & vbTab & strSave & vbTab & STRBUILD & vbTab & _
                     StrWARD & vbTab & strDept & vbTab & StrACC & vbTab & strOrdNm & vbTab & STRREAN & vbTab & STRDISEA & vbCr
        Next ii
    End With
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)

    strRfile = InstallDir & "BBS\Rpt" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\Rpt" & "\frmBBS102.rpt"
    
    Crystal_Print CReport, strTmp, strRfile, strRptPath
    Me.MousePointer = 0
End Sub

'Private Sub PrintTransReport()
''�������.....ũ����Ż
'    Dim strPtID As String, strPtNm As String, strABO As String, strOrdDt As String, STRUNIT As String, strReqDt As String
'    Dim strStat As String, STRDCFG As String, STRSTS As String, strSpcNo As String, strSave As String, STRBUILD As String
'    Dim StrWARD As String, STRDEPT As String, StrACC As String, strOrdNm As String, STRREAN As String, STRDISEA As String
'
'    Dim ii         As Integer
'
'    Dim sDupChk    As String
'    Dim sICSStr    As String
'
'
'    Dim objPrint   As clsBBSPrint
'
'    Dim strHeader1 As String
'    Dim strHeader2 As String
'    Dim strHeader3 As String
'    Dim strBody    As String
'
'    If tblPtList.MaxRows = 0 Then Exit Sub
'    Me.MousePointer = 11
'    Set objPrint = New clsBBSPrint
'
'    strHeader1 = "����ó�����"
'    strHeader2 = "�� ����� : " & ObjSysInfo.EmpNm & Space(5) & "�� ����� : " & Format(Now, "YYYY-MM-DD HH:MM") & COL_DIV & "5" & COL_DIV & "1"
'    strHeader3 = "��ȣ" & COL_DIV & "5" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "ȯ��ID" & COL_DIV & "15" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "ȯ�ڸ�" & COL_DIV & "35" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "������" & COL_DIV & "75" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "ó������" & COL_DIV & "90" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "����" & COL_DIV & "120" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "�����Ͻ�" & COL_DIV & "130" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "����" & COL_DIV & "170" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "D/C" & COL_DIV & "180" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "����" & COL_DIV & "190" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "��ü��ȣ" & COL_DIV & "205" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "�������" & COL_DIV & "230" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "�˻����" & COL_DIV & "250" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "Location" & COL_DIV & "270" & COL_DIV & "1"
'    strHeader3 = strHeader3 & vbTab & "������ȣ" & COL_DIV & "15" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "ó���" & COL_DIV & "35" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "����" & COL_DIV & "120" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "���ܸ�" & COL_DIV & "205" & COL_DIV & "1"
'
'    With tblPtList
'        For ii = 1 To .MaxRows
'            .Row = ii
'            STRDISEA = ""
'            .Col = TblColumn.tcPTID:    strPtID = .value
'            If sDupChk <> strPtID Then
'                sICSStr = ICSPatientString(strPtID, enICSNum.BBS_ALL)
'                .Col = TblColumn.tcPTNM:    strPtNm = Trim(.value) & sICSStr
'            Else
'                .Col = TblColumn.tcPTNM:    strPtNm = Trim(.value) & sICSStr
'            End If
'            sICSStr = ""
'            sDupChk = strPtID
'            .Col = TblColumn.TcABO:     strABO = Trim(.value)
'            .Col = TblColumn.tcORDNM:   strOrdNm = Trim(.value)
'            .Col = TblColumn.tcORDDT:   strOrdDt = Trim(.value)
'            .Col = TblColumn.tcUNITQTY: STRUNIT = Trim(.value)
'            .Col = TblColumn.tcREASON:  STRREAN = Trim(.value)
'            .Col = TblColumn.tcREQDT:    strReqDt = Trim(.value)
'            .Col = TblColumn.tcSTATnm:   strStat = Trim(.value)
'            .Col = TblColumn.tcDCNM:    STRDCFG = Trim(.value)
'
'            .Col = TblColumn.tcSTSNM:   STRSTS = Trim(.value)
'
'            .Col = TblColumn.tcDISEASE: STRDISEA = Trim(.value)
'
'            If STRDISEA <> "" Then
'                .Col = TblColumn.tcDISEASE2
'                If .value <> "" Then
'                    STRDISEA = STRDISEA & vbTab & Trim(.value)
'                Else
'                    STRDISEA = STRDISEA
'                End If
'                .Col = TblColumn.tcDISEASE3
'                If .value <> "" Then
'                    STRDISEA = STRDISEA & vbTab & Trim(.value)
'                Else
'                    STRDISEA = STRDISEA
'                End If
'                .Col = TblColumn.tcDISEASE4
'                If .value <> "" Then
'                    STRDISEA = STRDISEA & vbTab & Trim(.value)
'                Else
'                    STRDISEA = STRDISEA
'                End If
'            End If
'
'
'            .Col = TblColumn.tcSPCNO:    strSpcNo = Trim(.value)
'            .Col = TblColumn.tcSTORE:    strSave = Trim(.value)
'            .Col = TblColumn.tcACCNO:    StrACC = Trim(.value)
'
'            .Col = TblColumn.tcCENTERNM: STRBUILD = Trim(.value)
'            .Col = TblColumn.tcWARD:     StrWARD = Trim(.value)
'            .Col = TblColumn.tcDEPT:     STRDEPT = Trim(.value)
'            If StrWARD <> "" Then
'                StrWARD = StrWARD & "-" & STRDEPT
'            Else
'                StrWARD = STRDEPT
'            End If
'
'            strBody = strBody & ii & COL_DIV & "5" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strPtID & COL_DIV & "15" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strPtNm & COL_DIV & "35" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strABO & COL_DIV & "75" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strOrdDt & COL_DIV & "90" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRUNIT & COL_DIV & "120" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strReqDt & COL_DIV & "130" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strStat & COL_DIV & "170" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRDCFG & COL_DIV & "180" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRSTS & COL_DIV & "190" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strSpcNo & COL_DIV & "205" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strSave & COL_DIV & "230" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRBUILD & COL_DIV & "250" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & StrWARD & COL_DIV & "270" & COL_DIV & "1" & COL_DIV & "0"
'            strBody = strBody & vbTab & StrACC & COL_DIV & "15" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strOrdNm & COL_DIV & "35" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRREAN & COL_DIV & "120" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRDISEA & COL_DIV & "205" & COL_DIV & "1" & COL_DIV & "1" & vbTab
'
'        Next ii
'    End With
'    strBody = Mid(strBody, 1, Len(strBody) - 1)
'
'    With objPrint
'        .Header1 = strHeader1
'        .Header2 = strHeader2
'        .Header3 = strHeader3
'        .Body = strBody
'        Call .CallPrint("����")
'    End With
'
'    Set objPrint = Nothing
'
'    Me.MousePointer = 0
'End Sub

'2001-11-30�߰�
Private Sub PrintDeliveryList(Optional ByVal blnReprint As Boolean = False)

'�������.....ũ����Ż
    Dim strPtid As String, strPtnm As String, strABO As String, STRUNIT As String, strReqDt As String
    Dim StrWARD As String, strDept As String, strOrdNm As String, STRDISEA As String
    Dim strTmp  As String, strDoct As String, strTransDt As String
    
    Dim strRfile   As String
    Dim strRptPath As String
    Dim intFNum    As Integer
    Dim ii         As Integer
    Dim jj         As Integer
    Dim lngCnt     As Long
    
    Dim sDupChk     As String
    Dim sICSStr     As String
    

    If tblPtList.MaxRows = 0 Then Exit Sub
    Me.MousePointer = 11
    lngCnt = 0
    STRDISEA = ""
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            If ii = 1 Then
                .Col = TblColumn.tcPTID:    strPtid = .value
                
                If sDupChk <> strPtid Then
                    sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
                    .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
                Else
                    .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
                End If
                sICSStr = ""
                sDupChk = strPtid
                
                .Col = TblColumn.tcABO:     strABO = Trim(.value)
                .Col = TblColumn.tcREQDT:   strReqDt = Trim(.value)
                .Col = TblColumn.tcDISEASE: STRDISEA = Trim(.value)
                .Col = TblColumn.tcWARD:    StrWARD = Trim(.value)
                .Col = TblColumn.tcDEPT:    strDept = Trim(.value)
                .Col = TblColumn.tcDOCT:    strDoct = Trim(.value)
                .Col = TblColumn.tcTRANSDT: strTransDt = Trim(.value)
                
                strDoct = GetDoctNm(strDoct)
                strDept = GetDeptNm(strDept)
                
                If STRDISEA <> "" Then
                    .Col = TblColumn.tcDISEASE2
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                    .Col = TblColumn.tcDISEASE3
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                    .Col = TblColumn.tcDISEASE4
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                End If
            End If
            .Col = TblColumn.tcORDNM:   strOrdNm = Trim(.value)
            .Col = TblColumn.tcUNITQTY: STRUNIT = Trim(.value)
            
'            If Not blnReprint Then
                For jj = 1 To Val(STRUNIT)
                    strTmp = strTmp & "" & vbTab & strOrdNm & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                             "" & vbTab & "" & vbTab & "" & vbTab & "" & vbCr
                    lngCnt = lngCnt + 1
                Next
'            End If
        Next ii
    End With

'    If blnReprint Then
'        strTmp = String(23, vbCr)
'    Else
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1) & String(24 - lngCnt, vbCr)
'    End If

    strRfile = InstallDir & "BBS\RPT\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\RPT\frmBBS102_1.rpt"

    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    With CReport
        .ReportFileName = strRptPath
        .ParameterFields(0) = "ptid;" & strPtid & ";TRUE"
        .ParameterFields(1) = "ptnm;" & strPtnm & ";TRUE"
        .ParameterFields(2) = "ward;" & StrWARD & ";TRUE"
        .ParameterFields(3) = "abo;" & strABO & ";TRUE"
        .ParameterFields(4) = "sicknm;" & STRDISEA & ";TRUE"
        .ParameterFields(5) = "doct;" & strDoct & ";TRUE"
        .ParameterFields(6) = "dept;" & strDept & ";TRUE"
        .ParameterFields(7) = "hostnm;" & HOSPITAL_NAME & ";TRUE"
        .ParameterFields(8) = "transdt;" & Format(strTransDt, CS_DateLongMask) & ";TRUE"
        .ParameterFields(9) = "sexage;" & lblSex.Caption & " / " & lblAge.Caption & ";TRUE"
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        .Reset
    End With
    Me.MousePointer = 0
End Sub

Private Sub PrintIntialize()
    PrtLeft = 5
    LineSpace = 6
    lngCurYPos = 10
    
    
    Printer.Font = "����ü"
    Printer.FontSize = 9
    Printer.Orientation = vbPRORPortrait '/* ����
    Printer.ScaleMode = vbMillimeters
    

    Twidth = Printer.ScaleWidth

    LastLineYpos = Printer.ScaleHeight             '����������Y��ġ

End Sub

Private Sub PrintTrans(ByVal vRow As Long)
'����Ʈ ������Ʈ�� ����� ��쿡�� ��
    Dim lngX1 As Long
    Dim lngX2 As Long
    Dim lngX3 As Long
    
    Dim i As Long
    Dim strPtnm As String
    Dim StrWARD As String
    Dim strPtid As String
    Dim strDiease As String
    Dim strABO As String
    Dim strTrans As String
    Dim strDoct As String
    Dim strDept As String
    Dim strSexAge As String
    
    
'ó���� �ٸ� ��쿡 ���
'������ȣ, ��ü��ȣ���� ��������..


    With tblPtList
        For i = 1 To .DataRowCnt
            .Col = TblColumn.tcPTNM: strPtnm = .value
            .Col = TblColumn.tcWARD: StrWARD = .value
            .Col = TblColumn.tcPTID: strPtid = .value
'            .Col = "" 'Sex
            .Col = TblColumn.tcDISEASE: strDiease = .value
            .Col = TblColumn.tcABO: strABO = .value
            .Col = TblColumn.tcTRANSDT: strTrans = .value
'            .Col = "" 'IM
            .Col = TblColumn.tcDOCT: strDoct = .value
            .Col = TblColumn.tcDEPT: strDept = .value
            
            strDoct = GetDoctNm(strDoct)
            strDept = GetDeptNm(strDept)
            
            If strDiease <> "" Then
                .Col = TblColumn.tcDISEASE2
                If .value <> "" Then
                    strDiease = strDiease & "," & Trim(.value)
                Else
                    strDiease = strDiease
                End If
                .Col = TblColumn.tcDISEASE3
                If .value <> "" Then
                    strDiease = strDiease & "," & Trim(.value)
                Else
                    strDiease = strDiease
                End If
                .Col = TblColumn.tcDISEASE4
                If .value <> "" Then
                    strDiease = strDiease & "," & Trim(.value)
                Else
                    strDiease = strDiease
                End If
            End If
            
            Call PrintIntialize
        Next
    End With
    
    
    lngX1 = 10
    lngX2 = lngX1 + Printer.TextWidth("��    �� : ")
    lngX3 = lngX1 + 70
    
    Printer.FontSize = 16: Printer.FontBold = True
    Call Print_Setting("���� ��û �� ��� ��ǥ", PrtLeft, lngCurYPos, Twidth, "C", "C", False)
    Printer.FontSize = 13: Printer.FontBold = False
    
    lngCurYPos = lngCurYPos + 20
    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos + 70), , B 'Box �׸���
    
    '����, ����, ������ ���� Top�� �׸���
    lngCurYPos = lngCurYPos + LineSpace
    Call Print_Setting("��    �� : " & strPtnm, lngX1, LineSpace, , , "C", False)
    Call Print_Setting("��    �� : " & StrWARD, lngX3, LineSpace, , , "C", False)
    Call Print_Setting("   ������ ", 130, LineSpace, , "L", "C", False)
    
    '��Ϲ�ȣ, ����/����, �������� ���� Top�� �׸���
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("��Ϲ�ȣ : " & strPtid, lngX1, LineSpace, , , "C", False)
    Call Print_Setting("����/���� : " & strSexAge, lngX3, LineSpace, , , "C", False)
    Printer.FontBold = True: Printer.FontSize = 40
    Call Print_Setting(strABO, 135, LineSpace, , , "C", False)
    Printer.FontBold = False: Printer.FontSize = 13
    
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("�� �� �� : " & strDiease, lngX1, 10, , , "C", False)
    
    lngCurYPos = lngCurYPos + 10
'    Call Print_Setting("�� �� �� :     �� ��      �� �� " & pTrans, lngX1, 10, , , "C", False)
'    lngCurYPos = lngCurYPos + 10
'    Call Print_Setting("�� �� �� :     �� ��      �� ��  (     ��)" & pIM, lngX1, 10, , , "C", False)
'    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("����ǻ� : " & strDoct, lngX1, 10, , , "C", False)
    Call Print_Setting("�� �� �� : " & strDept, lngX3, 10, , , "C", False)
    
'    lngCurYPos = lngCurYPos + 10
'
'    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
'    Dim ii As Integer
'
'    lngCurYPos = lngCurYPos + 2
'
'    For ii = 1 To 12
'        Printer.Line (PrtLeft, lngCurYPos + 8 * ii)-(Twidth - PrtLeft, lngCurYPos + 8 * ii)
'    Next
'
''���׺���
'    Printer.Line (PrtLeft, lngCurYPos - 2)-(PrtLeft, lngCurYPos + 8 * 12)
'
'    '���׹�ȣ
'    Printer.Line (lngX2, lngCurYPos + 8)-(lngX2, lngCurYPos + 8 * 12)
'    '��������
'    Printer.Line (lngX2 + 30, lngCurYPos + 8)-(lngX2 + 30, lngCurYPos + 8 * 12)
'    '������
'    Printer.Line (lngX2 + 45, lngCurYPos + 8)-(lngX2 + 45, lngCurYPos + 8 * 12)
'    'ä����
'    Printer.Line (lngX2 + 60, lngCurYPos + 8)-(lngX2 + 60, lngCurYPos + 8 * 12)
'
'    '�������۽ð�
'
'    Printer.Line (lngX2 + 75, lngCurYPos + 8)-(lngX2 + 75, lngCurYPos + 8 * 12)
'
'    Printer.Line (lngX2 + 90, lngCurYPos - 2)-(lngX2 + 90, lngCurYPos + 8 * 12)
'
'
'
'
'    Printer.Line (lngX2 + 105, lngCurYPos + 8)-(lngX2 + 105, lngCurYPos + 8 * 12)
'
'
'
'    '�������ð�
'    Printer.Line (lngX2 + 120, lngCurYPos + 8)-(lngX2 + 120, lngCurYPos + 8 * 12)
'    'Dr
'    Printer.Line (lngX2 + 130, lngCurYPos + 8)-(lngX2 + 130, lngCurYPos + 8 * 12)
'    'Nr
'    Printer.Line (lngX2 + 140, lngCurYPos + 8)-(lngX2 + 140, lngCurYPos + 8 * 12)
'    '�������ۿ�
'    'Printer.Line (lngX2 + 165, lngCurYPos + 8)-(lngX2 + 142, lngCurYPos + 8 * 12)
'
'    '������
'    Printer.Line (Twidth - PrtLeft, lngCurYPos - 2)-(Twidth - PrtLeft, lngCurYPos + 8 * 12)
'
'    Printer.FontSize = 10
'
'    Call Print_Setting("���׺�����", PrtLeft, 8, , , "C", False)
'    Call Print_Setting("�������", lngX2 + 90, 8, , , "C", False)
'
'    lngCurYPos = lngCurYPos + LineSpace
'
'    Call Print_Setting("���׺���ð�", PrtLeft, 12, lngX2 - PrtLeft, "C", "C", False)
'    Call Print_Setting("���׹�ȣ", lngX2, 12, 30, "C", "C", False)
'    Call Print_Setting("��������", lngX2 + 30, 12, 15, "C", "C", False)
'    Call Print_Setting("������", lngX2 + 45, 12, 15, "C", "C", False)
'    Call Print_Setting("ä����", lngX2 + 60, 12, 15, "L", "C", False)
'    Call Print_Setting("�����", lngX2 + 75, 12, 27, "L", "C", False)
'    Call Print_Setting("�����ð�", lngX2 + 90, 12, 20, "L", "C", False)
'    Call Print_Setting("������", lngX2 + 105, 12, 20, "L", "C", False)
'    Call Print_Setting("Dr.", lngX2 + 120, 12, 10, "C", "C", False)
'
'    Call Print_Setting("Nr.", lngX2 + 130, 12, 10, "C", "C", False)
'    Call Print_Setting("�������ۿ�", lngX2 + 140, 12, 20, "C", "C", False)
'
'    lngCurYPos = lngCurYPos + 8 * 12
'    Printer.FontBold = True
'    Call Print_Setting("Memo (Special v/s �� ȯ�ڻ��±��)", PrtLeft, LineSpace, , , "C")
'
'    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos + 50), , B
'
'
'    Printer.Line (PrtLeft, lngCurYPos + 55)-(Twidth - PrtLeft, lngCurYPos + 55)
'
'    lngCurYPos = lngCurYPos + 60
'
'    Call Print_Setting(HOSPITAL_NAME, PrtLeft, LineSpace, Twidth, "C", "C", False)
'    Printer.FontBold = False
    
    Printer.EndDoc
End Sub

Private Sub PrintTransList(ByVal vRow As String)
'ũ�ν� ��Ī�� ��ǥ �ۼ�..
'ȯ�� ������ ó�� ������ �Ķ���ͷ� �ѱ��
'���ð˻�� �ʵ�� �Ѱ��ش�.
'�⺻���� ������ �Ϸ�� �� �ڵ� ���� (������ �ο쿡 ���� ����� ���)
    
    '�Ķ���Ϳ� ���� ����
    Dim strPtnm As String
    Dim strWardNm As String
    Dim strABO As String
    Dim strStat As String
    Dim strPtid As String
    Dim strSexAge As String
    Dim strOrdDoct As String
    Dim strDept As String
    Dim strDisease As String
    Dim strOrdDt As String
    Dim strOrdNo As String
    Dim strOrdNm As String
    Dim strColdttm As String
    Dim strColNm As String
    Dim strUnitQty As String
    Dim strSpcNo As String
    Dim strStore As String
    Dim strAccNo As String
    Dim strAccdttm As String
    Dim strAccNm As String
    Dim strRelTest As String
    Dim aryRelTest() As String
    Dim strTemp As String
    Dim aryRow() As String
    Dim strabScreen As String
    Dim strdCoombs As String
    
    Dim strRfile   As String
    Dim strRptPath As String
    Dim lngFileNo As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
'    Dim objDisease As clsDisease
    
    If tblPtList.MaxRows = 0 Then Exit Sub
        
    aryRow = Split(vRow, COL_DIV)
    
    Me.MousePointer = vbHourglass
    
    With tblPtList
        For i = LBound(aryRow) To UBound(aryRow)
            If aryRow(i) <> "" Then
                .Row = Val(aryRow(i))
                
                .Col = TblColumn.tcPTNM: strPtnm = .value
                .Col = TblColumn.tcWARD:
                If .value <> "" Then
                    strWardNm = GetWardNm(.value)
                Else
                    strWardNm = "�ܷ�"
                End If
                .Col = TblColumn.tcABO: strABO = .value
                .Col = TblColumn.tcSTATFG: strStat = IIf(.value = "1", "����", "")
                .Col = TblColumn.tcPTID: strPtid = .value
                .Col = 0 'SexAge
                .Col = TblColumn.tcDOCT: strOrdDoct = GetDoctNm(.value)
                .Col = TblColumn.tcDEPT: strDept = GetDeptNm(.value)
                If strDept = "�������а�" Then
                    strWardNm = "EM"
                End If
                .Col = TblColumn.tcDISEASE: strDisease = Trim(.value)
                If strDisease <> "" Then
                    .Col = TblColumn.tcDISEASE2
                    If .value <> "" Then
                        strDisease = strDisease & "," & Trim(.value)
                    Else
                        strDisease = strDisease
                    End If
                    .Col = TblColumn.tcDISEASE3
                    If .value <> "" Then
                        strDisease = strDisease & "," & Trim(.value)
                    Else
                        strDisease = strDisease
                    End If
                    .Col = TblColumn.tcDISEASE4
                    If .value <> "" Then
                        strDisease = strDisease & "," & Trim(.value)
                    Else
                        strDisease = strDisease
                    End If
                End If
                .Col = TblColumn.tcORDDT: strOrdDt = .value
                .Col = TblColumn.tcORDNO: strOrdNo = .value
                .Col = TblColumn.tcORDNM: strOrdNm = .value
                .Col = 0 'Coldttm
                .Col = 0 'Colnm
                .Col = TblColumn.tcUNITQTY: strUnitQty = .value
                .Col = TblColumn.tcSPCNO: strSpcNo = .value
                .Col = TblColumn.tcSTORE: strStore = .value
                .Col = TblColumn.tcACCNO: strAccNo = .value
                .Col = TblColumn.tcACCDTTM: strAccdttm = .value
                .Col = 0 'Accnm
                                
'                '�󺴺ҷ����� ���� �󺴸� �ҷ��´�.
'                Set objDisease = Nothing
'                Set objDisease = New clsDisease
'
'                objDisease.Clear
'                objDisease.PtId = strPtid
'                objDisease.orddt = Format(strOrdDt, "yyyyMMdd")
'                objDisease.ordno = strOrdNo
'
'                If objDisease.GetDisease Then
'                    strDisease = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
'                End If
'
'                Set objDisease = Nothing
                
                'ȯ�ڸ����Ϳ��� SexAge�� ���Ѵ�.
                strSexAge = GetSexAge(strPtid)
                
                'ä��, ���������� �д´�.
                Call GetColAccInfo(strSpcNo, strColdttm, strColNm, strAccdttm, strAccNm)
                
                '���ð˻簡 �ִ°�� ��ȸ
                strRelTest = GetRelTest(strPtid)
                If strRelTest <> "" Then aryRelTest = Split(strRelTest, vbTab)
                
                strRfile = InstallDir & "BBS\RPT\CrystalReport.txt"
                strRptPath = InstallDir & "BBS\RPT\frmBBS102_2.rpt"
            
                lngFileNo = FreeFile
                Open strRfile For Output As #lngFileNo
                Print #lngFileNo, strRelTest
                Close #lngFileNo
                With CReport
                    .ReportFileName = strRptPath
                    
                    .ParameterFields(0) = "ptnm;" & strPtnm & ";TRUE"
                    .ParameterFields(1) = "wardnm;" & strWardNm & ";TRUE"
                    .ParameterFields(2) = "abo;" & strABO & ";TRUE"
                    .ParameterFields(20) = "stat;" & strStat & ";TRUE"
                    .ParameterFields(3) = "ptid;" & strPtid & ";TRUE"
                    .ParameterFields(4) = "sexage;" & strSexAge & ";TRUE"
                    .ParameterFields(5) = "orddoct;" & strOrdDoct & ";TRUE"
                    .ParameterFields(6) = "dept;" & strDept & ";TRUE"
                    .ParameterFields(7) = "disease;" & strDisease & ";TRUE"
                    .ParameterFields(8) = "orddt;" & strOrdDt & ";TRUE"
                    .ParameterFields(9) = "ordnm;" & strOrdNm & ";TRUE"
                    .ParameterFields(10) = "coldttm;" & strColdttm & ";TRUE"
                    .ParameterFields(11) = "colnm;" & strColNm & ";TRUE"
                    .ParameterFields(12) = "unitqty;" & strUnitQty & ";TRUE"
                    .ParameterFields(13) = "spcno;" & strSpcNo & ";TRUE"
                    .ParameterFields(14) = "store;" & strStore & ";TRUE"
                    .ParameterFields(15) = "accno;" & strAccNo & ";TRUE"
                    .ParameterFields(16) = "accdttm;" & strAccdttm & ";TRUE"
                    .ParameterFields(17) = "accnm;" & strAccNm & ";TRUE"
                    .ParameterFields(18) = "hostnm;" & HOSPITAL_NAME & ";TRUE"
                    .ParameterFields(19) = "reltest;" & IIf(strRelTest = "", "(����)", "") & ";TRUE"
                    .ParameterFields(21) = "prtnm;" & GetEmpNm(ObjSysInfo.EmpId) & ";TRUE"
                    
                    If strRelTest <> "" Then
                        strabScreen = ""
                        strdCoombs = ""
                        For j = LBound(aryRelTest) To UBound(aryRelTest)
                            If aryRelTest(j) <> "" Then
                                k = k + 1

' 2009.06.16. �缺�� And strdCoombs = ""  �߰�
' ���ں��� ��ȸ�Ǳ⶧���� ���� ó���� ���õǴ� ���� ���� �ֱٰ��̴�.
' �����˻���ȸ �Ⱓ�� ���� ������ ���� 3600������
'                                If k < 13 Then Exit For
'                                .ParameterFields(22 + j) = "reltest" & (j + 1) & ";" & aryRelTest(j) & ";TRUE"
                                If k < 13 Then .ParameterFields(22 + j) = "reltest" & (j + 1) & ";" & aryRelTest(j) & ";TRUE"
                                strTemp = Trim(medGetP(Mid(aryRelTest(j), 13), 1, ":"))

'2015.09.15 �½�ȣ Ab Screening �ֱ� ��� ��ȸ
'Ab-id �˻��׸���� ����
'                                If Mid(strTemp, 1, 3) = "Ab " And strabScreen = "" Then
                                If InStr(strTemp, "Ab ") > 0 And strabScreen = "" Then
                                    If Val(Trim(medGetP(strabScreen, 2, "-"))) < Val(Trim(medGetP(aryRelTest(j), 2, "-"))) Then
                                        strabScreen = Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":"))
                                    End If

                                End If
                                If Mid(strTemp, 1, 4) = "Coom" And strdCoombs = "" Then
                                    If Val(Trim(medGetP(strdCoombs, 2, "-"))) < Val(Trim(medGetP(aryRelTest(j), 2, "-"))) Then
                                        strdCoombs = Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":"))
                                    End If

'                                    strdCoombs = Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":"))
'                                    .ParameterFields(23) = "dCooms;" & Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":")) & ";TRUE"
'                                    Debug.Print strTemp
                                End If
                            End If
                        Next j
                        .ParameterFields(22 + j) = "abScreen;" & strabScreen & ";TRUE"
                        .ParameterFields(22 + j + 1) = "dCooms;" & strdCoombs & ";TRUE"
                    End If
                    
                    .RetrieveDataFiles
    '                .WindowState = crptMaximized
                    .Destination = crptToPrinter
                    .Action = 1
                    .Reset
                End With
            End If
        Next i
    End With
    
    Me.MousePointer = 0
End Sub

Private Sub GetColAccInfo(ByVal vSpcNo As String, _
                          ByRef pColDtTm As String, ByRef pColId As String, _
                          ByRef pAccDtTm As String, ByRef pAccId As String)

    'S2bbs201���� ä��, ���������� �д´�
    
    Dim RS As Recordset
    Dim strSQL As String
    
    Set RS = New Recordset
    strSQL = " select * from " & T_BBS201 & _
             " where " & DBW("spcyy=", medGetP(vSpcNo, 1, "-")) & _
             " and " & DBW("spcno=", medGetP(vSpcNo, 2, "-"))
    RS.Open strSQL, DBConn
    
    If RS.EOF = False Then
        pColDtTm = Format(RS.Fields("coldt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("coltm").value & "", 1, 4), "00:00")
        If RS.Fields("colid").value & "" <> "" Then
            pColId = GetEmpNm(RS.Fields("colid").value & "")
        End If
        pAccDtTm = Format(RS.Fields("rcvdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("rcvtm").value & "", 1, 4), "00:00")
        If RS.Fields("rcvid").value & "" <> "" Then
            pAccId = GetEmpNm(RS.Fields("rcvid").value & "")
        End If
    End If
    
    Set RS = Nothing
End Sub

Private Function GetSexAge(ByVal vPtID As String)
    Dim objPt As clsPatient
    
    Set objPt = New clsPatient
    
    Call objPt.GETPatient(vPtID)
    
    GetSexAge = objPt.sexage
    
    Set objPt = Nothing
End Function

Private Function GetRelTest(ByVal vPtID As String) As String
'ũ����Ż ����Ʈ ��¿� ��Ʈ�� �����..

    Dim RS As Recordset
    Dim objSql As clsCrossMatching
    Dim strTmp As String
    Dim lngCnt As Long
    Dim strRstCd As String
    
    Set objSql = New clsCrossMatching
    Set RS = New Recordset
    
    RS.Open objSql.TestResultXM(vPtID), DBConn
        
    If RS.EOF Then
        GetRelTest = ""
    Else
        Do Until RS.EOF
            lngCnt = lngCnt + 1

' 2009.06.16. �缺�� �������� �˻���  Ab Screen ���� �������� ���ؼ� ���ƹ���.
'
'            If lngCnt > 12 Then Exit Do
'
            If RS.Fields("rstcdnm").value & "" = "" Then
                strRstCd = RS.Fields("rstcd").value & ""
            Else
                strRstCd = RS.Fields("rstcdnm").value & ""
            End If
            
            strTmp = strTmp & Format(RS.Fields("workarea").value & "" & "-" & _
                                     Mid(RS.Fields("accdt").value & "", 3) & "-" & _
                                     RS.Fields("accseq").value & "", "!" & String(17, "@")) & _
                     Format(RS.Fields("abbrnm10").value & "", "!" & String(11, "@")) & " : " & _
                     strRstCd & vbTab
            RS.MoveNext
        Loop
        
        strTmp = strTmp & vbNewLine
        
        GetRelTest = strTmp
    End If
    
    Set RS = Nothing
    Set objSql = Nothing
End Function

Private Sub LoadBuilding()
    
    Dim objcom003   As clsCom003
    Dim RS          As Recordset
    Dim i           As Long
    Dim itmX        As ListItem
    
    Set objcom003 = New clsCom003
    Set RS = objcom003.OpenRecordSet(BC2_CENTER)
    Set objcom003 = Nothing
    
    cboBuilding.Clear
    cboBuilding.AddItem "(��ü)"
    If Not RS.EOF Then
        With RS
            For i = 1 To .RecordCount
                cboBuilding.AddItem .Fields("cdval1").value & " " & .Fields("field1").value & ""
                .MoveNext
            Next i
        End With
    End If
    Set RS = Nothing
    If cboBuilding.ListCount > 1 Then
        cboBuilding.ListIndex = medComboFind(cboBuilding, ObjSysInfo.BuildingCd)
    Else
        cboBuilding.ListIndex = 0
    End If
    
End Sub

'2001-11-30 �߰�
'�����ǥ ����� ���� Query
Private Sub QueryForReport()
    Dim i           As Long
    Dim j           As Long
    
    Dim RS        As Recordset
    Dim RsTime      As Recordset
    Dim QueryOrder  As clsQueryOrder
    Dim objDisease  As clsDisease
    Dim ObjABO      As clsABO
    
    Dim accno       As String
    Dim reason      As String
    Dim status      As String
    Dim spcno       As String
    Dim storeleg    As String
    Dim storerow    As String
    Dim storecol    As String
    Dim center      As String
    
    Dim strLeg      As String
    Dim strRow      As String
    Dim strCol      As String
    Dim inout       As String
    Dim MaxRowCnt   As Long
    Dim TestDiv     As String
'    Dim blnComplete As Boolean
    
    Dim objPrgBar   As clsProgress
    
    Dim otherCenter As Boolean
    
    '���ٰ� ���������̸� ���ڸ� ���߱� ���Ѻ�����
    Dim bkPtId      As String
    Dim bkReason    As String
    Dim bkReqDt     As String
    Dim bkOrdDt     As String
    Dim bkRoomid    As String
    Dim bkWard      As String
    Dim bkDept      As String
    
    Dim strDc       As String
    
    Dim blnCompleted As Boolean
    Dim blnAccomplished As Boolean
    
    tblPtList.MaxRows = 0
    
    Call Save_LegRowCol
    
    Set QueryOrder = New clsQueryOrder
    
    If cboOrd.ListIndex <> 0 Then TestDiv = medGetP(cboOrd.Text, 1, " ")
    '-----------
    '���º� ��ȸ
    '-----------

    QueryOrder.stscd = "'3','4'"

    '------------------------------------
    '�����ǥ��¿� �´� �������� �ʱ�ȭ
    cboInOut.ListIndex = 0
    chkDc.value = 0
    chkStat.value = 0
    txtWardId = ""
    cboOrd.ListIndex = -1
    
    inout = ""
    strDc = ""
    TestDiv = ""
    '------------------------------------
    
        
    
    
    Set RS = QueryOrder.QueryRequest(Format(dtpFrDt, PRESENTDATE_FORMAT), Format(dtpToDt, PRESENTDATE_FORMAT), _
                                      chkStat.value, txtPtId, inout, strDc, txtWardId, TestDiv)
    
    If RS Is Nothing Then
        Set RS = Nothing
        Set QueryOrder = Nothing
        Exit Sub
    End If
    
    Set objDisease = New clsDisease
    Set ObjABO = New clsABO
    
    Set objPrgBar = New clsProgress
    objPrgBar.Container = medMain.stsBar
    
    objPrgBar.Min = 1
    objPrgBar.Max = RS.RecordCount
    
    
    With tblPtList
        bkPtId = ""
        .ReDraw = False
        For i = 1 To RS.RecordCount
        
            objPrgBar.value = i
'            blnComplete = CompleteOrderChk(Rs.Fields("accdt").value & "", Rs.Fields("accseq").value & "", Rs.Fields("unitqty").value & "")
            Call CheckCompleted(RS.Fields("accdt").value & "", RS.Fields("accseq").value & "", RS.Fields("unitqty").value & "", _
                                blnCompleted, blnAccomplished)
            If blnCompleted = True Then GoTo Skip

Skip1:
            
            MaxRowCnt = MaxRowCnt + 1
            .MaxRows = MaxRowCnt
            .Row = MaxRowCnt
            accno = Trim(RS.Fields("accdt").value & "") & "-" & Val(Trim(RS.Fields("accseq").value & ""))
            If accno = "-0" Then accno = "" 'accno = "������"
            
            '�������� ���ϱ�...
            reason = QueryOrder.GetTransReason(RS.Fields("ptid").value & "", RS.Fields("orddt").value & "", RS.Fields("ordno").value & "")
            
            
            If reason = "" Then reason = "(����)"
            

            
            .Col = TblColumn.tcACCNO:      .value = accno
            .Col = TblColumn.tcPTID:       .value = RS.Fields("ptid").value & ""
            
            .Col = TblColumn.tcPTNM:       .value = RS.Fields("ptnm").value & ""
            .Col = TblColumn.tcORDNM:      .value = RS.Fields("testnm").value & ""
            .Col = TblColumn.tcORDDT:      .value = Format(RS.Fields("orddt").value & "", "####-##-##")
            '.Col = TblColumn.tcUNITQTY:    .value = RS.Fields("unitqty").value & ""
            .Col = TblColumn.tcUNITQTY:    .value = RS.Fields("reqcnt").value & ""
            .Col = TblColumn.tcREASON:     .value = Trim(Trim0(reason))
            .Col = TblColumn.tcREQDT:      .value = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
            '2001-11-30�߰�
            '�����ǥ�� ����ǻ�/�ֱټ����� ����ϱ�����
            .Col = TblColumn.tcDOCT:       .value = RS.Fields("orddoct").value & ""
            .Col = TblColumn.tcTRANSDT:    .value = QueryOrder.GetLatestTrandDt(RS.Fields("ptid").value & "")
'
            .Col = TblColumn.tcWARD:       .value = RS.Fields("wardid").value & ""
            .Col = TblColumn.tcROOM:       .value = RS.Fields("hosilid").value & ""
            
            .Col = TblColumn.tcDEPT:       .value = RS.Fields("deptcd").value & ""
            .Col = TblColumn.tcBUSSDIV:    .value = RS.Fields("bussdiv").value & ""
            .Col = TblColumn.tcORDDTDB:    .value = RS.Fields("orddt").value & ""
            .Col = TblColumn.tcORDNO:      .value = Val(RS.Fields("ordno").value & "")
            .Col = TblColumn.tcORDSEQ:     .value = Val(RS.Fields("ordseq").value & "")
            .Col = TblColumn.tcSTATFG:     .value = RS.Fields("statfg").value & ""
            .Col = TblColumn.tcSTATnm:     .value = IIf(RS.Fields("statfg").value & "" = "1", "Y", "")
                                           .ForeColor = vbRed
                                           .FontBold = True
            .Col = TblColumn.tcBedInDT:    .value = RS.Fields("bedindt").value & ""
            .Col = TblColumn.tcDCFG:       .value = RS.Fields("dcfg").value & ""
            .Col = TblColumn.tcDCNM:       .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", "")
                                           .ForeColor = vbBlue
                                           .FontBold = True
            '.Col = TblColumn.tcCENTERCD:   .value = center
            .Col = TblColumn.tcPHERESIS:   .value = RS.Fields("testdiv").value & ""
            .Col = TblColumn.tcSTSCD:      .value = RS.Fields("stscd").value & ""
            .Col = TblColumn.tcSTSNM
                                            If TRANS_REQUIRE_USED Then
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER: .ForeColor = DCM_Gray '"ó��"
                                                         Case "1": .value = STS_NM_COLLECT '"ä��"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"����"
                                                         '2001-11-15 ���� : '��û' Status �߰�
                                                         Case "3": .value = STS_NM_REQUEST: .ForeColor = DCM_Red '"��û"
                                                         Case "4": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS) '"����","�Ϸ�","�˻���"
                                                                   .ForeColor = IIf(blnCompleted, &H8000&, DCM_Brown)
                                                         'Case "3": .value = "�˻���"
                                                         Case Else: .value = ""
                                                    End Select
                                            Else
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER '"ó��"
                                                         Case "1": .value = STS_NM_COLLECT: .ForeColor = DCM_LightRed '"ä��"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"����"
                                                         Case "3": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS): .ForeColor = DCM_Brown ''"����","�Ϸ�","�˻���"
                                                         'Case "3": .value = "�˻���"
                                                         Case Else: .value = ""
                                                    End Select
                                            End If
            .Col = TblColumn.TcMESG: .value = RS.Fields("mesg").value & ""
            '�������� ���Ѵ�.
            ObjABO.Ptid = RS.Fields("ptid").value & ""
            
            If ObjABO.GetABO = False Then
                .Col = TblColumn.tcABO:    .value = ""
            Else
                .Col = TblColumn.tcABO:    .value = ObjABO.ABO & ObjABO.Rh
            End If
            
            '���ܸ��� ���Ѵ�.
            With objDisease
                .Clear
                .Ptid = RS.Fields("ptid").value & ""
                .OrdDt = RS.Fields("orddt").value & ""
                .ordno = RS.Fields("ordno").value & ""
            End With
            
            
            If objDisease.GetDisease = False Then
                .Col = TblColumn.tcDISEASE: .value = ""
                .Col = TblColumn.tcDISEASE2: .value = ""
                .Col = TblColumn.tcDISEASE3: .value = ""
                .Col = TblColumn.tcDISEASE4: .value = ""
            Else
                j = 0
                Do
                    If objDisease.EOF Then Exit Do
                    
                    If objDisease.DiseaseCd <> "" Then
                        j = j + 1
                        Select Case j
                            Case 1: .Col = TblColumn.tcDISEASE
                            Case 2: .Col = TblColumn.tcDISEASE2
                            Case 3: .Col = TblColumn.tcDISEASE3
                            Case 4: .Col = TblColumn.tcDISEASE4
                        End Select
                        .value = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                    End If
                    objDisease.MoveNext
                Loop
            End If
            
            otherCenter = False
            
            
            '-------------------------------------------
            'ó���� irradiation ó���� �ƴ� ó���ϰ�츸
            '-------------------------------------------
            Call QueryOrder.GetSpcNoAndStore(RS.Fields("ptid").value & "", spcno, storeleg, storerow, storecol, center)
            
            '--------------------------------------------------------------------------------------
            '2001-11-23 �߰� :
            '�ǹ������� ����� ���, �׸��� (��ü)�� �ƴҰ�� �ش� �ǹ��� ����Ÿ�� ���÷���
            If ObjSysInfo.UseBuildingInfo = 1 And cboBuilding.ListIndex <> 0 Then
                If medGetP(center, 1, vbTab) <> ObjSysInfo.BuildingCd Then
                    MaxRowCnt = MaxRowCnt - 1
                    .MaxRows = MaxRowCnt
                    GoTo Skip
                End If
            End If
            '--------------------------------------------------------------------------------------
            
            If center = "" Then center = ObjSysInfo.BuildingCd & vbTab & ObjSysInfo.BuildingNm
            .Col = TblColumn.tcCENTERNM:    .value = medGetP(center, 2, vbTab) 'GetCenterNm(medGetP(center, 1, vbTab))
            .Col = TblColumn.tcCENTERCD:    .value = medGetP(center, 1, vbTab)
            
            If medGetP(center, 1, vbTab) <> ObjSysInfo.BuildingCd Then
                '��ü�� �ٸ� ���Ϳ� �ִ�.
                .Col = TblColumn.tcSTORE:   .value = medGetP(center, 2, vbTab) & "(" & medGetP(center, 1, vbTab) & ")"
                otherCenter = True
            End If
            
            .Col = TblColumn.tcORDDIV:      .value = RS.Fields("orddiv").value & ""
            If .value = C_WORKAREA Then
                '--------------------------
                '��ü��ȣ�� ������� ���ϱ�
                '--------------------------
                If storerow = "0" Then storerow = ""
                If storecol = "0" Then storecol = ""
                
                .Col = TblColumn.tCLegRowCol:   .value = storeleg & ";" & storerow & ";" & storecol
                
                .Col = TblColumn.tcSPCNO:       .value = spcno
                
                If spcno = "" Then
                    .Col = TblColumn.tcSTORE:   .value = "" '.value = "��ä��"
                Else
                    If storeleg = "" Then
                        .Col = TblColumn.tcSTORE:    .value = ""
                        .Col = TblColumn.tcNOACCSSS: .value = "1"
                    Else
                        .Col = TblColumn.tcSTORE:    .value = storeleg & "(" & storerow & "," & storecol & ")"
                        .Col = TblColumn.tcNOACCSSS: .value = "0"
                    End If
                End If
                '----------------------------
                '��ü��� �ð��� ���ϱ����ؼ�
                '----------------------------
                Dim today   As Date
                Dim coldttm As String
                today = GetSystemDate
                
                If spcno <> "" Then
                    If Val(RS.Fields("stscd").value & "") > 2 Then
                        If QueryOrder.Get_ExistSPC(medGetP(spcno, 1, "-"), medGetP(spcno, 2, "-")) <> "1" Then
                            .Col = TblColumn.tcSPCNO: .ForeColor = DCM_LightGray
                            .Col = TblColumn.tcSTORE: .ForeColor = DCM_LightGray
                        End If
                    End If
                    Set RsTime = Nothing
                    Set RsTime = New Recordset
                    RsTime.Open QueryOrder.Get_spcTime(medGetP(spcno, 1, "-"), medGetP(spcno, 2, "-")), DBConn
                    
                    If Not RsTime.EOF Then
                        If Len(RsTime.Fields("coltm").value & "") = 4 Then
                            coldttm = RsTime.Fields("coltm").value & "" & "00"
                            coldttm = Format(RsTime.Fields("coldt").value & "", "0###-##-##") & " " & Format(coldttm, "0#:##:##")
                        Else
                            coldttm = Format(RsTime.Fields("coldt").value & "", "0###-##-##") & " " & Format(RsTime.Fields("coltm").value & "", "0#:##:##")
                        End If
                        
                       ' coldttm = Format(RsTime.Fields("coldt").value, "0###-##-##") & " " & Format(coldttm, "0#:##:##")
                        .Col = TblColumn.tcTime: .value = DateDiff("h", coldttm, today)
                    End If
                End If
            End If
            
            
            .Col = TblColumn.tcDUPCHK: .value = RS.Fields("ptid").value & "" & COL_DIV & RS.Fields("orddt").value & ""
            
            '-------------------------
            '�ߺ��Ǵ� ���� �Ⱥ��̰�...
            '-------------------------
            
            If bkPtId <> RS.Fields("ptid").value & "" Then
                bkPtId = RS.Fields("ptid").value & ""
                bkReason = reason
                bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                bkRoomid = RS.Fields("hosilid").value & ""
                bkWard = RS.Fields("wardid").value & ""
                bkDept = RS.Fields("deptcd").value & ""
                
            Else
                .Row = i - 1
                .Col = TblColumn.tcWARD: bkWard = .value
                .Col = TblColumn.tcDEPT: bkDept = .value
                
                .Row = i
                .Col = TblColumn.tcPTID: .ForeColor = .BackColor
                .Col = TblColumn.tcPTNM: .ForeColor = .BackColor
                If bkReason = reason Then
                    If reason <> "(����)" Then .Col = TblColumn.tcREASON: .ForeColor = .BackColor
                Else
                    bkReason = reason
                End If
                If bkWard = RS.Fields("wardid").value & "" Then
                    .Col = TblColumn.tcWARD: .ForeColor = .BackColor
                End If
                If bkDept = RS.Fields("deptcd").value & "" Then
                    .Col = TblColumn.tcDEPT: .ForeColor = .BackColor
                End If
                
                If bkRoomid = RS.Fields("hosilid").value & "" Then
                    .Col = TblColumn.tcROOM: .ForeColor = .BackColor
                Else
                    bkRoomid = RS.Fields("hosilid").value & ""
                End If
                If bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00") Then
                    .Col = TblColumn.tcREQDT: .ForeColor = .BackColor
                Else
                    bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                End If
                If bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##") Then
                    .Col = TblColumn.tcORDDT: .ForeColor = .BackColor
                Else
                    bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                End If
            End If
Skip:
            '---------------------
            '������ �� �ִ� ������
            '---------------------
            If MaxRowCnt > 0 Then
                If CanSelect(1, MaxRowCnt) Then
                    .Row = MaxRowCnt
                    .Col = TblColumn.tcSEL
                    .CellType = CellTypeCheckBox
                    .TypeCheckCenter = True
                Else
                    .Row = MaxRowCnt
                    
                    .Col = TblColumn.tcSEL
                    .CellType = CellTypeStaticText
                    .Col = TblColumn.tcSTSNM
                    If .value = STS_NM_DONE Or .value = STS_NM_END Then
                        .Col = TblColumn.tcSEL
                        .Text = "��"
                        .ForeColor = vbRed
                    End If
                End If
            End If
            RS.MoveNext
        Next i
        .ReDraw = True
    End With
     
    Set RS = Nothing
    Set ObjABO = Nothing
    Set objPrgBar = Nothing
    Set objDisease = Nothing
    Set QueryOrder = Nothing
End Sub


