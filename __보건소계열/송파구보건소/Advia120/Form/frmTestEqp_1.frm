VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmTestEqp_1 
   Caption         =   " ��� VS �˻��ڵ� ����"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11970
   WindowState     =   2  '�ִ�ȭ
   Begin VB.TextBox txtOutseq 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   9585
      MaxLength       =   3
      TabIndex        =   19
      Top             =   975
      Width           =   735
   End
   Begin VB.TextBox txtRefH 
      Appearance      =   0  '���
      Height          =   270
      Left            =   6750
      MaxLength       =   10
      TabIndex        =   23
      Text            =   "1234567890"
      Top             =   1305
      Width           =   1020
   End
   Begin VB.TextBox txtRefL 
      Appearance      =   0  '���
      Height          =   270
      Left            =   5490
      MaxLength       =   10
      TabIndex        =   22
      Text            =   "1234567890"
      Top             =   1305
      Width           =   1020
   End
   Begin VB.TextBox txtAuto 
      Appearance      =   0  '���
      Height          =   270
      Left            =   10350
      MaxLength       =   10
      TabIndex        =   21
      Text            =   "1234567890"
      Top             =   975
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSpccd 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   9585
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp_1.frx":0000
            Key             =   "TST_E"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp_1.frx":059A
            Key             =   "TST_M"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwTstListEqp 
      Height          =   5850
      Left            =   45
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   585
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   10319
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�˻��ڵ�(���)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�˻��(���)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�˻� �ڵ�(������)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�˻� �� (������)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtTestNm 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   975
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.TextBox txtTestCD 
      Appearance      =   0  '���
      Height          =   270
      Left            =   5490
      MaxLength       =   100
      TabIndex        =   16
      Text            =   "1234567890"
      Top             =   975
      Width           =   4080
   End
   Begin VB.TextBox txtVIndex 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      Height          =   270
      Left            =   11010
      MaxLength       =   5
      TabIndex        =   12
      Top             =   660
      Width           =   675
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '�� ����
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmTestEqp_1.frx":0B34
      Caption         =   " Instruments Test Item Link ."
      SubCaption      =   "�˻�� �˻��׸�� ��� �˻��׸��� ���� �մϴ�."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   360
         Left            =   11520
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.TextBox txtTstcdEqp 
      Appearance      =   0  '���
      Height          =   270
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   0
      Top             =   675
      Width           =   1425
   End
   Begin VB.TextBox txtTstnmEqp 
      Appearance      =   0  '���
      Height          =   270
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1020
      Width           =   1650
   End
   Begin VB.TextBox lblTstcdEqp 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   5490
      MaxLength       =   20
      TabIndex        =   13
      Text            =   "1234567890"
      Top             =   660
      Width           =   1005
   End
   Begin VB.TextBox lblTstnmEqp 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   6495
      MaxLength       =   40
      TabIndex        =   14
      Top             =   660
      Width           =   3075
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "����"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   15
      TabIndex        =   5
      Top             =   6450
      Width           =   11940
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Print"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   1
         Left            =   1515
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   2
         Left            =   2895
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Clear"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   3
         Left            =   4260
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
   End
   Begin HSCotrol.CButton cmdEqpItm_Add 
      Height          =   300
      Left            =   3015
      TabIndex        =   3
      Top             =   675
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp_1.frx":1DB6
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdEqpItm_Del 
      Height          =   300
      Left            =   3015
      TabIndex        =   4
      Top             =   1005
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      Caption         =   "Del"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp_1.frx":1F10
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.UserPanel pnlTestitem 
      Height          =   5160
      Left            =   6930
      TabIndex        =   10
      Top             =   1890
      Visible         =   0   'False
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   9102
      Bevel           =   2
      Moveble         =   -1  'True
      CloseEnabled    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView lvwTestitem 
         Height          =   4815
         Left            =   105
         TabIndex        =   11
         Top             =   270
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   8493
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin HSCotrol.CButton cmdAdd 
      Height          =   300
      Left            =   10380
      TabIndex        =   34
      Top             =   1290
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   529
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp_1.frx":24AA
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdDel 
      Height          =   300
      Left            =   11160
      TabIndex        =   35
      Top             =   1290
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   529
      Caption         =   "Del"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp_1.frx":2604
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
      HoverColor      =   -2147483635
   End
   Begin HSCotrol.CButton cmdSerch 
      Height          =   300
      Left            =   6525
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   529
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestEqp_1.frx":2B9E
      MaskColor       =   0
      PicCapAlign     =   1
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin VB.Frame Frame5 
      Height          =   6015
      Left            =   4020
      TabIndex        =   33
      Top             =   450
      Width           =   30
   End
   Begin MSComctlLib.ListView lvwTestListLab 
      Height          =   4860
      Left            =   4080
      TabIndex        =   15
      Top             =   1605
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   8573
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�˻��ڵ�(���)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�˻��(���)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�˻� �ڵ�(������)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�˻� �� (������)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtDelta 
      Appearance      =   0  '���
      Height          =   270
      Left            =   5490
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox txtDeltagbn 
      Appearance      =   0  '���
      Height          =   270
      Left            =   6525
      MaxLength       =   10
      TabIndex        =   25
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtPanic 
      Appearance      =   0  '���
      Height          =   270
      Index           =   0
      Left            =   8235
      MaxLength       =   10
      TabIndex        =   26
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txtPanic 
      Appearance      =   0  '���
      Height          =   270
      Index           =   1
      Left            =   9405
      MaxLength       =   10
      TabIndex        =   27
      Text            =   "1234567890"
      Top             =   1620
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Label8 
      Caption         =   "~"
      Height          =   195
      Left            =   6570
      TabIndex        =   41
      Top             =   1350
      Width           =   240
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "�˻�����ġ :"
      Height          =   180
      Left            =   4185
      TabIndex        =   40
      Top             =   1350
      Width           =   1020
   End
   Begin VB.Label Label6 
      Caption         =   "~"
      Height          =   195
      Left            =   9225
      TabIndex        =   39
      Top             =   1665
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Panic value : "
      Height          =   180
      Left            =   7065
      TabIndex        =   38
      Top             =   1665
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Delta value : "
      Height          =   180
      Left            =   4185
      TabIndex        =   37
      Top             =   1665
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "�� �� �� �� :"
      Height          =   180
      Left            =   4170
      TabIndex        =   36
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "View Index :"
      Height          =   180
      Left            =   9915
      TabIndex        =   31
      Top             =   705
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��� �˻�� :"
      Height          =   180
      Left            =   60
      TabIndex        =   30
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��� �˻� �ڵ� :"
      Height          =   180
      Left            =   60
      TabIndex        =   29
      Top             =   735
      Width           =   1320
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "�� �� �� �� :"
      Height          =   180
      Left            =   4170
      TabIndex        =   28
      Top             =   705
      Width           =   1020
   End
End
Attribute VB_Name = "frmTestEqp_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const OBJTAG_EQP    As String = "EQP"
Private Const OBJTAG_TST    As String = "TST"
Private Const AUTO_VEFY     As String = "YES"
Private Const AUTO_VEFN     As String = "NO"

Private Const TLB_TEMP      As String = "TEMPTEABLE"
Private Const TLB_RESULT    As String = "INTERFACE003"

Private mAdoRs              As ADODB.Recordset
Private WithEvents PopUp_List As Listview
Attribute PopUp_List.VB_VarHelpID = -1

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
        Case 0: Call cmdPrint
        Case 1: Call cmdSave
        Case 2: Call cmdClear
        Case 3: Call cmdClose
        Case Else
    End Select
    
End Sub

Private Sub cmdPrint()

    Call PrintFrom(lvwTestListLab.ListItems)

End Sub

Private Sub cmdAdd_Click()
    
    Dim itemX As ListItem
    Dim itemS As ListItem
    Dim itemZ As ListSubItem
    If Trim(lblTstcdEqp) = "" Then
        Call ShowMessage("��� �˻��ڵ尡 �����ϴ�. �ڵ带 ���� �Ͻÿ�.   ")
        Exit Sub
    End If
    
    If Trim(lblTstnmEqp) = "" Then
        Call ShowMessage("��� �˻��ڵ尡 �����ϴ�. �ڵ带 ���� �Ͻÿ�.   ")
        Exit Sub
    End If
    
    If Trim(txtTestCd) = "" Then
        Call ShowMessage("���˻��ڵ�� ������ �˻��ڵ尡 �����ϴ�. �ڵ带 ���� �Ͻÿ�.   ")
        Exit Sub
    End If
    
    Set itemS = lvwTestListLab.FindItem(Trim(lblTstcdEqp), lvwText, , lvwWhole)
    
    If Not itemS Is Nothing Then
        If vbYes = MsgBox(Trim(lblTstcdEqp) & " ���˻� �ڵ�� �̹� �ֽ��ϴ�. �ٲٽðڽ��ϱ�?", vbExclamation + vbYesNo) Then
            Call lvwTestListLab.ListItems.Remove(itemS.Index)
            Set itemX = lvwTestListLab.ListItems.Add(, , Trim(lblTstcdEqp), , "TST_M")
            With itemX
'                .SubItems(1) = Trim(lblTstnmEqp)
                .SubItems(1) = Trim(txtTestCd)
                .SubItems(2) = Trim(lblTstnmEqp)
                '.SubItems(3) = Trim(txtSpccd)
                '.SubItems(4) = Trim(txtAuto)
                .SubItems(3) = Trim(txtRefL)
                .SubItems(4) = Trim(txtRefH)
                .SubItems(5) = Trim$(txtOutseq)
                '.SubItems(7) = Trim(txtDelta)
                '.SubItems(8) = Trim(txtDeltagbn)
                '.SubItems(9) = Trim(txtPanic(0))
                '.SubItems(10) = Trim(txtPanic(1))
            End With
        End If
    Else
        Set itemX = lvwTestListLab.ListItems.Add(, , Trim(lblTstcdEqp), , "TST_M")
        With itemX
'            .SubItems(1) = Trim(lblTstnmEqp)
            .SubItems(1) = Trim(txtTestCd)
            .SubItems(2) = Trim(lblTstnmEqp)
            '.SubItems(3) = Trim(txtSpccd)
            '.SubItems(4) = Trim(txtAuto)
            .SubItems(3) = Trim(txtRefL)
            .SubItems(4) = Trim(txtRefH)
            .SubItems(5) = Trim$(txtOutseq)
            '.SubItems(7) = Trim(txtDelta)
            '.SubItems(8) = Trim(txtDeltagbn)
            '.SubItems(9) = Trim(txtPanic(0))
            '.SubItems(10) = Trim(txtPanic(1))
        End With
    End If
    lblTstcdEqp = ""
    lblTstnmEqp = ""
    txtTestCd = ""
    txtTestNm = ""
    txtSpccd = ""
    txtVIndex = ""
    txtAuto = ""
    txtDelta = ""
    txtDeltagbn = ""
    txtPanic(0) = ""
    txtPanic(1) = ""
    txtRefL = ""
    txtRefH = ""
    txtOutseq = ""
    Set itemX = Nothing
    Set itemS = Nothing
    
    lvwTstListEqp.SetFocus
    
End Sub

Private Sub cmdClose()
    Unload Me
End Sub

Private Sub cmdClear()
    
    Call f_subClear_Form

End Sub

Private Sub cmdDel_Click()
    Dim itemX   As ListItem
    Dim itemXs  As ListItems
    Dim i       As Long
    
    Set itemX = lvwTestListLab.SelectedItem
    
    If itemX Is Nothing Then
        Call ShowMessage("���õ� �׸��� �����ϴ�. ���� �Ϸ��� �׸��� ������ �����Ͻÿ�.")
        Exit Sub
    Else
        Set itemXs = lvwTestListLab.ListItems
        For i = itemXs.Count To 1 Step -1
           If itemXs(i).Selected = True Then
              lvwTestListLab.ListItems.Remove i
           End If
        Next
    End If
    Set itemX = Nothing
    
   lvwTstListEqp.SetFocus
End Sub

Private Sub cmdEqpItm_Add_Click()
    Dim objEqpItem  As clsCommon
    Dim strTemp     As String
        
    If Trim(txtTstcdEqp) = "" Then
        Call ShowMessage("��� �˻��ڵ尡 �����ϴ�. �ڵ带 �Է� �Ͻÿ�.   ")
        txtTstcdEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTstnmEqp) = "" Then
        Call ShowMessage("��� �˻���� �����ϴ�. �˻���� �Է� �Ͻÿ�.   ")
        txtTstnmEqp.SetFocus
        Exit Sub
    End If
    
    Set objEqpItem = New clsCommon
    
    With objEqpItem
        .SetAdoCn AdoCn_Jet
        If .Let_EqpTestItem(INS_CODE, Trim(txtTstcdEqp), Trim(txtTstnmEqp)) Then
            txtTstcdEqp = ""
            txtTstnmEqp = ""
            txtTstcdEqp.SetFocus
        Else
            Call ShowMessage("�������־� ���� ���� ���߽��ϴ�.")
        End If
    End With
    
    Set objEqpItem = Nothing
    Call f_subSet_EqpData(INS_CODE)
    txtTstcdEqp.SetFocus
End Sub

Private Sub cmdEqpItm_Del_Click()
    Dim itemX As ListItem
    Dim objEqpItem As clsCommon
    
    Set itemX = lvwTstListEqp.SelectedItem
    
    If itemX Is Nothing Then
        Call ShowMessage("���õ� �׸��� �����ϴ�. ���� �Ϸ��� �׸��� ������ �����Ͻÿ�.")
        Exit Sub
    Else
        Set objEqpItem = New clsCommon
        With objEqpItem
            .SetAdoCn AdoCn_Jet
            If Not .Del_EqpTestItem(INS_CODE, Trim(itemX.Text)) Then
                Call ShowMessage("�������־� ���� ���� ���߽��ϴ�.")
            End If
        End With
    End If
    Set itemX = Nothing
    Set objEqpItem = Nothing
    Call f_subSet_EqpData(INS_CODE)
End Sub

Private Sub cmdSave()

    On Error GoTo frmTestEqp_Add_Error
    
    Dim sqlDoc  As String, sqlRet   As Integer
    Dim itemX   As ListItem
    
    If lvwTestListLab.ListItems.Count < 1 Then
        sqlDoc = "Update INTERFACE002" & _
                 "   set OUT_SEQ = 0, TESTCD = '',   TESTNM = '', AUTOVERIFY = '', REMARK = ''," & _
                 "       DELTA = '',  DELTAGBN = '', PANICL = '', PANICH = ''" & _
                 " where EQP_CD = '" & INS_CODE & "'"
        AdoCn_Jet.Execute sqlDoc
    Else
        sqlDoc = "Update INTERFACE002" & _
                 "   set OUT_SEQ = 0, TESTCD = '',   TESTNM = '', AUTOVERIFY = '', REMARK = ''," & _
                 "       DELTA = '',  DELTAGBN = '', PANICL = '', PANICH = ''" & _
                 " where EQP_CD = '" & INS_CODE & "'"
        AdoCn_Jet.Execute sqlDoc
        
        For Each itemX In lvwTestListLab.ListItems
            sqlDoc = "Update INTERFACE002" & _
                     "   set TESTNM_EQP = '" & Trim$(itemX.SubItems(2)) & "'," & _
                     "       OUT_SEQ    = " & Val(itemX.SubItems(5)) & "," & _
                     "       TESTCD     = '" & Trim$(itemX.SubItems(1)) & "'," & _
                     "       TESTNM     = '" & Trim$(itemX.SubItems(2)) & "'," & _
                     "       AUTOVERIFY = ''," & _
                     "       REMARK     = ''," & _
                     "       DELTA      = ''," & _
                     "       DELTAGBN   = ''," & _
                     "       PANICL     = ''," & _
                     "       PANICH     = ''" & _
                     " where EQP_CD     = '" & INS_CODE & "'" & _
                     "   and TESTCD_EQP = '" & Trim$(itemX.Text) & "'"
            AdoCn_Jet.Execute sqlDoc, sqlRet
            If sqlRet = 0 Then
                sqlDoc = "Insert into INTERFACE002(" & _
                         "            EQP_CD, TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD," & _
                         "            TESTNM, AUTOVERIFY, REMARK,     DELTA,   DELTAGBN," & _
                         "            PANICL, PANICH)" & _
                         "    values( '" & INS_CODE & "', '" & Trim$(itemX.Text) & "'," & _
                         "            '" & Trim$(itemX.SubItems(2)) & "'," & _
                         "             " & Val(itemX.SubItems(5)) & "," & _
                         "            '" & Trim$(itemX.SubItems(1)) & "'," & _
                         "            '" & Trim$(itemX.SubItems(2)) & "'," & _
                         "            '', '', '', '', '', '')"
                AdoCn_Jet.Execute sqlDoc, sqlRet
            End If
        Next itemX
    End If
    Call f_subSet_EqpData(INS_CODE)

    Exit Sub
frmTestEqp_Add_Error:

    Call ErrMsgProc("frmTestEqp - Private Sub cmdSave()")

End Sub

Private Sub cmdSerch_Click()

    Dim objTestItem As clsCommon
    
    Set objTestItem = New clsCommon
    
    With objTestItem
        Call .SetAdoCn(AdoCn_SQL)
        Set mAdoRs = .Get_TestItem("")
    End With
    
    Set objTestItem = Nothing
    
    Call PopUp_List.ListItems.Clear
    If Not mAdoRs Is Nothing Then
        If Not mAdoRs.EOF Then
            Call DataLoadLvw(PopUp_List, vbCr, vbTab, mAdoRs.GetString)
            Call PopUp_List.ListItems.Remove(PopUp_List.ListItems.Count)
            
            With pnlTestitem
                .Visible = True
                .ZOrder
            End With
            PopUp_List.SetFocus
        End If
    Else
        Call ShowMessage("��ϵ� �˻��׸��� �����ϴ�.")
    End If
    
    Set mAdoRs = Nothing
End Sub

Private Sub Form_Load()
    
    CaptionBar1.Caption = INS_NAME & " Instruments Test Item Link ."
    Call cmdClear
    Call f_subSet_ListView
    Call f_subSet_EqpData(INS_CODE)
    
    With pnlTestitem
        .Moveble = True
    End With
    
    Set PopUp_List = lvwTestitem
    
    With PopUp_List
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            .Add 1, , "�˻��ڵ�", (PopUp_List.Width - 310) * 0.7
            .Add 2, , "�˻��׸�", (PopUp_List.Width - 310) * 0.3
            .Add 3, , "��ü", (PopUp_List.Width - 310) * 0.15
            .Add 4, , "����", (PopUp_List.Width - 310) * 0.15
'            .Add 4, , "", 0
'            .Add 5, , "", 0
'            .Add 6, , "Delta", (PopUp_List.Width - 310) * 0.1
'            .Add 7, , "DeltaGgn", (PopUp_List.Width - 310) * 0.1
'            .Add 8, , "Panic Low", (PopUp_List.Width - 310) * 0.1
'            .Add 9, , "Panic High", (PopUp_List.Width - 310) * 0.1
        End With
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If PopUp_List Is Nothing Then Set PopUp_List = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub

Private Sub Image1_DblClick()
    If lvwTstListEqp.Top > txtTstnmEqp.Top Then
        Call lvwTstListEqp.Move(Label1.left, CaptionBar1.Height, lvwTstListEqp.Width, ScaleHeight - (CaptionBar1.Height + fraCmdBar.Height + 30))
        txtTstcdEqp.Enabled = False
        txtTstnmEqp.Enabled = False
        cmdEqpItm_Add.Enabled = False
        cmdEqpItm_Del.Enabled = False
        Call lvwTstListEqp.ZOrder
    Else
        Call lvwTstListEqp.Move(Label1.left, lvwTestListLab.Top, lvwTstListEqp.Width, lvwTestListLab.Height)
        txtTstcdEqp.Enabled = True
        txtTstnmEqp.Enabled = True
        cmdEqpItm_Add.Enabled = True
        cmdEqpItm_Del.Enabled = True
    End If
End Sub

Private Sub lvwTestListLab_Click()
    Dim itemX As ListItem

    Set itemX = lvwTestListLab.SelectedItem
    If Not itemX Is Nothing Then
        With itemX
            lblTstcdEqp = .Text             '���˻� �ڵ�
            lblTstnmEqp = .SubItems(2)      '���˻� �̸�
            txtTestCd = .SubItems(1)        '�ӻ�˻� �ڵ�
            txtTestNm = .SubItems(2)        '�ӻ�˻� �̸�
            'txtSpccd = .SubItems(3)         '��ü��ȣ
            'txtAuto = .SubItems(4)          '
            txtRefL = .SubItems(3)
            txtRefH = .SubItems(4)
            txtOutseq = .SubItems(5)
            'txtDelta = .SubItems(7)
            'txtDeltagbn = .SubItems(8)
            'txtPanic(0) = .SubItems(9)
            'txtPanic(1) = .SubItems(10)
        End With
    End If
End Sub

Private Sub lvwTestListLab_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SetListView_Sort(lvwTestListLab, ColumnHeader)
End Sub

Private Sub lvwTstListEqp_Click()

    Dim itemX As ListItem
    
    Set itemX = lvwTstListEqp.SelectedItem
    
    If itemX Is Nothing Then
        Exit Sub
    Else
        lblTstcdEqp = Trim(itemX.Text)
        lblTstnmEqp = Trim(itemX.SubItems(1))
        txtTestCd.Text = ""
        txtTestNm.Text = ""
        txtOutseq.Text = ""
        txtRefL.Text = ""
        txtRefH.Text = ""
        
        txtTestCd.SetFocus
    End If
    
    Set itemX = Nothing

End Sub

Private Sub lvwTstListEqp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Call SetListView_Sort(lvwTstListEqp, ColumnHeader)

End Sub

Private Sub lvwTstListEqp_DblClick()
    
    On Error GoTo lvwTstListEqp_DblClick
    
    If MsgBox("���˻��ڵ带 �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Dim itemX       As ListItem
    Dim strTestEqp  As String, intRow   As Integer
    
    Set itemX = lvwTstListEqp.SelectedItem
    
    If Not itemX Is Nothing Then
        AdoCn_Jet.Execute "delete from INTERFACE002 where EQP_CD = '" & INS_CODE & "' and TESTCD_EQP = '" & Trim$(itemX.Text) & "'"
    
        lblTstcdEqp = "":   lblTstnmEqp = ""
    End If
    Set itemX = Nothing
       
    Call f_subSet_EqpData(INS_CODE)
    
    Exit Sub
    
lvwTstListEqp_DblClick:
    Set itemX = Nothing
    Call ErrMsgProc("frmTestEqp - Private Sub lvwTstListEqp_DblClick()")

End Sub

Private Sub lvwTstListEqp_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call lvwTstListEqp_Click
        KeyAscii = 0
        Exit Sub
    End If

End Sub

Private Sub pnlTestitem_CloseMe()
    pnlTestitem.Visible = False
End Sub

Private Sub f_subSet_EqpData(ByVal strEqp_Cd As String)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemX   As ListItem
    
    lvwTstListEqp.ListItems.Clear
    lvwTestListLab.ListItems.Clear
    
    sqlDoc = "select TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM," & _
             "       AUTOVERIFY, REMARK,     REFL,    REFH,   DELTA," & _
             "       DELTAGBN,   PANICL,     PANICH" & _
             "  from INTERFACE002" & _
             " where EQP_CD = '" & INS_CODE & "'" & _
             " order by TESTCD_EQP, TESTCD"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        Set itemX = lvwTstListEqp.ListItems.Add(, , Trim(adoRS("TESTCD_EQP") & ""), , "TST_E")
            itemX.SubItems(1) = Trim(adoRS("TESTNM_EQP") & "")
        Set itemX = Nothing
        
        If Trim$(adoRS("TESTCD") & "") <> "" Then
            Set itemX = lvwTestListLab.ListItems.Add(, , Trim(adoRS("TESTCD_EQP") & ""), , "TST_M")
            With itemX
                .SubItems(1) = Trim$(adoRS("TESTCD") & "")
                .SubItems(2) = Trim$(adoRS("TESTNM") & "")
                .SubItems(3) = Trim$(adoRS("REFL") & "")
                .SubItems(4) = Trim$(adoRS("REFH") & "")
                .SubItems(5) = Trim$(adoRS("OUT_SEQ") & "")
            End With
        End If
        Set itemX = Nothing
        
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subSet_ListView()
    
    Dim lvwWidth    As Long
    
    With lvwTstListEqp
        .View = lvwReport
        Set .SmallIcons = imlList
        Set .ColumnHeaderIcons = imlList
        
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        lvwWidth = .Width - 310
        With .ColumnHeaders
            .Clear
            Call .Add(1, "Code", "�˻��ڵ�(����)", lvwWidth * 0.4)
            Call .Add(2, "Name", "�˻��", lvwWidth * 0.6)
        End With
    End With
    
    With lvwTestListLab
        .View = lvwReport
        Set .SmallIcons = imlList
        Set .ColumnHeaderIcons = imlList
        
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .MultiSelect = True
        lvwWidth = .Width - 310
        With .ColumnHeaders
            .Clear
            Call .Add(1, "CodeEqp", "����ڵ�", lvwWidth * 0.15)
            Call .Add(2, "Code", "�˻��ڵ�", lvwWidth * 0.5)
            Call .Add(3, "Name", "�˻��", lvwWidth * 0.15)
            Call .Add(4, "RefL", "����ġ(L)", lvwWidth * 0.1, lvwColumnCenter)
            Call .Add(5, "RefH", "����ġ(H)", lvwWidth * 0.1, lvwColumnCenter)
            Call .Add(6, "Prtno", "����", lvwWidth * 0.05, lvwColumnCenter)
        End With
    End With
    
End Sub

Private Sub f_subClear_Form()
    
    txtTstcdEqp = ""
    txtTstnmEqp = ""
    lblTstcdEqp = ""
    lblTstnmEqp = ""
    txtTestCd = ""
    txtTestNm = ""
    txtSpccd = ""
    txtAuto = ""
    txtRefL = "":   txtRefH = ""
    txtDeltagbn = ""
    txtDelta = ""
    txtPanic(0) = ""
    txtPanic(1) = ""
    txtOutseq = ""
    
End Sub

Private Sub PopUp_List_DblClick()
    Dim itemX As ListItem
    Set itemX = PopUp_List.SelectedItem
    
    If Not itemX Is Nothing Then
        txtTestCd = itemX.Text
        txtTestNm = itemX.SubItems(1)
        txtSpccd = itemX.SubItems(2)
        
        Call pnlTestitem_CloseMe
        txtVIndex.SetFocus
    End If
End Sub

Private Sub PopUp_List_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call PopUp_List_DblClick
        KeyAscii = 0
    End If
End Sub

Private Sub txtDelta_GotFocus()

    With txtDelta
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtDelta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub


Private Sub txtDeltagbn_GotFocus()

    With txtDeltagbn
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtDeltagbn_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub


Private Sub txtPanic_GotFocus(Index As Integer)

    With txtPanic(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtPanic_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub


Private Sub txtRefH_GotFocus()

    With txtRefH
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtRefH_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub


Private Sub txtRefL_GotFocus()

    With txtRefL
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtRefL_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub


Private Sub txtTestCD_Change()
    txtTestNm = ""
    txtSpccd = ""
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)

    'txtTestCD.Locked = False
    If KeyAscii = vbKeyReturn Then
        'Call cmdSerch_Click
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub

Private Sub txtTstnmEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdEqpItm_Add_Click
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtVIndex_GotFocus()
    Call TextBoxs_GotFocus(txtVIndex)
End Sub

Private Sub txtVIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txtVIndex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtVIndex.IMEMode = 8
End Sub
