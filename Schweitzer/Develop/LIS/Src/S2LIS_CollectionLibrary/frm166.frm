VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm166OgyCollect 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14670
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox Picture1 
      Height          =   7035
      Left            =   75
      ScaleHeight     =   6975
      ScaleWidth      =   8355
      TabIndex        =   32
      Top             =   1980
      Width           =   8415
      Begin FPSpread.vaSpread tblPtList 
         Height          =   6810
         Left            =   0
         TabIndex        =   33
         Tag             =   "15109"
         Top             =   0
         Width           =   8340
         _Version        =   196608
         _ExtentX        =   14711
         _ExtentY        =   12012
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   23
         MaxRows         =   26
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frm166.frx":0000
         VisibleCols     =   3
         VisibleRows     =   25
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   6240
      Left            =   8550
      ScaleHeight     =   6180
      ScaleWidth      =   5880
      TabIndex        =   23
      Top             =   2265
      Width           =   5940
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   345
         TabIndex        =   24
         Top             =   555
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         BackColor       =   13752531
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtCount 
         Height          =   330
         Left            =   345
         TabIndex        =   25
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   13752531
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
         LeftGab         =   100
      End
      Begin FPSpread.vaSpread tblCount 
         Height          =   5970
         Left            =   2415
         TabIndex        =   26
         Tag             =   "15109"
         Top             =   0
         Width           =   3465
         _Version        =   196608
         _ExtentX        =   6112
         _ExtentY        =   10530
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   14737632
         MaxCols         =   3
         MaxRows         =   18
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frm166.frx":07D6
         VisibleCols     =   3
         VisibleRows     =   15
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   345
         TabIndex        =   27
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
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
         Caption         =   "�� ä����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   345
         TabIndex        =   28
         Top             =   1065
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
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
         Caption         =   "�� ȯ�ڼ�"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��"
         Height          =   255
         Left            =   1620
         TabIndex        =   29
         Tag             =   "20104"
         Top             =   1515
         Width           =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2400
         X2              =   2400
         Y1              =   0
         Y2              =   4770
      End
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   22
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   21
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   20
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CheckBox chkPrintFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��¾���"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8940
      TabIndex        =   1
      Top             =   465
      Width           =   1305
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��ü����(&A)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1455
      TabIndex        =   0
      Top             =   1665
      Width           =   1560
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   8550
      TabIndex        =   2
      Top             =   45
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "��� �ɼ�"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   75
      TabIndex        =   10
      Top             =   45
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "����� ����"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   4
      Left            =   75
      TabIndex        =   19
      Top             =   1605
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
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
      Caption         =   "ȯ�ڸ���Ʈ"
      Appearance      =   0
   End
   Begin VB.Frame fraPrtOption 
      BackColor       =   &H00DBE6E6&
      Height          =   1320
      Left            =   8550
      TabIndex        =   3
      Top             =   285
      Width           =   5925
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ�Lable And ä�� ����Ʈ"
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Top             =   420
         Width           =   3180
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   1  '������ ����
         Height          =   345
         Left            =   3255
         TabIndex        =   4
         Top             =   915
         Width           =   750
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   4020
         TabIndex        =   7
         Top             =   900
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MedControls1.LisLabel lblColList 
         Height          =   255
         Left            =   855
         TabIndex        =   8
         Top             =   945
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   450
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
         Caption         =   "ä������Ʈ ������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPage 
         Height          =   255
         Left            =   4335
         TabIndex        =   9
         Top             =   975
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
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
         Caption         =   "��"
         Appearance      =   0
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ� Only"
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   675
         Width           =   3180
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1320
      Left            =   75
      TabIndex        =   11
      Top             =   285
      Width           =   8340
      Begin VB.CommandButton cmdGetOrders 
         BackColor       =   &H00F4F0F2&
         Caption         =   "��ȸ(&Q)"
         Height          =   510
         Left            =   6885
         Style           =   1  '�׷���
         TabIndex        =   14
         Tag             =   "15101"
         Top             =   690
         Width           =   1320
      End
      Begin VB.TextBox txtDeptCd 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1065
         TabIndex        =   13
         Top             =   240
         Width           =   1110
      End
      Begin VB.CommandButton cmdWardList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2175
         MousePointer    =   14  'ȭ��ǥ�� ����ǥ
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   240
         Width           =   315
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   360
         Left            =   2490
         TabIndex        =   15
         Top             =   255
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   635
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   375
         Left            =   1065
         TabIndex        =   16
         Top             =   720
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
         Format          =   38141952
         CurrentDate     =   36342.5951388889
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   17
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
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
         Caption         =   "�μ��ڵ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   18
         Top             =   720
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
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
         Caption         =   "ó����"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ProgressBar pbrPtCnt 
      Height          =   150
      Left            =   8655
      TabIndex        =   30
      Top             =   2025
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   300
      Left            =   8550
      TabIndex        =   31
      Top             =   1620
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "���� ��Ȳ"
      LeftGab         =   100
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00808080&
      FillColor       =   &H00D8DEDA&
      FillStyle       =   0  '�ܻ�
      Height          =   330
      Index           =   1
      Left            =   8550
      Shape           =   4  '�ձ� �簢��
      Top             =   1935
      Width           =   5910
   End
End
Attribute VB_Name = "frm166OgyCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'---- Collect
Private objSQL                  As clsLISSqlCollection
Private objCollect              As clsLISCollectioin
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private CleanFg                 As Boolean
Private blnInitFg               As Boolean
Private intPtCount              As Integer
Private intErrCount             As Integer

Private Const lngMaxRows = 25
Private Const lngRowHeight = 12
Public Event LastFormUnload()
 
Private Sub cmdClear_Click()
    Call ClearRtn(1)
    txtDeptCd.SetFocus
End Sub

Private Sub dtpToTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub Form_Activate()

    If blnInitFg Then Exit Sub
    
    txtCopy.Text = 1
    dtpToTime.Value = Format(GetSystemDate, "YYYY-MM-DD HH:MM:SS")
    CleanFg = True
    intErrCount = 0
    txtDeptCd.Text = ""
    txtDeptCd.SetFocus
    chkPrintFg.Value = 0
    optOption(1).Value = True
    
    blnInitFg = True
    
End Sub

Private Sub Form_Deactivate()
    Set objMyList = Nothing
End Sub

Private Sub Form_Load()

    Me.Show
    blnInitFg = False
    Set objSQL = New clsLISSqlCollection
    Set objCollect = New clsLISCollectioin
End Sub
Private Sub chkAll_Click()
    With tblPtList
        .Col = 1: .Col2 = 1
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .Value = chkAll.Value
        .BlockMode = False
    End With
End Sub

'& ��� Option ����
Private Sub chkPrintFg_Click()
    If chkPrintFg.Value = 1 Then
        optOption(0).Value = False
        optOption(1).Value = False
        fraPrtOption.Enabled = False
    Else
        optOption(1).Value = True
        fraPrtOption.Enabled = True
    End If
End Sub

'% ����
Private Sub cmdExit_Click()
    Unload Me
    Set objMyList = Nothing
    Set objSQL = Nothing
    Set objCollect = Nothing
    If IsLastForm Then RaiseEvent LastFormUnload
    
End Sub

'% �ϰ�ä�� ����
Private Sub cmdGenerate_Click()
    Dim Resp        As VbMsgBoxResult
    Dim SelCount    As Integer

    Dim SavePtId    As String
    Dim sWorkarea   As String
    Dim sAccdt      As String
    
    Dim sBuildCd    As String
    Dim sBuildNm    As String
    Dim sWorkDt     As String
    Dim sWorkTm     As String
    Dim iAccseq     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    
    Set objCollect = New clsLISCollectioin

    sWorkDt = Format(GetSystemDate, CS_DateDbFormat)
    sWorkTm = Format(GetSystemDate, CS_TimeDbFormat)

    Call objCollect.SetWardCol(sWorkDt, sWorkTm, txtDeptCd.Text)

    tblCount.Row = 0
    intErrCount = 0
    SelCount = 0
    SavePtId = ""

    'Locking...
    txtDeptCd.Enabled = False
    txtDeptCd.BackColor = &H8000000F
    cmdWardList.Enabled = False
    dtpToTime.Enabled = False
    cmdGetOrders.Enabled = False

    MouseRunning  '13
    
'    Dim objBld As clsBasisData
    Dim strBld As String
    
    With tblPtList
        For i = 1 To intPtCount
            .Row = i
            If pbrPtCnt.Value >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 1
            pbrPtCnt.Value = pbrPtCnt.Value + 1
            DoEvents
            '* ���ܹ�ư Check
            .Col = 1: If .Value = 1 Then GoTo Skip
            

            SelCount = SelCount + 1

            '* ä������
            .Col = 15: If Trim(.Value) <> "" Then Call DoCollection(i)
            DoEvents
            '* Delivery Location �� Count
            .Col = 2
            For j = 1 To objCollect.ColCount
                Call objCollect.GetLabNumbers(j, sWorkarea, sAccdt, iAccseq, sBuildCd)
'                Call ObjLISComCode.Building.KeyChange(sBuildCd)
'                sBuildNm = ObjLISComCode.Building.Fields("buildnm")
'                Set objBld = Nothing
'                Set objBld = New clsBasisData
                sBuildNm = GetBuildNm(sBuildCd)
'                Set objBld = Nothing
                For k = 1 To tblCount.DataRowCnt
                    tblCount.Row = k
                    tblCount.Col = 3
                    If tblCount.Value = sBuildCd Then
                        '* ��ü�� Count
                        tblCount.Col = 2
                        tblCount.Text = CStr(Val(tblCount.Text) + 1)
                        GoTo NextCol
                    End If
                Next

                If tblCount.DataRowCnt >= tblCount.MaxRows Then tblCount.MaxRows = tblCount.MaxRows + 1
                tblCount.Row = tblCount.DataRowCnt + 1
                tblCount.Col = 1: tblCount.Value = sBuildNm
                tblCount.Col = 2: tblCount.Text = "1"
                tblCount.Col = 3: tblCount.Value = sBuildCd
NextCol:
            Next

            '* ȯ�ڼ� Count
            .Row = i
            .Col = 3
            If SavePtId <> Trim(.Value) Then
                lblPtCount.Caption = Val(lblPtCount.Caption) + 1
                SavePtId = .Value
            End If
            '* ä�� Class Initialize
            Call objCollect.InitRtn
            DoEvents
Skip:
        Next

        'ä����
        lblColNm.Caption = ObjSysInfo.EmpId

    End With

    If SelCount = 0 Then
        MouseDefault  '0
        Call cmdClear_Click
        MsgBox "ó���� ����Ÿ�� �����ϴ�..", vbInformation, "Message"
        Exit Sub
    End If

    pbrPtCnt.Value = pbrPtCnt.Max
    DoEvents

    MouseDefault  '0

    If intErrCount > 0 Then
        MsgBox CStr(intErrCount) & "���� ������ �߻��߽��ϴ�.."
    Else
        If optOption(0).Value Then
            Resp = MsgBox("��� ���������� ä��ó�� �Ǿ����ϴ�.." & vbCrLf & _
                                    "ä������Ʈ�� ���� ����Ͻðڽ��ϱ� ? ", vbYesNo, "ä������Ʈ ���")
            If Resp = vbYes Then
                For i = 1 To tblCount.DataRowCnt
                    tblCount.Row = i
                    tblCount.Col = 3:  sBuildCd = tblCount.Value
                    tblCount.Col = 1:  sBuildNm = tblCount.Value
                    For j = 1 To Val(txtCopy.Text)
                        Call PrintColList(txtDeptCd.Text, lblWardNm.Caption, sWorkDt, sWorkTm, sBuildCd, sBuildNm)
                    Next
                Next
            End If
        Else
            Call MsgBox("��� ���������� ä��ó�� �Ǿ����ϴ�..", vbInformation, "�޼���")
        End If

        Call ClearRtn(0)
        txtDeptCd.SetFocus
   End If

End Sub

'& ä�� Ŭ���� objCollect �� �̿��Ͽ� �ش� ȯ�ڵ��� ó���� ä�������Ѵ�.
Private Sub DoCollection(ByVal Row As Long)
    Dim Rs          As Recordset
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim SqlStmt     As String
    Dim tmpData()   As String
    
    Dim Success     As Boolean
    Dim i           As Integer
    
    ReDim tmpData(0 To 16)

    With tblPtList
        .Row = Row
                    tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)
        .Col = 3:   tmpData(1) = .Value           'ȯ��ID
        .Col = 4:   tmpData(2) = .Value           'ȯ�ڸ�
        .Col = 14:  tmpData(3) = .Value           'ȯ�ڼ���
        .Col = 7:
            If IsDate(Format(.Value, CS_DateMask)) Then
               tmpData(4) = DateDiff("y", Format(.Value, CS_DateMask), GetSystemDate)  'ȯ���Ϸ�
            Else
               tmpData(4) = 50000       '��������� ��Ȯ���� ������� Max��
            End If
        .Col = 8:   tmpData(5) = .Value       '�Կ���
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)                    '�Է���
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)                    '�Է½ð�
        tmpData(8) = ObjSysInfo.EmpId                                               '�Է���
        tmpData(9) = ""                                                             '��������ȣ
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat)                    'ä����
        objCollect.ColTm = Format(GetSystemDate, CS_TimeDbFormat)               'ä����
        tmpData(11) = ObjSysInfo.EmpId                                              'ä����
        .Col = 9:   tmpData(12) = .Value                                            '����ID
        .Col = 12: tmpData(13) = .Value                                             '����ID
        .Col = 12: tmpData(14) = .Value                                             'ȣ��ID
        tmpData(15) = ""                                                            'ħ��ID
        tmpData(16) = ObjSysInfo.BuildingCd                                         '** ä���� ����Ǵ� �ǹ��ڵ�
        
        Call objCollect.SetColData(tmpData)
    
    End With

    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)

    ' ó�泻�� �˻�
    SqlStmt = objSQL.OGySQlOrderRead(tmpData(1), tmpDate, tmpTime, txtDeptCd.Text)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn

    ReDim tmpData(0 To 20)

    With Rs
        For i = 1 To .RecordCount
            tmpData(0) = ObjSysInfo.BuildingCd
            tmpData(1) = Trim("" & .Fields("WorkArea").Value)   'WorkArea
            tmpData(2) = Trim("" & .Fields("SpcCd").Value)      'SpcCd
            tmpData(3) = Trim("" & .Fields("StoreCd").Value)    'StoreCd
            tmpData(4) = Trim("" & .Fields("StatFg").Value)
            tmpData(5) = Format("" & Rs.Fields("ReqDt").Value, CS_DateLongMask) & " " & _
                         Format("" & Rs.Fields("ReqTm").Value, CS_TimeLongMask)        '���ä���Ͻ�
            
            tmpData(6) = Trim("" & .Fields("TestDiv").Value)    'TestDiv
            tmpData(7) = Trim("" & .Fields("MultiFg").Value)    'MultiFg
            tmpData(8) = Trim("" & .Fields("SpcGrp").Value)     'SpcGrp
            tmpData(9) = Trim("" & .Fields("OrdDt").Value)      'OrdDt
            tmpData(10) = Trim("" & .Fields("OrdNo").Value)     'OrdNo
            tmpData(11) = Trim("" & .Fields("OrdSeq").Value)    'OrdSeq
            tmpData(12) = Trim("" & .Fields("OrdCd").Value)     'OrdCd
            tmpData(13) = Trim("" & .Fields("DeptCd").Value)    'DeptCd
            tmpData(14) = Trim("" & .Fields("OrdDoct").Value)   'OrdCd
            tmpData(15) = Trim("" & .Fields("MajDoct").Value)   'OrdCd
            tmpData(16) = Trim("" & .Fields("AbbrNm5").Value)   'ó�� ����
            tmpData(17) = Trim("" & .Fields("LabelCnt").Value)  '��������
            
'            Call ObjLISComCode.LisItem.KeyChange(Trim("" & .Fields("TestCd").Value))
            tmpData(18) = GetLabDiv(Trim("" & .Fields("TestCd").Value)) 'ObjLISComCode.LisItem.Fields("labdiv")    'LabDiv
            
            Call GetSpcInfo(tmpData(2), tmpData(19), tmpData(20))
'            Call ObjLISComCode.LisSpc.KeyChange(tmpData(2))         '��ü�ڵ�
'            tmpData(19) = ObjLISComCode.LisSpc.Fields("spcabbr")    '��ü����
'            tmpData(20) = ObjLISComCode.LisSpc.Fields("labrange")   '�̻���������ȣ����
            
            Call objCollect.SetAddLabCollect(tmpData)
            .MoveNext
        Next
    End With

    ' ä�� ����
    Success = objCollect.DoCollection
    If Not Success Then
        tblPtList.Row = Row: tblPtList.Row2 = Row
        tblPtList.Col = -1
        tblPtList.BlockMode = True
        tblPtList.ForeColor = &HFF&       '������
        tblPtList.BlockMode = False
        intErrCount = intErrCount + 1
    End If
    Set Rs = Nothing
End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b"
    strSQL = strSQL & " where " & DBW("b.cdindex=", LC3_WorkArea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
    GetLabDiv = Rs.Fields("field2").Value & ""
    End If
    Set Rs = Nothing
End Function

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
    vSpcAbbr = Rs.Fields("spcabbr").Value & ""
    vLabRng = Rs.Fields("labrange").Value & ""
    End If
    Set Rs = Nothing
End Sub

'% �������� ���� �Կ����� ȯ�ڵ��� ó���� �˻��Ѵ�.
Private Sub cmdGetOrders_Click()
    Dim Rs          As Recordset
    Dim SqlStmt     As String
    Dim tmpPtId     As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpOrdDiv   As String
    Dim tmpSpcCd    As String
    Dim i           As Integer
    
    If Trim(txtDeptCd.Text) = "" Then
        MsgBox "�μ��ڵ带 �Է��ϼ���.", vbInformation, "���������"
        txtDeptCd.SetFocus
        Exit Sub
    End If
    
    '2001-11-07 : ������ �����ϰ�ä�� ���� ���� --------------------------------------------------
    Dim objStatus As New jProgressBar.clsProgress
    With objStatus
        .Container = Me
        .Left = LisLabel1.Left
        .Top = LisLabel1.Top
        .Width = LisLabel1.Width
        .Height = 280
        .Message = "����ΰ� ä�� ����� ��ȸ�ϰ� �ֽ��ϴ�..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = LisLabel1.Width
'        .XPos = LisLabel1.Left
'        .YPos = LisLabel1.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "����ΰ� ä�� ����� ��ȸ�ϰ� �ֽ��ϴ�."
'        .Max = 100
'        .Value = 50
    End With

    Set objCollect = New clsLISCollectioin
    If Not objCollect.Archive_WardColData(txtDeptCd.Text) Then
        MsgBox "����ΰ� �ϰ�ä�� ���� Archive�� ������ �߻��߽��ϴ�." & vbCrLf & _
                "����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�. (��" & ObjSysInfo.HelpLine & ")", vbCritical, "�����߻�"
    End If
    Set objStatus = Nothing
    Set objCollect = Nothing
    '---------------------------------------------------------------------------------------------
    
    Call TableClear(1)
    
    MouseRunning
    
    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)

    Set Rs = New Recordset
    Rs.Open objSQL.OGYOutOrder(tmpDate, tmpTime, txtDeptCd.Text), DBConn
    
    If Rs.EOF Then
        MsgBox "ä������� �����ϴ�..", vbInformation, "�ܷ�ä��"
        cmdGenerate.Enabled = False
        MouseDefault
        Exit Sub
    End If

    With tblPtList
        .MaxRows = 0
        If Rs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
        Else
            .MaxRows = Rs.RecordCount
        End If
        .Row = 1
        intPtCount = 0

        For i = 1 To Rs.RecordCount
            If tmpPtId <> Trim(Rs.Fields("PtId").Value & "") Then
                intPtCount = intPtCount + 1
                .Row = intPtCount
                .Col = 2: .Text = "" & Rs.Fields("DeptCd").Value    '����ID
                .Col = 3: .Text = "" & Rs.Fields("PtId").Value     'ȯ��ID
                .Col = 4: .Text = "" & Rs.Fields("PtNm").Value   '����
                .Col = 7: .Text = "" & Rs.Fields("DOB").Value    '�������
                
               .Col = 14: .Text = Trim("" & Rs.Fields("sex").Value)
                If IsNumeric(.Text) Then
                    .Text = Choose((Val(.Text) Mod 2) + 1, "F", "M")
                End If
                tmpPtId = "" & Rs.Fields("PtId").Value

            End If
            .Col = 9: .Text = "" & Rs.Fields("DeptCd").Value  '�����
            .Col = 10: .Text = "" & Rs.Fields("OrdDoct").Value 'ó����
            .Col = 11: .Text = "" & Rs.Fields("MajDoct").Value '��ġ��
            tmpStatFg = "" & Rs.Fields("StatFg").Value      '���޿���
            tmpOrdDiv = "" & Rs.Fields("orddiv").Value             'ó�汸��
            tmpSpcCd = "" & Rs.Fields("SpcCd").Value               '��ü


            If tmpStatFg = "1" Then
                .Col = 5
                If InStr(1, .Value, Rs.Fields("SpcNm").Value) = 0 Then
                    .Value = .Value & Rs.Fields("SpcNm").Value & ", "     '���ް�ü
                End If
            Else
                .Col = 6
                If InStr(1, .Value, Rs.Fields("SpcNm").Value) = 0 Then
                    .Value = .Value & Rs.Fields("SpcNm").Value & ", "
                End If
            End If

            .Col = 15: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
            .Col = 19: .Text = Format(GetSystemDate, "YY-MM-DD")
            .Col = 20: .Text = Format(GetSystemDate, "HH:MM")

            Rs.MoveNext
        Next

        pbrPtCnt.Min = 0
        pbrPtCnt.Max = .DataRowCnt + 10
        pbrPtCnt.Value = 0

        .Row = 1: .Row2 = .MaxRows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False

    End With

    cmdGenerate.Enabled = True
    CleanFg = False
    Set Rs = Nothing

    MouseDefault

End Sub

' ���ؽð��� ����Ǹ� Clear
Private Sub dtpToTime_Change()
    If Not CleanFg Then Call TableClear(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objSQL = Nothing
    Set objCollect = Nothing
End Sub


Private Sub optOption_Click(Index As Integer)
    
    Select Case Index
        Case 0, 2: txtCopy.Text = 1
                   txtCopy.Enabled = True
        Case 1:    txtCopy.Text = 0
                   txtCopy.Enabled = False
    End Select

End Sub

Private Sub cmdWardList_Click()
'% �����ڵ� ����Ʈ�� �˾��Ѵ�.

    Set objMyList = New clsPopUpList
    With objMyList
        .Connection = DBConn
        .FormCaption = "�μ��ڵ� ��ȸ"
        .ColumnHeaderText = "�μ��ڵ�;�μ���"
         Call .LoadPopUp(objSQL.SqlGetBatchDept) ', 2700, cmdWardList.Left)
        If .SelectedString <> "" Then
            txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    Set objMyList = Nothing
End Sub


Private Sub ClearRtn(ByVal intOpt As Integer)
    'Unlocking...
    txtDeptCd.Enabled = True
    txtDeptCd.BackColor = vbWhite
    cmdWardList.Enabled = True
    dtpToTime.Enabled = True
    cmdGetOrders.Enabled = True
    cmdGenerate.Enabled = False

    txtDeptCd.Text = ""
    lblWardNm.Caption = ""
    dtpToTime.Value = Format(Now, "YYYY/MM/DD HH:MM:SS")
    pbrPtCnt.Value = 0
    chkPrintFg = 0
    optOption(1).Value = True
    intErrCount = 0
    Call TableClear(intOpt)

End Sub


'% Table���� Clear�Ѵ�
Private Sub TableClear(ByVal intOpt As Integer)
    tblPtList.MaxRows = 0
    tblPtList.MaxRows = 50
    If intOpt = 1 Then
        lblColNm.Caption = ""
        lblPtCount.Caption = ""
        tblCount.MaxRows = 0
        tblCount.MaxRows = 50
        CleanFg = True
    End If
End Sub

Private Sub PrintColList(ByVal pDeptCd As String, ByVal pWardNm As String, ByVal pWorkDt As String, _
                        ByVal pWorkTm As String, ByVal pBuildCd As String, ByVal pBuildNm As String)

    Dim MyReport As clsWardColList
    Dim strTitleNm As String
    
    Set MyReport = New clsWardColList
    
    strTitleNm = "�ܷ� ä�� ����Ʈ"

    With MyReport
        .WardId = pDeptCd
        .WardNm = pWardNm
        .WorkDt = pWorkDt
        .WorkTm = pWorkTm
        .BuildCd = pBuildCd
        .BuildNm = pBuildNm
        .TestDiv = "0"  'chkTestdiv.Value
        .TitleNm = strTitleNm
        .SetCrpt CReport
        Call .Print_ColList
    End With

    Set MyReport = Nothing
End Sub


Private Sub txtDeptCd_Change()
    If Not CleanFg Then Call TableClear(1)
End Sub

Private Sub txtDeptCd_GotFocus()
    With txtDeptCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDeptCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objMyList Is Nothing Then Call cmdWardList_Click
    End If
End Sub

Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If txtDeptCd.Text = "" Then
            txtDeptCd.SetFocus
            Exit Sub
        Else
'            Set objDept = New clsBasisData
            strDept = GetDeptNm(txtDeptCd.Text)
'            Set objDept = Nothing
            
            If strDept = "" Then
                MsgBox "�μ� �ڵ带 Ȯ���ϼ���.."
                txtDeptCd.Text = ""
                Call cmdWardList_Click
                Exit Sub
            Else
                lblWardNm.Caption = strDept
                SendKeys "{TAB}"
            End If
            
'            If Not ObjLISComCode.DeptCd.Exists(txtDeptCd.Text) Then
'                MsgBox "�μ� �ڵ带 Ȯ���ϼ���.."
'                txtDeptCd.Text = ""
'                Call cmdWardList_Click
'                Exit Sub
'            Else
'                ObjLISComCode.DeptCd.KeyChange txtDeptCd.Text
'                lblWardNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'                SendKeys "{TAB}"
'            End If
        End If
    End If
End Sub
