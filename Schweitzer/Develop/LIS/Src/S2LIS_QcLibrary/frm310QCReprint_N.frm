VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm310QCReprint_N 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   15240
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   960
      Left            =   11160
      TabIndex        =   9
      Top             =   -45
      Width           =   3300
      Begin VB.OptionButton optWorkTime 
         BackColor       =   &H00F4F0F2&
         Caption         =   "�����"
         Height          =   285
         Index           =   4
         Left            =   60
         Style           =   1  '�׷���
         TabIndex        =   32
         Top             =   480
         Width           =   765
      End
      Begin VB.OptionButton optWorkTime 
         BackColor       =   &H00F4F0F2&
         Caption         =   "�߰�"
         Height          =   285
         Index           =   3
         Left            =   2400
         Style           =   1  '�׷���
         TabIndex        =   15
         Tag             =   "200001:235959"
         Top             =   150
         Width           =   765
      End
      Begin VB.OptionButton optWorkTime 
         BackColor       =   &H00F4F0F2&
         Caption         =   "����"
         Height          =   285
         Index           =   2
         Left            =   1620
         Style           =   1  '�׷���
         TabIndex        =   14
         Tag             =   "120001:200000"
         Top             =   150
         Width           =   765
      End
      Begin VB.OptionButton optWorkTime 
         BackColor       =   &H00F4F0F2&
         Caption         =   "����"
         Height          =   285
         Index           =   1
         Left            =   840
         Style           =   1  '�׷���
         TabIndex        =   13
         Tag             =   "080001:120000"
         Top             =   150
         Width           =   765
      End
      Begin VB.OptionButton optWorkTime 
         BackColor       =   &H00F4F0F2&
         Caption         =   "��ü"
         Height          =   285
         Index           =   0
         Left            =   60
         Style           =   1  '�׷���
         TabIndex        =   12
         Tag             =   "000001:235959"
         Top             =   150
         Value           =   -1  'True
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpFromTime 
         Height          =   300
         Left            =   855
         TabIndex        =   10
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   67174403
         CurrentDate     =   36859.8743055556
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   300
         Left            =   2115
         TabIndex        =   11
         Top             =   480
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   67174403
         CurrentDate     =   36859.1243055556
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "~"
         Height          =   180
         Left            =   1920
         TabIndex        =   16
         Tag             =   "30310"
         Top             =   540
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   960
      Left            =   75
      TabIndex        =   4
      Top             =   -45
      Width           =   11055
      Begin VB.ComboBox cboLevel 
         Height          =   300
         ItemData        =   "frm310QCReprint_N.frx":0000
         Left            =   9225
         List            =   "frm310QCReprint_N.frx":0010
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   22
         Top             =   135
         Width           =   1785
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ü"
         Height          =   180
         Left            =   9840
         TabIndex        =   21
         Top             =   630
         Width           =   660
      End
      Begin VB.CommandButton cmdControl 
         BackColor       =   &H00F4F0F2&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2835
         Picture         =   "frm310QCReprint_N.frx":00A3
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   525
         Width           =   330
      End
      Begin VB.TextBox txtCtrlCd 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1425
         TabIndex        =   17
         Top             =   540
         Width           =   1395
      End
      Begin VB.ComboBox cboSection 
         Height          =   300
         Left            =   5895
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   5
         Top             =   135
         Width           =   1950
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   300
         Left            =   1425
         TabIndex        =   6
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   67174403
         CurrentDate     =   36545
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   300
         Left            =   3075
         TabIndex        =   7
         Top             =   135
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   67174403
         CurrentDate     =   36545
      End
      Begin MedControls1.LisLabel lblCtrlNm 
         Height          =   360
         Left            =   3180
         TabIndex        =   19
         Top             =   540
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   635
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   45
         TabIndex        =   38
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "��    ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   45
         TabIndex        =   39
         Top             =   540
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "Control"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   4515
         TabIndex        =   40
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "Section"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   7845
         TabIndex        =   41
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "Level"
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "~"
         Height          =   180
         Left            =   2865
         TabIndex        =   8
         Tag             =   "30310"
         Top             =   180
         Width           =   135
      End
   End
   Begin MedControls1.LisLabel lblWarn 
      Height          =   240
      Left            =   9765
      TabIndex        =   34
      Top             =   1440
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   423
      BackColor       =   8388608
      ForeColor       =   12632319
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
      Alignment       =   2
      Caption         =   "�� �˻������� ����Ǿ� ���� ��ȸ�� ����� �ٸ��ϴ�."
      Appearance      =   0
   End
   Begin VB.CommandButton cmdConfig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��������(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   20
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��   ȸ"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   915
      Width           =   1320
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ä  ��"
      Height          =   510
      Left            =   9180
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   11160
      TabIndex        =   23
      Top             =   840
      Width           =   1965
      Begin VB.CheckBox chkBarPrint 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ����"
         Height          =   180
         Left            =   390
         TabIndex        =   24
         ToolTipText     =   "ä��� ���ÿ� ���ڵ带 ����ϸ� �ټ� ó���ӵ��� ���� �ɼ� �ֽ��ϴ�."
         Top             =   240
         Value           =   1  'Ȯ��
         Width           =   1200
      End
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   6750
      Left            =   75
      TabIndex        =   33
      Top             =   1695
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   11906
      _StockProps     =   64
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
      MaxCols         =   20
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frm310QCReprint_N.frx":0155
      TextTip         =   2
   End
   Begin MedControls1.LisLabel lblDesc 
      Height          =   240
      Left            =   75
      TabIndex        =   35
      Top             =   1440
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   423
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "�� ������ �˻������� 1991/12/12-2000/02/02,��ü,��ü,20:40:00~02:00:00, ä�����(����,�κ�,�������) �Դϴ�."
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   8790
      TabIndex        =   36
      Top             =   840
      Width           =   2340
      Begin VB.CheckBox chkSchedule 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�����ٳ�����"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Value           =   1  'Ȯ��
         Width           =   1380
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   75
      TabIndex        =   25
      Top             =   840
      Width           =   8715
      Begin VB.Frame fraQryKey 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '����
         Height          =   315
         Left            =   3540
         TabIndex        =   28
         Tag             =   "'3','4'"
         Top             =   165
         Width           =   5100
         Begin VB.CheckBox chkQryKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��������"
            Height          =   180
            Index           =   0
            Left            =   375
            TabIndex        =   31
            Tag             =   "2"
            Top             =   75
            Width           =   1020
         End
         Begin VB.CheckBox chkQryKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "�κа��"
            Height          =   180
            Index           =   1
            Left            =   2055
            TabIndex        =   30
            Top             =   75
            Width           =   1020
         End
         Begin VB.CheckBox chkQryKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "�������"
            Height          =   180
            Index           =   2
            Left            =   3675
            TabIndex        =   29
            Tag             =   "'5','6'"
            Top             =   75
            Width           =   1020
         End
      End
      Begin VB.OptionButton optQryKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "����� ���"
         Height          =   180
         Index           =   1
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   1470
      End
      Begin VB.OptionButton optQryKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ä�� ���"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frm310QCReprint_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Coding By Legends
'2003/10/14
'QC ������ ä�� �� ���Է¸���Ʈ ��ȸ
'1   /       2/       3/       4/       5/     6/      7/    8/     9/      10/  11/
'����/ó������/ó��ð�/ä������/ä��ð�/ctrlcd/levelcd/lotno/����/������ȣ/����/
'12      /    13/    14/      15/16/   17/    18/19/    20
'����ڵ�/ctrlnm/sectcd/��ü��ȣ/wa/accdt/accseq/��/errmsg

Public Event LastFormUnload()

Private objQC As clsQcMst
Private objOrder As clsQcOrder

Private mvarParentHwnd As Long

Public Property Let ParentHwnd(ByVal vData As Long)
    mvarParentHwnd = vData
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = mvarParentHwnd
End Property

Private Sub cboLevel_Click()
    On Error Resume Next
    
    If Screen.ActiveControl.Name <> cboLevel.Name Then Exit Sub
    Call GetWarningMsg
End Sub

Private Sub cboSection_Click()
    On Error Resume Next
    
    If Screen.ActiveControl.Name <> cboSection.Name Then Exit Sub
    
    Call GetWarningMsg
    
    If lblCtrlNm.Caption = "" Then Exit Sub
    
    txtCtrlCd.Text = ""
    lblCtrlNm.Caption = ""
End Sub

Private Sub chkAll_Click()
'    On Error Resume Next
'    If Screen.ActiveControl.Name <> chkAll.Name Then Exit Sub
    
    If chkAll.Value = 1 Then
        txtCtrlCd.Text = ""
        txtCtrlCd.Enabled = False
        cmdControl.Enabled = False
        lblCtrlNm.Caption = ""
    ElseIf chkAll.Value = 0 Then
        txtCtrlCd.Enabled = True
        cmdControl.Enabled = True
    End If
End Sub

Private Sub chkQryKey_Click(Index As Integer)
    On Error Resume Next
    If Screen.ActiveControl.Name <> chkQryKey(Index).Name Then Exit Sub
    
    Call GetWarningMsg
End Sub

Private Sub chkSchedule_Click()
    On Error Resume Next
    If Screen.ActiveControl.Name <> chkSchedule.Name Then Exit Sub
    
    Call GetWarningMsg
End Sub

Private Sub cmdClear_Click()
    Call InitForm
    
    Call ReadConfig
End Sub

Private Sub cmdCollect_Click()
    Dim lngCnt As Long
    Dim i As Long
    
    cmdCollect.Enabled = False
    For i = 1 To tblResult.DataRowCnt
        tblResult.Col = 1
        tblResult.Row = i
        If tblResult.Value <> "1" Then
            lngCnt = lngCnt + 1
        End If
    Next
    
    If lngCnt = 0 Then
        MsgBox "ó���� �ڷᰡ ���ų� ��� '����'�� ���õǾ����ϴ�.", vbExclamation
        cmdCollect.Enabled = True
        Exit Sub
    End If
    
    If optQryKey(0).Value Then
        Call DoCollection(lngCnt)   'ä���۾� �� ���ڵ� ����
    Else
        Call DoRePrint(lngCnt)  '���ڵ� �����
    End If
    cmdCollect.Enabled = True
End Sub

Private Sub DoCollection(ByVal pProcessCnt As Long)
'ä�� �� ������ ����
'pCnt ó���� ���ڵ� ����

    Dim i As Long
    Dim objPro As clsProgress
    
    '���α׷��� ��
    Set objPro = Nothing
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Width = tblResult.Width
        .Left = tblResult.Left
        .Top = tblResult.Top
        .Height = 470
        .Message = "ä�� �۾��� �����ϰ� �ֽ��ϴ�."
        .Max = pProcessCnt
    End With
    
    For i = 1 To tblResult.DataRowCnt
        tblResult.Row = i
        tblResult.Col = 1
        If tblResult.Value <> "1" Then
            
            tblResult.Col = 4
            If tblResult.Value = "" Then    'ä�����ڰ� ���������� ��ŵ
                Set objQC = Nothing
                Set objOrder = Nothing
                
                Set objQC = New clsQcMst
                Set objOrder = New clsQcOrder
                
                DoEvents
                DBConn.BeginTrans
                If ReadyToCollect(i) And UpdateSchedule(i) Then
                    DBConn.CommitTrans
'                    DBConn.RollbackTrans    'test ��
                Else
                    DBConn.RollbackTrans
                    
                    '��ü��ȣ,������ȣ ����
                    tblResult.Col = 10
                    tblResult.Value = ""
                    tblResult.Col = 15
                    tblResult.Value = ""
                    
                    'ä������,�ð��� ��������..
                    tblResult.Col = 4
                    tblResult.Value = ""
                    tblResult.Col = 5
                    tblResult.Value = ""
                    tblResult.Col = 11
                    tblResult.Value = "ó��"
                    tblResult.ForeColor = vbRed
                    tblResult.Col = 20
                    tblResult.Value = IIf(tblResult.Value = "", "ä��ó���߿� ������ �߻��Ͽ����ϴ�.", tblResult.Value & vbNewLine & "ä��ó���߿� ������ �߻��Ͽ����ϴ�.")
                End If
            End If
            
            objPro.Value = i
        End If
    Next
    
    Set objQC = Nothing
    Set objOrder = Nothing
    Set objPro = Nothing
End Sub

Private Function ReadyToCollect(ByVal pRow As Long) As Boolean

    Dim i As Long
    Dim varCtrlCd As Variant
    Dim varCtrlNm As Variant
    Dim varLevelCd As Variant
    Dim varLotNo As Variant
    Dim varEqpNm As Variant
    Dim varWorkArea As Variant
    Dim varOrdDt As Variant
    
'1   /       2/       3/       4/       5/     6/      7/    8/     9/      10/  11/
'����/ó������/ó��ð�/ä������/ä��ð�/ctrlcd/levelcd/lotno/����/������ȣ/����/
'12      /    13/    14/      15/16/   17/    18/19/    20
'����ڵ�/ctrlnm/sectcd/��ü��ȣ/wa/accdt/accseq/��/errmsg

    Call tblResult.GetText(2, pRow, varOrdDt)
    Call tblResult.GetText(6, pRow, varCtrlCd)
    Call tblResult.GetText(7, pRow, varLevelCd)
    Call tblResult.GetText(8, pRow, varLotNo)
    Call tblResult.GetText(9, pRow, varEqpNm)
    Call tblResult.GetText(13, pRow, varCtrlNm)
    Call tblResult.GetText(16, pRow, varWorkArea)
'    varLotNo = objOrder.GetLastLotNo(varCtrlCd, varLevelCd)
    
    Call objQC.GetQcData(varCtrlCd, varLevelCd, varLotNo)
    Call objQC.GetQCItems(varCtrlCd, varLevelCd, varLotNo)
    
    '�˻��׸��� ���� ��� ������ ����
    If objQC.ItemCount = 0 Then
        ReadyToCollect = False
        Exit Function
    End If
    
    For i = 1 To objQC.ItemCount
        objQC.Item(i).Selected = True
    Next
   
    With objOrder
        Set .MyQc = objQC
    
        .SpcYY = LIS_BarDiv & Mid(Format(GetSystemDate, "YYYY"), 4) '��ü�⵵
        
        .PtId = 0                                     'ȯ��ID
        .PtNm = ""
        .Sex = ""                                     '����
        .AgeDay = 0                                   'ȯ���Ϸ�
        .BedInDt = ""                                 '�Կ���
        .OrdDt = Format(varOrdDt, CS_DateDbFormat)    'ó����
        
        .Controlcd = varCtrlCd                        'Control�ڵ�
        .ControlNm = varCtrlNm                        'Control��
        .EqpNm = varEqpNm                             '����
        .LevelCd = varLevelCd                         'Level�ڵ�
        .Lotno = varLotNo                             'Lot Number
        .WardId = ""                                  '����ID
        .EntDt = Format(GetSystemDate, CS_DateDbFormat)          '�Է���
        .DeptCd = ""
        .BuildCd = ObjSysInfo.BuildingCd
        .SpcCd = varLevelCd
        .MultiFg = ""
        .QcFg = "1"                                   '������������
        
        .EntTm = Format(GetSystemDate, CS_TimeDbFormat)          '�Է½ð�
        .EntId = ObjSysInfo.EmpId                         '�Է���
        .OrgAccNo = ""                                '��������ȣ
        .HosilId = ""                                 '����ID
        .RoomId = ""                                  '����ID
        .BedId = ""                                   'ħ��ID
        .ColDt = Format(GetSystemDate, CS_DateDbFormat)          'ä����
        .ColId = ObjSysInfo.EmpId                         'ä����
        .OrgBuildCd = ObjSysInfo.BuildingCd      '** ä���� ����Ǵ� �ǹ��ڵ�
        .WorkArea = varWorkArea
        
        .Trans = True   'Ʈ������� �ܺο��� ���հ����ϱ� ���ؼ�...
                
        If .DoCollection Then
            If chkBarPrint.Value = 1 Then
                'ä���ϸ鼭 �ٷ� ���ڵ� ����
                If .PrintBarcodeLabel(1) = False Then     '���ڵ� ��µ��� ������ ��� ä���� ���������� �ϰ� �޽����� ����ش�.
                    tblResult.Row = pRow
                    tblResult.Col = 11
                    tblResult.ForeColor = vbRed
                    tblResult.Col = 20
                    tblResult.Value = "���ڵ� ��µ��� ������ �߻��Ͽ����ϴ�."
                End If
            End If
            
            tblResult.Row = pRow
            tblResult.Col = 10
            tblResult.Value = .WorkArea & "-" & Mid(.AccDt, 3) & "-" & .AccSeq
            tblResult.Col = 15
            tblResult.Value = .SpcYY & "-" & .SpcNo
            
            'ä������,�ð��� ��������..
            tblResult.Col = 4
            tblResult.Value = Format(.ColDt, CS_DateMask)
            tblResult.Col = 5
            tblResult.Value = Format(.ColTm, CS_TimeLongMask)
            tblResult.Col = 11
            tblResult.Value = "����"
            
            ReadyToCollect = True
        Else
            ReadyToCollect = False
        End If
    End With
End Function

Private Function UpdateSchedule(ByVal pRow As Long) As Boolean

    Dim strSQL As String
    Dim strSpcYY As String
    Dim strSpcNum As String
    Dim strSpcNo As Variant
    
    Call tblResult.GetText(15, pRow, strSpcNo)
    
    strSpcYY = medGetP(strSpcNo, 1, "-")
    strSpcNum = medGetP(strSpcNo, 2, "-")
    
    strSQL = "update " & T_LAB025 & " set donefg = '1', prtfg = '1', " & _
             DBW("spcyy=", medGetP(strSpcNo, 1, "-"), 1) & _
             DBW("spcno=", medGetP(strSpcNo, 2, "-"))
             
    With tblResult
        .Row = pRow
        .Col = 2: strSQL = strSQL & " where " & DBW("dodt=", Format(.Value, CS_DateDbFormat))
        .Col = 3: strSQL = strSQL & " and " & DBW("dotm=", Format(.Value, CS_TimeDbFormat))
        .Col = 14: strSQL = strSQL & " and " & DBW("sectcd=", .Value)
        .Col = 6:: strSQL = strSQL & " and " & DBW("ctrlcd=", .Value)
        .Col = 7:: strSQL = strSQL & " and " & DBW("levelcd=", .Value)
    End With
    
    UpdateSchedule = True
    
    On Error GoTo ErrTrap
    DBConn.Execute strSQL
    
    Exit Function
    
ErrTrap:
    UpdateSchedule = False
End Function

Private Sub DoRePrint(ByVal pProcessCnt As Long)
    Dim lngCnt As Long
    Dim lngECnt As Long
    Dim lngSCnt As Long
    Dim i As Long
    Dim objPro As clsProgress
    
    '���α׷��� ��
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Width = tblResult.Width
        .Height = 470
        .Left = tblResult.Left
        .Top = tblResult.Top
        .Message = "���ڵ� ���̺��� ��������� �а� �ֽ��ϴ�."
        .Max = pProcessCnt
    End With
    
    lngCnt = 0
    For i = 1 To tblResult.DataRowCnt
        tblResult.Row = i
        tblResult.Col = 1
        If tblResult.Value <> 1 Then
            Set objOrder = Nothing
            Set objOrder = New clsQcOrder
            
            If SetBarcodeInfo(1, i) Then   '���ڵ� ��������� ��´�
                objOrder.ColCount = 1  '����� �Ǽ�
                If objOrder.PrintBarcodeLabel(1) = False Then
                    tblResult.Row = i
                    tblResult.Col = 11
                    tblResult.ForeColor = vbRed
                    tblResult.Col = 20
                    tblResult.Value = "���ڵ� ��µ��� ������ �߻��Ͽ����ϴ�."
                End If
            Else
                tblResult.Row = i
                tblResult.Col = 11
                tblResult.ForeColor = vbRed
                tblResult.Col = 20
                tblResult.Value = "���ڵ� ��µ��� ������ �߻��Ͽ����ϴ�."
            End If
            
            objPro.Message = "���ڵ� ���̺��� ����ϰ� �ֽ��ϴ�."
            objPro.Value = i
        End If
    Next
    Set objPro = Nothing
    
    Set objOrder = Nothing
End Sub

Private Function SetBarcodeInfo(ByVal pCnt As Long, ByVal pRow As Long) As Boolean
'����� �˻��׸��� ���� ��쿡�� ���ڵ� ������ �߻����Ѿ� �Ѵ�.
'pCnt : ���ڵ� ��� �Ǽ�
'pRow : ����� ������ ��� �ִ� ���������� Row

    Dim lngCnt As Integer
'1   /       2/       3/       4/       5/     6/      7/    8/     9/      10/  11/
'����/ó������/ó��ð�/ä������/ä��ð�/ctrlcd/levelcd/lotno/����/������ȣ/����/
'12      /    13/    14/      15/16/   17/    18/19/    20
'����ڵ�/ctrlnm/sectcd/��ü��ȣ/wa/accdt/accseq/��/errmsg
   
    With tblResult
        .Row = pRow
        objOrder.BarCount = 1
        objOrder.BuildNm = ObjSysInfo.BuildingNm
        .Col = 7: objOrder.SpcNm = .Value
        .Col = 16: objOrder.WorkArea = .Value
        .Col = 2: objOrder.OrdDt = Replace(.Value, "-", "")
        .Col = 17: objOrder.AccDt = .Value
        .Col = 18: objOrder.AccSeq = .Value
        .Col = 15: objOrder.SpcYY = medGetP(.Value, 1, "-")
        .Col = 15: objOrder.SpcNo = medGetP(.Value, 2, "-")
        .Col = 9: objOrder.EqpNm = .Value
        .Col = 13: objOrder.ControlNm = .Value
        .Col = 6: objOrder.Controlcd = .Value
        .Col = 7: objOrder.LevelCd = .Value
        objOrder.TestNames = Replace(objOrder.GetTestNames(objOrder.WorkArea, objOrder.AccDt, objOrder.AccSeq, lngCnt), vbTab, ",")
        
        If lngCnt = 0 Then GoTo Nodata  '����� �˻��׸��� ���� ��� ����ó��
        
        '���ڵ� ��������� ����ִ� �޼ҵ�
        Call objOrder.PrintBarcode(pCnt, Format(GetSystemDate, "YYYY-MM-DD HH:MM"))
    End With
    
    SetBarcodeInfo = True
    
    Exit Function
    
Nodata:
    SetBarcodeInfo = False
End Function

Private Sub cmdConfig_Click()
'����ں� ȭ�� ���������� ����
    Dim strMsg As VbMsgBoxResult
    Dim strMsg2 As VbMsgBoxResult
    Dim User As String
    
    dtpFromDate.Font.Italic = True
    dtpToDate.Font.Italic = True
    cboSection.FontItalic = True
    cboLevel.FontItalic = True
    txtCtrlCd.FontItalic = True
    chkAll.FontItalic = True
    optWorkTime(0).FontItalic = True
    optWorkTime(1).FontItalic = True
    optWorkTime(2).FontItalic = True
    optWorkTime(3).FontItalic = True
    optWorkTime(4).FontItalic = True
    dtpFromTime.Font.Italic = True
    dtpToTime.Font.Italic = True
    optQryKey(0).FontItalic = True
    optQryKey(1).FontItalic = True
    If optQryKey(1).Value Then
        chkQryKey(0).FontItalic = True
        chkQryKey(1).FontItalic = True
        chkQryKey(2).FontItalic = True
    End If
    chkSchedule.FontItalic = True
    chkBarPrint.FontItalic = True
    
    strMsg = MsgBox("������ �� �ִ� ������ ���Ÿ����� ǥ�õ� �͵��Դϴ�.." & vbNewLine & _
                    "�ٽ� ȭ���� ������ ������ ������ ����˴ϴ�." & vbNewLine & _
                    "(��, ��ȸ���ڴ� ���� ��¥���� ���ݸ��� �����մϴ�.)" & vbNewLine & vbNewLine & _
                    "���� ������ �����Ͻðڽ��ϱ�?" & vbNewLine & _
                    "(��:���� ����,�ƴϿ�:���� ����)", vbYesNoCancel + vbExclamation)
    
    If strMsg = vbCancel Then GoTo NoAction
    
    '��������
    
    User = GetSetting("Schweitzer2000 LIS\Config", "frm310QCReprint_N", "User", "")
        
    If strMsg = vbYes Then  '��������
        strMsg2 = MsgBox("Control�� �����ϸ� �����͸� ��ȸ�ϴ� ����� �ֽ��ϴ�." & vbNewLine & _
                         "�� ����� �̹� ������ ��ȸ������ �⺻������ ��ȸ�� �մϴ�." & vbNewLine & vbNewLine & _
                         "�� ����� ����Ͻðڽ��ϱ�?", vbExclamation + vbYesNo)
        
        If strMsg2 = vbYes Then
            Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "Refresh", "1")
        Else
            Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "Refresh", "0")
        End If
        
        If InStr(User, ObjSysInfo.logonid) = 0 Then '��������
            Call SaveSetting("Schweitzer2000 LIS\Config", "frm310QCReprint_N", "User", User & ObjSysInfo.logonid & ",")
        End If
        
        Call SaveSetting("Schweitzer2000 LIS\Config", "frm310QCReprint_N", "Desc", "QC ������ ó�� �� ������ȸ")
        
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "DuraDate", DateDiff("d", dtpFromDate.Value, dtpToDate.Value))
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "Section", cboSection.Text)
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "Level", cboLevel.Text)
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "CtrlCd", Trim(txtCtrlCd.Text))
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "CtrlNm", Trim(lblCtrlNm.Caption))
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "AllCtrl", IIf(chkAll.Value = 1, 1, 0))
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "WorkTime", IIf(optWorkTime(0).Value, 0, IIf(optWorkTime(1).Value, 1, IIf(optWorkTime(2).Value, 2, IIf(optWorkTime(3).Value, 3, 4)))))
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "DuraTime", dtpFromTime.Value & "," & dtpToTime.Value)
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "optQryKey", IIf(optQryKey(0).Value, 0, 1))
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "chkQryKey", chkQryKey(0).Value & "," & chkQryKey(1).Value & "," & chkQryKey(2).Value)
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "Schedule", IIf(chkSchedule.Value = 1, 1, 0))
        Call SaveSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "BarPrint", IIf(chkBarPrint.Value = 1, 1, 0))
        
        MsgBox "������ ������ �����Ǿ����ϴ�.", vbExclamation
    ElseIf strMsg = vbNo Then   '��������
        User = Replace(User, ObjSysInfo.logonid, "")
        User = Replace(User, ",,", ",")
        If Len(User) = 1 Then User = ""
        Call SaveSetting("Schweitzer2000 LIS\Config", "frm310QCReprint_N", "User", User)
        
        On Error Resume Next
        Call DeleteSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid)
        
        MsgBox "������ �ʱ�ȭ�Ǿ����ϴ�.", vbExclamation
    End If

NoAction:
    dtpFromDate.Font.Italic = False
    dtpToDate.Font.Italic = False
    cboSection.FontItalic = False
    cboLevel.FontItalic = False
    txtCtrlCd.FontItalic = False
    chkAll.FontItalic = False
    optWorkTime(0).FontItalic = False
    optWorkTime(1).FontItalic = False
    optWorkTime(2).FontItalic = False
    optWorkTime(3).FontItalic = False
    optWorkTime(4).FontItalic = False
    dtpFromTime.Font.Italic = False
    dtpToTime.Font.Italic = False
    optQryKey(0).FontItalic = False
    optQryKey(1).FontItalic = False
    If optQryKey(1).Value Then
        chkQryKey(0).FontItalic = False
        chkQryKey(1).FontItalic = False
        chkQryKey(2).FontItalic = False
    End If
    chkSchedule.FontItalic = False
    chkBarPrint.FontItalic = False
End Sub

Private Sub cmdControl_Click()
    Call LoadControlInfo
End Sub

Private Sub LoadControlInfo(Optional ByVal pCtrlCd As String = "")
    Dim objPop As clsPopUpList
    
    Set objPop = New clsPopUpList
    
    With objPop
        .Recordset = GetControlInfo(pCtrlCd)
        .ColumnHeaderText = "�ڵ�;��Ʈ�Ѹ�;Level"
        .ColumnHeaderWidth = "794.8347;2055.118;840.189"
        .ColumnHeaderAlign = "0;0;2"
        .FormWidth = 4140
        Call .LoadPopUp
        
        txtCtrlCd.Text = medGetP(.SelectedString, 1, .Delimiter)
        lblCtrlNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
        
        If txtCtrlCd.Text <> "" Then
            If cboLevel.ListIndex > 0 Then
                lblCtrlNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
            Else
                lblCtrlNm.Caption = GetControlName
                lblCtrlNm.ToolTipText = lblCtrlNm.Caption
            End If
        End If
    End With
    
    Set objPop = Nothing
End Sub

Private Function GetControlInfo(Optional ByVal pCtrlCd As String = "") As Recordset
    Dim strSQL As String
    
    strSQL = " select a.ctrlcd,a.ctrlnm,a.levelcd from " & T_LAB021 & " a " & _
            " where exists (select * from " & T_LAB023 & _
            "               where " & DBW("opendt<=", Format(dtpFromDate.Value, "yyyyMMdd")) & _
            "               and " & DBW("expdt>=", Format(dtpToDate.Value, "yyyyMMdd")) & _
            "               and a.ctrlcd=ctrlcd " & _
            "               and a.levelcd=levelcd) "
                
    If pCtrlCd <> "" Then
        strSQL = strSQL & " and " & DBW("a.ctrlcd=", pCtrlCd)
    End If
    
    If cboLevel.ListIndex > 0 Then
        strSQL = strSQL & " and " & DBW("a.levelcd=", Trim(Mid(cboLevel.Text, 20)))
    End If
    
    If cboSection.ListIndex > 0 Then
        strSQL = strSQL & " and " & DBW("a.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
    End If
            
    strSQL = strSQL & " order by a.ctrlcd,a.ctrlnm "
    
    
    Set GetControlInfo = New Recordset
    GetControlInfo.Open strSQL, DBConn
End Function

Private Function GetControlName() As String
    Dim Rs As Recordset
    Dim strTmp As String
                   
    Set Rs = GetControlInfo(Trim(txtCtrlCd.Text))
               
    Do Until Rs.EOF
        strTmp = strTmp & Rs.Fields("ctrlnm").Value & "" & ","
        
        Rs.MoveNext
    Loop
    
    GetControlName = Mid(strTmp, 1, Len(strTmp) - 1)
    
    Set Rs = Nothing
End Function

Private Sub cmdExit_Click()
    Unload Me
'    Unload frm311QCResultEntry
    If IsLastForm Then RaiseEvent LastFormUnload
'    If IsLastForm Then Call UnloadForm(Me)
End Sub

Private Sub cmdQuery_Click()
    Dim strTmp As String
    Dim i As Long
    
    If CheckValidation = False Then Exit Sub
    
    lblWarn.Caption = ""
    If optQryKey(0).Value Then
        strTmp = "ä�� ���"
    Else
        strTmp = "����� ���("
        For i = chkQryKey.LBound To chkQryKey.UBound
            If chkQryKey(i).Value = 1 Then
                strTmp = strTmp & Mid(chkQryKey(i).Caption, 1, 2) & "/"
            End If
        Next
        
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        strTmp = strTmp & ")"
    End If
    lblDesc.Caption = "�� ������ �˻����� " & _
                      Format(dtpFromDate.Value, "yyyy/MM/dd") & "~" & Format(dtpToDate.Value, "yyyy/MM/dd") & "," & _
                      IIf(cboSection.ListIndex = 0, "��ü", Trim(medGetP(cboSection.Text, 2, COL_DIV))) & "," & _
                      Trim(Mid(cboLevel.Text, 1, 5)) & "," & _
                      IIf(txtCtrlCd.Text = "", "��ü", txtCtrlCd.Text) & "," & _
                      Format(dtpFromTime.Value, "HH:mm:ss") & "~" & Format(dtpToTime.Value, "HH:mm:ss") & "," & _
                      strTmp & "," & _
                      IIf(chkSchedule.Value = 1, "�����ٳ�����", "��ü")
    
    Call LoadList
End Sub

Private Sub LoadList()
'optOption �� �Ǵ�
    Dim objPro As clsProgress
    Dim Rs As Recordset
    Dim strKey As String
    Dim strOrdDt As String
    Dim strOrdTm As String
    Dim strCtrlcd As String
    Dim strLevelcd As String
    Dim strLotNo As String

    If optQryKey(0).Value Then    'ä�����
        Set Rs = LoadForCollection
    ElseIf optQryKey(1).Value Then  '�������
        Set Rs = LoadForReprint
    End If
    
    cmdQuery.Enabled = False
    
    tblResult.MaxRows = 27
    Call medClearTable(tblResult)
    With tblResult
        .Row = -1
        .Col = 8: .Col2 = 8
        .BlockMode = True
        .CellType = CellTypeStaticText
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
    End With
    
    If Rs.EOF Then GoTo Nodata
    
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Width = tblResult.Width
        .Height = 470
        .Left = tblResult.Left
        .Top = tblResult.Top
        .Message = "�˻����ǿ� �ش��ϴ� �ڷḦ �а� �ֽ��ϴ�..."
        .Max = Rs.RecordCount
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = tblResult.Width
'        .YHeight = 470
'        .XPos = tblResult.Left
'        .YPos = tblResult.Top
'        .ForeColor = &H864B24
'        .Msg = "�˻����ǿ� �ش��ϴ� �ڷḦ �а� �ֽ��ϴ�..."
'        .Value = 1
'        .Max = Rs.RecordCount
    End With
    
    tblResult.ReDraw = False
    Do Until Rs.EOF
            With tblResult
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
'1   /       2/       3/       4/       5/     6/      7/    8/     9/      10/  11/
'����/ó������/ó��ð�/ä������/ä��ð�/ctrlcd/levelcd/lotno/����/������ȣ/����/
'12      /    13/    14/      15/16/   17/    18/19/    20
'����ڵ�/ctrlnm/sectcd/��ü��ȣ/wa/accdt/accseq/��/errmsg
                If strKey <> Format(Rs.Fields("orddt").Value, CS_DateMask) & Format(Rs.Fields("ordtm").Value, CS_TimeLongMask) & Rs.Fields("ctrlcd").Value & "" & Rs.Fields("levelcd").Value & "" Then
                    .Row = .DataRowCnt + 1
                    .Col = 2: .Value = Format(Rs.Fields("orddt").Value, CS_DateMask): strOrdDt = .Value
                    .Col = 3: .Value = Format(Rs.Fields("ordtm").Value, CS_TimeLongMask): strOrdTm = .Value
                    .Col = 4: .Value = Format(Rs.Fields("rcvdt").Value, CS_DateMask)
                    .Col = 5: .Value = Format(Rs.Fields("rcvtm").Value, CS_TimeLongMask)
                    .Col = 6: .Value = Rs.Fields("ctrlcd").Value & "": strCtrlcd = .Value
                    .Col = 7: .Value = Rs.Fields("levelcd").Value & "": strLevelcd = .Value
                    .Col = 8
                            .CellType = CellTypeStaticText
                            .TypeVAlign = TypeVAlignCenter
                            .TypeHAlign = TypeHAlignCenter
                            .Value = Rs.Fields("lotno").Value & "": strLotNo = .Value
                    .Col = 9: .Value = Rs.Fields("eqpnm").Value & ""
                    .Col = 10: .Value = IIf(Rs.Fields("accdt").Value & "" = "", "", Rs.Fields("workarea").Value & "" & "-" & _
                                       Mid(Rs.Fields("accdt").Value & "", 3) & "-" & _
                                       Rs.Fields("accseq").Value & "")
                    .Col = 11: .Value = GetStatus(Rs.Fields("workarea").Value & "", Rs.Fields("accdt").Value & "", Rs.Fields("accseq").Value & "")
                    .Col = 12: .Value = Rs.Fields("eqpcd").Value & ""
                    .Col = 13: .Value = Rs.Fields("ctrlnm").Value & ""
                    .Col = 14: .Value = Rs.Fields("sectcd").Value & ""
                    .Col = 15: .Value = Rs.Fields("spcyy").Value & "" & "-" & Rs.Fields("spcno").Value & ""
                    .Col = 16: .Value = Rs.Fields("workarea").Value & ""
                    .Col = 17: .Value = Rs.Fields("accdt").Value & ""
                    .Col = 18: .Value = Rs.Fields("accseq").Value & ""
                    
                    If optQryKey(1).Value Then
                        .Col = 11
                        
                        '����,�˻���,�߰��γѵ��� ������������ ȭ�� ����ش�.
                        If .Value = "����" Or .Value = "�κ�" Then
                            .Col = 19: .Value = "��"
                            .ForeColor = DCM_LightBlue
                        End If
                    End If
                Else
                    .Col = 8
                    strLotNo = strLotNo & vbTab & Rs.Fields("lotno").Value & ""
                    
                    .CellType = CellTypeComboBox
                    .TypeVAlign = TypeVAlignCenter
                    .TypeHAlign = TypeHAlignCenter
                    .Action = ActionComboClear
                    .TypeComboBoxList = strLotNo
                    .TypeComboBoxIndex = 0
                End If
                
                If optQryKey(0).Value Then
                    strKey = strOrdDt & strOrdTm & strCtrlcd & strLevelcd
                ElseIf optQryKey(1).Value Then
                    strKey = ""
                End If
                
                objPro.Value = objPro.Value + 1
            End With
        Rs.MoveNext
    Loop
    tblResult.ReDraw = True
    tblResult.TopRow = 1
    
Nodata:
    If tblResult.DataRowCnt = 0 Then
        MsgBox "���ǿ� �´� �ڷᰡ �����ϴ�.", vbExclamation
        lblWarn.Caption = ""
    End If
    
    cmdQuery.Enabled = True
    
    Set Rs = Nothing
    Set objPro = Nothing
End Sub

Private Function CheckValidation() As Boolean
'
    CheckValidation = True
    
    If dtpFromDate.Value > dtpToDate.Value Then
        MsgBox "�Ⱓ������ �߸��Ǿ����ϴ�.", vbExclamation
        CheckValidation = False
        Exit Function
    End If
    
    If dtpFromTime.Value > dtpToTime.Value Then
        MsgBox "�ð������� �߸��Ǿ����ϴ�.", vbExclamation
        CheckValidation = False
        Exit Function
    End If
    
    If chkAll.Value = 0 Then
        If Trim(txtCtrlCd.Text) = "" Then
            MsgBox "��Ʈ���� �����Ͻʽÿ�.", vbExclamation
            CheckValidation = False
            Exit Function
        End If
    End If
    
    If optQryKey(0).Value = False And optQryKey(1).Value = False Then
        MsgBox "��ȸ������ �����Ͻʽÿ�.", vbExclamation
        CheckValidation = False
        Exit Function
    End If
    
    If optQryKey(1).Value Then
        If chkQryKey(0).Value = 0 And chkQryKey(1).Value = 0 And chkQryKey(2).Value = 0 Then
            MsgBox "����� ����� �����Ͻʽÿ�.", vbExclamation
            CheckValidation = False
            Exit Function
        End If
    End If
End Function

Private Function LoadForCollection() As Recordset
'ä�����
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.dodt orddt,a.dotm ordtm,'' rcvdt,'' rcvtm,a.ctrlcd,z.ctrlnm,a.levelcd,c.lotno,z.eqpcd,d.eqpnm," & _
            " z.workarea, '' accdt, 0 accseq,'1' stscd,z.sectcd,'' spcyy, 0 spcno,'Y' falg " & _
            " from " & T_LAB025 & " a, " & T_LAB021 & " z, " & T_LAB023 & " c, " & T_LAB006 & " d " & _
            " where " & DBW("a.dodt>=", Format(dtpFromDate.Value, CS_DateDbFormat)) & _
            " and " & DBW("a.dodt<=", Format(dtpToDate.Value, CS_DateDbFormat)) & _
            " and " & DBW("a.dotm>=", Format(dtpFromTime.Value, CS_TimeDbFormat)) & _
            " and " & DBW("a.dotm<=", Format(dtpToTime.Value, CS_TimeDbFormat)) & _
            " and (a.donefg='' or a.donefg is null) " & _
            " and a.ctrlcd=z.ctrlcd " & _
            " and a.levelcd=z.levelcd " & _
            " and " & DBJ("z.eqpcd*=d.eqpcd") & _
            " and a.ctrlcd=c.ctrlcd " & _
            " and a.levelcd=c.levelcd " & _
            " and " & DBW("c.opendt<=", Format(GetSystemDate, CS_DateDbFormat)) & _
            " and " & DBW("c.expdt>=", Format(GetSystemDate, CS_DateDbFormat))

'    strSQL = " select a.dodt as orddt, a.dotm as ordtm, '' coldt, '' coltm,a.sectcd, a.ctrlcd, a.levelcd,d.lotno, " & _
'             "        a.donefg, '' accdt, 0 accseq, " & _
'             "        b.ctrlnm, b.workarea, b.eqpcd, c.eqpnm " & _
'             " from " & T_LAB023 & " d," & T_LAB006 & " c," & T_LAB021 & " b," & T_LAB025 & " a " & _
'             " where " & DBW("a.dodt>=", Format(dtpFromDate.Value, CS_DateDbFormat)) & _
'             " and " & DBW("a.dodt<=", Format(dtpToDate.Value, CS_DateDbFormat)) & _
'             " and " & DBW("a.dotm>=", Format(dtpFromTime.Value, CS_TimeDbFormat)) & _
'             " and " & DBW("a.dotm<=", Format(dtpToTime.Value, CS_TimeDbFormat)) & _
'             " and (a.donefg = '' or a.donefg is null)" & _
'             " and b.ctrlcd = a.ctrlcd   and   b.levelcd = a.levelcd " & _
'             " and " & DBW("b.buildcd=", ObjSysInfo.BuildingCd) & _
'             " and " & DBJ("b.eqpcd *= c.eqpcd") & _
'             " and b.ctrlcd=d.ctrlcd " & _
'             " and b.levelcd=d.levelcd " & _
'             " and " & DBW("d.opendt<=", Format(GetSystemDate, CS_DateDbFormat)) & _
'             " and " & DBW("d.expdt>=", Format(GetSystemDate, CS_DateDbFormat))

    If chkAll.Value = 0 Then
        strSQL = strSQL & " and " & DBW("z.ctrlcd=", Trim(txtCtrlCd.Text))
    End If
    
    If cboSection.ListIndex > 0 Then
        strSQL = strSQL & " and " & DBW("z.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
    End If
    
    If cboLevel.ListIndex > 0 Then
        strSQL = strSQL & " and " & DBW("z.levelcd=", Trim(Mid(cboLevel.Text, 31)))
    End If
    
    strSQL = strSQL & " order by orddt,ordtm,ctrlcd,levelcd asc, opendt desc "
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    Set LoadForCollection = Rs
    
End Function

Private Function LoadForReprint() As Recordset
'�������

    Dim Rs As Recordset
    Dim strSQL As String
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim strTmp As String
    
    '�������� �Ǵ�...
    If chkQryKey(0).Value = 1 And chkQryKey(1).Value = 1 And chkQryKey(2).Value = 1 Then
        strTmp = ""
    ElseIf chkQryKey(0).Value = 1 And chkQryKey(1).Value = 0 And chkQryKey(2).Value = 0 Then
        strTmp = " and y.stscd='2' "
    ElseIf chkQryKey(0).Value = 0 And chkQryKey(1).Value = 1 And chkQryKey(2).Value = 0 Then
        strTmp = " and y.stscd in ('3','4') "
    ElseIf chkQryKey(0).Value = 0 And chkQryKey(1).Value = 0 And chkQryKey(2).Value = 1 Then
        strTmp = " and y.stscd >='5' "
    ElseIf chkQryKey(0).Value = 0 And chkQryKey(1).Value = 1 And chkQryKey(2).Value = 1 Then
        strTmp = " and y.stscd >='3' "
    ElseIf chkQryKey(0).Value = 1 And chkQryKey(1).Value = 0 And chkQryKey(2).Value = 1 Then
        strTmp = " and (y.stscd ='2' or y.stscd >='5') "
    ElseIf chkQryKey(0).Value = 1 And chkQryKey(1).Value = 1 And chkQryKey(2).Value = 0 Then
        strTmp = " and y.stscd in ('2','3','4') "
    End If
    
    strSQL1 = " select distinct a.dodt orddt,a.dotm ordtm,y.rcvdt,y.rcvtm,a.ctrlcd,z.ctrlnm,a.levelcd,b.lotno,z.eqpcd,e.eqpnm," & _
            " b.workarea,b.accdt,b.accseq,y.stscd,z.sectcd,y.spcyy,y.spcno,'Y' flag " & _
            " from " & T_LAB025 & " a, " & T_LAB026 & " b, " & T_LAB201 & " y, " & T_LAB021 & " z, " & T_LAB006 & " e " & _
            " where " & DBW("a.dodt>=", Format(dtpFromDate.Value, CS_DateDbFormat)) & _
            " and " & DBW("a.dodt<=", Format(dtpToDate.Value, CS_DateDbFormat)) & _
            " and " & DBW("a.dotm>=", Format(dtpFromTime.Value, CS_TimeDbFormat)) & _
            " and " & DBW("a.dotm<=", Format(dtpToTime.Value, CS_TimeDbFormat)) & _
            " and (a.donefg <>'' or a.donefg is not null) " & _
            " and a.ctrlcd=b.ctrlcd " & _
            " and a.levelcd=b.levelcd " & _
            " and a.spcyy=y.spcyy " & _
            " and a.spcno=y.spcno " & _
            " and b.workarea=y.workarea " & _
            " and b.accdt=y.accdt " & _
            " and b.accseq=y.accseq " & _
            " and a.ctrlcd=z.ctrlcd " & _
            " and a.levelcd=z.levelcd " & _
            " and z.eqpcd=e.eqpcd "
    strSQL1 = strSQL1 & strTmp
                
    strSQL2 = " select distinct y.rcvdt orddt,y.rcvtm ordtm,y.rcvdt,y.rcvtm,a.ctrlcd,z.ctrlnm,a.levelcd,a.lotno,z.eqpcd,d.eqpnm," & _
            " a.workarea,a.accdt,a.accseq,y.stscd,z.sectcd,y.spcyy,y.spcno,'' flag " & _
            " from " & T_LAB026 & " a, " & T_LAB201 & " y, " & T_LAB021 & " z, " & T_LAB006 & " d " & _
            " where " & DBW("y.rcvdt>=", Format(dtpFromDate.Value, CS_DateDbFormat)) & _
            " and " & DBW("y.rcvdt<=", Format(dtpToDate.Value, CS_DateDbFormat)) & _
            " and " & DBW("y.rcvtm>=", Format(dtpFromTime.Value, CS_TimeDbFormat)) & _
            " and " & DBW("y.rcvtm<=", Format(dtpToTime.Value, CS_TimeDbFormat)) & _
            " and a.workarea=y.workarea " & _
            " and a.accdt=y.accdt " & _
            " and a.accseq=y.accseq " & _
            " and a.ctrlcd=z.ctrlcd " & _
            " and a.levelcd=z.levelcd " & _
            " and z.eqpcd=d.eqpcd " & _
            " and not exists ( select * from " & T_LAB025 & _
            "                 where " & DBW("dodt>=", Format(dtpFromDate.Value, CS_DateDbFormat)) & _
            "                 and " & DBW("dodt<=", Format(dtpToDate.Value, CS_DateDbFormat)) & _
            "                 and " & DBW("dotm>=", Format(dtpFromTime.Value, CS_TimeDbFormat)) & _
            "                 and " & DBW("dotm<=", Format(dtpToTime.Value, CS_TimeDbFormat)) & _
            "                 and (donefg <>'' or donefg is not null) " & _
            "                 and ctrlcd=a.ctrlcd " & _
            "                 and levelcd=a.levelcd " & _
            "                 and spcyy=y.spcyy " & _
            "                 and spcno=y.spcno) "
    strSQL2 = strSQL2 & strTmp
'order by a.dodt,a.dotm,a.ctrlcd,a.levelcd
    
'    strSQL1 = " select a.dodt as orddt, a.dotm as ordtm,b.coldt,b.coltm,b.stscd, a.sectcd, a.ctrlcd, a.levelcd,'' as lotno, a.prtfg, " & _
'             "        a.spcyy, a.spcno, a.donefg, b.workarea, b.accdt, b.accseq, " & _
'             "        c.ctrlnm, c.eqpcd, d.eqpnm " & _
'             " from " & T_LAB201 & " b, " & _
'                        T_LAB021 & " c, " & T_LAB006 & " d, " & T_LAB025 & " a " & _
'             " where " & DBW("a.dodt>=", Format(dtpFromDate.Value, CS_DateDbFormat)) & _
'             " and " & DBW("a.dodt<=", Format(dtpToDate.Value, CS_DateDbFormat)) & _
'             " and " & DBW("a.dotm>=", Format(dtpFromTime.Value, CS_TimeDbFormat)) & _
'             " and " & DBW("a.dotm<=", Format(dtpToTime.Value, CS_TimeDbFormat)) & _
'             " and a.donefg = '1' " & _
'             " and b.spcyy = a.spcyy     and   b.spcno = a.spcno " & _
'             " and c.ctrlcd = a.ctrlcd   and   c.levelcd = a.levelcd " & _
'             " and " & DBW("c.buildcd=", ObjSysInfo.BuildingCd) & _
'             " and " & DBJ("d.eqpcd =* c.eqpcd ") & strTmp & _
'             " and b.reqtotcnt<>0 "
'
'    strSQL2 = " select distinct b.coldt as orddt,b.coltm as ordtm, b.coldt,b.coltm,b.stscd,a.sectcd,a.ctrlcd,a.levelcd,c.lotno,'1' ptrfg, " & _
'              " b.spcyy," & FUNC_CONVERT("char", "b.spcno") & " as spcno,'1' donefg,b.workarea,b.accdt,b.accseq, " & _
'              " a.ctrlnm , a.eqpcd, d.eqpnm " & _
'              " from " & T_LAB021 & " a, " & T_LAB026 & " c, " & T_LAB201 & " b, " & T_LAB006 & " d " & _
'              " where a.ctrlcd = c.ctrlcd " & _
'              " and a.levelcd=c.levelcd " & _
'              " and " & DBW("a.buildcd=", ObjSysInfo.BuildingCd) & _
'              " and d.eqpcd(+) = a.eqpcd " & strTmp & _
'              " and b.workarea=c.workarea " & _
'              " and b.accdt=c.accdt " & _
'              " and b.accseq=c.accseq " & _
'              " and " & DBW("b.coldt>=", Format(dtpFromDate.Value, CS_DateDbFormat)) & _
'              " and " & DBW("b.coldt<=", Format(dtpToDate.Value, CS_DateDbFormat)) & _
'              " and " & DBW("b.coltm>=", Format(dtpFromTime.Value, CS_TimeDbFormat)) & _
'              " and " & DBW("b.coltm<=", Format(dtpToTime.Value, CS_TimeDbFormat)) & _
'              " and not exists (select * from " & T_LAB025 & " where b.spcyy=spcyy and b.spcno=spcno) "

    If chkAll.Value = 0 Then
        strSQL1 = strSQL1 & " and " & DBW("z.ctrlcd=", Trim(txtCtrlCd.Text))
        strSQL2 = strSQL2 & " and " & DBW("z.ctrlcd=", Trim(txtCtrlCd.Text))
    End If
    
    If cboSection.ListIndex > 0 Then
        strSQL1 = strSQL1 & " and " & DBW("z.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
        strSQL2 = strSQL2 & " and " & DBW("z.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
    End If
             
    If cboLevel.ListIndex > 0 Then
        strSQL1 = strSQL1 & " and " & DBW("z.levelcd=", Trim(Mid(cboLevel.Text, 31)))
        strSQL2 = strSQL2 & " and " & DBW("z.levelcd=", Trim(Mid(cboLevel.Text, 31)))
    End If
                     
    If chkSchedule.Value = 1 Then
        strSQL = strSQL1 & " order by a.dodt,a.dotm,a.ctrlcd,a.levelcd "
    ElseIf chkSchedule.Value = 0 Then
        strSQL = strSQL1 & " union " & strSQL2 & " order by orddt,ordtm,ctrlcd,levelcd "
    End If
                 
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    Set LoadForReprint = Rs
End Function

Private Function GetStatus(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String) As String
'����
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select stscd,reqtotcnt,reqinputcnt from " & T_LAB201 & _
             " where " & DBW("workarea=", pWorkArea) & _
             " and " & DBW("accdt=", pAccDt) & _
             " and " & DBW("accseq=", pAccSeq) & _
             " and (qcfg='1' or qcfg='2') "
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    With tblResult
        .Row = .DataRowCnt
        .Col = 11
        If Rs.EOF Then
            GetStatus = "ó��"
            .ForeColor = vbBlack
        ElseIf Rs.Fields("stscd").Value & "" = "2" Then
            If Rs.Fields("reqinputcnt").Value & "" = "0" Then
                GetStatus = "����"
                .ForeColor = vbBlack
            ElseIf Rs.Fields("reqinputcnt").Value & "" <> "0" Then
                GetStatus = "�κ�"
                .ForeColor = vbBlue
            End If
        ElseIf Rs.Fields("stscd").Value & "" = "3" Or Rs.Fields("stscd").Value & "" = "4" Then
            GetStatus = "�κ�"
            .ForeColor = vbBlue
        ElseIf Rs.Fields("stscd").Value & "" = "5" Then
            GetStatus = "����"
            .ForeColor = vbGreen
        ElseIf Rs.Fields("stscd").Value & "" = "6" Then
            GetStatus = "����"
            .ForeColor = vbRed
        End If
    End With
    
    Set Rs = Nothing
End Function

Private Sub GetWarningMsg()
'��ȸ�Ⱓ,section,level,��ȸ�ð�,��Ʈ��,��󿩺�,

    Dim strFromDate As String
    Dim strToDate As String
    Dim strSection As String
    Dim strLevel As String
    Dim strCtrlcd As String
    Dim strFromTime As String
    Dim strToTime As String
    Dim strQryKey As String
    Dim strSchedule As String
    Dim strDesc As String
    Dim i As Long
        
    strFromDate = Format(dtpFromDate.Value, "yyyy/MM/dd")
    strToDate = Format(dtpToDate.Value, "yyyy/MM/dd")
    strSection = IIf(cboSection.ListIndex = 0, "��ü", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
    strLevel = Trim(Mid(cboLevel.Text, 1, 5))
    strCtrlcd = IIf(txtCtrlCd.Text = "", "��ü", txtCtrlCd.Text)
    strFromTime = Format(dtpFromTime.Value, "HH:mm:ss")
    strToTime = Format(dtpToTime.Value, "HH:mm:ss")
        
    If optQryKey(0).Value Then
        strQryKey = "ä�� ���"
    Else
        strQryKey = "����� ���("
        For i = chkQryKey.LBound To chkQryKey.UBound
            If chkQryKey(i).Value = 1 Then
                strQryKey = strQryKey & Mid(chkQryKey(i).Caption, 1, 2) & "/"
            End If
        Next
        strQryKey = Mid(strQryKey, 1, Len(strQryKey) - 1)
        strQryKey = strQryKey & ")"
    End If
   
    strSchedule = IIf(chkSchedule.Value = 1, "�����ٳ�����", "��ü")
   
    strDesc = lblDesc.Caption
    
    If (InStr(medGetP(medGetP(strDesc, 1, ","), 1, "~"), strFromDate) > 0) And _
       (InStr(medGetP(medGetP(strDesc, 1, ","), 2, "~"), strToDate) > 0) And _
       (InStr(medGetP(strDesc, 2, ","), strSection) > 0) And _
       (InStr(medGetP(strDesc, 3, ","), strLevel) > 0) And _
       (InStr(medGetP(strDesc, 4, ","), strCtrlcd) > 0) And _
       (InStr(medGetP(medGetP(strDesc, 5, ","), 1, "~"), strFromTime) > 0) And _
       (InStr(medGetP(medGetP(strDesc, 5, ","), 2, "~"), strToTime) > 0) And _
       (InStr(medGetP(strDesc, 6, ","), strQryKey) > 0) And _
       (InStr(medGetP(strDesc, 7, ","), strSchedule) > 0) Then
        lblWarn.Caption = ""
    Else
        lblWarn.Caption = "�� �˻������� ����Ǿ� ���� ��ȸ�� ����� �ٸ��ϴ�." '"�� ��ȸ�� ����� �˻������� ���� �ٸ��ϴ�."
    End If
End Sub

Private Sub dtpFromDate_Change()
    Call GetWarningMsg
End Sub

Private Sub dtpFromTime_Change()
    Call GetTimeType
    Call GetWarningMsg
End Sub

Private Sub dtpToDate_Change()
    Call GetWarningMsg
End Sub

Private Sub dtpToTime_Change()
    Call GetTimeType
    Call GetWarningMsg
End Sub

Private Sub GetTimeType()
    Dim strFromTime As String
    Dim strToTime As String
    
    strFromTime = Format(dtpFromTime.Value, "HHMMss")
    strToTime = Format(dtpToTime.Value, "HHMMss")

    If strFromTime = "000001" And strToTime = "235959" Then
        optWorkTime(0).Value = True
    ElseIf strFromTime = "080001" And strToTime = "120000" Then
        optWorkTime(1).Value = True
    ElseIf strFromTime = "120001" And strToTime = "200000" Then
        optWorkTime(2).Value = True
    ElseIf strFromTime = "200001" And strToTime = "235959" Then
        optWorkTime(3).Value = True
    Else
        optWorkTime(4).Value = True
    End If
End Sub

Private Sub Form_Load()
    
    cboSection.Clear
    Call InitForm
    
    DoEvents
    Call LoadSection
    
    DoEvents
    Call ReadConfig
End Sub

Private Sub ReadConfig()
    Dim User As String
    
    User = GetSetting("Schweitzer2000 LIS\Config", "frm310QCReprint_N", "User", "")
    
    If InStr(User, ObjSysInfo.logonid) = 0 Then Exit Sub
             
    dtpToDate.Value = GetSystemDate
    
    If GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "DuraDate", "") = "" Then
        dtpFromDate.Value = DateAdd("d", -7, dtpToDate.Value)
    Else
        dtpFromDate.Value = DateAdd("d", "-" & Val(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "DuraDate", "")), dtpToDate.Value)
    End If
    
    cboSection.ListIndex = medComboFind(cboSection, GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "Section", ""))
    cboLevel.ListIndex = medComboFind(cboLevel, GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "Level", ""))
    txtCtrlCd.Text = GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "CtrlCd", "")
    lblCtrlNm.Caption = GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "CtrlNm", "")
    chkAll.Value = Val(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "AllCtrl", ""))
    optWorkTime(Val(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "WorkTime", ""))).Value = 1
    dtpFromTime.Value = medGetP(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "DuraTime", ""), 1, ",")
    dtpToTime.Value = medGetP(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "DuraTime", ""), 2, ",")
    optQryKey(Val(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "optQryKey", ""))).Value = 1
    chkQryKey(0).Value = Val(medGetP(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "chkQryKey", ""), 1, ","))
    chkQryKey(1).Value = Val(medGetP(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "chkQryKey", ""), 2, ","))
    chkQryKey(2).Value = Val(medGetP(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "chkQryKey", ""), 3, ","))
    chkBarPrint.Value = Val(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "BarPrint", ""))
End Sub

Private Sub InitForm()
    dtpFromDate.Value = GetSystemDate
    dtpToDate.Value = GetSystemDate
    
    cboLevel.ListIndex = 0
    txtCtrlCd.Text = ""
    lblCtrlNm.Caption = ""
    optWorkTime(4).Value = True
    dtpFromTime.Value = GetSystemDate
    dtpToTime.Value = DateAdd("h", 8, dtpFromTime.Value)
    tblResult.MaxRows = 0
    tblResult.MaxRows = 27
    Call medClearTable(tblResult)
    
    lblWarn.Caption = ""
    lblDesc.Caption = ""
End Sub

Private Sub LoadSection()
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = "select * from " & T_LAB032 & " where " & DBW("cdindex=", LC3_Section)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    cboSection.Clear
    cboSection.addItem "��ü"
    Do Until Rs.EOF
        cboSection.addItem Format(Rs.Fields("field1").Value & "", "!" & String(50, "@")) & COL_DIV & _
                           Rs.Fields("cdval1").Value & ""
        Rs.MoveNext
    Loop
        
    cboSection.ListIndex = 0
        
    Set Rs = Nothing
End Sub

Private Sub optQryKey_Click(Index As Integer)
'    On Error Resume Next
'    If Screen.ActiveControl.Name <> optQryKey(Index).Name Then Exit Sub
    
    
    If optQryKey(0).Value Then
        chkQryKey(0).Value = 0
        chkQryKey(1).Value = 0
        chkQryKey(2).Value = 0
        
        fraQryKey.Enabled = False
        chkSchedule.Value = 0
        chkSchedule.Enabled = False
        
'        chkBarPrint.Value = 1   '������ �о� ��..
        chkBarPrint.Enabled = True
        
        cmdCollect.Caption = "ä  ��"
        
        Dim User As String

        User = GetSetting("Schweitzer2000 LIS\Config", "frm401ResultView_N", "User", "")

        If InStr(User, ObjSysInfo.logonid) <> 0 Then
            chkBarPrint.Value = Val(GetSetting("Schweitzer2000 LIS\Config\frm310QCReprint_N", ObjSysInfo.logonid, "BarPrint", ""))
        End If
    ElseIf optQryKey(1).Value Then
        fraQryKey.Enabled = True
        chkSchedule.Enabled = True
        chkQryKey(0).Value = 1
        
        chkBarPrint.Value = 1
        chkBarPrint.Enabled = False
        
        cmdCollect.Caption = "�����"
    End If
    
    If tblResult.DataRowCnt = 0 Then Exit Sub
    Call GetWarningMsg
End Sub

Private Sub optWorkTime_Click(Index As Integer)
    On Error Resume Next
    If Screen.ActiveControl.Name <> optWorkTime(Index).Name Then Exit Sub
    
    Dim strFrTime As String, strToTime As String
    
    If optWorkTime(4).Value Then
        dtpFromTime.Value = GetSystemDate
        dtpToTime.Value = DateAdd("h", 8, dtpFromTime.Value)
        
        Call GetWarningMsg
        Exit Sub
    End If
    
    strFrTime = medGetP(optWorkTime(Index).Tag, 1, ":") '& "00"
    strToTime = medGetP(optWorkTime(Index).Tag, 2, ":") '& "00"
    dtpFromTime.Value = Format(strFrTime, "0#:##:##")
    dtpToTime.Value = Format(strToTime, "0#:##:##")
    
    Call GetWarningMsg
End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Long
    Static blnToggle As Boolean
    
    If tblResult.DataRowCnt = 0 Then Exit Sub
    
    If Col = 1 And Row = 0 Then
        blnToggle = IIf(blnToggle, False, True)
        
        For i = 1 To tblResult.DataRowCnt
            tblResult.Col = 1
            tblResult.Row = i
            tblResult.Value = IIf(blnToggle, 1, 0)
        Next
    End If
    
    '���� �ϴ��� ���� �������.. �Ʊ� �Ӹ���~~
    '�����°� ���� ������̾�!
    'frm311QCResultEntry���� �׻� ��ε� �ϰ� ���� ����� �ϴ°�? ������� ���� �����ø�....
    
    If Col = 19 And Row <> 0 Then
        tblResult.Col = 19
        tblResult.Row = Row
        If tblResult.Value = "��" Then  '������������ ȭ���� ����ش�.
            Dim strWorkArea As String
            Dim strAccDt As String
            Dim strAccSeq As String

            tblResult.Col = 16: strWorkArea = tblResult.Value
            tblResult.Col = 17: strAccDt = Mid(tblResult.Value, 3)
            tblResult.Col = 18: strAccSeq = tblResult.Value
            Call LoadForm(frm311QCResultEntry_N, Me)
            Call frm311QCResultEntry_N.CallByExternal(strWorkArea & "-" & strAccDt & "-" & strAccSeq)
'            Dim frm As Form
'            Dim blnExist As Boolean
'            Dim strWorkArea As String
'            Dim strAccDt As String
'            Dim strAccSeq As String
'
'            tblResult.Col = 16
'            strWorkArea = tblResult.Value
'            tblResult.Col = 17
'            strAccDt = Mid(tblResult.Value, 3)
'            tblResult.Col = 18
'            strAccSeq = tblResult.Value
'
'            frm311QCResultEntry.ParentHwnd = GetAncestor(Me.hwnd, 1)
'
'            Unload frm311QCResultEntry
'
'            DoEvents
'            Call SetParent(frm311QCResultEntry.hwnd, GetAncestor(Me.hwnd, 1))
'            frm311QCResultEntry.WindowState = 2
'            frm311QCResultEntry.Show
'            frm311QCResultEntry.ZOrder 0
'
'            frm311QCResultEntry.mskAccNo.Text = strWorkArea & "-" & strAccDt & "-" & strAccSeq & String(4 - Len(strAccSeq), "_")
'            frm311QCResultEntry.Call_mskAccNo_LostFocus
        End If
    End If
End Sub

Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    With tblResult
        If .DataRowCnt = 0 Then ShowTip = False: Exit Sub
        
        If Col = 11 And Row <> 0 Then
            .Col = Col
            .Row = Row
            
            If .ForeColor = vbRed Then
                .Col = 20
                MultiLine = 1
                TipText = vbNewLine & " ���� ���� : " & .Value & vbNewLine
                TipWidth = 4000
                .TextTipDelay = 1000
                Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
                ShowTip = True
            Else
                ShowTip = False
            End If
        Else
            ShowTip = False
        End If
    End With
End Sub

Private Sub txtCtrlCd_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtCtrlCd.Name Then Exit Sub
    
    If lblCtrlNm.Caption <> "" Then lblCtrlNm.Caption = ""
    Call GetWarningMsg
End Sub

Private Sub txtCtrlCd_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCtrlCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCtrlCd_LostFocus()
'    Dim Rs As Recordset
    Dim strSQL As String
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If lblCtrlNm.Caption <> "" Then Exit Sub
    
    Call LoadControlInfo(Trim(txtCtrlCd.Text))
'    strSQL = " select distinct a.ctrlcd,a.ctrlnm from " & T_LAB021 & " a, " & T_LAB023 & " b " & _
'             " where a.ctrlcd = b.ctrlcd " & _
'             " and a.levelcd=b.levelcd " & _
'             " and " & DBW("b.opendt<=", Format(GetSystemDate, CS_DateDbFormat)) & _
'             " and " & DBW("b.expdt>=", Format(GetSystemDate, CS_DateDbFormat)) & _
'             " and " & DBW("a.buildcd=", ObjSysInfo.BuildingCd) & _
'             " and " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text))
'
'    If cboSection.ListIndex > 0 Then
'        strSQL = strSQL & " and " & DBW("a.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
'    End If
'
'    If cboLevel.ListIndex > 0 Then
'        strSQL = strSQL & " and " & DBW("a.levelcd=", Trim(Mid(cboLevel.Text, 20)))
'    End If
'
'    Set Rs = OpenRecordSet(strSQL)
'
'    If Rs.EOF Then
'        MsgBox "�ش� ��Ʈ���� �������� �ʽ��ϴ�.", vbExclamation
'        txtCtrlCd.Text = ""
'        txtCtrlCd.SetFocus
'    Else
'        lblCtrlNm.Caption = Rs.Fields("ctrlnm").Value & ""
'    End If
'
'    Set Rs = Nothing
End Sub
