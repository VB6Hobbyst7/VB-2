VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS408 
   BackColor       =   &H00DBE6E6&
   Caption         =   "���� ���� ����"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "frmBBS408.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14715
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�԰����"
      Height          =   510
      Left            =   8100
      Style           =   1  '�׷���
      TabIndex        =   8
      Top             =   7320
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   10740
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   7320
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   12060
      Style           =   1  '�׷���
      TabIndex        =   7
      Top             =   7320
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&S)"
      Height          =   510
      Left            =   9420
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   7320
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1020
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "���� ����"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   4605
      TabIndex        =   10
      Top             =   1020
      Width           =   8760
      _ExtentX        =   15452
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
      Caption         =   "���� ���� ����"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2940
      Left            =   1680
      TabIndex        =   11
      Top             =   1260
      Width           =   2910
      Begin MSComCtl2.MonthView mvRcvDt 
         Height          =   2220
         Left            =   315
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   585
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   14411494
         Appearance      =   1
         StartOfWeek     =   66650113
         CurrentDate     =   36853
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   315
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "��������"
         Appearance      =   0
      End
      Begin VB.Label lblRcvDt 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1395
         TabIndex        =   14
         Tag             =   "103"
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   3060
      Left            =   1680
      TabIndex        =   12
      Top             =   4125
      Width           =   2910
      Begin VB.ListBox lstRcvDt 
         Height          =   2400
         Left            =   300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   2235
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   0
         Left            =   300
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   120
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   582
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
         Caption         =   "�������� ��� ����"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   4170
      Left            =   4605
      TabIndex        =   16
      Top             =   1260
      Width           =   8790
      Begin VB.TextBox txtBldYY 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         TabIndex        =   1
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtBldSrc 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtFrNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         MaxLength       =   10
         TabIndex        =   2
         Top             =   300
         Width           =   1005
      End
      Begin VB.TextBox txtToNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   1005
      End
      Begin MedControls1.LisLabel lblBldSrcNm 
         Height          =   315
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3600
         Width           =   6555
         _ExtentX        =   11562
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
      Begin MedControls1.LisLabel lblTotCnt 
         Height          =   330
         Left            =   6975
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         Width           =   1110
         _ExtentX        =   1958
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
      Begin FPSpread.vaSpread tblEnter 
         Height          =   2790
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   7860
         _Version        =   196608
         _ExtentX        =   13864
         _ExtentY        =   4921
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
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
         GridShowVert    =   0   'False
         MaxCols         =   9
         MaxRows         =   5
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS408.frx":076A
      End
      Begin MedControls1.LisLabel lblSum 
         Height          =   315
         Left            =   6840
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MedControls1.LisLabel lblBldSrc 
         Height          =   330
         Left            =   3420
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   300
         Width           =   375
         _ExtentX        =   661
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
      Begin MedControls1.LisLabel lblBldYY 
         Height          =   330
         Left            =   3810
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   300
         Width           =   375
         _ExtentX        =   661
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   5895
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   300
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "���μ���"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   300
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "���� ��ȣ"
         Appearance      =   0
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3195
         TabIndex        =   23
         Tag             =   "103"
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   1815
      Left            =   4605
      TabIndex        =   24
      Top             =   5370
      Width           =   8790
      Begin VB.TextBox txtRemark 
         Height          =   1155
         Left            =   225
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   4
         Top             =   540
         Width           =   8355
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   225
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   195
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
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
         Caption         =   "���� Remark"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS408"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ListPop As clsPopUpList
Attribute ListPop.VB_VarHelpID = -1
Private objProgress As clsProgress

Private isQuery As Boolean



Private Sub cmdBldSrc_Click()
    Dim objSQL  As clsGetSqlStatement
    
    Set objSQL = New clsGetSqlStatement
    ListPop.Connection = DBConn
    
    Call ListPop.LoadPopUp(objSQL.GetBldSrcList)
    
    Set objSQL = Nothing
End Sub

Private Sub cmdCancel_Click()
    Dim objBDP As clsBloodDonationPaper
    Dim strMsg As String
    
    Set objBDP = New clsBloodDonationPaper
    If objBDP.IsExistUseable(Format(mvRcvDt, PRESENTDATE_FORMAT), "<") = False Then
        MsgBox "�� �ڷḦ �����ϸ� ����� ���������� �����ϴ�.", vbCritical, Me.Caption
        Set objBDP = Nothing
        Exit Sub
    End If
    Set objBDP = Nothing
    
    strMsg = "�ѹ� ������ �ڷ�� ������ �� �����ϴ�." & vbNewLine & _
             "��� �Ͻð����ϱ�?"
    If MsgBox(strMsg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    If Delete = True Then
        Call Clear
        isQuery = False
        SetLstRcvDt
    End If
End Sub

Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strMsg As String
    
    If chkValid = False Then Exit Sub
    If isQuery = True Then
        strMsg = "�����ϸ� ������ ������ ������ �� �����ϴ�." & vbNewLine & _
                 "�����Ͻð����ϱ�?"
        If MsgBox(strMsg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    End If
    
    Set objProgress = New clsProgress
    
    If Save = True Then
        Call Query
        isQuery = True
        SetLstRcvDt
    End If
    
    Set objProgress = Nothing
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim objSQL  As clsGetSqlStatement
    Dim code    As String
    Dim name    As String
    
    '�޷¿� ���� ���ڸ� �⺻���� ����...
    mvRcvDt = GetSystemDate
    
    '����ŷ����� ���׿��� �����ش�.
    Set objSQL = New clsGetSqlStatement
    Call objSQL.GetActiveBldSrc(code, name)
    txtBldSrc = code
    txtBldSrc_LostFocus
    Set objSQL = Nothing
    'ȭ������
    Call ClearAll
    
    '�������� ��� ����
    SetLstRcvDt
End Sub

'Private Sub ListPop_SendCode(ByVal SelString As String)
'    Dim Row As Long
'
'    If ListPop.tag = "BLDSRC" Then
'        txtBldSrc = medGetP(SelString, 1, ";")
'        lblBldSrcNm.Caption = medGetP(SelString, 2, ";")
'    Else
'        Row = Val(ListPop.tag)
'        With tblEnter
'            .Row = Row
'            .Col = 5: .value = medGetP(SelString, 2, ";") '�������̸�
'            .Col = 8: .value = medGetP(SelString, 1, ";") '������ ID
'        End With
'    End If
'End Sub

Private Sub ListPop_SelectedItem(ByVal pSelectedItem As String)
    Dim Row As Long
    
    If ListPop.tag = "BLDSRC" Then
        txtBldSrc = medGetP(pSelectedItem, 1, ";")
        lblBldSrcNm.Caption = medGetP(pSelectedItem, 2, ";")
    Else
        Row = Val(ListPop.tag)
        With tblEnter
            .Row = Row
            .Col = 5: .value = medGetP(pSelectedItem, 2, ";") '�������̸�
            .Col = 8: .value = medGetP(pSelectedItem, 1, ";") '������ ID
        End With
    End If
End Sub

Private Sub lstRcvDt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If lstRcvDt.Text <> "" Then
            mvRcvDt = lstRcvDt.Text
            Call mvRcvDt_DateClick(mvRcvDt)
        End If
    End If
End Sub

Private Sub mvRcvDt_DateClick(ByVal DateClicked As Date)
    Dim objBDP As clsBloodDonationPaper
    Dim i As Long
    
    lblRcvDt = Format(DateClicked, "YYYY-MM-DD")
    
    lstRcvDt.ListIndex = -1
    For i = 0 To lstRcvDt.ListCount - 1
        If lblRcvDt = lstRcvDt.List(i) Then
            lstRcvDt.ListIndex = i
            Exit For
        End If
    Next i
    
    
    If IsEnter(Format(DateClicked, PRESENTDATE_FORMAT)) = True Then
        '���ų�����ȸ
        Call Query
        isQuery = True
    Else
        Clear
        isQuery = False
    End If
    
    '�԰���Ҹ� �� �� �ִ��� ����(����)---------------------------------------------------------
    If isQuery = True Then
        If Format(DateClicked, "YYYY-MM-DD") < lstRcvDt.List(0) Then
            cmdCancel.Enabled = False
        Else
            Set objBDP = New clsBloodDonationPaper
            cmdCancel.Enabled = objBDP.IsExistUseable(Format(DateClicked, PRESENTDATE_FORMAT), "<")
            Set objBDP = Nothing
        End If
    End If
    
    '�԰�ó���� �� �� �ִ��� ���θ� ����(�ű��Է� Ȥ�� ����)------------------------------------
    If lstRcvDt.ListCount > 1 Then
        If Format(DateClicked, "YYYY-MM-DD") < lstRcvDt.List(0) Then
            '�԰�ó���� �� ����
            txtFrNo.Enabled = False
            txtToNo.Enabled = False
            tblEnter.Enabled = False
            tblEnter.Col = 1: tblEnter.Row = 1: tblEnter.Action = ActionActiveCell
            cmdSave.Enabled = False
        ElseIf Format(DateClicked, "YYYY-MM-DD") > lstRcvDt.List(0) Then
            '�԰�ó���� �� ����
            txtFrNo.Enabled = True
            txtToNo.Enabled = True
            tblEnter.Enabled = True
            tblEnter.Col = 1: tblEnter.Row = 1: tblEnter.Action = ActionActiveCell
            cmdSave.Enabled = True
        ElseIf Format(DateClicked, "YYYY-MM-DD") = lstRcvDt.List(0) Then
            Set objBDP = New clsBloodDonationPaper
            
            If objBDP.GetNotUsedCnt(Format(DateClicked, PRESENTDATE_FORMAT)) = Val(lblTotCnt.Caption) Then
                '�԰�ó���� �� ����
                txtFrNo.Enabled = True
                txtToNo.Enabled = True
                tblEnter.Enabled = True
                tblEnter.Col = 1: tblEnter.Row = 1: tblEnter.Action = ActionActiveCell
                cmdSave.Enabled = True
            Else
                '�԰�ó���� �� ����
                txtFrNo.Enabled = False
                txtToNo.Enabled = False
                tblEnter.Enabled = False
                tblEnter.Col = 1: tblEnter.Row = 1: tblEnter.Action = ActionActiveCell
                cmdSave.Enabled = False
            End If
            
            Set objBDP = Nothing
        End If
    Else
        '�԰�ó���� �� ����
        txtFrNo.Enabled = True
        txtToNo.Enabled = True
        tblEnter.Enabled = True
        tblEnter.Col = 1: tblEnter.Row = 1: tblEnter.Action = ActionActiveCell
        cmdSave.Enabled = True
    End If

End Sub

Private Sub mvRcvDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call mvRcvDt_DateClick(mvRcvDt)
    End If
End Sub

Private Sub tblEnter_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Set ListPop = New clsPopUpList
    
    ListPop.tag = Row
    ListPop.Connection = DBConn
    Call ListPop.LoadPopUp(GetSQLEmpList)
    
End Sub

Private Sub tblEnter_Change(ByVal Col As Long, ByVal Row As Long)
    Dim frno As Double
    Dim tono As Double
    Dim r As Long
    Dim sum As Double
    
        
    With tblEnter
        If Col <> 2 And Col <> 3 Then Exit Sub
        
        .Row = Row
        
        .Col = 2: frno = Val(.value)
        .Col = 3: tono = Val(.value)
        .Col = 4: .value = tono - frno + 1
        
        sum = 0
        For r = 1 To .MaxRows
            .Row = r
            .Col = 4: sum = sum + Val(.value)
        Next r
        
        lblSum.Caption = IIf(sum = 0, "", sum)
    End With

End Sub

Private Sub txtBldSrc_Change()
    lblBldSrc.Caption = txtBldSrc
End Sub

Private Sub txtBldSrc_LostFocus()
    Dim objSQL  As clsGetSqlStatement
    Dim nm      As String
    
    Set objSQL = New clsGetSqlStatement
    
    nm = objSQL.GetBldSrcNm(txtBldSrc)
    If nm <> "" Then
        lblBldSrcNm.Caption = "���׿� (" & nm & ") ���� ���� �������� �԰� �۾��Դϴ�."
    Else
        lblBldSrcNm.Caption = ""
    End If
    
    Set objSQL = Nothing
End Sub

Private Sub txtBldYY_Change()
    lblBldYY.Caption = txtBldYY
End Sub

Private Sub txtFrNo_Change()
    If txtFrNo = "" Or txtToNo = "" Then lblTotCnt.Caption = ""
    
    lblTotCnt.Caption = Val(txtToNo) - Val(txtFrNo) + 1
End Sub

Private Sub txtToNo_Change()
    If txtFrNo = "" Or txtToNo = "" Then lblTotCnt.Caption = ""
    
    lblTotCnt.Caption = Val(txtToNo) - Val(txtFrNo) + 1
End Sub










Private Sub ClearAll()
'    lblRcvDt = ""
    Clear

'    txtFrNo.Enabled = False
'    txtToNo.Enabled = False
'    tblEnter.Enabled = False
    tblEnter.Col = 1: tblEnter.Row = 1: tblEnter.Action = ActionActiveCell
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub Clear()
    txtFrNo = ""
    txtToNo = ""
    lblTotCnt.Caption = ""
    lblSum.Caption = ""
    ClearTblEnter
End Sub

Private Sub ClearTblEnter()
    Dim Rs As Recordset
    Dim r As Long
    Dim i As Long
    
    
    Set Rs = New Recordset
    
    
    r = 0
    With tblEnter
        .ReDraw = False
        
        .MaxRows = 0
        
        ' ���� ����Ʈ-----------------------------------------------------
        Call Rs.Open(GetCom003(BC2_CENTER), DBConn)
        If Rs.EOF Then
'            'dbconn.DisplayErrors
        Else
            .MaxRows = Rs.RecordCount
            For i = 1 To Rs.RecordCount
                r = r + 1
                .Row = r
                .Col = 1: .value = Rs.Fields("field1").value & ""    '���͸�
                .Col = 2: .value = ""
                .Col = 3: .value = ""
                .Col = 4: .value = ""
                .Col = 5: .value = ObjMyUser.EmpLngNm
                .Col = 7: .value = Rs.Fields("cdval1").value & ""   '�����ڵ�
                .Col = 8: .value = ObjMyUser.EmpId
                .Col = 9: .value = "0"                      '����
                
                Rs.MoveNext
            Next i
        End If
        
        ' �ں��� ����Ʈ---------------------------------------------------
        Set Rs = Nothing
        Set Rs = New Recordset
        Call Rs.Open(GetCom003(BC2_BRANCH), DBConn)
        If Rs.EOF Then
'            'dbconn.DisplayErrors
        Else
            .MaxRows = .MaxRows + Rs.RecordCount
            For i = 1 To Rs.RecordCount
                r = r + 1
                .Row = r
                .Col = 1: .value = Rs.Fields("field1").value & ""    '�ں�����
                .Col = 2: .value = ""
                .Col = 3: .value = ""
                .Col = 4: .value = ""
                .Col = 5: .value = ObjMyUser.EmpLngNm
                .Col = 7: .value = Rs.Fields("cdval1").value & ""    '�ں����ڵ�
                .Col = 8: .value = ObjMyUser.EmpId
                .Col = 9: .value = "1"
                Rs.MoveNext
            Next i
        End If
    
        .ReDraw = True
    End With
    
    
    Set Rs = Nothing

End Sub

Private Sub SetLstRcvDt()
    Dim objBDP As clsBloodDonationPaper
    Dim astrRcvDt() As String
    Dim Cnt As Long
    Dim i As Long
    
    
    '���ſ� �԰�ó���� ���ڸ���Ʈ
    Set objBDP = New clsBloodDonationPaper
    Cnt = objBDP.GetRcvDtList(astrRcvDt)
    lstRcvDt.Clear
    For i = 0 To Cnt - 1
        lstRcvDt.AddItem Format(astrRcvDt(i), "####-##-##")
    Next i
    Set objBDP = Nothing
End Sub

Private Function IsEnter(ByVal rcvdt As String) As Boolean
    Dim i As Long
    
    If lstRcvDt.ListCount <= 0 Then IsEnter = False
    
    For i = 0 To lstRcvDt.ListCount - 1
        If rcvdt = Format(lstRcvDt.List(i), PRESENTDATE_FORMAT) Then IsEnter = True: Exit Function
    Next i
    
    IsEnter = False
End Function

Private Sub Query()
    Dim r As Long
    Dim Cnt As Long
    Dim sum As Double
    Dim rcvnm As String
    Dim centernm As String
    Dim Rs As Recordset
    Dim objBDP As clsBloodDonationPaper
    
    
    Clear
    
    Set objBDP = New clsBloodDonationPaper
    
    Set Rs = New Recordset
    Call Rs.Open(objBDP.GetEnterList(Format(mvRcvDt, PRESENTDATE_FORMAT)), DBConn)
    If Rs.EOF Then
'        'dbconn.DisplayErrors
        Set Rs = Nothing
        Set objBDP = Nothing
        Exit Sub
    End If
    
    
    sum = 0
    With tblEnter
        .ReDraw = False
        
        .MaxRows = Rs.RecordCount
        For r = 1 To Rs.RecordCount
            .Row = r
            
            If Rs.Fields("divcd").value & "" = "0" Then
                centernm = GetCenterNm(Rs.Fields("centercd").value & "")
                rcvnm = GetEmpNm(Rs.Fields("rcvid").value & "")
            Else
                centernm = GetBranchNm(Rs.Fields("centercd").value & "")
                rcvnm = Rs.Fields("rcvnm").value & ""
            End If
            
            .Col = 1: .value = centernm '����,�ں�����Ī
            .Col = 2: .value = Rs.Fields("frno").value & ""
            .Col = 3: .value = Rs.Fields("tono").value & ""
            .Col = 4: .value = Val(Rs.Fields("tono").value & "") - Val(Rs.Fields("frno").value & "") + 1
            .Col = 5: .value = rcvnm
            .Col = 7: .value = Rs.Fields("centercd").value & ""
            .Col = 8: .value = Rs.Fields("rcvid").value & ""
            .Col = 9: .value = Rs.Fields("divcd").value & ""
            
            If r = 1 Then txtBldSrc = Rs.Fields("bldsrc").value & ""

            If txtFrNo = "" Then txtFrNo = Rs.Fields("frno").value & ""
            If txtToNo = "" Then txtToNo = Rs.Fields("tono").value & ""
            
            If Val(Rs.Fields("frno")) < Val(txtFrNo) Then txtFrNo = Rs.Fields("frno").value & ""
            If Val(Rs.Fields("tono")) > Val(txtToNo) Then txtToNo = Rs.Fields("tono").value & ""
            sum = sum + Val(Rs.Fields("tono").value & "") - Val(Rs.Fields("frno").value & "") + 1
            
            Rs.MoveNext
        Next r
    
        .ReDraw = True
    End With
    txtBldYY.Text = Format(GetSystemDate, "yy")
    lblSum.Caption = IIf(sum = 0, "", sum)

    Set objBDP = Nothing
End Sub

Private Function chkDup(ByVal Prow As Long, ByVal pfrno As Double, ByVal ptono As Double) As Boolean
    Dim r As Long
    Dim frno As Double
    Dim tono As Double
    
    With tblEnter
        For r = Prow + 1 To .MaxRows
            .Row = r
            .Col = 2: frno = Val(.value)
            .Col = 3: tono = Val(.value)
            If pfrno >= frno And pfrno <= tono Then chkDup = False: Exit Function
            If ptono >= frno And ptono <= tono Then chkDup = False: Exit Function
        Next r
    End With
    
    chkDup = True
End Function

Private Function chkValid() As Boolean
    '�Է� ���� ��ȿ�� ����
    Dim chkCenter As Boolean
    Dim chkMsg As String
    Dim r As Long
    Dim frno As Double, tono As Double
    
    ' ----------------------------------------------------------
    ' step 1 : ���� ��ü �����Ͱ� �ԷµǾ����� �˻�
    ' ----------------------------------------------------------
    If txtFrNo = "" Or txtToNo = "" Then
        MsgBox "������ �������� ������ �Է��Ͻʽÿ�.", vbCritical, Me.Caption
        chkValid = False
        Exit Function
    End If
    
    ' ----------------------------------------------------------
    ' step 2 : ���Ϳ� �ں����� ��� �ԷµǾ����� �˻�
    ' ----------------------------------------------------------
'    chkCenter = True
'    With tblEnter
'        For r = 1 To .MaxRows
'            .Row = r
'            .Col = 2    'frno
'            If .value = "" Then chkCenter = False: Exit For
'            .Col = 3    'tono
'            If .value = "" Then chkCenter = False: Exit For
'        Next r
'    End With
'
'
'    If chkCenter = False Then
'        chkMsg = "������ �Ҵ��� �ȵ� ���ͳ� �ں����� �����ϴ�." & vbNewLine & "��� �����Ͻð����ϱ�?"
'        If MsgBox(chkMsg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then
'            chkValid = False
'            Exit Function
'        End If
'    End If
    
    ' ----------------------------------------------------------
    ' step 3 : �� ������ ���ͺ� ���� + �ں��� ������ �´��� �˻�
    ' ----------------------------------------------------------
    If lblTotCnt.Caption <> lblSum.Caption Then
        MsgBox "�Ѽ����� ���ͺ�,�ں����� ������ �հ谡 �ٸ��ϴ�.", vbCritical, Me.Caption
        chkValid = False
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' step 3 : ���ͺ� ��ȣ�� �ں��� ��ȣ�� �ߺ��� ���� �ִ��� �˻�
    ' ------------------------------------------------------------
    With tblEnter
        For r = 1 To .MaxRows
            .Row = r
            .Col = 2
            If .value <> "" Then
                .Col = 2: frno = Val(.value)
                .Col = 3: tono = Val(.value)
                
                If chkDup(r, frno, tono) = False Then
                    MsgBox "���ͺ�,�ں����� ������ȣ�� �ߺ��� �����ϴ�.", vbCritical, Me.Caption
                    chkValid = False
                    Exit Function
                End If
            End If
        Next r
    End With
    
    chkValid = True
End Function

Private Function Save() As Boolean
    Dim r As Long
    Dim no As Long
    Dim frno As Long
    Dim tono As Long
    Dim objBloodDonationPaper As clsBloodDonationPaper
    
'    Set objProgress.StatusBar = medMain.stsBar
    objProgress.Container = MainFrm.stsBar
    objProgress.Min = 1
    objProgress.Max = Val(lblTotCnt.Caption)
    objProgress.value = 0
    
On Error GoTo SAVE_ERROR

    DBConn.BeginTrans
    
    Set objBloodDonationPaper = New clsBloodDonationPaper
    
    If objBloodDonationPaper.Delete(txtBldSrc, Format(mvRcvDt, PRESENTDATE_FORMAT)) = False Then
        GoTo SAVE_ERROR
    End If
    
    With objBloodDonationPaper
        .BldSrc = txtBldSrc
        .BldYY = txtBldYY
        .rcvdt = Format(mvRcvDt, PRESENTDATE_FORMAT)
        .returndt = ""
        .returnid = ""
        .usedt = ""
        .useid = ""
    
        For r = 1 To tblEnter.MaxRows
            tblEnter.Row = r
            tblEnter.Col = 2
            If tblEnter.value <> "" Then
                tblEnter.Col = 2: frno = Val(tblEnter.value)
                tblEnter.Col = 3: tono = Val(tblEnter.value)
                tblEnter.Col = 9: .divcd = tblEnter.value
                tblEnter.Col = 8: .RcvID = tblEnter.value
                tblEnter.Col = 5: .rcvnm = tblEnter.value
                tblEnter.Col = 7: .CenterCd = tblEnter.value
                For no = frno To tono
                    .BldNo = no
                    If .Insert() = False Then GoTo SAVE_ERROR
                    
                    objProgress.value = objProgress.value + 1
                    
                Next no
            End If
        Next r
    End With
    
    Set objBloodDonationPaper = Nothing
    
    DBConn.CommitTrans
    Save = True
    
    Exit Function
    
SAVE_ERROR:
    Set objBloodDonationPaper = Nothing
    DBConn.RollbackTrans
    Save = False
    MsgBox Err.Description, vbExclamation
End Function

Private Function Delete() As Boolean
    Dim r As Long
    Dim no As Long
    Dim frno As Long
    Dim tono As Long
    Dim objBloodDonationPaper As clsBloodDonationPaper
    
    
On Error GoTo Delete_Error

    DBConn.BeginTrans
    
    Set objBloodDonationPaper = New clsBloodDonationPaper
    
    If objBloodDonationPaper.Delete(txtBldSrc, Format(mvRcvDt, PRESENTDATE_FORMAT)) = False Then
        GoTo Delete_Error
    End If
    
    Set objBloodDonationPaper = Nothing
    
    DBConn.CommitTrans
    Delete = True
    
    Exit Function
    
Delete_Error:
    Set objBloodDonationPaper = Nothing
    DBConn.RollbackTrans
    Delete = False
    MsgBox Err.Description, vbExclamation
End Function
