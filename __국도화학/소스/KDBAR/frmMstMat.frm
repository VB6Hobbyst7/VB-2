VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMstMat 
   Caption         =   "�����ڵ� ���"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16800
   LinkTopic       =   "Form1"
   ScaleHeight     =   11220
   ScaleWidth      =   16800
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " �����ڵ� ���� �Է� "
      Height          =   1275
      Left            =   90
      TabIndex        =   2
      Top             =   8010
      Width           =   15225
      Begin VB.TextBox txtMatCd 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   450
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "123456"
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtMatNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "����Ŭ����"
         Top             =   690
         Width           =   3765
      End
      Begin VB.TextBox txtUserPW 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5490
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "0001"
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtUserRegID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8010
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "123456"
         Top             =   690
         Width           =   1245
      End
      Begin VB.CheckBox chkUsedYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         Height          =   255
         Left            =   6990
         TabIndex        =   7
         Top             =   720
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�ݱ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   12390
         Style           =   1  '�׷���
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  '���
         BackColor       =   &H00C0FFFF&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10110
         Style           =   1  '�׷���
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  '���
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   11250
         Style           =   1  '�׷���
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   13530
         Style           =   1  '�׷���
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   450
         Top             =   390
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   1710
         Top             =   390
         Width           =   3765
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   2
         Left            =   5490
         Top             =   390
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   5
         Left            =   6750
         Top             =   390
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   6
         Left            =   8010
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFC0&
         Caption         =   "�����ڵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   450
         TabIndex        =   16
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFC0&
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1710
         TabIndex        =   15
         Top             =   420
         Width           =   3765
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFC0&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   5490
         TabIndex        =   14
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFC0&
         Caption         =   "��뿩��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   6750
         TabIndex        =   13
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFC0&
         Caption         =   "�Է���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   8010
         TabIndex        =   12
         Top             =   420
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " �����ڵ� ����Ʈ "
      Height          =   7905
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   15225
      Begin FPSpread.vaSpread spdMat 
         Height          =   7545
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   14925
         _Version        =   393216
         _ExtentX        =   26326
         _ExtentY        =   13309
         _StockProps     =   64
         ColsFrozen      =   8
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
         MaxCols         =   10
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   15400934
         SpreadDesigner  =   "frmMstMat.frx":0000
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmMstMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmMstMat.frm
'   �ۼ���  : ������
'   ��  ��  : �����ڵ���
'   �ۼ���  : 2020-02-10
'   ��  ��  : 1.0.0
'   ��  ��  : ����ȭ��
'-----------------------------------------------------------------------------'

Private Sub cmdClear_Click()
    
    txtMatCd.Text = ""
    txtMatNm.Text = ""
    chkUsedYN.Value = "1"
    txtUserRegID.Text = gKUKDO.USERID

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


'-- �����ڿ�
Private Sub cmdDelete_Click()
    
    gMAT.CD = txtUserID.Text
    gMAT.NAME = txtUserNm.Text
    
    If chkUsedYN.Value = "1" Then
        gMAT.YN = "Y"
    Else
        gMAT.YN = "N"
    End If
    
    If Set_User("DEL") Then
        Call CtlInitializing
        Call GetUserList
    End If
    
End Sub

Private Sub cmdOK_Click()

    Call SetUser
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    Call GetUserList
    
End Sub

Private Sub GetUserList()

    Set AdoRs = Get_UserList
    
    If AdoRs Is Nothing Then
        '��ϵ� ���� ����
    Else
        Do Until AdoRs.EOF
            With spdUser
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdUser, AdoRs.Fields("USER_CD").Value & "", .MaxRows, 1)
                Call SetText(spdUser, AdoRs.Fields("USER_NAME").Value & "", .MaxRows, 2)
                Call SetText(spdUser, AdoRs.Fields("USER_PW").Value & "", .MaxRows, 3)
                Call SetText(spdUser, AdoRs.Fields("USER_DEPART").Value & "", .MaxRows, 4)
                
                If AdoRs.Fields("USER_COMP").Value & "" = "1" Then
                    Call SetText(spdUser, "������", .MaxRows, 5)
                Else
                    Call SetText(spdUser, "�����", .MaxRows, 5)
                End If
                
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdUser, "1", .MaxRows, 6)
                Else
                    Call SetText(spdUser, "0", .MaxRows, 6)
                End If
                
                Call SetText(spdUser, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 7)
                Call SetText(spdUser, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 8)
                Call SetText(spdUser, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 9)
                Call SetText(spdUser, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 10)
            End With
            
            AdoRs.MoveNext
        Loop
    
    End If
    
    AdoRs.Close
    
End Sub

Private Sub SetUser()
    
    '�ʼ��Է� üũ
    If txtUserID.Text = "" Then
        MsgBox "�����ID�� �Է��ϼ���", vbOKOnly + vbCritical, Me.Caption
        txtUserID.SetFocus
        Exit Sub
    End If
        
    If txtUserNm.Text = "" Then
        MsgBox "����ڸ��� �Է��ϼ���", vbOKOnly + vbCritical, Me.Caption
        txtUserNm.SetFocus
        Exit Sub
    End If
        
    If txtUserPW.Text = "" Then
        MsgBox "����� ��й�ȣ�� �Է��ϼ���", vbOKOnly + vbCritical, Me.Caption
        txtUserPW.SetFocus
        Exit Sub
    End If
        
    '-- ���
    gMAT.ID = txtUserID.Text
    gMAT.NAME = txtUserNm.Text
    gMAT.PW = txtUserPW.Text
    gMAT.DEPT = txtUserDepart.Text
    If cboUserComp.Text = "�����" Then
        gMAT.COMP = "2"
    Else
        gMAT.COMP = "1"
    End If
    If chkUsedYN.Value = "1" Then
        gMAT.YN = "Y"
    Else
        gMAT.YN = "N"
    End If
    
    '-- Insert / Update ã�ƿ���
    Set AdoRs = Get_UserList(txtUserID.Text)
        
    '-- ����
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_User("IN") Then
            Call CtlInitializing
            Call GetUserList
        End If
    Else
        'UPDATE
        If Set_User("UP") Then
            Call CtlInitializing
            Call GetUserList
        End If
    End If
    
End Sub

'-- ��Ʈ���ʱ�ȭ
Private Sub CtlInitializing()
    
    With spdUser
        Call SetText(spdUser, "�����ID", 0, 1):    .ColWidth(1) = 10
        Call SetText(spdUser, "����ڸ�", 0, 2):    .ColWidth(2) = 10
        Call SetText(spdUser, "���", 0, 3):        .ColWidth(3) = 8
        Call SetText(spdUser, "�μ�", 0, 4):        .ColWidth(4) = 8
        Call SetText(spdUser, "����", 0, 5):        .ColWidth(5) = 8
        Call SetText(spdUser, "��뿩��", 0, 6):    .ColWidth(6) = 10
        Call SetText(spdUser, "�Է���", 0, 7):      .ColWidth(7) = 10
        Call SetText(spdUser, "�Է��Ͻ�", 0, 8):    .ColWidth(8) = 20
        Call SetText(spdUser, "������", 0, 9):      .ColWidth(9) = 10
        Call SetText(spdUser, "�����Ͻ�", 0, 10):   .ColWidth(10) = 20
    
        .MaxRows = 0
    End With
    
    txtUserID.Text = ""
    txtUserNm.Text = ""
    txtUserPW.Text = ""
    txtUserDepart.Text = ""
    cboUserComp.ListIndex = 0
    chkUsedYN.Value = "1"
    txtUserRegID.Text = gKUKDO.USERID
    
    If gKUKDO.USERGRD = "1" Then
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    
    gSORT = 0

End Sub

'-- ����� ����
Private Sub spdUser_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SetSpreadSort(spdUser)
        Exit Sub
    End If
    
    txtUserID.Text = GetText(spdUser, Row, 1)
    txtUserNm.Text = GetText(spdUser, Row, 2)
    txtUserPW.Text = GetText(spdUser, Row, 3)
    txtUserDepart.Text = GetText(spdUser, Row, 4)
    If GetText(spdUser, Row, 5) = "�����" Then
        cboUserComp.ListIndex = 0
    Else
        cboUserComp.ListIndex = 1
    End If
    If GetText(spdUser, Row, 6) = "1" Then
        chkUsedYN.Value = "1"
    Else
        chkUsedYN.Value = "0"
    End If
    txtUserRegID.Text = GetText(spdUser, Row, 7)
    
End Sub

