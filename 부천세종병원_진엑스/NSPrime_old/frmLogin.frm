VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "����� �α���"
   ClientHeight    =   3375
   ClientLeft      =   3240
   ClientTop       =   2925
   ClientWidth     =   5760
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5760
   StartUpPosition =   1  '������ ���
   Begin VB.CheckBox chkPW 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      Caption         =   "���̵�����"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4350
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ȯ��"
      Height          =   405
      Left            =   3840
      MaskColor       =   &H00000000&
      TabIndex        =   16
      Top             =   2910
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���"
      Height          =   405
      Left            =   4740
      MaskColor       =   &H00000000&
      TabIndex        =   15
      Top             =   2910
      Width           =   825
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2850
      TabIndex        =   13
      Top             =   2490
      Width           =   1425
   End
   Begin VB.Timer Timer1 
      Left            =   1170
      Top             =   2280
   End
   Begin VB.TextBox txtTemp 
      Height          =   495
      Left            =   -1170
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtPW 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '��� ����
      Left            =   7410
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2850
      TabIndex        =   3
      Top             =   2130
      Width           =   1425
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblSite 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�� ���ó : ��õ��������"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   3030
      TabIndex        =   17
      Top             =   150
      Width           =   2355
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "����ڸ� :"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1590
      TabIndex        =   14
      Top             =   2490
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "* ���̵� �Է��� �� �α����ϼ���"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   900
      TabIndex        =   12
      Top             =   1710
      Width           =   4515
   End
   Begin VB.Image imgNet3 
      Height          =   240
      Left            =   210
      Picture         =   "frmLogin.frx":058A
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet2 
      Height          =   240
      Left            =   210
      Picture         =   "frmLogin.frx":06D4
      Top             =   2940
      Width           =   240
   End
   Begin VB.Image imgNet1 
      Height          =   240
      Left            =   210
      Picture         =   "frmLogin.frx":081E
      Top             =   2940
      Width           =   240
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����� ID�� �Է� �Ͻʽÿ�."
      Height          =   180
      Left            =   480
      TabIndex        =   11
      Top             =   2970
      Width           =   2205
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   465
      Left            =   360
      Top             =   900
      Width           =   105
   End
   Begin VB.Label lblHospNm 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "���ٿ��κ���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   525
      Left            =   330
      TabIndex        =   10
      Top             =   180
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H008080FF&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   90
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   30
      Top             =   2130
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "���ܰ˻����а� "
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   3690
      TabIndex        =   8
      Top             =   330
      Width           =   3915
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '����
      Caption         =   "* ����� ID�� Password �� �߸��Ǿ����ϴ�."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   570
      TabIndex        =   7
      Top             =   1410
      Width           =   4515
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   7470
      TabIndex        =   6
      Top             =   2220
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblCommit 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Ȯ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6690
      TabIndex        =   5
      Top             =   2220
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPW 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��й�ȣ :"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6150
      TabIndex        =   2
      Top             =   2730
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblID 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "���̵� :"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1590
      TabIndex        =   1
      Top             =   2130
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FF8080&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   -30
      Top             =   2130
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblEquipName 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "ABL 800 Basic Interface"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   885
      Left            =   600
      TabIndex        =   0
      Top             =   900
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   0
      Picture         =   "frmLogin.frx":0968
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   5745
   End
   Begin VB.Image Image3 
      Height          =   2010
      Left            =   0
      Picture         =   "frmLogin.frx":18F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5745
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gwTmp1 As String

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOk_Click()
Dim blnUser As Boolean
Dim strUser As String

    blnUser = False

    If Trim(txtID.Text) = "" Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtID.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtPW.Text) <> NUAPI.PW Then
'        blnUser = False
'    Else
'        blnUser = True
'    End If
     
    If Trim(txtUserName.Text) = "" Then
        blnUser = False
    Else
        blnUser = True
    End If
     
    If blnUser = False Then
        lblErr = "* ��й�ȣ�� ��ġ���� �ʽ��ϴ�."
        'txtID.Text = ""
        txtID.SetFocus
    Else
        If chkPW.Value = 1 Then
            Call WritePrivateProfileString("Assay", "SAVEPW", "1", App.Path & "\Interface.ini")
            Call WritePrivateProfileString("Assay", "UID", txtID.Text, App.Path & "\Interface.ini")
            Call WritePrivateProfileString("Assay", "PW", txtPW.Text, App.Path & "\Interface.ini")
        Else
            Call WritePrivateProfileString("Assay", "SAVEPW", "0", App.Path & "\Interface.ini")
            'Call WritePrivateProfileString("Assay", "UID", txtID.Text, App.Path & "\Interface.ini")
            Call WritePrivateProfileString("Assay", "UID", "", App.Path & "\Interface.ini")
            'Call WritePrivateProfileString("Assay", "PW", "", App.Path & "\Interface.ini")
        End If
    
        lblErr = ""
        gIFUser = Trim(txtID.Text)
        frmInterface.StatusBar1.Panels(1).Text = gIFUser & " " & strUser
        frmInterface.Show 0
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    lblErr = ""
    
    GetSetup
    
    lblHospNm.Caption = App.ProductName
    lblEquipName.Caption = App.EXEName
    
    imgNet1.ZOrder 0
'    Timer1.Interval = 500
'    Timer1.Enabled = True

    txtID.Text = NUAPI.UID
    txtPW.Text = NUAPI.PW
    chkPW.Value = NUAPI.SAVEPW
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    End
End Sub


Private Sub lblCancel_Click()

'    Unload Me
    End
    
End Sub

Private Sub lblCommit_Click()
'Dim lsWK As Integer
Dim blnUser As Boolean

    blnUser = False

    If Trim(txtID.Text) = "" Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtID.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtPW.Text) = "" Then
'        lblErr = "* ��й�ȣ�� �Է��ϼ���."
'        txtPW.SetFocus
'        Exit Sub
'    End If
    
    'blnUser = GetUser(Trim(txtID.Text), Trim(txtPW.Text))
    gUserID = Trim(txtID.Text)
'    gDB_Parm.User = gUserID
     
'    If blnUser = False Then
'        lblErr = "* ���̵� �н����尡 ��ġ���� �ʽ��ϴ�."
'        txtPW.Text = ""
'        txtID.Text = ""
'        txtID.SetFocus
'    Else
        lblErr = ""
        'frmInterface.lblUser.Caption = gUserID
        gIFUser = Trim(txtID.Text)
        frmInterface.StatusBar1.Panels(1).Text = gIFUser

        frmInterface.Show 0
        Unload Me
'    End If
    
    
    
'    If Trim(gWorker_Info.WK_PW) = Trim(txtPW.Text) And Trim(gWorker_Info.WK_ID) = Trim(txtID.Text) Then
'        lblErr = ""
'        frmInterface.lblUser.Caption = "����� : " & gWorker_Info.WK_NM
'        frmInterface.Show 0
'        Me.Hide
'
'    Else
'        lblErr = "* ��й�ȣ�� Ȯ���ϼ���."
'        txtPW.Text = ""
'        txtPW.SetFocus
'    End If
End Sub

Private Sub Timer1_Timer()
    DoEvents

    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call lblCancel_Click
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        Call txtID_LostFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtID_LostFocus()
    Dim Ret As Boolean
    Dim sHtmlLine
    Dim sUrl, sPost, sParam As String
    Dim sRcvData, sData As String
        
On Error GoTo ErrorTrap

    If ActiveControl.Name = "cmdOk" Then Exit Sub
    
    If ActiveControl.Name = "cmdCancel" Then Exit Sub
     

'GoTo RST

    If txtID.Text = "" Then
        MsgBox "�α׿� ID�� �Է��ϼ���. ", vbOKOnly + vbExclamation
        txtID.SetFocus
        Exit Sub
    End If

    labMsg.Caption = "����Ÿ ���̽��� ������ ...."
    Screen.MousePointer = vbArrowHourglass
    
    'http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00104&business_id=lis&ex_interface=12345678|012&
             sParam = "submit_id=TRLII00104&"
    sParam = sParam & "business_id=lis&"
    sParam = sParam & "ex_interface=" & Trim(txtID.Text) & "|" & NUAPI.HOSPCD & "&"  '�����ID|����ڵ�
    sParam = sParam & "instcd=" & NUAPI.INSTCD & "&" '����ڵ�
    sParam = sParam & "userid=" & Trim(txtID.Text) '�����ID
    
'''                 sParam = "submit_id=TRLII00104&"
'''        sParam = sParam & "business_id=lis&"
'''        sParam = sParam & "ex_interface=" & Trim(txtUserID.Text) & "|012&" '�����ID|����ڵ�
'''        sParam = sParam & "instcd=012&" '����ڵ�
'''        sParam = sParam & "userid=" & Trim(txtUserID.Text) '�����ID
    
    'SetRawData "[�α���1]" & NUAPI.APIURL & sParam
    
    sRcvData = OpenURLWithIE2(NUAPI.APIURL & sParam, Inet1)
            
    Call SetSQLData("�α���", sRcvData)
            
    
    If InStr(1, sRcvData, "<?xml version") > 0 Then
        gwTmp1 = ""
    End If
    
    gwTmp1 = gwTmp1 & sRcvData
                
    'sData = mGetP(gwTmp1, 1, "usernm")
    sData = mGetP(mGetP(mGetP(gwTmp1, 2, "usernm"), 2, ">"), 1, "<")

    Screen.MousePointer = vbDefault
    labMsg.Caption = "����Ÿ ���̽��� ���� �Ǿ����ϴ�."

    If sData = "" Then
        MsgBox "��ϵ��� ���� ID�Դϴ�. �α��� ID�� Ȯ���ϼ���. ", vbOKOnly + vbExclamation
        With txtID
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Else
        Timer1.Enabled = False
        With CurrUser
            .CuUserID = Trim(txtID.Text)
            .CuUserNM = sData
            .CuUserPW = ""
            txtUserName = .CuUserNM
            cmdOK.SetFocus
        End With
    End If
        
    Exit Sub
    
ErrorTrap:
    labMsg.Caption = "����� ID�� Ȯ���ϼ���"
    
End Sub

Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtPW.Text) = "" Then
            lblErr = "* ��й�ȣ�� �Է��ϼ���."
            txtPW.SetFocus
            Exit Sub
        Else
            lblErr = ""
            lblCommit_Click
            
        End If
        
    End If
End Sub


