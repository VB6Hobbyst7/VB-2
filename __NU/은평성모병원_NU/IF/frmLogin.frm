VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  '���� ����
   Caption         =   " �α���"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5670
   StartUpPosition =   1  '������ ���
   Begin VB.PictureBox Picture2 
      Align           =   2  '�Ʒ� ����
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   5670
      TabIndex        =   4
      Top             =   2235
      Width           =   5670
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
         Left            =   2550
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   450
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         Height          =   405
         Left            =   4080
         MaskColor       =   &H00000000&
         TabIndex        =   14
         Top             =   900
         Width           =   825
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ȯ��"
         Height          =   405
         Left            =   3180
         MaskColor       =   &H00000000&
         TabIndex        =   13
         Top             =   900
         Width           =   855
      End
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "���̵� ����"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4260
         TabIndex        =   2
         Top             =   510
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Left            =   150
         Top             =   600
      End
      Begin VB.TextBox txtID 
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
         Left            =   2550
         TabIndex        =   0
         Top             =   60
         Width           =   1575
      End
      Begin VB.TextBox txtUserName 
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
         Left            =   4140
         TabIndex        =   1
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label lblUserNm 
         BackStyle       =   0  '����
         Caption         =   "ȫ�浿"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1980
         TabIndex        =   10
         Top             =   1050
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblID 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���̵� :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1290
         TabIndex        =   6
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lblPW 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "��й�ȣ :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1290
         TabIndex        =   5
         Top             =   540
         Width           =   1155
      End
      Begin VB.Image Image2 
         Height          =   1245
         Left            =   0
         Picture         =   "frmLogin.frx":000C
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   5895
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   5670
      TabIndex        =   3
      Top             =   0
      Width           =   5670
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   210
         Top             =   510
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label lblErr 
         BackStyle       =   0  '����
         Caption         =   "* ����� ID�� Password �� �߸��Ǿ����ϴ�."
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   630
         TabIndex        =   15
         Top             =   1710
         Width           =   4515
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmLogin.frx":0F96
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmLogin.frx":10E0
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label labMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����� ID�� �Է� �Ͻʽÿ�."
         Height          =   180
         Left            =   660
         TabIndex        =   12
         Top             =   2010
         Width           =   2205
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmLogin.frx":122A
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label lblPartNm 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '����
         Caption         =   "��ȭ�а˻��"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   3780
         TabIndex        =   11
         Top             =   780
         Width           =   1755
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00400000&
         BorderWidth     =   2
         Height          =   375
         Left            =   390
         Top             =   1320
         Width           =   105
      End
      Begin VB.Label lblMachNm 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '����
         Caption         =   "ABL 800 Basic "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   1260
         Width           =   3975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabNm 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '����
         Caption         =   "���ܰ˻����а�"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   3450
         TabIndex        =   8
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label lblHospNm 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '����
         Caption         =   "�������б� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   3465
      End
      Begin VB.Image Image3 
         Height          =   2100
         Left            =   30
         Picture         =   "frmLogin.frx":1374
         Stretch         =   -1  'True
         Top             =   30
         Width           =   5745
      End
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

    If txtUserName.Text = "" Then
        Call txtPW_LostFocus
    End If
    
    blnUser = False

    If Trim(txtID.Text) = "" Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtID.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPW.Text) = "" Then
        lblErr = "* ��й�ȣ�� �Է��ϼ���."
        txtPW.SetFocus
        Exit Sub
    End If
     
    If Trim(txtUserName.Text) = "" Then
        blnUser = False
    Else
        blnUser = True
    End If
     
    If blnUser = False Then
        lblErr.Caption = "* ��й�ȣ�� ��ġ���� �ʽ��ϴ�."
        'txtID.Text = ""
        txtID.SetFocus
    Else
        If chkSave.Value = 1 Then
            Call WritePrivateProfileString("HOSP", "SAVEPW", "1", App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("HOSP", "USERID", txtID.Text, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("HOSP", "USERNM", txtUserName.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Else
            Call WritePrivateProfileString("HOSP", "SAVEPW", "0", App.PATH & App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("HOSP", "USERID", "", App.PATH & App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("HOSP", "USERNM", txtUserName.Text, App.PATH & "\INI\" & gMACH & ".ini")
        End If
    
        lblErr = ""
        gHOSP.USERID = Trim(txtID.Text)
        'frmInterface.StatusBar1.Panels(1).Text = gIFUser & " " & strUser
        Screen.MousePointer = 0
        frmMain.Show 0
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        End
    End If
    
End Sub

Private Sub Form_Load()

    imgNet1.ZOrder 0
    Timer1.Interval = 500
    Timer1.Enabled = True
    
    Call CtlInitializing
    
'    txtID.SetFocus
    
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


Private Sub txtID_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = vbKeyReturn Then
        'Call txtID_LostFocus
        Call txtPW.SetFocus
        KeyAscii = 0
    End If

End Sub


Public Sub CtlInitializing()
    Dim i           As Integer
    Dim strPW       As String
    Dim strOrgPW    As String
    
    lblHospNm.Caption = gHOSP.HOSPNM
    lblLabNm.Caption = gHOSP.LABNM
    lblPartNm.Caption = gHOSP.PARTNM
    lblMachNm.Caption = gHOSP.MACHNM
    lblErr.Caption = ""
    If gHOSP.SAVEPW = "1" Then
'        If gHOSP.USERPW <> "" Then
'            strPW = Mid(gHOSP.USERPW, 2)
'            strPW = Mid(strPW, 1, Len(strPW) - 2)
'            For i = 1 To Len(strPW) Step 2
'                strOrgPW = strOrgPW & Chr(Mid(strPW, i, 2))
'            Next
            
            chkSave.Value = "1"
            txtID.Text = gHOSP.USERID
            txtUserName.Text = gHOSP.USERNM
'        End If
    End If
    
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call txtPW_LostFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtPW_LostFocus()
    Dim Ret As Boolean
    Dim sHtmlLine
    Dim sUrl, sPost, sParam As String
    Dim sRcvData, sData As String
        
    Screen.MousePointer = 11

On Error GoTo ErrorTrap

    If ActiveControl.NAME = "cmdOk" Then Exit Sub
    
    If ActiveControl.NAME = "cmdCancel" Then Exit Sub
     

    If txtID.Text = "" Then
        MsgBox "�α׿� ID�� �Է��ϼ���. ", vbOKOnly + vbExclamation
        txtID.SetFocus
        Exit Sub
    End If

    If txtPW.Text = "" Then
        MsgBox "��й�ȣ�� �Է��ϼ���. ", vbOKOnly + vbExclamation
        txtPW.SetFocus
        Exit Sub
    End If

    labMsg.Caption = "����Ÿ ���̽��� ������ ...."
    Screen.MousePointer = vbArrowHourglass
    
    '���ٿ��κ���
    'http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00104&business_id=lis&ex_interface=12345678|012&
             
'             sParam = "submit_id=TRLII00104&"
'    sParam = sParam & "business_id=lis&"
'    sParam = sParam & "ex_interface=" & Trim(txtID.Text) & "|" & gHOSP.HOSPCD & "&"  '�����ID|����ڵ�
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"  '����ڵ�
'    sParam = sParam & "userid=" & Trim(txtID.Text) '�����ID
    
    '���ٽ� ��亴��
    'strURL = SERVERIP + "/himed2/.live?submit_id=TRLII00000&business_id=lis&jobkind=E&userid=" + argId + "&instcode=his053&password=" + argPass;
                                       'submit_id=TRLII00000&business_id=lis&jobkind=E&instcode=H1&userid=1password=1

    
'             sParam = "submit_id=TRLII00000&"
'    sParam = sParam & "business_id=lis&"
'    sParam = sParam & "jobkind=E&"
'    sParam = sParam & "instcode=" & gHOSP.HOSPCD & "&"  '����ڵ�
'    sParam = sParam & "userid=" & Trim(txtID.Text)      '�����ID
'    sParam = sParam & "password=" & Trim(txtPW.Text)    '��й�ȣ
'
        
             sParam = "submit_id=TRLII00104&"
    sParam = sParam & "business_id=li&"
    sParam = sParam & "ex_interface=" & Trim(txtID.Text) & "|" & gHOSP.HOSPCD & "&"  '�����ID|����ڵ�
    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"  '����ڵ�
    sParam = sParam & "userid=" & Trim(txtID.Text) '�����ID
        
    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, Inet1)
            
    Call SetSQLData("�α���", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'SetRawData "[�α���]" & sRcvData
    
    If InStr(1, sRcvData, "<?xml version") > 0 Then
        gwTmp1 = ""
    End If
    
    gwTmp1 = gwTmp1 & sRcvData
                
    'MsgBox gwTmp1
    
    sData = mGetP(mGetP(mGetP(gwTmp1, 2, "usernm"), 2, ">"), 1, "<")
    'MsgBox sData
    
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
'RST:
        Timer1.Enabled = False
        txtUserName.Text = sData
        
        With gHOSP
            .USERID = Trim(txtID.Text)
            .USERNM = sData
            cmdOk.SetFocus
        End With
    End If
        
    Screen.MousePointer = 0
    
    Exit Sub
    
ErrorTrap:
    Screen.MousePointer = 0
    labMsg.Caption = "����� ID�� ��й�ȣ�� Ȯ���ϼ���"
    
End Sub
