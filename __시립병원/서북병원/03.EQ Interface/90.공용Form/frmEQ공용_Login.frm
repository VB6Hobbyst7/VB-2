VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEQ����_Login 
   BorderStyle     =   1  '���� ����
   Caption         =   "�α���"
   ClientHeight    =   3735
   ClientLeft      =   6135
   ClientTop       =   2595
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ����_Login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6675
   Begin VB.TextBox txtUserPW 
      Height          =   315
      IMEMode         =   3  '��� ����
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      IMEMode         =   8  '����
      Left            =   5160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Timer tmr��� 
      Interval        =   2
      Left            =   4980
      Top             =   2820
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   435
      Left            =   5520
      TabIndex        =   5
      Top             =   2820
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackStyle       =   1
      FullWidth       =   65
      FullHeight      =   29
   End
   Begin VB.Label lblUserPW 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "User PW"
      Height          =   180
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "User ID"
      Height          =   180
      Left            =   4440
      TabIndex        =   12
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Interface For Medical Machine"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   540
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "HIS DataBase Info"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   540
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Shape shpNo2 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '����
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   255
      Left            =   180
      Shape           =   3  '����
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Local DataBase Info"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   540
      TabIndex        =   9
      Top             =   1680
      Width           =   2565
   End
   Begin VB.Shape shpNo1 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '����
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   255
      Left            =   180
      Shape           =   3  '����
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '����
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   255
      Left            =   4860
      Shape           =   3  '����
      Top             =   360
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1560
      X2              =   4920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "DataBase ���� ��..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   900
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Interface EQ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   435
      Index           =   6
      Left            =   1560
      TabIndex        =   7
      Top             =   0
      Width           =   2715
   End
   Begin VB.Label lblȸ���̸� 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�޵����Ʈ(��)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      X1              =   1500
      X2              =   6600
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   1680
      X2              =   6600
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   2400
      X2              =   6600
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   180
      X2              =   6480
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   2160
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   1920
      X2              =   6600
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Hi"
      BeginProperty Font 
         Name            =   "����"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1440
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   1410
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Version ?.?"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5460
      TabIndex        =   3
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Copyright �� 2010 Medimate Co., Ltd."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   3360
      Width           =   3165
   End
   Begin VB.Shape Shp��¦�� 
      BackColor       =   &H00C000C0&
      BackStyle       =   1  '�������� ����
      Height          =   315
      Index           =   0
      Left            =   6300
      Top             =   60
      Width           =   315
   End
End
Attribute VB_Name = "frmEQ����_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbl������� As Double
Dim dbl��ϼӵ� As Double

'/�� ���� ȿ��----------------------------------------------------------------------------------------------------------------------------------------------------------------/
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
'/�� ���� ȿ��----------------------------------------------------------------------------------------------------------------------------------------------------------------/

'/�� ���� ȿ��
Private Function MakeLayeredWnd(hwnd As Long) As Long
     Dim WndStyle As Long

     WndStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
     WndStyle = WndStyle Or WS_EX_LAYERED
     MakeLayeredWnd = SetWindowLong(hwnd, GWL_EXSTYLE, WndStyle)
End Function

Public Sub SUB_MM_INITIAL()
    '/STEP1.Local DataBase ����
    GoSub RTN_LOCALDB_CONNECT
    
    '/STEP2.HIS DataBase ����
    GoSub RTN_HISDB_CONNECT
Exit Sub

'/----------------------------------------------------------------------------------------------------/

'/STEP1.Local DataBase ����
'/- 1.INTERFACE.MDB ȭ���� �������� Ȯ���Ѵ�.
'/    -> ȭ�� ������ ���α׷� ����
'/    -> ȭ�� ������ ���� Step ����
'/- 2.Local DataBase Connect �� �����Ѵ�
'/    -> ������� �� ���α׷� ����
'/    -> ���Ἲ�� �� �⺻���� �ν�

RTN_LOCALDB_CONNECT:
    lblState = "Local DataBase ���� ��": DoEvents
    
    If Dir(App.Path & "\INTERFACE.MDB") = "" Then
        MsgBox "INTERFACE.MDB ȭ���� ����ȭ�ϰ� ���� ������ �������� �ʽ��ϴ�" & vbCrLf & _
               "����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbCritical, "���α׷� ����"
        
        End
    End If

    If ConnDB_LOC = False Then
        lblState = "Local DataBase ���� ����!!!": DoEvents
            
        MsgBox "������Local DataBase Info" & vbCrLf & vbCrLf & _
               "Local DataBase �� ������ �� �����ϴ�." & vbCrLf & _
               "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
    Else
        shpNo1.FillColor = RGB(0, 0, 255)
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM CUS_MST "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then End
        
        If Not ADR_LOC Is Nothing Then
            gtypHIS_CNN_INFO.ID = Trim(ADR_LOC!HISDB_ID & "")
            gtypHIS_CNN_INFO.PW = Trim(ADR_LOC!HISDB_PW & "")
            gtypHIS_CNN_INFO.SV = Trim(ADR_LOC!HISDB_SERVER & "")
            gtypHIS_CNN_INFO.DBNM = Trim(ADR_LOC!HISDB_DBNM & "") '/DBNM ��(SQL Server �� ���)
            gtypHIS_CNN_INFO.TYPE = Trim(ADR_LOC!HISDB_TYPE & "") '/DB ����
        
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_CONF "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then End
        
        If Not ADR_LOC Is Nothing Then
            gtypEQ_INFO.SERIALPORT = Trim(ADR_LOC!SERIALPORT & "")
            gtypEQ_INFO.SERIALBAUD = Trim(ADR_LOC!SERIALBAUD & "")
            gtypEQ_INFO.SERIALDATABIT = Trim(ADR_LOC!SERIALDATABIT & "")
            gtypEQ_INFO.SERIALSTARTBIT = Trim(ADR_LOC!SERIALSTARTBIT & "")
            gtypEQ_INFO.SERIALSTOPBIT = Trim(ADR_LOC!SERIALSTOPBIT & "")
            gtypEQ_INFO.SERIALPARITY = Trim(ADR_LOC!SERIALPARITY & "")
            gtypEQ_INFO.SERIALRTS = Trim(ADR_LOC!SERIALRTS & "")
            gtypEQ_INFO.SERIALDTR = Trim(ADR_LOC!SERIALDTR & "")
            gtypEQ_INFO.WORKLISTGB = Trim(ADR_LOC!WORKLISTGB & "")
            gtypEQ_INFO.AUTOGB = Trim(ADR_LOC!AUTOGB & "")
        
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        Call CloseDB_LOC
    End If
Return

'/----------------------------------------------------------------------------------------------------/

'/STEP2.HIS DataBase ����
'/- 1.Local DataBase Connect �� �����Ѵ�
'/    -> ������� �� ���α׷� ����
'/    -> ���Ἲ�� �� Login Process ����

RTN_HISDB_CONNECT:
    lblState = "HIS Database ���� ��": DoEvents

    If ConnDB_HIS = True Then
        shpNo2.FillColor = RGB(0, 0, 255)
        
        Call CloseDB_HIS
    Else
        lblState = "HIS Database ����!!!": DoEvents
        
        If MsgBox("������HIS DataBase Info" & vbCrLf & vbCrLf & _
                  "HIS DB Connection Information�� �ùٸ��� �ʽ��ϴ�." & vbCrLf & _
                  "(��)�����ϰڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbNo Then
            
            MsgBox "������HIS DataBase Info" & vbCrLf & vbCrLf & _
                   "HIS DataBase �� ������ �� �����ϴ�." & vbCrLf & _
                   "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
                
            '/ID�� ��ȣ�� �ڻ� ���� ������ �� ��� End�� ���� �Ʒ� �ּ��� Ǭ��.
            '''MsgBox "������HIS DataBase Info" & vbCrLf & vbCrLf & _
                   "��� ������ ��� �Ϻ� ����� ���ѵ˴ϴ�." & vbCrLf & _
                   "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
                
            End
            '''GoTo RTN_HISDB_CONNECT_SKIP'/ID�� ��ȣ�� �ڻ� ���� ������ �� ��� End�� ���� �� ������ Ǭ��.
        Else
            gstrArgTemp1 = "HIS": frmEQ����_Set_DB.Show vbModal
        End If

        GoTo RTN_HISDB_CONNECT
        
RTN_HISDB_CONNECT_SKIP:
    
    End If
Return
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '/�ߺ����� ���� ��ƾ----------------------------------------------------------------------------------------------------/
'    If PrevInstance Then
'        MsgBox "���α׷��� �̹� �������Դϴ�", vbExclamation, "�̹� ������"
'        End
'    End If
    '/�ߺ����� ���� ��ƾ----------------------------------------------------------------------------------------------------/

    Me.Height = 4215
    Me.Width = 6795
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    txtUserID = ""
    txtUserPW = ""
    
    lblState = ""
    lbl���� = "Interface For " & App.FileDescription
    
    lblȸ���̸� = App.CompanyName
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    shpNo1.FillColor = RGB(255, 0, 0)
    shpNo2.FillColor = RGB(255, 0, 0)
    
    
On Error Resume Next
    Animation1.Open App.Path & "\Login1.avi"
On Error GoTo 0

    DoEvents
    DoEvents
    DoEvents
    
    '/ȭ�� ��ϰ��� ��ƾ----------------------------------------------------------------------------------------------------/
    Me.Visible = False
    tmr���.Enabled = False
    
    MakeLayeredWnd Me.hwnd
    SetLayeredWindowAttributes Me.hwnd, 0, 255 * (0), LWA_ALPHA
    
    dbl��ϼӵ� = 0.01
    
    tmr���.Enabled = True
    tmr���.Interval = 2
    
    Me.Visible = True
        
    DoEvents
    DoEvents
    DoEvents
    '/ȭ�� ��ϰ��� ��ƾ----------------------------------------------------------------------------------------------------/
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Call CloseDB_HIS
    Call CloseDB_ETC
    
    Set frmEQ����_Login = Nothing
End Sub

Private Sub tmr���_Timer()
    dbl������� = dbl������� + dbl��ϼӵ� '0.03
    
    If dbl������� > 1 Then
        dbl������� = 1
    
        tmr���.Enabled = False
        tmr���.Interval = 0
    
        Call SUB_MM_INITIAL
        
        'txtUserID = "800042"
        'txtUserPW = "1"
        
        lblState = "ID �� Password �� �Է��Ͻʽÿ�!": DoEvents
        txtUserID.SetFocus
    Else
        MakeLayeredWnd Me.hwnd
        SetLayeredWindowAttributes Me.hwnd, 0, 255 * (dbl�������), LWA_ALPHA
    End If
End Sub

Private Sub txtUserID_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtUserPW_GotFocus()
    
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtUserPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtUserID) = "" Then
            MsgBox "User ID�� (��)�Է��Ͻʽÿ�!", vbCritical, "�α��� ����"
            txtUserID.SetFocus
            Exit Sub
        End If
        If Trim(txtUserPW) = "" Then
            MsgBox "User PW�� (��)�Է��Ͻʽÿ�!", vbCritical, "�α��� ����"
            txtUserPW.SetFocus
            Exit Sub
        End If
        
        With gtypUSER
            .USERID = "" '/�����ID
            .USERNM = "" '/����ڸ�
            .USERPW = "" '/�����PW
        End With
        
'''        '/----------------------------------------------------------------------------------------------------/
'''        '/�⺻ Login Pass �κ�
'''        '/----------------------------------------------------------------------------------------------------/
'''        If ConnDB_LOC = True Then
'''            gstrQuy = "SELECT USER_ID, USER_PW, USER_NM "
'''            gstrQuy = gstrQuy & vbCrLf & "  FROM USER_MST " '/HIS ����ڸ����� ���̺�(���Ϻ���)
'''            gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID = '" & Trim(txtUserID) & "' "
'''            If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
'''
'''            If Not ADR_LOC Is Nothing Then
'''                If Trim(txtUserPW) = Trim(ADR_LOC!USER_PW & "") Then
'''                    gtypUSER.USERID = Trim(ADR_LOC!USER_ID & "")
'''                    gtypUSER.USERNM = Trim(ADR_LOC!USER_NM & "")
'''                    gtypUSER.USERPW = Trim(ADR_LOC!USER_PW & "")
'''
'''                    ADR_LOC.Close: Set ADR_LOC = Nothing
'''
'''                    Unload Me
'''
'''                    Call Main
'''                Else
'''                    ADR_LOC.Close: Set ADR_LOC = Nothing
'''
'''                    MsgBox "User PW�� ���� �ʽ��ϴ�!", vbCritical, "�α��� ����": Exit Sub
'''                End If
'''            Else
'''                MsgBox "��ϵ��� ���� ID �Դϴ�!", vbCritical, "�α��� ����": Exit Sub
'''            End If
'''
'''            Call CloseDB_LOC
'''        End If
'''        '/----------------------------------------------------------------------------------------------------/
'''        '/�⺻ Login Pass �κ�
'''        '/----------------------------------------------------------------------------------------------------/
        
        
        '/----------------------------------------------------------------------------------------------------/
        '/������� Login Pass �κ�
        '/----------------------------------------------------------------------------------------------------/
        If ConnDB_HIS = True Then
            gstrQuy = "SELECT UID_1,USERENAME,UPASSWD "
            gstrQuy = gstrQuy & vbCrLf & "FROM USERMASTER"
            gstrQuy = gstrQuy & vbCrLf & "WHERE UID_1 = '" & Trim(txtUserID) & "' "
'            gstrQuy = "SELECT USER_ID, USER_NM, PWD "
'            gstrQuy = gstrQuy & vbCrLf & "  FROM TZUSERMSTN " '/HIS ����ڸ����� ���̺�(���Ϻ���)
'            gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID = '" & Trim(txtUserID) & "' "
            If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End

            If Not ADR_HIS Is Nothing Then
                If Trim(txtUserPW) = Trim(ADR_HIS!UPASSWD & "") Then
                    gtypUSER.USERID = Trim(ADR_HIS!UID_1 & "")
                    gtypUSER.USERNM = Trim(ADR_HIS!USERENAME & "")
                    gtypUSER.USERPW = Trim(ADR_HIS!UPASSWD & "")

                    ADR_HIS.Close: Set ADR_HIS = Nothing

                    Unload Me

                    Call Main
                Else
                    ADR_HIS.Close: Set ADR_HIS = Nothing

                    MsgBox "User PW�� ���� �ʽ��ϴ�!", vbCritical, "�α��� ����": Exit Sub
                End If
            Else
                MsgBox "��ϵ��� ���� ID �Դϴ�!", vbCritical, "�α��� ����": Exit Sub
            End If

            Call CloseDB_HIS
        End If
        '/----------------------------------------------------------------------------------------------------/
        '/������� Login Pass �κ�
        '/----------------------------------------------------------------------------------------------------/
    End If
End Sub
