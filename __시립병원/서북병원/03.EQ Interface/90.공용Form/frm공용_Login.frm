VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm����_Login 
   BorderStyle     =   1  '���� ����
   Caption         =   "�α���"
   ClientHeight    =   3735
   ClientLeft      =   14550
   ClientTop       =   1065
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
   Icon            =   "frm����_Login.frx":0000
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
      TabIndex        =   14
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "User ID"
      Height          =   180
      Left            =   4440
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   540
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ComPort Info"
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
      Index           =   5
      Left            =   540
      TabIndex        =   11
      Top             =   2400
      Width           =   1620
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
   Begin VB.Shape shpNo3 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '����
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   255
      Left            =   180
      Shape           =   3  '����
      Top             =   2400
      Width           =   255
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
Attribute VB_Name = "frm����_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbl������� As Double
Dim dbl��ϼӵ� As Double

Private MMFTP   As New cls����_FTP
Private MMSFTP  As New cls����_SFTP

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

Public Sub MM_INITIAL()

    '/STEP1.�۾���� Setting(�⺻�� ǥ�ظ��� �� �� HIS DataBase ���� STEP���� �۾���带 �� �����Ѵ�.
    gstrJobMode = "1" '/�۾����(1.ǥ��(��񿬵�,HIS���� ����), 2.�ӽ�(��񿬵��� ����))
    
    '/STEP2.Local DataBase ����
    GoSub RTN_LOCALDB_CONNECT
    
    '/STEP3.HIS DataBase ����
    GoSub RTN_HISDB_CONNECT
    
    '/STEP4.ComPort �ν�
    GoSub RTN_EQUIPMENT_INFO

    '/----------ID�� PW�ް� ó���� ��� ���´�.
'''    '/STEP5.Login ȭ�� �ݱ�
'''    Unload Me
'''
'''    Call Main
    '/----------ID�� PW�ް� ó���� ��� ���´�.
Exit Sub

'/----------------------------------------------------------------------------------------------------/

'/STEP1.�۾���� Setting(�⺻�� ǥ�ظ��� �� �� DB���� STEP���� �۾���带 �� �����Ѵ�.
RTN_LOCALDB_CONNECT:
    '/1.DB ConnectString�� ���� ���� (��)�Է��ϰ� �Ѵ�.(����ڰ� �Է��� �ź��ϸ� �۾���带 "2"(�ӽø��)�� ��ȯ�Ѵ�.)
    '/2.�Էµ� DB ConnectString���� ������ �ȵ� ���� �� �Է��ϰ� �Ѵ�.(����ڰ� �Է��� �ź��ϸ� �۾���带 "2"(�ӽø��)�� ��ȯ�Ѵ�.)
    
RTN_REPEAT1:

    lblState = "DataBase ���� ��": DoEvents
    gstrREG_DB_CONSTR = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_DB_INFO, REG_DB_CONSTR)
    '''gstrREG_DB_CONSTR = "Provider=msdaora;Data Source=phis;User Id=Phis_lis;Password=Phis_lis;" '/��õ�Ƿ��
    If Len(gstrREG_DB_CONSTR) = 0 Then
        lblState = "DataBase ���� ����!!!": DoEvents
        
        If MsgBox("������DataBase Info" & vbCrLf & vbCrLf & _
                  "DataBase Connect String ������ �ùٸ��� �ʽ��ϴ�." & vbCrLf & _
                  "DataBase Info Setting �� (��)�����ϰڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbNo Then
            
            MsgBox "������DataBase Info" & vbCrLf & vbCrLf & _
                   "��� ������ ��� Local Image Capture �۾��� �����մϴ�." & vbCrLf & _
                   "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
            
            gstrJobMode = "2" '/�۾����(1.ǥ��(DB,FTP ���ᰡ��), 2.�ӽ�(ImageCapture�� ����))
            
            GoTo DB_JUMP_RTN
        Else
            frm����_Set_DataBase.Show vbModal
        End If

        GoTo RTN_REPEAT1
    Else
        If OpenDB(gstrREG_DB_CONSTR) = False Then
            lblState = "DataBase ���� ����!!!": DoEvents
            If MsgBox("������DataBase Info" & vbCrLf & vbCrLf & _
                      "DataBase Connect String ������ �ùٸ��� �ʽ��ϴ�." & vbCrLf & _
                      "DataBase Info Setting �� (��)�����ϰڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbNo Then
                
                MsgBox "������DataBase Info" & vbCrLf & vbCrLf & _
                       "��� ������ ��� Local Image �۾��� �����մϴ�." & vbCrLf & _
                       "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
                    
                gstrJobMode = "2" '/�۾����(1.ǥ��(DB,FTP ���ᰡ��), 2.�ӽ�(ImageCapture�� ����))
                
                GoTo DB_JUMP_RTN
            Else
                frm����_Set_DataBase.Show vbModal
            End If

            GoTo RTN_REPEAT1
        Else
            gstrSTAUS_DB = "Y" '/DB �������(Y/N)
            
            shpNo1.FillColor = RGB(0, 0, 255)
            Call CloseDB
        End If
    End If

DB_JUMP_RTN:

Return

'/----------------------------------------------------------------------------------------------------/

RTN_HISDB_CONNECT:
    '/1.FTP Server ������ DB�� ������� ���� �����Ѵ�.
    '/2.FTP Server�� ������ �ȵǴ��� ���α׷��� ����Ǿ߸� �Ѵ�. ������ Local Image �۾��� �����ؾ��ϱ� �����̴�.
    
    If gstrJobMode = "1" Then '/�۾����(1.ǥ��(DB,FTP ���ᰡ��), 2.�ӽ�(ImageCapture�� ����))
        lblState = "FTP ���� ��": DoEvents
        gstrFTP_RH = ""
        gstrFTP_RP = ""
        gstrFTP_UN = ""
        gstrFTP_PW = ""
    
        If OpenDB(gstrREG_DB_CONSTR) = True Then
            gstrQuy = "SELECT * "
            gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_HOS "
            If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
            If Not ADR Is Nothing Then
                gstrHOS_CUSCD = Trim(ADR!CUSCD & "")
                gstrFTP_RH = Trim(ADR!RemoteHost & "")
                gstrFTP_RP = Trim(ADR!RemotePort & "")
                gstrFTP_UN = Trim(ADR!USERID & "")
                gstrFTP_PW = Trim(ADR!Password & "")
        
                ADR.Close: Set ADR = Nothing
        
                If Len(gstrFTP_RH) = 0 Or Len(gstrFTP_RP) = 0 Or Len(gstrFTP_UN) = 0 Or Len(gstrFTP_PW) = 0 Then
                    lblState = "FTP ���� ����!!!": DoEvents
                    
                    MsgBox "������FTP Info" & vbCrLf & vbCrLf & _
                           "EMR_Image File FTP Information �� �ùٸ��� �ʽ��ϴ�." & vbCrLf & _
                           "��� ������ ��� Image Server ���� �����۾��� �Ұ����մϴ�." & vbCrLf & _
                           "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
                Else
                    '/FTP ���� �õ�
                    If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
                        lblState = "FTP ���� ����!!!": DoEvents

                        MsgBox "������FTP Info" & vbCrLf & vbCrLf & _
                               "EMR_Image File FTP Information �� �ùٸ��� �ʽ��ϴ�." & vbCrLf & _
                               "��� ������ ��� Image Server ���� �����۾��� �Ұ����մϴ�." & vbCrLf & _
                               "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
                    Else
                        gstrSTAUS_FTP = "Y" '/FTP �������(Y/N)

                        shpNo2.FillColor = RGB(0, 0, 255)
                        '/FTP ���� ����
                        '''Call MMSFTP.CloseConnection
                    End If
                End If
            Else
                lblState = "FTP ���� ����!!!": DoEvents
                MsgBox "������FTP Info" & vbCrLf & vbCrLf & _
                       "EMR_Image File FTP Information ������ �����ϴ�." & vbCrLf & _
                       "��� ������ ��� Image Server ���� �����۾��� �Ұ����մϴ�." & vbCrLf & _
                       "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
            End If
        
            Call CloseDB
        Else
            lblState = "FTP ���� ����!!!": DoEvents
            MsgBox "������FTP Info" & vbCrLf & vbCrLf & _
                   "DataBase�� ������� �ʾ� EMR_Image File FTP Information ������ �ν��� �� �����ϴ�." & vbCrLf & _
                   "��� ������ ��� Image Server ���� �����۾��� �Ұ����մϴ�." & vbCrLf & _
                   "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
        End If
    End If
Return

'/----------------------------------------------------------------------------------------------------/

RTN_EQUIPMENT_INFO:

    Dim strEQCD             As String
    Dim strEQSEQ            As String
    
    Dim strEQCD_Array
    Dim strEQSEQ_Array

RTN_REPEAT3:

    lblState = "��� �Ƿ���� �ν� ��": DoEvents

    '/��� �Ƿ���� ����(��������) ��������
    strEQCD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQCD)
    strEQSEQ = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQSEQ)

    If Len(strEQCD) = 0 Or Len(strEQSEQ) = 0 Then '/��� �Ƿ���� ������ ���� ��...
        lblState = "��� �Ƿ���� ����!!!": DoEvents
        
        If gstrJobMode = "1" Then '/�۾����(1.ǥ��(DB,FTP ���ᰡ��), 2.�ӽ�(ImageCapture�� ����))
            If MsgBox("EMR Interface Medical Equipment ������ �����ϴ�." & vbCrLf & _
                      "Medical Equipment Info Setting �� (��)�����ϰڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbNo Then
                
                MsgBox "������Client Info" & vbCrLf & vbCrLf & _
                       "�Ƿ���� �������� �ʾҽ��ϴ�." & vbCrLf & _
                       "�ʱ� ���α׷� ���� �� DataBase �� ����� ���¿��� �۾��� �Ƿ���� �����ؾ��մϴ�." & vbCrLf & _
                       "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�." & vbCrLf & vbCrLf & _
                       "���α׷��� �����մϴ�.", vbInformation, "���α׷� ����"
                End
            Else
                frm����_Set_Equipment_List.Show vbModal '/�ش� ���� DB ���ᰡ�� �� ����.
            End If
    
            GoTo RTN_REPEAT3 '/������ ��� �Ƿ���� ���� �� �ν�
        Else
            '/��� �Ƿ���� ������ ���� ��Ȳ���� DB������ �ȵ��ִٸ� ���α׷��� ������ �� ����.
            MsgBox "������Client Info" & vbCrLf & vbCrLf & _
                   "�Ƿ���� �������� �ʾҽ��ϴ�." & vbCrLf & _
                   "�ʱ� ���α׷� ���� �� DataBase �� ����� ���¿��� �۾��� �Ƿ���� �����ؾ��մϴ�." & vbCrLf & _
                   "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�." & vbCrLf & vbCrLf & _
                   "���α׷��� �����մϴ�.", vbInformation, "���α׷� ����"
            End
        End If
    Else '/��� �Ƿ���� ������ ���� ��...
        If InStr(strEQCD, ",") = 0 Then '/������ ��� 1�� �̸�...
            If gstrJobMode = "1" Then '/ǥ�ظ���...
                Call GET_EQUIPMENT_INFO(strEQCD, strEQSEQ)
                
                If gtypEQ_INFO.EQUIPCODE = "" Then
                    If MsgBox("�������Ϳ� ������ ��� �Ƿ���� ������ DataBase�� �����ϴ�." & vbCrLf & _
                              "Medical Equipment Info Setting �� (��)�����ϰڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbNo Then
                        
                        MsgBox "������Client Info" & vbCrLf & vbCrLf & _
                               "������� �ùٸ��� �ʽ��ϴ�." & vbCrLf & vbCrLf & _
                               "���α׷��� �����մϴ�.", vbInformation, "���α׷� ����"
                        End
                    Else
                        '/Register�� ������ ����ڵ� �� ���SEQ�� DataBase �� ���� ����
                        '/���List�� ���� �� �� Setting�ϰ� �Ѵ�.
                        frm����_Set_Equipment_List.Show vbModal
                    End If

                    GoTo RTN_REPEAT3
                End If
            Else '/�ӽø���...
                '/��� �Ƿ���� ����(��������) �������� Setting
                gtypEQ_INFO.EQUIPCODE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQCD)
                gtypEQ_INFO.EQUIPNM = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQNM)
                gtypEQ_INFO.EQUIPSEQ = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQSEQ)
                gtypEQ_INFO.DEPTCODE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQPOS)
                gtypEQ_INFO.EQUIPTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQTYPE)
                gtypEQ_INFO.RECEIVETYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_RECEIVETYPE)
                gtypEQ_INFO.EQUIPPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQUIPPORT)
                gtypEQ_INFO.ORDYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ORDYN)
                gtypEQ_INFO.QUERYTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_QUERYTYPE)
                gtypEQ_INFO.ZIPYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ZIPYN)
                gtypEQ_INFO.ZIPNM = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ZIPNM)
                gtypEQ_INFO.SERIALYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALYN)
                gtypEQ_INFO.SERIALPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPORT)
                gtypEQ_INFO.SERIALBAUD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALBAUD)
                gtypEQ_INFO.SERIALDATABIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDATABIT)
                gtypEQ_INFO.SERIALSTARTBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTARTBIT)
                gtypEQ_INFO.SERIALSTOPBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTOPBIT)
                gtypEQ_INFO.SERIALPARITY = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPARITY)
                gtypEQ_INFO.SERIALRTS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALRTS)
                gtypEQ_INFO.SERIALDTR = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDTR)
                gtypEQ_INFO.EQIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQIMGFILEPATH)
                gtypEQ_INFO.FTPIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_FTPIMGFILEPATH)
            End If
        Else '/������ ��� 2�� �̻� �̸�...
            frm����_Set_Equipment.Show vbModal
        End If
        
        shpNo3.FillColor = RGB(0, 0, 255)
    End If
Return
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 4215
    Me.Width = 6795
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
'''    Me.Show
    
    txtUserID = ""
    txtUserPW = ""
    
    lblUserID.Visible = False
    lblUserPW.Visible = False
    txtUserID.Visible = False
    txtUserPW.Visible = False
    
    lblState = ""
    lbl���� = App.Comments
    lblȸ���̸� = App.CompanyName
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    shpNo1.FillColor = RGB(255, 0, 0)
    shpNo2.FillColor = RGB(255, 0, 0)
    shpNo3.FillColor = RGB(255, 0, 0)

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
    Set MMFTP = Nothing
    Call CloseDB
    Set frm����_Login = Nothing
End Sub

Private Sub tmr���_Timer()
    dbl������� = dbl������� + dbl��ϼӵ� '0.03
    
    If dbl������� > 1 Then
        dbl������� = 1
    
        tmr���.Enabled = False
        tmr���.Interval = 0
    
        Call MM_INITIAL
        
        lblUserID.Visible = True
        lblUserPW.Visible = True
        txtUserID.Visible = True
        txtUserPW.Visible = True
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
        
        If OpenDB(gstrREG_DB_CONSTR) = True Then
            gstrQuy = "SELECT USER_ID, USER_NM, PWD "
            gstrQuy = gstrQuy & vbCrLf & "  FROM TZUSERMSTN " '/HIS ����ڸ�����
            gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID = '" & Trim(txtUserID) & "' "
            If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
                        
            If Not ADR Is Nothing Then
                If Trim(txtUserPW) = Trim(ADR!PWD & "") Then
                    gtypUSER.USERID = Trim(ADR!USER_ID & "")
                    gtypUSER.USERNM = Trim(ADR!USER_NM & "")
                    gtypUSER.USERPW = Trim(ADR!PWD & "")
                    
                    ADR.Close: Set ADR = Nothing
                
                    '/STEP5.Login ȭ�� �ݱ�
                    Unload Me
                
                    Call Main
                Else
                    ADR.Close: Set ADR = Nothing
                
                    MsgBox "User PW�� ���� �ʽ��ϴ�!", vbCritical, "�α��� ����": Exit Sub
                End If
            Else
                MsgBox "��ϵ��� ���� ID �Դϴ�!", vbCritical, "�α��� ����": Exit Sub
            End If
        
            Call CloseDB
        End If
    End If
End Sub
