VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm공용_Login 
   BorderStyle     =   1  '단일 고정
   Caption         =   "로그인"
   ClientHeight    =   3735
   ClientLeft      =   14550
   ClientTop       =   1065
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm공용_Login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6675
   Begin VB.TextBox txtUserPW 
      Height          =   315
      IMEMode         =   3  '사용 못함
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
   Begin VB.Timer tmr명암 
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
      BackStyle       =   0  '투명
      Caption         =   "User PW"
      Height          =   180
      Left            =   4440
      TabIndex        =   14
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "User ID"
      Height          =   180
      Left            =   4440
      TabIndex        =   13
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lbl설명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Interface For Medical Machine"
      BeginProperty Font 
         Name            =   "굴림"
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
      BackStyle       =   0  '투명
      Caption         =   "ComPort Info"
      BeginProperty Font 
         Name            =   "굴림체"
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
      BackStyle       =   0  '투명
      Caption         =   "HIS DataBase Info"
      BeginProperty Font 
         Name            =   "굴림체"
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
      BorderStyle     =   0  '투명
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   180
      Shape           =   3  '원형
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape shpNo2 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '투명
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   180
      Shape           =   3  '원형
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Local DataBase Info"
      BeginProperty Font 
         Name            =   "굴림체"
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
      BorderStyle     =   0  '투명
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   180
      Shape           =   3  '원형
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '투명
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   4860
      Shape           =   3  '원형
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
      BackStyle       =   0  '투명
      Caption         =   "DataBase 접속 중..."
      BeginProperty Font 
         Name            =   "굴림"
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
      BackStyle       =   0  '투명
      Caption         =   "Interface EQ"
      BeginProperty Font 
         Name            =   "굴림"
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
   Begin VB.Label lbl회사이름 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "메디메이트(주)"
      BeginProperty Font 
         Name            =   "굴림체"
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
         Name            =   "굴림"
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
      BackStyle       =   0  '투명
      Caption         =   "Version ?.?"
      BeginProperty Font 
         Name            =   "굴림"
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
      BackStyle       =   0  '투명
      Caption         =   "Copyright ⓒ 2010 Medimate Co., Ltd."
      BeginProperty Font 
         Name            =   "굴림"
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
   Begin VB.Shape Shp반짝이 
      BackColor       =   &H00C000C0&
      BackStyle       =   1  '투명하지 않음
      Height          =   315
      Index           =   0
      Left            =   6300
      Top             =   60
      Width           =   315
   End
End
Attribute VB_Name = "frm공용_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbl명암증가 As Double
Dim dbl명암속도 As Double

Private MMFTP   As New cls공용_FTP
Private MMSFTP  As New cls공용_SFTP

'/폼 투명 효과----------------------------------------------------------------------------------------------------------------------------------------------------------------/
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
'/폼 투명 효과----------------------------------------------------------------------------------------------------------------------------------------------------------------/

'/폼 투명 효과
Private Function MakeLayeredWnd(hwnd As Long) As Long
     Dim WndStyle As Long

     WndStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
     WndStyle = WndStyle Or WS_EX_LAYERED
     MakeLayeredWnd = SetWindowLong(hwnd, GWL_EXSTYLE, WndStyle)
End Function

Public Sub MM_INITIAL()

    '/STEP1.작업모드 Setting(기본은 표준모드로 한 후 HIS DataBase 연결 STEP에서 작업모드를 재 결정한다.
    gstrJobMode = "1" '/작업모드(1.표준(장비연동,HIS연동 가능), 2.임시(장비연동만 가능))
    
    '/STEP2.Local DataBase 연결
    GoSub RTN_LOCALDB_CONNECT
    
    '/STEP3.HIS DataBase 연결
    GoSub RTN_HISDB_CONNECT
    
    '/STEP4.ComPort 인식
    GoSub RTN_EQUIPMENT_INFO

    '/----------ID와 PW받고 처리할 경우 막는다.
'''    '/STEP5.Login 화면 닫기
'''    Unload Me
'''
'''    Call Main
    '/----------ID와 PW받고 처리할 경우 막는다.
Exit Sub

'/----------------------------------------------------------------------------------------------------/

'/STEP1.작업모드 Setting(기본은 표준모드로 한 후 DB연결 STEP에서 작업모드를 재 결정한다.
RTN_LOCALDB_CONNECT:
    '/1.DB ConnectString이 없을 때는 (재)입력하게 한다.(사용자가 입력을 거부하면 작업모드를 "2"(임시모두)로 전환한다.)
    '/2.입력된 DB ConnectString으로 연결이 안될 때는 재 입력하게 한다.(사용자가 입력을 거부하면 작업모드를 "2"(임시모두)로 전환한다.)
    
RTN_REPEAT1:

    lblState = "DataBase 접속 중": DoEvents
    gstrREG_DB_CONSTR = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_DB_INFO, REG_DB_CONSTR)
    '''gstrREG_DB_CONSTR = "Provider=msdaora;Data Source=phis;User Id=Phis_lis;Password=Phis_lis;" '/인천의료원
    If Len(gstrREG_DB_CONSTR) = 0 Then
        lblState = "DataBase 접속 실패!!!": DoEvents
        
        If MsgBox("▶▶▶DataBase Info" & vbCrLf & vbCrLf & _
                  "DataBase Connect String 정보가 올바르지 않습니다." & vbCrLf & _
                  "DataBase Info Setting 을 (재)설정하겠습니까?", vbQuestion + vbYesNo, "질의") = vbNo Then
            
            MsgBox "▶▶▶DataBase Info" & vbCrLf & vbCrLf & _
                   "계속 진행할 경우 Local Image Capture 작업만 가능합니다." & vbCrLf & _
                   "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
            
            gstrJobMode = "2" '/작업모드(1.표준(DB,FTP 연결가능), 2.임시(ImageCapture만 가능))
            
            GoTo DB_JUMP_RTN
        Else
            frm공용_Set_DataBase.Show vbModal
        End If

        GoTo RTN_REPEAT1
    Else
        If OpenDB(gstrREG_DB_CONSTR) = False Then
            lblState = "DataBase 접속 실패!!!": DoEvents
            If MsgBox("▶▶▶DataBase Info" & vbCrLf & vbCrLf & _
                      "DataBase Connect String 정보가 올바르지 않습니다." & vbCrLf & _
                      "DataBase Info Setting 을 (재)설정하겠습니까?", vbQuestion + vbYesNo, "질의") = vbNo Then
                
                MsgBox "▶▶▶DataBase Info" & vbCrLf & vbCrLf & _
                       "계속 진행할 경우 Local Image 작업만 가능합니다." & vbCrLf & _
                       "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
                    
                gstrJobMode = "2" '/작업모드(1.표준(DB,FTP 연결가능), 2.임시(ImageCapture만 가능))
                
                GoTo DB_JUMP_RTN
            Else
                frm공용_Set_DataBase.Show vbModal
            End If

            GoTo RTN_REPEAT1
        Else
            gstrSTAUS_DB = "Y" '/DB 연결상태(Y/N)
            
            shpNo1.FillColor = RGB(0, 0, 255)
            Call CloseDB
        End If
    End If

DB_JUMP_RTN:

Return

'/----------------------------------------------------------------------------------------------------/

RTN_HISDB_CONNECT:
    '/1.FTP Server 연결은 DB가 연결됬을 때만 수행한다.
    '/2.FTP Server가 연결이 안되더라도 프로그램은 수행되야만 한다. 이유는 Local Image 작업이 가능해야하기 때문이다.
    
    If gstrJobMode = "1" Then '/작업모드(1.표준(DB,FTP 연결가능), 2.임시(ImageCapture만 가능))
        lblState = "FTP 접속 중": DoEvents
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
                    lblState = "FTP 접속 실패!!!": DoEvents
                    
                    MsgBox "▶▶▶FTP Info" & vbCrLf & vbCrLf & _
                           "EMR_Image File FTP Information 이 올바르지 않습니다." & vbCrLf & _
                           "계속 진행할 경우 Image Server 와의 연동작업이 불가능합니다." & vbCrLf & _
                           "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
                Else
                    '/FTP 접속 시도
                    If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
                        lblState = "FTP 접속 실패!!!": DoEvents

                        MsgBox "▶▶▶FTP Info" & vbCrLf & vbCrLf & _
                               "EMR_Image File FTP Information 이 올바르지 않습니다." & vbCrLf & _
                               "계속 진행할 경우 Image Server 와의 연동작업이 불가능합니다." & vbCrLf & _
                               "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
                    Else
                        gstrSTAUS_FTP = "Y" '/FTP 연결상태(Y/N)

                        shpNo2.FillColor = RGB(0, 0, 255)
                        '/FTP 접속 해제
                        '''Call MMSFTP.CloseConnection
                    End If
                End If
            Else
                lblState = "FTP 접속 실패!!!": DoEvents
                MsgBox "▶▶▶FTP Info" & vbCrLf & vbCrLf & _
                       "EMR_Image File FTP Information 정보가 없습니다." & vbCrLf & _
                       "계속 진행할 경우 Image Server 와의 연동작업이 불가능합니다." & vbCrLf & _
                       "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
            End If
        
            Call CloseDB
        Else
            lblState = "FTP 접속 실패!!!": DoEvents
            MsgBox "▶▶▶FTP Info" & vbCrLf & vbCrLf & _
                   "DataBase가 연결되지 않아 EMR_Image File FTP Information 정보를 인식할 수 없습니다." & vbCrLf & _
                   "계속 진행할 경우 Image Server 와의 연동작업이 불가능합니다." & vbCrLf & _
                   "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
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

    lblState = "대상 의료장비 인식 중": DoEvents

    '/대상 의료장비 정보(레지스터) 가져오기
    strEQCD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQCD)
    strEQSEQ = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQSEQ)

    If Len(strEQCD) = 0 Or Len(strEQSEQ) = 0 Then '/대상 의료장비 정보가 없을 때...
        lblState = "대상 의료장비 없음!!!": DoEvents
        
        If gstrJobMode = "1" Then '/작업모드(1.표준(DB,FTP 연결가능), 2.임시(ImageCapture만 가능))
            If MsgBox("EMR Interface Medical Equipment 정보가 없습니다." & vbCrLf & _
                      "Medical Equipment Info Setting 을 (재)설정하겠습니까?", vbQuestion + vbYesNo, "질의") = vbNo Then
                
                MsgBox "▶▶▶Client Info" & vbCrLf & vbCrLf & _
                       "의료장비가 설정되지 않았습니다." & vbCrLf & _
                       "초기 프로그램 가동 시 DataBase 가 연결된 상태에서 작업할 의료장비를 선택해야합니다." & vbCrLf & _
                       "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다." & vbCrLf & vbCrLf & _
                       "프로그램을 종료합니다.", vbInformation, "프로그램 종료"
                End
            Else
                frm공용_Set_Equipment_List.Show vbModal '/해당 폼은 DB 연결가능 시 실행.
            End If
    
            GoTo RTN_REPEAT3 '/설정된 대상 의료장비 정보 재 인식
        Else
            '/대상 의료장비 정보가 없는 상황에서 DB연결이 안되있다면 프로그램을 실행할 수 없다.
            MsgBox "▶▶▶Client Info" & vbCrLf & vbCrLf & _
                   "의료장비가 설정되지 않았습니다." & vbCrLf & _
                   "초기 프로그램 가동 시 DataBase 가 연결된 상태에서 작업할 의료장비를 선택해야합니다." & vbCrLf & _
                   "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다." & vbCrLf & vbCrLf & _
                   "프로그램을 종료합니다.", vbInformation, "프로그램 종료"
            End
        End If
    Else '/대상 의료장비 정보가 있을 때...
        If InStr(strEQCD, ",") = 0 Then '/설정된 장비가 1대 이면...
            If gstrJobMode = "1" Then '/표준모드면...
                Call GET_EQUIPMENT_INFO(strEQCD, strEQSEQ)
                
                If gtypEQ_INFO.EQUIPCODE = "" Then
                    If MsgBox("레지스터에 설정된 대상 의료장비 정보가 DataBase에 없습니다." & vbCrLf & _
                              "Medical Equipment Info Setting 을 (재)설정하겠습니까?", vbQuestion + vbYesNo, "질의") = vbNo Then
                        
                        MsgBox "▶▶▶Client Info" & vbCrLf & vbCrLf & _
                               "장비설정이 올바르지 않습니다." & vbCrLf & vbCrLf & _
                               "프로그램을 종료합니다.", vbInformation, "프로그램 종료"
                        End
                    Else
                        '/Register에 설정된 장비코드 및 장비SEQ가 DataBase 에 없을 때는
                        '/장비List를 보인 후 재 Setting하게 한다.
                        frm공용_Set_Equipment_List.Show vbModal
                    End If

                    GoTo RTN_REPEAT3
                End If
            Else '/임시모드면...
                '/대상 의료장비 정보(레지스터) 광역변수 Setting
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
        Else '/설정된 장비가 2대 이상 이면...
            frm공용_Set_Equipment.Show vbModal
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
    lbl설명 = App.Comments
    lbl회사이름 = App.CompanyName
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
    
    '/화면 명암관련 루틴----------------------------------------------------------------------------------------------------/
    Me.Visible = False
    tmr명암.Enabled = False
    
    MakeLayeredWnd Me.hwnd
    SetLayeredWindowAttributes Me.hwnd, 0, 255 * (0), LWA_ALPHA
    
    dbl명암속도 = 0.01
    
    tmr명암.Enabled = True
    tmr명암.Interval = 2
    
    Me.Visible = True
        
    DoEvents
    DoEvents
    DoEvents
    '/화면 명암관련 루틴----------------------------------------------------------------------------------------------------/
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MMFTP = Nothing
    Call CloseDB
    Set frm공용_Login = Nothing
End Sub

Private Sub tmr명암_Timer()
    dbl명암증가 = dbl명암증가 + dbl명암속도 '0.03
    
    If dbl명암증가 > 1 Then
        dbl명암증가 = 1
    
        tmr명암.Enabled = False
        tmr명암.Interval = 0
    
        Call MM_INITIAL
        
        lblUserID.Visible = True
        lblUserPW.Visible = True
        txtUserID.Visible = True
        txtUserPW.Visible = True
        txtUserID.SetFocus
    Else
        MakeLayeredWnd Me.hwnd
        SetLayeredWindowAttributes Me.hwnd, 0, 255 * (dbl명암증가), LWA_ALPHA
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
            MsgBox "User ID를 (재)입력하십시오!", vbCritical, "로그인 실패"
            txtUserID.SetFocus
            Exit Sub
        End If
        If Trim(txtUserPW) = "" Then
            MsgBox "User PW를 (재)입력하십시오!", vbCritical, "로그인 실패"
            txtUserPW.SetFocus
            Exit Sub
        End If
        
        If OpenDB(gstrREG_DB_CONSTR) = True Then
            gstrQuy = "SELECT USER_ID, USER_NM, PWD "
            gstrQuy = gstrQuy & vbCrLf & "  FROM TZUSERMSTN " '/HIS 사용자마스터
            gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID = '" & Trim(txtUserID) & "' "
            If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
                        
            If Not ADR Is Nothing Then
                If Trim(txtUserPW) = Trim(ADR!PWD & "") Then
                    gtypUSER.USERID = Trim(ADR!USER_ID & "")
                    gtypUSER.USERNM = Trim(ADR!USER_NM & "")
                    gtypUSER.USERPW = Trim(ADR!PWD & "")
                    
                    ADR.Close: Set ADR = Nothing
                
                    '/STEP5.Login 화면 닫기
                    Unload Me
                
                    Call Main
                Else
                    ADR.Close: Set ADR = Nothing
                
                    MsgBox "User PW가 맞지 않습니다!", vbCritical, "로그인 실패": Exit Sub
                End If
            Else
                MsgBox "등록되지 않은 ID 입니다!", vbCritical, "로그인 실패": Exit Sub
            End If
        
            Call CloseDB
        End If
    End If
End Sub
