VERSION 5.00
Begin VB.Form frmHospInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "frmHospInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   915
      Left            =   4230
      TabIndex        =   32
      Top             =   990
      Visible         =   0   'False
      Width           =   1215
      Begin VB.OptionButton optDB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용안함"
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   35
         Top             =   600
         Width           =   1125
      End
      Begin VB.OptionButton optDB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Oracle"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Top             =   30
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optDB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MS-SQL"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   33
         Top             =   300
         Width           =   1125
      End
   End
   Begin VB.TextBox txtColWidth 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   30
      Top             =   4320
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   345
      Left            =   2910
      TabIndex        =   27
      Top             =   4770
      Width           =   2565
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "팝업"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   29
         Top             =   30
         Width           =   1125
      End
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "메인"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   28
         Top             =   30
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "로그기록"
      Height          =   345
      Left            =   690
      TabIndex        =   25
      Top             =   5760
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdDBSet 
      Caption         =   "서버 DB"
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   5730
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdLocalDBSet 
      Caption         =   "로컬 DB"
      Height          =   375
      Left            =   2910
      TabIndex        =   23
      Top             =   5730
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPartNm 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      TabIndex        =   22
      Top             =   2740
      Width           =   1785
   End
   Begin VB.TextBox txtLabNm 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      TabIndex        =   21
      Top             =   2345
      Width           =   1785
   End
   Begin VB.TextBox txtHospNm 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      TabIndex        =   20
      Top             =   1950
      Width           =   1785
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   345
      Left            =   2910
      TabIndex        =   19
      Top             =   5220
      Visible         =   0   'False
      Width           =   2565
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "미사용"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   90
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   30
         Width           =   1125
      End
   End
   Begin VB.TextBox txtUserNm 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   6
      Top             =   3925
      Width           =   2565
   End
   Begin VB.TextBox txtUserID 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   5
      Top             =   3530
      Width           =   2565
   End
   Begin VB.TextBox txtMachNm 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3690
      TabIndex        =   4
      Top             =   3135
      Width           =   1785
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  '아래 맞춤
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   6375
      TabIndex        =   11
      Top             =   6195
      Width           =   6375
      Begin VB.Image imgMenuInsert 
         Height          =   375
         Left            =   1950
         Picture         =   "frmHospInfo.frx":000C
         Top             =   300
         Width           =   1725
      End
      Begin VB.Image imgMenuCancel 
         Height          =   375
         Left            =   3780
         Picture         =   "frmHospInfo.frx":0E08
         Top             =   300
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   6375
      TabIndex        =   9
      Top             =   0
      Width           =   6375
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "병원정보 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   540
         Width           =   3135
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmHospInfo.frx":1B60
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.TextBox txtLabCd 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   1
      Top             =   2345
      Width           =   765
   End
   Begin VB.TextBox txtPartCd 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   2
      Top             =   2740
      Width           =   765
   End
   Begin VB.TextBox txtHospCd 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   0
      Top             =   1950
      Width           =   765
   End
   Begin VB.TextBox txtMachCd 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   3
      Top             =   3135
      Width           =   765
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   8
      Left            =   690
      Picture         =   "frmHospInfo.frx":32A3
      Top             =   5280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사명 넓이 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   5
      Left            =   1530
      TabIndex        =   31
      Top             =   4425
      Width           =   1275
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   7
      Left            =   690
      Picture         =   "frmHospInfo.frx":368D
      Top             =   4862
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "워크조회 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   1800
      TabIndex        =   26
      Top             =   4860
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "로그인 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   1995
      TabIndex        =   18
      Top             =   5340
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "당당자명 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   1800
      TabIndex        =   17
      Top             =   4005
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   6
      Left            =   690
      Picture         =   "frmHospInfo.frx":3A77
      Top             =   4446
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "당당자ID : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   1800
      TabIndex        =   16
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   5
      Left            =   690
      Picture         =   "frmHospInfo.frx":3E61
      Top             =   4030
      Width           =   150
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   4
      Left            =   690
      Picture         =   "frmHospInfo.frx":424B
      Top             =   3614
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사파트 코드/명 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   10
      Left            =   1020
      TabIndex        =   15
      Top             =   2834
      Width           =   1770
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   2
      Left            =   690
      Picture         =   "frmHospInfo.frx":4635
      Top             =   2782
      Width           =   150
   End
   Begin VB.Label 사용자명 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "사용부서 코드/명 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   9
      Left            =   1035
      TabIndex        =   14
      Top             =   2422
      Width           =   1770
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   690
      Picture         =   "frmHospInfo.frx":4A1F
      Top             =   2366
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "병원 코드/명 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   8
      Left            =   1410
      TabIndex        =   13
      Top             =   2010
      Width           =   1380
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   690
      Picture         =   "frmHospInfo.frx":4E09
      Top             =   1950
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "장비 코드/명 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   1425
      TabIndex        =   12
      Top             =   3210
      Width           =   1380
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   3
      Left            =   690
      Picture         =   "frmHospInfo.frx":51F3
      Top             =   3198
      Width           =   150
   End
End
Attribute VB_Name = "frmHospInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub cmdDBSet_Click()
    
    If gDBTYPE = "99" Then
        
        Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        If gDBTYPE = "1" Then
            frmDB_Oracle.Show vbModal
        ElseIf gDBTYPE = "2" Then
            frmDB_MSSQL.Show vbModal
        Else
            MsgBox App.PATH & "\OKSOFT.ini 파일에서" & vbNewLine & vbNewLine & "DBTYPE을 먼저 설정하세요 ", vbOKOnly + vbInformation, "DB TYPE 설정"
        End If
    End If
    
End Sub

Private Sub cmdLocalDBSet_Click()
    
    frmDB_Local.Show vbModal

End Sub

Private Sub Form_Load()
    
    Call CtlInitializing
    
End Sub

Public Sub CtlInitializing()
     
    txtHospCd.Text = gHOSP.HOSPCD
    txtHospNm.Text = gHOSP.HOSPNM
    txtLabCd.Text = gHOSP.LABCD
    txtLabNm.Text = gHOSP.LABNM
    txtPartCd.Text = gHOSP.PARTCD
    txtPartNm.Text = gHOSP.PARTNM
    txtMachCd.Text = gHOSP.MACHCD
    txtMachNm.Text = gHOSP.MACHNM
    txtUserID.Text = gHOSP.USERID
    txtUserNm.Text = gHOSP.USERNM
    If gHOSP.LOGINYN = "Y" Then
        optLoginUse(1).Value = True
    Else
        optLoginUse(0).Value = True
    End If
    If gHOSP.LOQWRITE = "1" Then
        chkLog.Value = "1"
    Else
        chkLog.Value = "0"
    End If
    
    If gWORKPOS = "P" Then
        optWorkPos(1).Value = True
    Else
        optWorkPos(0).Value = True
    End If
    txtColWidth.Text = gCOLWIDTH
    
    Select Case gDBTYPE
        Case "1": optDB(0).Value = True
        Case "2": optDB(1).Value = True
        Case "99": optDB(2).Value = True
    End Select

End Sub

Private Sub imgMenuCancel_Click()
    Unload Me
    'End
End Sub

Private Sub imgMenuInsert_Click()
    Dim strDBType   As String
    
    Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "HOSPNM", txtHospNm.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "LABCD", txtLabCd.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "LABNM", txtLabNm.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "PARTCD", txtPartCd.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "PARTNM", txtPartNm.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "MACHCD", txtMachCd.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "MACHNM", txtMachNm.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "USERID", txtUserID.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtUserNm.Text, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    If optLoginUse(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "LOGINYN", "N", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("HOSP", "LOGINYN", "Y", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    If chkLog.Value = "1" Then
        Call WritePrivateProfileString("HOSP", "LOGWRITE", "1", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("HOSP", "LOGWRITE", "0", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    If optWorkPos(0).Value = True Then
        Call WritePrivateProfileString("VIEW", "WORKPOS", "M", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    Else
        Call WritePrivateProfileString("VIEW", "WORKPOS", "P", App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    End If
    
    If optDB(0).Value = True Then
        strDBType = "1"
    ElseIf optDB(1).Value = True Then
        strDBType = "2"
    Else
        strDBType = "99"
    End If
    
    Call WritePrivateProfileString("DATABASE", "DBTYPE", strDBType, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
    
    Call LetEqpMaster

    GetSetup
    
    Unload Me

    Call Main

End Sub


Private Sub LetEqpMaster()

    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET EQUIPCD = '" & Trim(txtMachCd.Text) & "'"
                          
    Call DBExec(AdoCn_Local, SQL)

End Sub

Private Sub Label2_DblClick()
    
    If Frame2.Visible = False Then
        Frame2.Visible = True
        Image5(8).Visible = True
        Label1(4).Visible = True
        Frame4.Visible = True
        cmdLocalDBSet.Visible = True
        cmdDBSet.Visible = True
        chkLog.Visible = True
    Else
        Frame2.Visible = False
        Image5(8).Visible = False
        Label1(4).Visible = False
        Frame4.Visible = False
        cmdLocalDBSet.Visible = False
        cmdDBSet.Visible = False
        chkLog.Visible = False
    End If
    
End Sub
