VERSION 5.00
Begin VB.Form frmHospInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "병원정보 설정"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "frmHospInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
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
      Top             =   2610
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
      Top             =   2160
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
      Top             =   1710
      Width           =   1785
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   555
      Left            =   2910
      TabIndex        =   19
      Top             =   4320
      Width           =   2565
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "미사용"
         Height          =   315
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
         Top             =   90
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
      Top             =   3930
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
      Top             =   3480
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
      Top             =   3060
      Width           =   1785
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  '아래 맞춤
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   6960
      TabIndex        =   11
      Top             =   5445
      Width           =   6960
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
      ScaleWidth      =   6960
      TabIndex        =   9
      Top             =   0
      Width           =   6960
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "의료기관정보 설정"
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
      Top             =   2160
      Width           =   765
   End
   Begin VB.TextBox txtPartCd 
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
      Top             =   2625
      Width           =   765
   End
   Begin VB.TextBox txtHospCd 
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
      TabIndex        =   0
      Top             =   1710
      Width           =   765
   End
   Begin VB.TextBox txtMachCd 
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
      Top             =   3060
      Width           =   765
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
      Top             =   4440
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
      Top             =   3990
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   6
      Left            =   690
      Picture         =   "frmHospInfo.frx":32A3
      Top             =   4440
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
      Top             =   3540
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   5
      Left            =   690
      Picture         =   "frmHospInfo.frx":368D
      Top             =   3990
      Width           =   150
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   4
      Left            =   690
      Picture         =   "frmHospInfo.frx":3A77
      Top             =   3540
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
      Top             =   2655
      Width           =   1770
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   2
      Left            =   690
      Picture         =   "frmHospInfo.frx":3E61
      Top             =   2625
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
      Top             =   2220
      Width           =   1770
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   690
      Picture         =   "frmHospInfo.frx":424B
      Top             =   2190
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
      Top             =   1770
      Width           =   1380
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   690
      Picture         =   "frmHospInfo.frx":4635
      Top             =   1740
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
      Top             =   3120
      Width           =   1380
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   3
      Left            =   690
      Picture         =   "frmHospInfo.frx":4A1F
      Top             =   3090
      Width           =   150
   End
End
Attribute VB_Name = "frmHospInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
End Sub

Private Sub imgMenuCancel_Click()
    End
End Sub

Private Sub imgMenuInsert_Click()

    Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "HOSPNM", txtHospNm.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "LABCD", txtLabCd.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "LABNM", txtLabNm.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "PARTCD", txtPartCd.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "PARTNM", txtPartNm.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "MACHCD", txtMachCd.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "MACHNM", txtMachNm.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "USERID", txtUserID.Text, App.PATH & "\OKSOFT.ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtUserNm.Text, App.PATH & "\OKSOFT.ini")
    If optLoginUse(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "LOGINYN", "Y", App.PATH & "\OKSOFT.ini")
    Else
        Call WritePrivateProfileString("HOSP", "LOGINYN", "N", App.PATH & "\OKSOFT.ini")
    End If
    
    GetSetup
    
    Unload Me

    Call Main

End Sub

