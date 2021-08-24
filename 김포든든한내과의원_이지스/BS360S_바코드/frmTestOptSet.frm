VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmTestOptSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "검사옵션 설정"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   6735
   Icon            =   "frmTestOptSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame6 
      BackColor       =   &H00BF8B59&
      Caption         =   " 로그기록 "
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   6780
      TabIndex        =   22
      Top             =   5190
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox chkLog 
         BackColor       =   &H00BF8B59&
         Caption         =   "로그기록"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3030
         TabIndex        =   23
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00BF8B59&
      Caption         =   " 워크리스트 조회화면 "
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   6780
      TabIndex        =   17
      Top             =   4260
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame Frame5 
         BackColor       =   &H00BF8B59&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   2970
         TabIndex        =   19
         Top             =   360
         Width           =   2235
         Begin VB.OptionButton optWorkPos 
            BackColor       =   &H00BF8B59&
            Caption         =   "메인"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   21
            Top             =   30
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optWorkPos 
            BackColor       =   &H00BF8B59&
            Caption         =   "팝업"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   20
            Top             =   30
            Width           =   1125
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회화면"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   1560
         TabIndex        =   18
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 전달결과 적용 "
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   690
      TabIndex        =   10
      Top             =   4740
      Width           =   5415
      Begin VB.OptionButton optSaveResult 
         BackColor       =   &H00FFFFFF&
         Caption         =   "장비결과"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   3030
         TabIndex        =   16
         Top             =   390
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optSaveResult 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LIS결과"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   4140
         TabIndex        =   15
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "적용결과"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   4
         Left            =   1560
         TabIndex        =   14
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 결과전송 형태 "
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   690
      TabIndex        =   9
      Top             =   3390
      Width           =   5415
      Begin VB.OptionButton optAutoSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "자동"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   3060
         TabIndex        =   13
         Top             =   390
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optAutoSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "수동"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   4170
         TabIndex        =   12
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과전송"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   3
         Left            =   1590
         TabIndex        =   11
         Top             =   435
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 인터페이스 형태 "
      ForeColor       =   &H00404040&
      Height          =   1875
      Left            =   690
      TabIndex        =   0
      Top             =   990
      Width           =   5415
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   3
         Left            =   3090
         MaskColor       =   &H00404040&
         TabIndex        =   8
         Top             =   1440
         Width           =   1125
      End
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   2
         Left            =   3090
         TabIndex        =   7
         Top             =   1080
         Width           =   1125
      End
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   3090
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   3090
         TabIndex        =   1
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "체크순"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   2
         Left            =   1590
         TabIndex        =   6
         Top             =   1530
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Rack/Pos"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   1
         Left            =   1590
         TabIndex        =   5
         Top             =   1155
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "순번 [SEQ]"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   0
         Left            =   1590
         TabIndex        =   4
         Top             =   795
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "바코드"
         ForeColor       =   &H00404040&
         Height          =   180
         Index           =   8
         Left            =   1590
         TabIndex        =   3
         Top             =   420
         Width           =   540
      End
   End
   Begin HSCotrol.CButton cmdSave 
      Height          =   495
      Left            =   3240
      TabIndex        =   24
      Top             =   6090
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " 설정저장"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestOptSet.frx":08CA
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.CButton cmdCancel 
      Height          =   495
      Left            =   4710
      TabIndex        =   25
      Top             =   6090
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " 닫    기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTestOptSet.frx":0A24
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.HSLabel HSLabel1 
      Height          =   345
      Left            =   150
      Top             =   150
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   609
      BackColor       =   16777215
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ▶ 검사옵션설정"
      BevelOut        =   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   7095
      Left            =   30
      Top             =   30
      Width           =   6705
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   405
      Left            =   120
      Top             =   120
      Width           =   6525
   End
End
Attribute VB_Name = "frmTestOptSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If optUse(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\INI\" & gMACH & ".ini")
    
    ElseIf optUse(1).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optUse(2).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optUse(3).Value = True Then
        Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If optAutoSend(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optAutoSend(1).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    If optSaveResult(0).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")
    ElseIf optSaveResult(1).Value = True Then
        Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    End If
    
    MsgBox "검사옵션정보가 변경되었습니다.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub Form_Load()

    Call GetTestOption
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub GetTestOption()

    '-- 바코드사용
    If gHOSP.BARUSE = "Y" Then
        optUse(0).Value = True
    Else
        If gHOSP.RSTTYPE = "1" Then
            optUse(1).Value = True
        ElseIf gHOSP.RSTTYPE = "2" Then
            optUse(2).Value = True
        ElseIf gHOSP.RSTTYPE = "3" Then
            optUse(3).Value = True
        End If
    End If
    
    '-- 결과전송
    If gHOSP.SAVEAUTO = "Y" Then
        optAutoSend(0).Value = True
    Else
        optAutoSend(1).Value = True
    End If
    
    '-- 적용결과
    If gHOSP.SAVELIS = "Y" Then
        optSaveResult(1).Value = True
    Else
        optSaveResult(0).Value = True
    End If
    
    
End Sub
