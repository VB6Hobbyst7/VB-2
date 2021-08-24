VERSION 5.00
Begin VB.Form frmTestOptSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   " ◈ 설정 ◈"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2700
      TabIndex        =   27
      Top             =   7320
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4350
      TabIndex        =   26
      Top             =   7320
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6675
      TabIndex        =   24
      Top             =   0
      Width           =   6675
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "검사옵션 설정"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   210
         TabIndex        =   25
         Top             =   180
         Width           =   2625
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 로그기록 "
      Height          =   765
      Left            =   600
      TabIndex        =   22
      Top             =   6150
      Width           =   5415
      Begin VB.CheckBox chkLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "로그기록"
         Height          =   345
         Left            =   3030
         TabIndex        =   23
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 워크리스트 조회화면 "
      Height          =   855
      Left            =   600
      TabIndex        =   17
      Top             =   5190
      Width           =   5415
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   345
         Left            =   2970
         TabIndex        =   19
         Top             =   360
         Width           =   2235
         Begin VB.OptionButton optWorkPos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "메인"
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   21
            Top             =   30
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optWorkPos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "팝업"
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
         ForeColor       =   &H80000008&
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
      Height          =   855
      Left            =   600
      TabIndex        =   10
      Top             =   4230
      Width           =   5415
      Begin VB.OptionButton optSaveResult 
         BackColor       =   &H00FFFFFF&
         Caption         =   "장비결과"
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
         ForeColor       =   &H80000008&
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
      Height          =   855
      Left            =   600
      TabIndex        =   9
      Top             =   3240
      Width           =   5415
      Begin VB.OptionButton optAutoSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "자동"
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
         ForeColor       =   &H80000008&
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
      Height          =   1875
      Left            =   600
      TabIndex        =   0
      Top             =   1230
      Width           =   5415
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   315
         Index           =   3
         Left            =   3090
         TabIndex        =   8
         Top             =   1440
         Width           =   1125
      End
      Begin VB.OptionButton optUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
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
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   1590
         TabIndex        =   3
         Top             =   420
         Width           =   540
      End
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
