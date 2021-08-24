VERSION 5.00
Begin VB.Form frmComment 
   BackColor       =   &H00FFFFFF&
   Caption         =   "코멘트 설정"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765
   Icon            =   "frmComment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9765
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2190
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4590
      Width           =   5000
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   6090
      Width           =   5000
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6300
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7470
      TabIndex        =   8
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox txtSGPB5 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2430
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3420
      Width           =   5000
   End
   Begin VB.TextBox txtSGRV16 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2820
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2580
      Width           =   5000
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9765
      TabIndex        =   1
      Top             =   0
      Width           =   9765
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "코멘트 설정"
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
         Index           =   12
         Left            =   210
         TabIndex        =   2
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.TextBox txtHeader 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   6435
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "INCONCLUSIVE 코멘트"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   390
      TabIndex        =   13
      Top             =   3660
      Width           =   1950
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "FOOTER"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   11
      Top             =   6210
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "POSITIVE 코멘트"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "NEGATIVE 코멘트"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "HEADER"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   1290
      Width           =   720
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub cmdConfirm_Click()
'    Dim strSGCovid19    As String
'    Dim strSDCovid19    As String
'    Dim strSGRV16       As String
'    Dim strSGPB5        As String
'
'
'    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
'        Unload Me
'        Exit Sub
'    Else
'        strSGCovid19 = Replace(txtSGCovid19.Text, vbCrLf, "CHR(10)CHR(13)")
'        strSGRV16 = Replace(txtSGRV16.Text, vbCrLf, "CHR(10)CHR(13)")
'        strSGPB5 = Replace(txtSGPB5.Text, vbCrLf, "CHR(10)CHR(13)")
'
'        Call WritePrivateProfileString("COMMENT", "SGCOVID", strSGCovid19, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("COMMENT", "RV16", strSGRV16, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("COMMENT", "PB5", strSGPB5, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("COMMENT", "PATH", txtPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
'
'        Unload Me
'    End If
'
'End Sub
'
'Private Sub cmdExit_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Dim strHeader  As String
'
'    strHeader = ""
'    strHeader = strHeader & "[검사명] 2019-nCoV, real-time RT PCR" & vbCrLf
'    strHeader = strHeader & "" & vbCrLf
'
'    strFooter = ""
'    strFooter = strFooter & "[검사방법] Real-time RT-PCR" & vbCrLf
'    strFooter = strFooter & "[검사개요]" & vbCrLf
'    strFooter = strFooter & ". 유전자와 검사 : 2019-nCoV (2019-Novel Coronavirus)의 E segment, RdRP segment를 real-time RT PCR로 검출함." & vbCrLf
'    strFooter = strFooter & ". 관련질환과 의의: 발열, 기침, 호흡곤란, 두통, 오한, 인후통, 호흡부전, 패혈성 쇼크, 다발성 장기 부전 등의 원인인 2019-nCoV를 검출함"
'    strFooter = strFooter & "" & vbCrLf
'    strFooter = strFooter & "--"
'    strFooter = strFooter & "[검사문의] 서울대병원 진단검사의학과 분자유전검사실, T. (02) 2072-2937, 0883"
'    strFooter = strFooter & "" & vbCrLf
'    strFooter = strFooter & "-------------------------------------------------------------------------" & vbCrLf
'    strFooter = strFooter & "보고자(판독의) : 조성임 M.T / 박수용 M.D / 김택수 M.D / 박성섭 M.D / 성문우 M.D" & vbCrLf
'    strFooter = strFooter & "-------------------------------------------------------------------------" & vbCrLf
'
'        If UCase(pResult) = "POSITIVE" Then
'            strPCmnt = ""
'            strPCmnt = strPCmnt & "[검사결과 및 의견]" & vbCrLf
'            strPCmnt = strPCmnt & "2019-nCoV:Positive(양성) (검체: " & strSpcNm & ")" & vbCrLf
'            strPCmnt = strPCmnt & "(E: " & strEval & ", RdRp: " & strRdRpVal & " )" & vbCrLf
'            strPCmnt = strPCmnt & "" & vbCrLf
'            strPCmnt = strPCmnt & "* Comment: 2020.04.08 1PM 부터 검사 kit 변경되었으니 참고하시기 바랍니다." & vbCrLf
'            strPCmnt = strPCmnt & "" & vbCrLf
'            strPCmnt = strPCmnt & "--" & vbCrLf
'
'            strCmnt = strPCmnt
'        End If
'
'        If UCase(pResult) = "NEGATIVE" Then
'            strNCmnt = ""
'            strNCmnt = strNCmnt & "[검사결과 및 의견]" & vbCrLf
'            strNCmnt = strNCmnt & "2019-nCoV:Negative(음성) (검체: " & strSpcNm & ")" & vbCrLf
'            strNCmnt = strNCmnt & "" & vbCrLf
'            strNCmnt = strNCmnt & "--" & vbCrLf
'
'            strCmnt = strNCmnt
'        End If
'
'        If UCase(pResult) = "INCONCLUSIVE" Then
'            strICmnt = ""
'            strICmnt = strICmnt & "[검사결과 및 의견]" & vbCrLf
'            strICmnt = strICmnt & "2019-nCoV:Inconclusive(미결정) (검체: " & strSpcNm & ")" & vbCrLf
'            strICmnt = strICmnt & "(E: " & strEval & ", RdRp: " & strRdRpVal & " )" & vbCrLf
'            strICmnt = strICmnt & "" & vbCrLf
'            If strEval <> "" And strRdRpVal = "" Then
'                strICmnt = strICmnt & "* Comment: E segment가 약한증폭이 확인되었고, RdRp segmentsms는 음성입니다." & vbCrLf
'            ElseIf strEval = "" And strRdRpVal <> "" Then
'                strICmnt = strICmnt & "* Comment: E segment는 음성이나, RdRp segment가 약한 증폭이 확인되었습니다." & vbCrLf
'            End If
'            strICmnt = strICmnt & "추가 확인을 위하여 새로 채취한 검체를 분자검사실로 보내주시면 검사 진행하겠습니다." & vbCrLf
'            strICmnt = strICmnt & "" & vbCrLf
'            strICmnt = strICmnt & "--" & vbCrLf
'
'            strCmnt = strICmnt
'        End If
'
'
'    txtHeader.Text = strHeader
'
'
'    txtSGCovid19.Text = Replace(gCFXCmnt.SGCOVID, "CHR(10)CHR(13)", vbCrLf)
'    txtSGRV16.Text = Replace(gCFXCmnt.RV16, "CHR(10)CHR(13)", vbCrLf)
'    txtSGPB5.Text = Replace(gCFXCmnt.PB5, "CHR(10)CHR(13)", vbCrLf)
'    txtPath.Text = gCFXCmnt.PATH
'
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyEscape Then
'        Unload Me
'    End If
'
'End Sub
'
