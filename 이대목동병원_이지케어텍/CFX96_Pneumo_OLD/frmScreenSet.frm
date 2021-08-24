VERSION 5.00
Begin VB.Form frmScreenSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   " ◈ 설정 ◈"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "frmScreenSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtColWidth 
      Alignment       =   2  '가운데 맞춤
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3570
      TabIndex        =   47
      Top             =   8010
      Width           =   975
   End
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
      Left            =   1860
      TabIndex        =   46
      Top             =   8640
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
      Left            =   3510
      TabIndex        =   45
      Top             =   8640
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6285
      TabIndex        =   37
      Top             =   0
      Width           =   6285
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
         Caption         =   "화면 설정"
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
         Index           =   2
         Left            =   210
         TabIndex        =   38
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.Frame fraView 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   7695
      Left            =   780
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   20
         Left            =   2790
         TabIndex        =   44
         Top             =   7080
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   20
         Left            =   540
         TabIndex        =   43
         Top             =   7150
         Width           =   1125
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   19
         Left            =   2790
         TabIndex        =   42
         Top             =   6750
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   19
         Left            =   540
         TabIndex        =   41
         Top             =   6827
         Width           =   1125
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   18
         Left            =   2790
         TabIndex        =   40
         Top             =   6420
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   18
         Left            =   540
         TabIndex        =   39
         Top             =   6504
         Width           =   1125
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   36
         Top             =   690
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   1
         Left            =   540
         TabIndex        =   35
         Top             =   1013
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   34
         Top             =   1336
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   3
         Left            =   540
         TabIndex        =   33
         Top             =   1659
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   4
         Left            =   540
         TabIndex        =   32
         Top             =   1982
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   5
         Left            =   540
         TabIndex        =   31
         Top             =   2305
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   6
         Left            =   540
         TabIndex        =   30
         Top             =   2628
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   7
         Left            =   540
         TabIndex        =   29
         Top             =   2951
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   8
         Left            =   540
         TabIndex        =   28
         Top             =   3274
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   9
         Left            =   540
         TabIndex        =   27
         Top             =   3597
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   10
         Left            =   540
         TabIndex        =   26
         Top             =   3920
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   11
         Left            =   540
         TabIndex        =   25
         Top             =   4243
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   12
         Left            =   540
         TabIndex        =   24
         Top             =   4566
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   13
         Left            =   540
         TabIndex        =   23
         Top             =   4889
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   14
         Left            =   540
         TabIndex        =   22
         Top             =   5212
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   15
         Left            =   540
         TabIndex        =   21
         Top             =   5535
         Width           =   1995
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   0
         Left            =   2790
         TabIndex        =   20
         Top             =   630
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   1
         Left            =   2790
         TabIndex        =   19
         Top             =   945
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   2
         Left            =   2790
         TabIndex        =   18
         Top             =   1260
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   3
         Left            =   2790
         TabIndex        =   17
         Top             =   1575
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   4
         Left            =   2790
         TabIndex        =   16
         Top             =   1890
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   5
         Left            =   2790
         TabIndex        =   15
         Top             =   2205
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   6
         Left            =   2790
         TabIndex        =   14
         Top             =   2520
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   7
         Left            =   2790
         TabIndex        =   13
         Top             =   2835
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   8
         Left            =   2790
         TabIndex        =   12
         Top             =   3135
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   9
         Left            =   2790
         TabIndex        =   11
         Top             =   3450
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   10
         Left            =   2790
         TabIndex        =   10
         Top             =   3765
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   11
         Left            =   2790
         TabIndex        =   9
         Top             =   4080
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   12
         Left            =   2790
         TabIndex        =   8
         Top             =   4395
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   13
         Left            =   2790
         TabIndex        =   7
         Top             =   4710
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   14
         Left            =   2790
         TabIndex        =   6
         Top             =   5055
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   15
         Left            =   2790
         TabIndex        =   5
         Top             =   5400
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   16
         Left            =   540
         TabIndex        =   4
         Top             =   5858
         Width           =   1125
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   17
         Left            =   540
         TabIndex        =   3
         Top             =   6181
         Width           =   1125
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   16
         Left            =   2790
         TabIndex        =   2
         Top             =   5730
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   17
         Left            =   2790
         TabIndex        =   1
         Top             =   6090
         Width           =   1515
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사항목 넓이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   7
      Left            =   2070
      TabIndex        =   48
      Top             =   8070
      Width           =   1230
   End
End
Attribute VB_Name = "frmScreenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i          As Integer
    Dim strSPDView As String
    Dim strSPDSize As String
    
    strSPDView = ""
    
    For i = 0 To 20
        strSPDView = strSPDView & IIf(chkColumn(i).Value = "1", "1", "0")
        strSPDSize = strSPDSize & txtColumn(i).Text & "|"
    Next
    
    Call WritePrivateProfileString("VIEW", "SPDVIEW", strSPDView, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("VIEW", "SPDSIZE", strSPDSize, App.PATH & "\INI\" & gMACH & ".ini")

    '-- 컬럼보이기설정
    Call SetColumnView(frmMain.spdOrder)
    
    Call WritePrivateProfileString("VIEW", "COLWIDTH", txtColWidth.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    MsgBox "컬럼정보가 변경되었습니다.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub Form_Load()

    '-- 화면설정
    Call SetColumnName
    
    Call SetColumnView(frmMain.spdOrder)
    
    txtColWidth.Text = gCOLWIDTH

End Sub

Private Sub SetColumnName()
    Dim i       As Integer
    Dim varView As Variant
    Dim varSize As Variant
    
    chkColumn(0).Caption = "선택"
    chkColumn(1).Caption = "검사일시"
    chkColumn(2).Caption = "검사시간"
    chkColumn(3).Caption = "검사순번"
    chkColumn(4).Caption = "ER"
    chkColumn(5).Caption = "RT"
    chkColumn(6).Caption = "접수일자"
    chkColumn(7).Caption = "검체번호"
    chkColumn(8).Caption = "검체"
    chkColumn(9).Caption = "RackNo"
    chkColumn(10).Caption = "TubePos"
    chkColumn(11).Caption = "SeqNo"
    chkColumn(12).Caption = "이름"
    chkColumn(13).Caption = "Sex"
    chkColumn(14).Caption = "Age"
    chkColumn(15).Caption = "병록번호"
    chkColumn(16).Caption = "챠트번호"
    chkColumn(17).Caption = "의뢰과"
    chkColumn(18).Caption = "입/외구분"
    chkColumn(19).Caption = "오더갯수"
    chkColumn(20).Caption = "결과갯수"
    
    
    For i = 0 To 20
        'If Mid(varViewi + 1, 1) = "1" Then
        chkColumn(i).Value = Mid(gCOLVIEW, i + 1, 1)
    Next
    
    
    varSize = Split(gCOLSIZE, "|")
    
    For i = 0 To 20
        txtColumn(i).Alignment = 2
        txtColumn(i).Text = varSize(i)
    Next

End Sub

'Private Sub SetColumnView()
'    Dim i       As Integer
'    Dim varSize As Variant
'
'    varSize = Split(gCOLSIZE, "|")
'
'    For i = 0 To UBound(varSize) - 1
'
'        If Mid(gCOLVIEW, i + 1, 1) = 1 Then
'            frmScreenSet.chkColumn(i).Value = "1"
'        Else
'            frmScreenSet.chkColumn(i).Value = "0"
'        End If
'        'spdOrder.ColWidth(i + 1) = varSize(i + 1)
'    Next
'
'
'End Sub

