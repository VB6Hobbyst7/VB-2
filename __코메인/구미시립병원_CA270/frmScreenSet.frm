VERSION 5.00
Begin VB.Form frmScreenSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   " ◈ 화면 설정 ◈"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   Icon            =   "frmScreenSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraView 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   6915
      Left            =   270
      TabIndex        =   0
      Top             =   -180
      Width           =   5175
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
         Top             =   1066
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   34
         Top             =   1442
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   3
         Left            =   540
         TabIndex        =   33
         Top             =   1818
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   4
         Left            =   540
         TabIndex        =   32
         Top             =   2194
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   5
         Left            =   540
         TabIndex        =   31
         Top             =   2570
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   6
         Left            =   540
         TabIndex        =   30
         Top             =   2946
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   7
         Left            =   540
         TabIndex        =   29
         Top             =   3322
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   8
         Left            =   540
         TabIndex        =   28
         Top             =   3698
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   9
         Left            =   540
         TabIndex        =   27
         Top             =   4074
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   10
         Left            =   540
         TabIndex        =   26
         Top             =   4450
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   11
         Left            =   540
         TabIndex        =   25
         Top             =   4826
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   12
         Left            =   540
         TabIndex        =   24
         Top             =   5202
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   13
         Left            =   540
         TabIndex        =   23
         Top             =   5578
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   16
         Left            =   540
         TabIndex        =   22
         Top             =   5954
         Width           =   1995
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   17
         Left            =   540
         TabIndex        =   21
         Top             =   6330
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
         Top             =   1006
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   2
         Left            =   2790
         TabIndex        =   18
         Top             =   1382
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   3
         Left            =   2790
         TabIndex        =   17
         Top             =   1758
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   4
         Left            =   2790
         TabIndex        =   16
         Top             =   2134
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   5
         Left            =   2790
         TabIndex        =   15
         Top             =   2510
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   6
         Left            =   2790
         TabIndex        =   14
         Top             =   2886
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   7
         Left            =   2790
         TabIndex        =   13
         Top             =   3262
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   8
         Left            =   2790
         TabIndex        =   12
         Top             =   3638
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   9
         Left            =   2790
         TabIndex        =   11
         Top             =   4014
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   10
         Left            =   2790
         TabIndex        =   10
         Top             =   4390
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   11
         Left            =   2790
         TabIndex        =   9
         Top             =   4766
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   12
         Left            =   2790
         TabIndex        =   8
         Top             =   5142
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   13
         Left            =   2790
         TabIndex        =   7
         Top             =   5518
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   16
         Left            =   2790
         TabIndex        =   6
         Top             =   5894
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   17
         Left            =   2790
         TabIndex        =   5
         Top             =   6240
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   14
         Left            =   420
         TabIndex        =   4
         Top             =   6900
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "저장순번"
         Height          =   180
         Index           =   15
         Left            =   420
         TabIndex        =   3
         Top             =   7140
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   14
         Left            =   1620
         TabIndex        =   2
         Top             =   6840
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
         Height          =   315
         Index           =   15
         Left            =   1620
         TabIndex        =   1
         Top             =   7140
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.Image imgMenuInsert 
      Height          =   375
      Left            =   1320
      Picture         =   "frmScreenSet.frx":000C
      Top             =   7290
      Width           =   1725
   End
   Begin VB.Image imgMenuCancel 
      Height          =   375
      Left            =   3150
      Picture         =   "frmScreenSet.frx":0E08
      Top             =   7290
      Width           =   1725
   End
End
Attribute VB_Name = "frmScreenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    '-- 화면설정
    Call SetColumnName
    
    Call SetColumnView

End Sub

Private Sub SetColumnName()
    Dim i As Integer
    
    chkColumn(0).Caption = "검사일시"
    chkColumn(1).Caption = "저장순번"
    chkColumn(2).Caption = "접수일자"
    chkColumn(3).Caption = "검체번호 (바코드)"
    chkColumn(4).Caption = "Seq"
    chkColumn(5).Caption = "RACK"
    chkColumn(6).Caption = "POS"
    chkColumn(7).Caption = "입원/외래"
    chkColumn(8).Caption = "챠트번호"
    chkColumn(9).Caption = "환자번호"
    chkColumn(10).Caption = "환자이름"
    chkColumn(11).Caption = "성별"
    chkColumn(12).Caption = "나이"
    chkColumn(13).Caption = "주민번호"
'    chkColumn(14).Caption = ""
'    chkColumn(15).Caption = ""
    chkColumn(16).Caption = "오더갯수"
    chkColumn(17).Caption = "결과갯수"
    
    For i = 0 To 17
        txtColumn(i).Alignment = 2
        txtColumn(i).Text = frmMain.spdOrder.ColWidth(i + 2)
    Next

End Sub

Private Sub SetColumnView()
    Dim i       As Integer
    Dim varSize As Variant
    
    varSize = Split(gCOLSIZE, "|")
    
    For i = 0 To UBound(varSize) - 1
    
        If Mid(gCOLVIEW, i + 1, 1) = 1 Then
            frmScreenSet.chkColumn(i).Value = "1"
        Else
            frmScreenSet.chkColumn(i).Value = "0"
        End If
        'spdOrder.ColWidth(i + 2) = varSize(i)
    Next


End Sub

Private Sub imgMenuCancel_Click()
    Unload Me
End Sub

Private Sub imgMenuInsert_Click()
    Dim i          As Integer
    Dim strSPDView As String
    Dim strSPDSize As String
    
    strSPDView = ""
    
    For i = 0 To 17
        strSPDView = strSPDView & IIf(chkColumn(i).Value = "1", "1", "0")
        strSPDSize = strSPDSize & txtColumn(i).Text & "|"
    Next
    
    Call WritePrivateProfileString("VIEW", "SPDVIEW", strSPDView, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("VIEW", "SPDSIZE", strSPDSize, App.PATH & "\INI\" & gMACH & ".ini")

    '-- 컬럼보이기설정
    Call SetColumnView
    
    MsgBox "컬럼정보가 변경되었습니다.", vbInformation + vbOKOnly, Me.Caption
End Sub
