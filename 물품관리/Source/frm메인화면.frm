VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frm메인화면 
   BackColor       =   &H8000000C&
   Caption         =   "물품관리시스템"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   19080
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  '위 맞춤
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   1191
      ButtonWidth     =   1455
      ButtonHeight    =   1138
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "구매요청"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imgTree 
         Left            =   12840
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm메인화면.frx":0000
               Key             =   "close"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm메인화면.frx":08DA
               Key             =   "open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm메인화면.frx":11B4
               Key             =   "choice"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgMenu 
         Left            =   11820
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm메인화면.frx":1EA6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  '아래 맞춤
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   10875
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2725
            MinWidth        =   2716
            Text            =   "(주)성원아이티"
            TextSave        =   "(주)성원아이티"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25241
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "홍길동"
            TextSave        =   "홍길동"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "2012-07-16"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu기초자료 
      Caption         =   "[&1]기초자료  "
      Begin VB.Menu mnu기초코드등록 
         Caption         =   "기초코드등록"
      End
      Begin VB.Menu mnu업체기초자료 
         Caption         =   "업체기초자료"
      End
      Begin VB.Menu mnu물품기초자료 
         Caption         =   "물품기초자료"
      End
      Begin VB.Menu mnu장비기초자료 
         Caption         =   "장비기초자료"
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnu검사별소요량 
         Caption         =   "검사항목별 소요량"
      End
      Begin VB.Menu mnu장비별운영소요량 
         Caption         =   "장비별운영 소요량"
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu mnu사용자기초자료 
         Caption         =   "사용자기초자료"
      End
   End
   Begin VB.Menu mnu물품요청 
      Caption         =   "[&2]물품요청  "
      Begin VB.Menu mnu물품요청서일반 
         Caption         =   "물품요청서 작성(일반)"
      End
      Begin VB.Menu mnu물품요청서분류 
         Caption         =   "물품요청서 작성(분류)"
      End
      Begin VB.Menu line20 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu발주구매 
      Caption         =   "[&3]발주구매  "
      Begin VB.Menu mnu발주서일반 
         Caption         =   "발주서 작성"
      End
      Begin VB.Menu mnu발주서업체 
         Caption         =   "발주서 작성(업체)"
      End
      Begin VB.Menu mnu발주서요청 
         Caption         =   "발주서 작성(요청)"
      End
      Begin VB.Menu line30 
         Caption         =   "-"
      End
      Begin VB.Menu mnu발주입고 
         Caption         =   "발주입고 처리"
      End
      Begin VB.Menu mnu일반구매서 
         Caption         =   "일반구매서 작성"
      End
   End
   Begin VB.Menu mnu물품출고 
      Caption         =   "[&4]물품출고  "
      Begin VB.Menu mnu물품출고서 
         Caption         =   "물품출고서 작성"
      End
      Begin VB.Menu line40 
         Caption         =   "-"
      End
      Begin VB.Menu mnu장비운영등록 
         Caption         =   "장비운영내역"
      End
      Begin VB.Menu mnu수동검사등록 
         Caption         =   "수동검사등록"
      End
      Begin VB.Menu line41 
         Caption         =   "-"
      End
      Begin VB.Menu mnu일마감 
         Caption         =   "일일마감작업"
      End
   End
   Begin VB.Menu mnu물품재고 
      Caption         =   "[&5]물품재고  "
      Begin VB.Menu mnu품목별재고현황 
         Caption         =   "품목별 재고현황"
      End
      Begin VB.Menu mnu수불현황 
         Caption         =   "품목별 수불현황"
      End
      Begin VB.Menu line50 
         Caption         =   "-"
      End
      Begin VB.Menu mnu기초재고등록 
         Caption         =   "기초재고등록"
      End
      Begin VB.Menu mnu재고년도이월 
         Caption         =   "재고년도 이월"
      End
   End
   Begin VB.Menu mnu환경설정 
      Caption         =   "[&6]환경설정  "
      Begin VB.Menu mnu사용자설정 
         Caption         =   "사용자 환경설정"
      End
      Begin VB.Menu line60 
         Caption         =   "-"
      End
      Begin VB.Menu mnu원격지원 
         Caption         =   "원격지원 서비스"
      End
      Begin VB.Menu mnu종료 
         Caption         =   "프로그램 종료"
      End
   End
End
Attribute VB_Name = "frm메인화면"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()

    Me.Height = 12000
    Me.Width = 19200

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If MsgBox("프로그램을 종료하시겠습니까 ?", vbQuestion + vbYesNo) <> vbYes Then
        Cancel = 1
    Else
        End
    End If

End Sub

Private Sub mnu검사별소요량_Click()

    With frm검사항목별시약기초
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu기초코드등록_Click()

    With frm기초코드
        Call psFormCenter(frm기초코드)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu물품기초자료_Click()

    With frm물품기초자료
        Call psFormCenter(frm물품기초자료)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu물품요청서분류_Click()

    With frm물품요청서분류
        Call psFormCenter(frm물품요청서분류)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu물품요청서일반_Click()

    With frm물품요청서일반
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub psFormCenter(ByVal brForm As Form)

    brForm.Top = (frm메인화면.ScaleHeight - brForm.Height) / 2
    brForm.Left = (frm메인화면.ScaleWidth - brForm.Width) / 2
    
    brForm.Height = brForm.Height - 120

End Sub

Private Sub mnu물품재고현황_Click()

End Sub

Private Sub mnu물품출고서_Click()

    With frm출고서일반
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu발주서업체_Click()

    With frm발주서업체
        Call psFormCenter(frm발주서업체)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu발주서요청_Click()

    With frm발주서요청
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu발주서일반_Click()

    With frm발주서일반
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu발주입고_Click()

    With frm구매서발주
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu사용자기초자료_Click()

    With frm사용자기초자료
        Call psFormCenter(frm사용자기초자료)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu사용자설정_Click()

    frm데이터베이스.Show vbModal
    Call gsRegisterApply

End Sub

Private Sub mnu업체기초자료_Click()

    With frm업체기초자료
        Call psFormCenter(frm업체기초자료)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu일반구매서_Click()

    With frm구매서일반
        .Show
        .Top = 0
        .Left = 0
        .ZOrder 0
    End With

End Sub

Private Sub mnu장비기초자료_Click()

    With frm장비기초자료
        Call psFormCenter(frm장비기초자료)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu장비별운영소요량_Click()

    With frm장비별운영시약기초
        Call psFormCenter(frm장비별운영시약기초)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu장비운영등록_Click()

    With frm장비운영내역서
        Call psFormCenter(frm장비운영내역서)
        .Show
        .ZOrder 0
    End With

End Sub

Private Sub mnu종료_Click()

    Unload Me
    
End Sub

Private Sub mnu품목별재고현황_Click()

    With frm품목별재고현황
        Call psFormCenter(frm품목별재고현황)
        .Show
        .ZOrder 0
    End With

End Sub
