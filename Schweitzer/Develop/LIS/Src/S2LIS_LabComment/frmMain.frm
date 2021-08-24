VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Schweitzer - 임상병리과 검사 종합검증/판독 보고 시스템"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15075
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '최대화
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   1530
      Top             =   1470
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   15045
      TabIndex        =   0
      Top             =   615
      Width           =   15075
      Begin MSComctlLib.ProgressBar pgrBar 
         Height          =   135
         Left            =   120
         TabIndex        =   3
         Top             =   345
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblMsg1 
         BackStyle       =   0  '투명
         Caption         =   "현재 선택된 환자 :  123456789  김아무개"
         ForeColor       =   &H00E0725F&
         Height          =   180
         Left            =   11130
         TabIndex        =   2
         Top             =   195
         Width           =   3825
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  '투명
         Caption         =   "박 필 환 선생님께선 2000년 3월 8일 현재 15 명의 환자에 대해 보고서를 작성하셨습니다."
         ForeColor       =   &H00E0725F&
         Height          =   165
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   10065
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   870
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0926
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":105E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":117A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":196A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrToolbar 
      Align           =   1  '위 맞춤
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   1085
      ButtonWidth     =   1455
      ButtonHeight    =   926
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "환자검색"
            Key             =   "PTS"
            Object.ToolTipText     =   "결과보고 대상 환자를 조회합니다."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "결과조회"
            Key             =   "QRY"
            Object.ToolTipText     =   "검사결과를 조회합니다"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "결과보고"
            Key             =   "RST"
            Object.ToolTipText     =   "Test별 결과를 일괄 입력합니다"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "출력"
            Key             =   "PRT"
            Object.ToolTipText     =   "판독보고서를 출력합니다."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "월별현황"
            Key             =   "MON"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "설정"
            Key             =   "CTL"
            Object.ToolTipText     =   "검사결과를 출력합니다"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "종료"
            Key             =   "EXT"
            Object.ToolTipText     =   "프로그램을 종료합니다"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '평면
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   13605
         ScaleHeight     =   495
         ScaleWidth      =   1335
         TabIndex        =   7
         Top             =   45
         Width           =   1365
         Begin VB.Image imgLogo 
            Appearance      =   0  '평면
            Height          =   570
            Left            =   30
            Picture         =   "frmMain.frx":1D06
            Top             =   -30
            Width           =   1275
         End
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   315
         Left            =   11640
         TabIndex        =   6
         Top             =   225
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   2
               AutoSize        =   2
               Object.Width           =   1799
               MinWidth        =   1411
               TextSave        =   "2020-10-18"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Object.Width           =   1764
               MinWidth        =   1764
               TextSave        =   "오후 1:46"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdShowMemo 
         Height          =   300
         Left            =   11280
         Picture         =   "frmMain.frx":21BA
         Style           =   1  '그래픽
         TabIndex        =   5
         ToolTipText     =   "중요한 약속이나 업무등을 Memo 해 놓을수 있습니다."
         Top             =   255
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FormShow(ByVal frmThis As Form)

   Dim i As Integer
   
   'frmThis.Top = 0
   'frmThis.Left = 0
   frmThis.Show
   frmThis.ZOrder 0
  
End Sub

Private Sub cmdShowMemo_Click()
    frmMemo.Show
    frmMemo.ZOrder 0
End Sub

Private Sub MDIForm_Load()
    
    Me.Top = 0
    Me.Left = 0
    Me.Caption = "Schweitzer - 임상병리과 검사 종합검증/판독 보고 시스템 " & App.Major & "." & App.Minor & "." & App.Revision
    lblMsg.Caption = ""
    lblMsg.Top = iMsgTop2
    lblMsg1.Caption = ""
    pgrBar.Visible = False
    
    Call InitRtn
    
    frmPtList.Show
    
End Sub

Private Sub InitRtn()

    lblMsg.Caption = gDoctNm & " 선생님의 기초자료를 로딩하고 있습니다..."
    'DoEvents
    
    Set objDoctor = New clsDoctor
    With objDoctor
        .DoctId = gDoctId
        .DoctNm = gDoctNm
        Call .GetDoctInfo
        Call RptStatus(.RptCount)
    End With
    

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim Resp As VbMsgBoxResult
    
    Resp = MsgBox("종합검증/판독 메뉴를 종료하시겠습니까?", vbQuestion + vbYesNo)
    If Resp = vbNo Then
        Cancel = True
        Exit Sub
    End If
        
    If gPtntId <> "" Then
        Call UnlockPtnt(gPtntId, gBedInDT)
    End If
'    dbconn.DbClose
'    Set dbconn = Nothing
    'End

End Sub

Private Sub tbrToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

      Select Case Button.Key
         Case "PTS": Call FormShow(frmPtList)
         Case "QRY": Call FormShow(frmResultReview)
         Case "RST":
                     If gPtntId = "" Then
                        MsgBox "환자가 선택되지 않았습니다.", vbExclamation, "메세지"
                        Exit Sub
                     End If
                     frmReport.ptid = gPtntId
                     frmReport.BedinDt = gBedInDT
                     frmReport.StartQuery
                     Call FormShow(frmReport)
         Case "MON": Call FormShow(frmMonthView)
         Case "PRT":
                     If gPtntId = "" Then
                        MsgBox "환자가 선택되지 않았습니다.", vbExclamation, "메세지"
                        Exit Sub
                     End If
                     frmReport.ptid = gPtntId
                     frmReport.BedinDt = gBedInDT
                     frmReport.StartQuery
                     DoEvents
                     frmReport.PrtReport (1)
         Case "CTL": Call FormShow(frmDoctSet)
         Case "EXT":
                     Dim Resp As VbMsgBoxResult
                     
                     Resp = MsgBox("종합검증/판독 메뉴를 종료하시겠습니까?", vbQuestion + vbYesNo, "메세지")
                     If Resp = vbNo Then
                         Exit Sub
                     End If
        
                     If gPtntId <> "" Then
                         Call UnlockPtnt(gPtntId, gBedInDT)
                     End If
                     
                     End
      End Select
End Sub

Private Sub Timer1_Timer()
    
    Static TimeCount As Long
    
    TimeCount = TimeCount + 1
    If TimeCount = 15 Then  '15분 간격
        Call objDoctor.GetRptCount
        Call RptStatus(objDoctor.RptCount)
        TimeCount = 0
    End If
    
End Sub

Public Sub RptStatus(ByVal iCnt As Integer)

    With objDoctor
        lblMsg.Caption = .DoctNm & " 선생님께선 " & Format(Now, CS_DateLongFormat) & " 현재 " & _
                         .RptCount & " 명의 환자에 대해 보고서를 작성하셨습니다."
    End With
    
End Sub
