VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "검체접수Main Screen"
   ClientHeight    =   5745
   ClientLeft      =   1920
   ClientTop       =   1650
   ClientWidth     =   6870
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6165
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":030A
            Key             =   "One"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0EDE
            Key             =   "Two"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":17BA
            Key             =   "Three"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2096
            Key             =   "Four"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":23B2
            Key             =   "Five"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":26CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":29EE
            Key             =   "Seven"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2D0E
            Key             =   "Eight"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":302E
            Key             =   "Nine"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":57E2
            Key             =   "Ten"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5B02
            Key             =   "Eleven"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":63DE
            Key             =   "Twelve"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7572
            Key             =   "Thirteen"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":83C6
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "One"
            Description     =   "외래검체접수"
            Object.ToolTipText     =   "외래검체접수"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Two"
            Description     =   "병동검체접수"
            Object.ToolTipText     =   "병동검체접수"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Three"
            Object.ToolTipText     =   "병동검체확인"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Four"
            Object.ToolTipText     =   "수작업 접수"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Five"
            Object.ToolTipText     =   "접수조회"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Twelve"
            Object.ToolTipText     =   "Order조회"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eight"
            Object.ToolTipText     =   "결과조회"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nine"
            Object.ToolTipText     =   "Barcode Label 재발행"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ten"
            Object.ToolTipText     =   "외부검사List"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eleven"
            Object.ToolTipText     =   "IDChange"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "End of Program"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5370
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6482
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "오후 12:07"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuEnrol 
      Caption         =   "채혈접수(&j)"
      Begin VB.Menu mnuOpd 
         Caption         =   "외래채혈"
      End
      Begin VB.Menu mnuIPD 
         Caption         =   "병동채혈"
      End
      Begin VB.Menu mnuVerify 
         Caption         =   "병동검체확인"
      End
      Begin VB.Menu mnuManual 
         Caption         =   "수작업접수"
      End
   End
   Begin VB.Menu mnuTotalQuery 
      Caption         =   "TotalQuery"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
' 기존에 이미 Window에 해당 Program이 Loading 되었을 경우
'        Loading 되어있는 Program이 Activate 되도록 하는 Routine
'        새로 Loading 하려는 Program 은 End 시킨다
    
    Dim Title$

    If App.PrevInstance Then
        Title$ = App.Title
        App.Title = "Temp"
        AppActivate Title$
        SendKeys "%{ENTER}{ENTER}"
        End
    End If
    
'/----------------------------------------------------

    frmSplash.Show
    
    DoEvents:  Call adoDbConnect("TW_MIS_EXAM", "HOSPITAL", "v2mts")
    DoEvents:  Unload frmSplash
    DoEvents:  FrmIdPass.Show vbModal

    stbMain.Panels(3).Text = GstrPassName
    SendKeys "%" & "J"
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call adoDbDisconnect
    
    
End Sub

Private Sub mnuExit_Click()
    
    If vbNo = MsgBox("작업을 종료하시겠습니까?", vbYesNo + vbQuestion, "작업종료확인") Then
        Exit Sub
    End If
    
    End

End Sub

Private Sub mnuIPD_Click()
    
    GstrIOGubun = "IPD"
    frmIPDMain.Show
    
End Sub

Private Sub mnuManual_Click()
    
    frmManual.Show
    
End Sub

Private Sub mnuOpd_Click()
    
    GstrIOGubun = "OPD"
    frmMain.Show
    frmMain.ZOrder 0
    
End Sub

Private Sub mnuTotalQuery_Click()
    
    frmTotalQry.Show
    frmTotalQry.ZOrder 0
    
End Sub

Private Sub mnuVerify_Click()
    
    frmReEnrol.Show
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    
    
    Select Case Button.Index
        Case 1: frmMain.Show
                frmMain.ZOrder 0
                
        Case 2: frmIPDMain.Show
                frmIPDMain.ZOrder 0
                
        'Case 3: frmReEnrol.Show
        '        frmReEnrol.ZOrder 0
                
        Case 3: frmReJupsu.Show
                frmReJupsu.ZOrder 0
        
        Case 4: frmManual.Show
                frmManual.ZOrder 0
                
        Case 6: frmQuery.Show
                frmQuery.ZOrder 0
                
        Case 7: frmQryOrder.Show
                frmQryOrder.ZOrder 0
                
        Case 8: frmResult.Show
                frmResult.ZOrder 0
                
        Case 10: frmBarno.Show
                 frmBarno.ZOrder 0
                 
        Case 11: frmExExam.Show
                 frmExExam.ZOrder 0
                 
        Case 12: frmIDChange.Show vbModal
        
        Case 14: If vbYes = MsgBox("작업을 종료하시겠습니까?", _
                                   vbYesNo + vbQuestion, _
                                  "작업종료확인") Then End

    End Select
    
    
    
End Sub
