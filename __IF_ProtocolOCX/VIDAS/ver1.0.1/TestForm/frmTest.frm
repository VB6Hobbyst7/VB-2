VERSION 5.00
Object = "*\A..\ACKVIDA01.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtOrdOK 
      Height          =   1890
      Left            =   7815
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   8
      Top             =   3435
      Width           =   3105
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Port Close"
      Height          =   375
      Left            =   270
      TabIndex        =   7
      Top             =   2640
      Width           =   1440
   End
   Begin VB.TextBox txtRst 
      Height          =   1890
      Left            =   7815
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   6
      Top             =   1425
      Width           =   3105
   End
   Begin VB.TextBox txtSLog 
      Height          =   2325
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   4
      Top             =   3120
      Width           =   5415
   End
   Begin VB.TextBox txtRLog 
      Height          =   2325
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   3
      Top             =   690
      Width           =   5415
   End
   Begin VB.TextBox txtID 
      Height          =   345
      Left            =   8640
      TabIndex        =   2
      Top             =   750
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND ORDER"
      Height          =   390
      Left            =   8055
      TabIndex        =   1
      Top             =   195
      Width           =   1740
   End
   Begin ACKVIDA01.VIDAS VIDAS1 
      Height          =   2190
      Left            =   225
      TabIndex        =   0
      Top             =   105
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   3863
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   300
      Left            =   7875
      TabIndex        =   5
      Top             =   825
      Width           =   660
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    '오더정보 조회/편집
    With VIDAS1
        .p_iOrdCnt = 2
        .p_sID = txtID
        .p_sSeq = "123"
        .p_sTIFCd = "IGE|FSH|"
        
        If .p_iOrdCnt > 0 Then
            Call .Send_Chr(5)
            .iPhase = 2
        Else
            .iPhase = 1
        End If
    End With
    
End Sub

Private Sub Command2_Click()

    VIDAS1.PortOpen = False
    
End Sub


Private Sub Form_Load()

    '장비 OCX 초기화
    With VIDAS1
        .OpenPW = "ACK"
        .EditPW = "MEDI@CK"
        .EqName = "VIDAS"
        .bUseBarcode = True
        .iPhase = 1
        .iSendPhase = 1
        .iFrameN = 1
        .sTestMode = "77"
        
        .CommPort = "1"
        .Settings = "9600,N,8,1"
        .RTSEnable = True
        .RThreshold = 1
        .PortOpen = True
    End With
    
End Sub


Private Sub VIDAS1_AppendData(sID As String, sSeq As String, sRack As String, sPos As String, iRstCnt As Integer, sTIFCd As String, sTRst1 As String, sTRst2 As String, sTUnit As String, sTFlag As String)

    txtRst = txtRst & "ID:" & sID & vbCrLf & iRstCnt & "," & sTIFCd & "," & sTRst1 & vbCrLf
    
End Sub

Private Sub VIDAS1_PrintRcvLog(sLog As String)

    txtRLog = txtRLog & sLog
    
End Sub


Private Sub VIDAS1_PrintSendLog(sLog As String)

    txtSLog = txtSLog & sLog
    
End Sub


Private Sub VIDAS1_SendOrderOK(sID As String, sSeq As String, sRack As String, sPos As String)

    txtOrdOK = txtOrdOK & sID & " Send OK!" & vbCrLf
    
End Sub


