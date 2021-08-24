VERSION 5.00
Object = "{C727F370-34DA-4F2D-B2E0-260AED72E823}#1.0#0"; "ACKCLAV100.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   8340
      TabIndex        =   3
      Top             =   465
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   3540
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      Top             =   240
      Width           =   5010
   End
   Begin ACKCLAV100.CLA CLA1 
      Height          =   2355
      Left            =   435
      TabIndex        =   0
      Top             =   300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4154
   End
   Begin VB.Label Label1 
      Height          =   360
      Left            =   450
      TabIndex        =   1
      Top             =   4080
      Width           =   8880
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CLA1_AppendData(sID As String, sSeq As String, sRack As String, sPos As String, iRstCnt As Integer, sTIFCd As String, sTRst1 As String, sTRst2 As String, sTUnit As String, sTFlag As String, sTRstDt As String)

'    Beep
    
End Sub

Private Sub CLA1_DispMsg(sMsg As String)
    
    Label1 = sMsg
    
End Sub

Private Sub CLA1_PrintRcvLog(sLog As String)
    
    Print #1, sLog;

    Text1 = Text1 & sLog
    
End Sub

Private Sub Command1_Click()
    Text1 = ""
End Sub

Private Sub Form_Load()
    
    Open App.Path & "\CLA_Dump_" & Format(Now, "HHMM") & ".log" For Output Shared As #1
    
    With Me.CLA1
        .OpenPW = "ACK"
        .EditPW = "MEDI@CK"
        .EqName = "CLA"
        .SiteNm = ""
        .UseBarcode = True
        .Phase = 1
        .SendPhase = 1
        .FrameNo = 1
        .TestMode = 77
        
        .CommPort = 1
        .Settings = "9600,n,8,1"
        .RTSEnable = True
        .RThreshold = 1
        .PortOpen = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Close 1#
    
End Sub
