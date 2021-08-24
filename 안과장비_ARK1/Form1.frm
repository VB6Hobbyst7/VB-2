VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmNIDEK 
   Caption         =   "NIDEK IF"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   6450
      TabIndex        =   10
      Top             =   6000
      Width           =   1725
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Port Open"
      Height          =   345
      Left            =   7260
      TabIndex        =   9
      Top             =   1830
      Width           =   1215
   End
   Begin VB.TextBox txtSettings 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   345
      Left            =   7230
      TabIndex        =   6
      Text            =   "9600,o,8,1"
      Top             =   1350
      Width           =   1275
   End
   Begin VB.TextBox txtComPort 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   345
      Left            =   7230
      TabIndex        =   5
      Text            =   "1"
      Top             =   990
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Left            =   7350
      Top             =   5160
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6780
      Top             =   4950
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   5190
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   1
      Top             =   540
      Width           =   5685
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   975
      Left            =   6120
      TabIndex        =   0
      Top             =   3420
      Width           =   1425
   End
   Begin VB.Label Label6 
      Caption         =   "Settings"
      Height          =   255
      Left            =   6210
      TabIndex        =   8
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "CommPort"
      Height          =   255
      Left            =   6210
      TabIndex        =   7
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   855
      Left            =   180
      TabIndex        =   4
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   405
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "Label1"
      Height          =   765
      Left            =   5910
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "frmNIDEK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimeupFlag As Boolean
Dim RecEotFlag As Boolean
Dim Buf, L_Buf As String
Dim CRM_SD, CRK_SD, CXX_RS As String

Private Sub cmClear_Click()
    
    Text1.Text = ""
    
End Sub

Private Sub Form_Load()

    Text1.Text = ""
    Label1.Caption = "Stand By (Port Close)"
    Label2.Caption = "Communication with AR/ARK"
    Label3.Caption = "Push Start Button"
    
    MSComm1.CommPort = txtComPort.Text
    MSComm1.Settings = txtSettings.Text
    MSComm1.RThreshold = 1
    
End Sub

Private Sub cmdOpen_Click()

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    
    MSComm1.CommPort = txtComPort.Text
    MSComm1.Settings = txtSettings.Text
    MSComm1.RThreshold = 1
    MSComm1.PortOpen = True

    If MSComm1.PortOpen = True Then
        Label1.Caption = "Receive Data (Port Open)"
    Else
        Label1.Caption = ""
    End If
    
End Sub

Private Sub cmdStart_Click()
    
    CRM_SD = Chr(1) + "CRM" + Chr(2) + "SD" + Chr(23) + Chr(4)
    CRK_SD = Chr(1) + "CRM" + Chr(2) + "SD" + Chr(23) + Chr(4)
    CXX_RS = Chr(1) + "C**" + Chr(2) + "RS" + Chr(23) + Chr(4)
    
    Text1.Text = ""
    Label3.Caption = "Initializing..."
    
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
        MSComm1.DTREnable = False
        Timer1.Enabled = True
        Timer1.Interval = 2000
        Timer1.Enabled = True: TimeupFlag = False
        
        Do  'waiting for DSR turn off
            DoEvents
        Loop Until TimeupFlag = True
        Label3.Caption = "Wating for RS command (DSR)"
        L_Buf = "": RecEotFlag = False
        
        Do  'check DSR
            DoEvents
        Loop Until MSComm1.DSRHolding = True
        Label3.Caption = "Wating for RS command (Text)"
    
        Timer1.Interval = 5000
        Timer1.Enabled = True: TimeupFlag = False
        MSComm1.DTREnable = True
        
        Do  'waiting for DSR command
            DoEvents
            If TimeupFlag = True Then
                Exit Do
            End If
        Loop Until RecEotFlag = True
    
        If L_Buf = CXX_RS Then
            Label3.Caption = "waiting for SD command transmitting (DSR)"
            L_Buf = "": RecEotFlag = False
            Do  'check DSR
                DoEvents
            Loop Until MSComm1.DSRHolding = True
            MSComm1.Output = CRM_SD
            Label3.Caption = "Wating for Data (Text)"
            Timer1.Interval = 5000
            Timer1.Enabled = True: TimeupFlag = False
            
            Do  'waiting for Data
                DoEvents
                If TimeupFlag = True Then
                    Exit Do
                End If
            Loop Until RecEotFlag = True
            
            If RecEotFlag = True Then
                Label3.Caption = "Finished Data receiving" & vbNewLine & "Push Start Button"
            Else
                Label3.Caption = "EOT error" & vbNewLine & "Push Start Button"
            End If
        Else
            Label3.Caption = "RS Command Eror(Terminated)" & vbNewLine & "Push Start Button"
        End If
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If
    End If
    
End Sub



Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
        Case comEvReceive
            Buf = MSComm1.Input
            L_Buf = L_Buf + Buf
            Text1.Text = Text1.Text + Buf
            If Right$(L_Buf, 1) = Chr(4) Then
                RecEotFlag = True
                Text1.Text = Text1.Text + vbNewLine
            End If
        Case Else
            MsgBox "Error"
    End Select
End Sub

Private Sub Timer1_Timer()
    TimeupFlag = True
    Timer1.Enabled = False
End Sub
