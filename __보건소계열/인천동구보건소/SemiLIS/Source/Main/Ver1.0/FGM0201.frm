VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form FGM0201 
   BackColor       =   &H00808080&
   BorderStyle     =   0  '없음
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Enabled         =   0   'False
   Icon            =   "FGM0201.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel3 
      Align           =   1  '위 맞춤
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12000
      _Version        =   65536
      _ExtentX        =   21167
      _ExtentY        =   529
      _StockProps     =   15
      ForeColor       =   16777215
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   2
      BevelInner      =   2
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  '아래 맞춤
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   12000
      _Version        =   65536
      _ExtentX        =   21167
      _ExtentY        =   794
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      Begin Threed.SSPanel pnlMsg 
         Height          =   375
         Left            =   2970
         TabIndex        =   4
         Top             =   30
         Width           =   6705
         _Version        =   65536
         _ExtentX        =   11827
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlUserNm 
         Height          =   375
         Left            =   930
         TabIndex        =   2
         Top             =   30
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "김태윤"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   375
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "User"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlDT 
         Height          =   375
         Left            =   9690
         TabIndex        =   3
         Top             =   30
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "1999-01-25   14:24:43"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
      End
   End
   Begin VB.Timer Timer1 
      Left            =   630
      Top             =   1740
   End
End
Attribute VB_Name = "FGM0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iMsgTimeCnt%
Dim iMsgInterval%

Private Sub Form_Load()
    Dim sBuf$
    Dim bRetVal As Boolean
    Dim i%
        
''''Width
'''    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\BottomMain.View", "Width")
'''
'''    If sBuf = "" Then
'''        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\BottomMain.View", "Width", "12000")
'''
'''        If bRetVal = True Then
'''            Me.Width = 12000
'''        Else
'''            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
'''            Me.Width = 12000
'''        End If
'''    Else
'''        Me.Width = CInt(sBuf)
'''    End If
    

'StatusPnl 초기화
    pnlUserNm = ""
    pnlMsg = ""
    
    FGM0201.Caption = "   " & FGM0101.Caption
    
'Msg Handle 레지스트리 입력
    bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "MsgHwnd", pnlMsg.hwnd)
        
    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
'UserNm 쓸곳의 Handle 레지스트리 입력
    bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "UserNmHwnd", pnlUserNm.hwnd)
        
    If bRetVal = True Then
    Else
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
    End If
    
    
'알림판 초기화 타이머 간격 설정
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\MsgTimerInterval", "Sec")

    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\MsgTimerInterval", "Sec", "5")

        If bRetVal = True Then
            iMsgInterval = 5
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            iMsgInterval = 10
        End If
    Else
        iMsgInterval = CInt(sBuf)
    End If
    
    
'Timer 초기화
    iMsgTimeCnt = 0
    Timer1.Interval = 1000
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    iMsgTimeCnt = iMsgTimeCnt + 1
    pnlDT = Format$(Now, "YYYY-MM-DD   HH:MM:SS")
    
    If iMsgTimeCnt = iMsgInterval Then
        ViewMsg ""
        DoEvents
        iMsgTimeCnt = 0
    End If
    
End Sub
