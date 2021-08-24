VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSetComm 
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신환경설정"
   ClientHeight    =   4770
   ClientLeft      =   2265
   ClientTop       =   1650
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "setcomm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   10.5
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4653
      TabIndex        =   7
      Top             =   96
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   10.5
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2422
      TabIndex        =   6
      Top             =   96
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   10.5
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   192
      TabIndex        =   5
      Top             =   96
      Width           =   1575
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3855
      Left            =   180
      TabIndex        =   8
      Top             =   690
      Width           =   6030
      _Version        =   65536
      _ExtentX        =   10636
      _ExtentY        =   6800
      _StockProps     =   14
      Caption         =   "통신환경정의"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbGb 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2175
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   480
         Width           =   3492
      End
      Begin VB.ComboBox cmbJSerial 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   930
         Width           =   3492
      End
      Begin VB.ComboBox cmbPort 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         Sorted          =   -1  'True
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   1410
         Width           =   3492
      End
      Begin VB.ComboBox cmbBaud 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   1875
         Width           =   3492
      End
      Begin VB.ComboBox cmbData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   2820
         Width           =   3492
      End
      Begin VB.ComboBox cmbParity 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   2355
         Width           =   3492
      End
      Begin VB.ComboBox cmbStop 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   3300
         Width           =   3492
      End
      Begin VB.Label Label9 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "검사항목코드"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   300
         TabIndex        =   17
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "장비일련번호"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   300
         TabIndex        =   15
         Top             =   990
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "포트(&C)"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   13
         Top             =   1425
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "전송속도(&B)"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   12
         Top             =   1905
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "데이타비트(&D)"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   11
         Top             =   2865
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "패리티(&P)"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   10
         Top             =   2385
         Width           =   1545
      End
      Begin VB.Label Label5 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "정지비트(&S)"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   9
         Top             =   3345
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmSetComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim sLen                As Integer
    Dim sB1                 As Integer
    Dim sB2                 As Integer
    Dim sP1                 As Integer
    Dim sP2                 As Integer
    Dim sD1                 As Integer
    Dim sD2                 As Integer
    Dim sS1                 As Integer
    Dim sS2                 As Integer
    Dim i                   As Integer
    
    Dim CfgPort             As String
    Dim CfgComm             As String
    
    Dim sComm               As String
    
    Dim sPort, sBaud, sParity, sData, sStop
'


'Private Sub cmdPc_Click()
'    txtPc.Visible = True
'    txtPc.Text = pnlPc.Caption
'
'End Sub


'Private Sub txtPc_GotFocus()
'    txtPc.SelStart = 0
'    txtPc.SelLength = Len(txtPc.Text)
'
'End Sub

'Private Sub txtPc_KeyPress(KeyAscii As Integer)
'    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32''
'
'    If KeyAscii <> 13 Then Exit Sub
'    If txtPc.Text < "CP001" Or txtPc.Text > "CP099" Then Exit Sub
'    txtPc.Visible = False
'    pnlPc.Caption = txtPc.Text
'    cmbJGb.ListIndex = Val(Mid(pnlPc.Caption, 3, 3)) - 1'
'
'End Sub


Private Sub Form_Load()

    Dim CodeCheck
 
    On Error Resume Next
    frmSetComm.Left = (Screen.Width - frmSetComm.Width) / 2
    frmSetComm.Top = 1300
    
    Call Set_Comm
    
    CodeCheck = GetSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE)
    GGCODE = Mid(CodeCheck, 1, 2)
    
    If Mid(CodeCheck, 6, 1) = "1" Then
        sPort = GetSetting("LabInterface", "SetComm", "ComPort1")
        sComm = GetSetting("LabInterface", "SetComm", "ComSettings1")
    ElseIf Mid(CodeCheck, 6, 1) = "2" Then
        sPort = GetSetting("LabInterface", "SetComm", "ComPort2")
        sComm = GetSetting("LabInterface", "SetComm", "ComSettings2")
    End If
    
    sLen = Len(sComm)

    strSQL = ""
    strSQL = strSQL & " SELECT Unique GCODE, GJITEM "
    strSQL = strSQL & "   FROM TWEXAM_LABINIT "
    strSQL = strSQL & "  WHERE ETC = '*' "
    strSQL = strSQL & "  ORDER BY GCODE "
    Result = adoSQL(strSQL)
    If Result = 0 And rowindicator > 1 Then
    
        Do Until Rs.EOF
            cmbGb.AddItem Format(Trim$(Rs.Fields("GCODE") & ""), "00") & " " & Trim$(Rs.Fields("GJITEM") & "")
            Rs.MoveNext
        Loop
    Else
        MsgBox "  Communication setting data 조회 ERROR  " & vbCrLf & vbCrLf & _
               "  장비 CODE가 없습니다.", vbCritical
        Unload Me
    End If
    
    AdoCloseSet Rs
    
    cmbJSerial.ListIndex = Mid(CodeCheck, 4, 1) - 1
    
    For i = 0 To 10
        If Mid(cmbGb.List(i), 1, 2) = GGCODE Then
            cmbGb.ListIndex = i
            Exit For
        End If
    Next i
    
    
    Call Set_Comm_Select

End Sub


Private Sub CmdOK_Click()
    
    Dim aaa
    
    If cmbJSerial.ListIndex <> cmbPort.ListIndex Then
        MsgBox "  장비일련 번호와 PORT 번호를 일치 시키십시요. "
        Exit Sub
    End If
    
    CfgPort = Mid(cmbPort.Text, 4, 1)
''''CfgComm = Baud & Parity & Data & Stop
    CfgComm = cmbBaud.Text & "," & MidH(cmbParity.Text, 6, 1) & "," & cmbData.Text & "," & cmbStop.Text
    
    aaa = Mid(cmbGb.Text, 1, 2) & "," & Trim(cmbJSerial.Text) & "," & CfgPort
    
    Call SaveSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE, aaa)
    If Trim(cmbJSerial.Text) = "1" Then
        Call SaveSetting("LabInterface", "SetComm", "ComPort1", CfgPort)
        Call SaveSetting("LabInterface", "SetComm", "ComSettings1", CfgComm)
    ElseIf Trim(cmbJSerial.Text) = "2" Then
        Call SaveSetting("LabInterface", "SetComm", "ComPort2", CfgPort)
        Call SaveSetting("LabInterface", "SetComm", "ComSettings2", CfgComm)
    End If
    
'    If Mid(GGJCODE, 6, 1) = "1" Then
'        sPort = GetSetting("LabInterface", "SetComm", "ComPort1")
'        sComm = GetSetting("LabInterface", "SetComm", "ComSettings1")
'    ElseIf Mid(GGJCODE, 6, 1) = "2" Then
'        sPort = GetSetting("LabInterface", "SetComm", "ComPort2")
'        sComm = GetSetting("LabInterface", "SetComm", "ComSettings2")
'    End If

    If MsgBox("변경된 내용은 프로그램을 다시 시작해야 적용됩니다" & vbNewLine & vbNewLine & _
            "지금 다시 시작하시겠습니까?, ", vbQuestion + vbYesNo, "변경확인") = vbYes Then
            End
    End If
    
End Sub


Private Sub cmdClear_Click()
    cmbPort.Clear
    cmbBaud.Clear
    cmbParity.Clear
    cmbData.Clear
    cmbStop.Clear
    'cmbJGb.Clear
    cmbGb.Clear
    cmbJSerial.Clear
    Call Form_Load

End Sub


Private Sub CmdExit_Click()
    Unload Me

End Sub


Sub Set_Comm()
    
    cmbJSerial.AddItem " 1"
    cmbJSerial.AddItem " 2"
    
    cmbPort.AddItem "COM1"        ' default value
    cmbPort.AddItem "COM2"
    cmbPort.AddItem "COM3"
    cmbPort.AddItem "COM4"
    
    cmbBaud.AddItem "1200"        ' default value
    cmbBaud.AddItem "2400"
    cmbBaud.AddItem "4800"
    cmbBaud.AddItem "9600"
'    cmbBaud.AddItem  "14400"
'    cmbBaud.AddItem  "19200"
'    cmbBaud.AddItem  "38400"

    cmbParity.AddItem "없슴(N)"   ' default value
    cmbParity.AddItem "짝수(E)"
    cmbParity.AddItem "홀수(O)"
    
    cmbData.AddItem "8"           ' default value
    cmbData.AddItem "7"

    cmbStop.AddItem "1"           ' default value
    cmbStop.AddItem "2"
  
End Sub


Sub Set_Comm_Select()
'   1234567890
'  "960,N,8,1"
'  "9600,N,8,1"
'  "96000,N,8,1"
'  "960000,N,8,1"
    
    '
    Select Case sLen
           Case 9
                   '+                 +                 +                  +
                    sB1 = 1: sB2 = 3: sP1 = 5: sP2 = 1: sD1 = 7:  sD2 = 1: sS1 = 9:  sS2 = 1
           Case 10
                    sB1 = 1: sB2 = 4: sP1 = 6: sP2 = 1: sD1 = 8:  sD2 = 1: sS1 = 10: sS2 = 1
           Case 11
                    sB1 = 1: sB2 = 5: sP1 = 7: sP2 = 1: sD1 = 9:  sD2 = 1: sS1 = 11: sS2 = 1
           Case 12
                    sB1 = 1: sB2 = 6: sP1 = 8: sP2 = 1: sD1 = 10: sD2 = 1: sS1 = 12: sS2 = 1
    End Select

    sBaud = Mid(sComm, sB1, sB2)
    sParity = Mid(sComm, sP1, sP2)
    sData = Mid(sComm, sD1, sD2)
    sStop = Mid(sComm, sS1, sS2)

'============================================================

    'Port Check
    cmbPort.ListIndex = sPort - 1
    
    ' Baud Check
    For i = 0 To 9
        If sBaud = cmbBaud.List(i) Then
            cmbBaud.ListIndex = i
            Exit For
        End If
    Next
    
    ' Parity Check
    For i = 0 To 2
        If sParity = MidH(cmbParity.List(i), 6, 1) Then
            cmbParity.ListIndex = i
            Exit For
        End If
    Next
    
    ' Data Check
    For i = 0 To 1
        If sData = cmbData.List(i) Then
            cmbData.ListIndex = i
            Exit For
        End If
    Next

    ' Stop Check
    For i = 0 To 1
        If sStop = cmbStop.List(i) Then
            cmbStop.ListIndex = i
            Exit For
        End If
    Next
  
End Sub
