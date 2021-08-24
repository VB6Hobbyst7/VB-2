VERSION 5.00
Begin VB.Form frmRegComm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "Åë½Å¼³Á¤"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   9510
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin VB.CommandButton cmdExit 
      Caption         =   "Ãë¼Ò"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6870
      TabIndex        =   20
      Top             =   5610
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ÀúÀå"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5700
      TabIndex        =   19
      Top             =   5610
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'À§ ¸ÂÃã
      BackColor       =   &H00808000&
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9510
      TabIndex        =   17
      Top             =   0
      Width           =   9510
      Begin VB.Label Label1 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Åë½ÅÁ¤º¸ ¼³Á¤"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   18
         Top             =   180
         Width           =   2625
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ½Ã¸®¾ó ¹æ½Ä "
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   4095
      Begin VB.CheckBox chkRTS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   8
         Top             =   3720
         Value           =   1  'È®ÀÎ
         Width           =   1785
      End
      Begin VB.CheckBox chkDTR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   7
         Top             =   4230
         Value           =   1  'È®ÀÎ
         Width           =   1785
      End
      Begin VB.ComboBox Combo_Parity 
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   6
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox Combo_Stopbit 
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   5
         Top             =   2790
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Startbit 
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   4
         Top             =   2340
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Databit 
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   3
         Top             =   1890
         Width           =   1605
      End
      Begin VB.ComboBox Combo_BPS 
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Top             =   1440
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Port 
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   1
         Top             =   990
         Width           =   1605
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "RTSEnable"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   330
         TabIndex        =   16
         Top             =   3780
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "DTREnable"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   48
         Left            =   330
         TabIndex        =   15
         Top             =   4290
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Parity"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   60
         TabIndex        =   14
         Top             =   3315
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Stop Bit"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   13
         Top             =   2865
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Start Bit"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   12
         Top             =   2415
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Data Bit"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   11
         Top             =   1965
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   75
         TabIndex        =   10
         Top             =   1515
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Com Port"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   11
         Left            =   75
         TabIndex        =   9
         Top             =   1065
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmRegComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click()
    
    On Error GoTo ErrorHandler
    
    If MsgBox("¼³Á¤À» ÀúÀåÇÏ½Ã°Ú½À´Ï±î?", vbCritical + vbOKCancel + vbDefaultButton2, "È®ÀÎ!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
'        If optUse(0).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\KDBAR.ini")
'        ElseIf optUse(1).Value = True Then
'            Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\KDBAR.ini")
'        End If
        
        gComm.COMPORT = Combo_Port.Text
        gComm.SPEED = Combo_BPS.Text
        gComm.DATABIT = Combo_Databit.Text
        gComm.STARTBIT = Combo_Startbit.Text
        gComm.STOPBIT = Combo_Stopbit.Text
        If Combo_Parity.ListIndex = 0 Then
           gComm.Parity = "N"
        End If
        If Combo_Parity.ListIndex = 1 Then
           gComm.Parity = "E"
        End If
        If Combo_Parity.ListIndex = 2 Then
           gComm.Parity = "O"
        End If
        
        Call WritePrivateProfileString("COMM", "COMPORT", gComm.COMPORT, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("COMM", "SPEED", gComm.SPEED, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("COMM", "PARITY", gComm.Parity, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("COMM", "DATABIT", gComm.DATABIT, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("COMM", "STOPBIT", gComm.STOPBIT, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("COMM", "STARTBIT", gComm.STARTBIT, App.PATH & "\KDBAR.ini")
        
        If chkRTS.Value = "1" Then
            Call WritePrivateProfileString("COMM", "RTSEnable", "True", App.PATH & "\KDBAR.ini")
        Else
            Call WritePrivateProfileString("COMM", "RTSEnable", "False", App.PATH & "\KDBAR.ini")
        End If
        If chkRTS.Value = "1" Then
            Call WritePrivateProfileString("COMM", "DTREnable", "True", App.PATH & "\KDBAR.ini")
        Else
            Call WritePrivateProfileString("COMM", "DTREnable", "False", App.PATH & "\KDBAR.ini")
        End If
            
'            If frmMain.comEqp.PortOpen = True Then
'                frmMain.comEqp.PortOpen = False
'            End If
'
'            frmMain.comEqp.CommPort = gComm.COMPORT
'            frmMain.comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
'
'            frmMain.comEqp.PortOpen = True
        
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("Åë½Å¼³Á¤ÀÌ ¸ÂÁö ¾Ê½À´Ï´Ù", vbCritical + vbOKCancel + vbDefaultButton2, "Á¾·á¹öÆ°") = vbCancel Then
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Dim intComPortExist As Long
    Dim i As Integer
    Dim Ret As Integer
    
    Combo_Port.Clear
    For i = 1 To 16
        intComPortExist = EnumSerPorts(i)
        If intComPortExist > 0 Then
            Combo_Port.AddItem Trim(Str(i))
        End If
    Next
    
    Combo_BPS.AddItem ("150")
    Combo_BPS.AddItem ("300")
    Combo_BPS.AddItem ("600")
    Combo_BPS.AddItem ("1200")
    Combo_BPS.AddItem ("2400")
    Combo_BPS.AddItem ("4800")
    Combo_BPS.AddItem ("9600")
    Combo_BPS.AddItem ("14400")
    Combo_BPS.AddItem ("19200")
    Combo_BPS.AddItem ("38400")
    Combo_BPS.AddItem ("115200")
    
    
    Combo_Databit.AddItem ("7")
    Combo_Databit.AddItem ("8")
    
    Combo_Startbit.AddItem ("1")
    Combo_Startbit.AddItem ("2")
    
    Combo_Stopbit.AddItem ("1")
    Combo_Stopbit.AddItem ("1.5")
    Combo_Stopbit.AddItem ("2")
    
    Combo_Parity.AddItem ("N")
    Combo_Parity.AddItem ("E")
    Combo_Parity.AddItem ("O")
   
    Ret = -1
    For i = 0 To Combo_Port.ListCount - 1
        If gComm.COMPORT = Trim(Combo_Port.List(i)) Then
            Combo_Port.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Port.ListIndex = 1
    End If
    
    Ret = -1
    For i = 0 To Combo_BPS.ListCount - 1
        If gComm.SPEED = Trim(Combo_BPS.List(i)) Then
            Combo_BPS.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_BPS.ListIndex = 4
    End If
    
    Ret = -1
    For i = 0 To Combo_Databit.ListCount - 1
        If gComm.DATABIT = Trim(Combo_Databit.List(i)) Then
            Combo_Databit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Databit.ListIndex = 1
    End If

    Ret = -1
    For i = 0 To Combo_Startbit.ListCount - 1
        If gComm.STARTBIT = Trim(Combo_Startbit.List(i)) Then
            Combo_Startbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Startbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To Combo_Stopbit.ListCount - 1
        If gComm.STOPBIT = Trim(Combo_Stopbit.List(i)) Then
            Combo_Stopbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Stopbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To Combo_Parity.ListCount - 1
        If gComm.Parity = Trim(Combo_Parity.List(i)) Then
            Combo_Parity.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Parity.ListIndex = 0
    End If
    
    If gComm.RTSEnable = True Then
        chkRTS.Value = "1"
    Else
        chkRTS.Value = "0"
    End If
    
    If gComm.DTREnable = True Then
        chkDTR.Value = "1"
    Else
        chkDTR.Value = "0"
    End If

End Sub

