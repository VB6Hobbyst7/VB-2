VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmConfig 
   BackColor       =   &H00BF8B59&
   Caption         =   "Åë½Å¼³Á¤"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   17070
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
   ScaleHeight     =   7680
   ScaleWidth      =   17070
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin VB.Frame Frame3 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00FFFFFF&
      Caption         =   " °á°ú °æ·Î "
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5565
      Left            =   8910
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   6585
      Begin VB.TextBox txtRstPath 
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   6345
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   165
         TabIndex        =   32
         Top             =   405
         Width           =   525
      End
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00C0C0C0&
      Caption         =   " ¼ÒÄÏ ¹æ½Ä »ç¿ë"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   1
      Left            =   4950
      TabIndex        =   29
      Top             =   390
      Width           =   3165
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "½Ã¸®¾ó ¹æ½Ä »ç¿ë"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   390
      Value           =   -1  'True
      Width           =   2685
   End
   Begin VB.TextBox txtSaveDay 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2460
      TabIndex        =   26
      Top             =   5670
      Width           =   975
   End
   Begin VB.Frame fraTcp 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '¾øÀ½
      Caption         =   " ¼ÒÄÏ ¹æ½Ä "
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   4590
      TabIndex        =   18
      Top             =   60
      Width           =   4245
      Begin VB.OptionButton optType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Client"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2250
         TabIndex        =   22
         Top             =   990
         Width           =   975
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1230
         TabIndex        =   21
         Top             =   990
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1260
         TabIndex        =   20
         Text            =   "5050"
         Top             =   1890
         Width           =   1815
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1260
         TabIndex        =   19
         Text            =   "127.0.0.1"
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   -270
         TabIndex        =   25
         Top             =   1950
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "IP"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -255
         TabIndex        =   24
         Top             =   1425
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   -255
         TabIndex        =   23
         Top             =   975
         Width           =   1305
      End
   End
   Begin VB.Frame fraSerial 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000005&
      BorderStyle     =   0  '¾øÀ½
      Caption         =   " ½Ã¸®¾ó ¹æ½Ä "
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4455
      Begin VB.CheckBox chkRTS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   9
         Top             =   3720
         Value           =   1  'È®ÀÎ
         Width           =   1785
      End
      Begin VB.CheckBox chkDTR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   330
         TabIndex        =   17
         Top             =   3780
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "DTREnable"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   48
         Left            =   330
         TabIndex        =   16
         Top             =   4290
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Parity"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   60
         TabIndex        =   15
         Top             =   3315
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Stop Bit"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   2865
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Start Bit"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   13
         Top             =   2415
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Data Bit"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   12
         Top             =   1965
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   75
         TabIndex        =   11
         Top             =   1515
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Com Port"
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
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
         TabIndex        =   10
         Top             =   1065
         Width           =   1305
      End
   End
   Begin BHButton.BHImageButton cmdConfirm 
      Height          =   555
      Left            =   5610
      TabIndex        =   33
      Top             =   5760
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   979
      Caption         =   "¼³Á¤ÀúÀå"
      CaptionChecked  =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmConfig.frx":08CA
      TransparentPicture=   "frmConfig.frx":0A24
      ButtonAttrib    =   2
      ButtonTransparent=   1
      ForeColor       =   16777215
      BackColor       =   16777215
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdExit 
      Height          =   555
      Left            =   7080
      TabIndex        =   34
      Top             =   5760
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   979
      Caption         =   "´Ý±â"
      CaptionChecked  =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmConfig.frx":3416
      TransparentPicture=   "frmConfig.frx":3570
      ButtonAttrib    =   2
      ButtonTransparent=   1
      ForeColor       =   16777215
      BackColor       =   16777215
      ImgOutLineSize  =   3
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Æò¸é
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°Ë»ç°á°ú ·ÎÄÃÀúÀå±â°£"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Index           =   7
      Left            =   390
      TabIndex        =   28
      Top             =   5730
      Width           =   2010
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Æò¸é
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÀÏ ÀúÀå"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Index           =   8
      Left            =   3570
      TabIndex        =   27
      Top             =   5730
      Width           =   600
   End
End
Attribute VB_Name = "frmConfig"
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
        If optUse(0).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
        ElseIf optUse(1).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
        End If
        
        If optUse(0).Value = True Then
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
            
            Call WritePrivateProfileString("COMM", "COMPORT", gComm.COMPORT, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "SPEED", gComm.SPEED, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "PARITY", gComm.Parity, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "DATABIT", gComm.DATABIT, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "STOPBIT", gComm.STOPBIT, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "STARTBIT", gComm.STARTBIT, App.PATH & "\INI\" & gMACH & ".ini")
            
            If chkRTS.Value = "1" Then
                Call WritePrivateProfileString("COMM", "RTSEnable", "True", App.PATH & "\INI\" & gMACH & ".ini")
            Else
                Call WritePrivateProfileString("COMM", "RTSEnable", "False", App.PATH & "\INI\" & gMACH & ".ini")
            End If
            If chkRTS.Value = "1" Then
                Call WritePrivateProfileString("COMM", "DTREnable", "True", App.PATH & "\INI\" & gMACH & ".ini")
            Else
                Call WritePrivateProfileString("COMM", "DTREnable", "False", App.PATH & "\INI\" & gMACH & ".ini")
            End If
            
            If frmInterface.comEqp.PortOpen = True Then
                frmInterface.comEqp.PortOpen = False
            End If
            
            frmInterface.comEqp.CommPort = gComm.COMPORT
            frmInterface.comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
            
            frmInterface.comEqp.PortOpen = True
        
        ElseIf optUse(1).Value = True Then
            gComm.TCPIP = txtIP.Text
            gComm.TCPPORT = txtPort.Text
            
            If optType(0).Value = True Then
                gComm.TCPTYPE = "SERVER"
            Else
                gComm.TCPTYPE = "CLIENT"
            End If
    
            Call WritePrivateProfileString("COMM", "TCPTYPE", gComm.TCPTYPE, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "TCPIP", gComm.TCPIP, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "TCPPORT", gComm.TCPPORT, App.PATH & "\INI\" & gMACH & ".ini")
        
        End If
        
        Call WritePrivateProfileString("COMM", "RSTPATH", txtRstPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        Call WritePrivateProfileString("HOSP", "SAVEDAY", txtSaveDay.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
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
    
    If gComm.COMTYPE = "1" Then
        optUse(0).Value = True
        
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
            
        fraSerial.BackColor = vbWhite
        
    ElseIf gComm.COMTYPE = "2" Then
        optUse(1).Value = True
        
        If gComm.TCPTYPE = "SERVER" Then
            optType(0).Value = True
        Else
            optType(1).Value = True
        End If
        
        txtIP.Text = gComm.TCPIP
        txtPort.Text = gComm.TCPPORT
        
        fraTcp.BackColor = &HC0C0C0
    End If
    
    txtRstPath.Text = gComm.RSTPATH
    
    txtSaveDay.Text = gHOSP.SAVEDAY
    
End Sub

Private Sub optUse_Click(Index As Integer)
    
    optUse(0).FontBold = False
    optUse(1).FontBold = False
    optUse(Index).FontBold = True
    
    If Index = 0 Then
        fraSerial.BackColor = vbWhite
        fraTcp.BackColor = &HC0C0C0
        
        optUse(0).BackColor = vbWhite
        optUse(1).BackColor = &HC0C0C0
        
        chkRTS.BackColor = vbWhite
        chkDTR.BackColor = vbWhite
        
        optType(0).BackColor = &HC0C0C0
        optType(1).BackColor = &HC0C0C0
    Else
        fraSerial.BackColor = &HC0C0C0
        fraTcp.BackColor = vbWhite
        
        optUse(0).BackColor = &HC0C0C0
        optUse(1).BackColor = vbWhite
        
        chkRTS.BackColor = &HC0C0C0
        chkDTR.BackColor = &HC0C0C0
        
        optType(0).BackColor = vbWhite
        optType(1).BackColor = vbWhite
    End If
End Sub
