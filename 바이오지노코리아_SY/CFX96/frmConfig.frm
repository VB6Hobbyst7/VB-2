VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Åë½Å¼³Á¤"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   13680
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin VB.Frame fraTcp 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00FFFFFF&
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
      Height          =   4665
      Left            =   4380
      TabIndex        =   30
      Top             =   1230
      Width           =   3825
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
         Height          =   330
         Left            =   1260
         TabIndex        =   35
         Text            =   "127.0.0.1"
         Top             =   1410
         Width           =   2175
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
         Height          =   330
         Left            =   1260
         TabIndex        =   34
         Text            =   "5050"
         Top             =   1890
         Width           =   2175
      End
      Begin VB.Frame fraType 
         BackColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   1260
         TabIndex        =   31
         Top             =   780
         Width           =   2175
         Begin VB.OptionButton optType 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   1140
            TabIndex        =   33
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   210
            Value           =   -1  'True
            Width           =   975
         End
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
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   0
         Left            =   -255
         TabIndex        =   38
         Top             =   975
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
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   37
         Top             =   1455
         Width           =   915
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
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   2
         Left            =   150
         TabIndex        =   36
         Top             =   1950
         Width           =   885
      End
   End
   Begin HSCotrol.CButton cmdConfirm 
      Height          =   495
      Left            =   10380
      TabIndex        =   28
      Top             =   6360
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " ¼³Á¤ÀúÀå"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmConfig.frx":08CA
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Share ¹æ½Ä"
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
      Index           =   2
      Left            =   8370
      TabIndex        =   27
      Top             =   780
      Width           =   4845
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Network (Socket) ¹æ½Ä"
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
      Index           =   1
      Left            =   4380
      TabIndex        =   26
      Top             =   780
      Width           =   3825
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Serial (RS232C) ¹æ½Ä"
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
      Index           =   0
      Left            =   390
      TabIndex        =   25
      Top             =   780
      Width           =   3825
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   6090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFile 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¾øÀ½
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
      Height          =   4665
      Left            =   8370
      TabIndex        =   20
      Top             =   1230
      Width           =   4845
      Begin VB.CommandButton cmdFind 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4260
         TabIndex        =   24
         Top             =   1290
         Width           =   375
      End
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
         Height          =   345
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1290
         Width           =   4035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Åõ¸í
         Caption         =   "°á°ú°æ·Î : "
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   12
         Left            =   210
         TabIndex        =   23
         Top             =   900
         Width           =   1305
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
         TabIndex        =   22
         Top             =   405
         Width           =   525
      End
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
      Left            =   3120
      TabIndex        =   17
      Top             =   6510
      Width           =   975
   End
   Begin VB.Frame fraSerial 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00FFFFFF&
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
      Height          =   4665
      Left            =   390
      TabIndex        =   0
      Top             =   1230
      Width           =   3825
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1830
         TabIndex        =   8
         Top             =   3720
         Value           =   1  'È®ÀÎ
         Width           =   1425
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1830
         TabIndex        =   7
         Top             =   4230
         Value           =   1  'È®ÀÎ
         Width           =   1365
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
         Width           =   1605
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
         Height          =   345
         Left            =   1800
         TabIndex        =   3
         Top             =   1890
         Width           =   1605
      End
      Begin VB.ComboBox Combo_BPS 
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
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Top             =   1440
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Port 
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
         Height          =   345
         Left            =   1800
         TabIndex        =   1
         Top             =   990
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
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
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   16
         Top             =   3780
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
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
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   48
         Left            =   60
         TabIndex        =   15
         Top             =   4290
         Width           =   1305
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
         ForeColor       =   &H00404040&
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
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
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
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
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
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
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
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
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
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   11
         Left            =   75
         TabIndex        =   9
         Top             =   1065
         Width           =   1305
      End
   End
   Begin HSCotrol.CButton cmdExit 
      Height          =   495
      Left            =   11850
      TabIndex        =   29
      Top             =   6360
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " ´Ý    ±â"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmConfig.frx":0A24
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.HSLabel HSLabel1 
      Height          =   345
      Left            =   150
      Top             =   150
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   609
      BackColor       =   16777215
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ¢º Åë½Å¼³Á¤"
      BevelOut        =   0
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   4755
      Left            =   4350
      Top             =   1200
      Width           =   3915
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   4755
      Left            =   360
      Top             =   1200
      Width           =   3915
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   4755
      Left            =   8340
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   405
      Left            =   8340
      Top             =   750
      Width           =   4935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   405
      Left            =   4350
      Top             =   750
      Width           =   3915
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   405
      Left            =   360
      Top             =   750
      Width           =   3915
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   7305
      Left            =   30
      Top             =   30
      Width           =   13635
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   405
      Left            =   120
      Top             =   120
      Width           =   13455
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
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   7
      Left            =   1050
      TabIndex        =   19
      Top             =   6570
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
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   8
      Left            =   4230
      TabIndex        =   18
      Top             =   6570
      Width           =   600
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub chkCommUse_Click(Index As Integer)
'    Dim i As Integer
'
'    If Index = 0 Then
'        fraSerial.BackColor = vbWhite
'        fraTcp.BackColor = &HC0C0C0
'        fraFile.BackColor = &HC0C0C0
'        chkRTS.BackColor = vbWhite
'        chkDTR.BackColor = vbWhite
'        optType(0).BackColor = &HC0C0C0
'        optType(1).BackColor = &HC0C0C0
'
'    ElseIf Index = 1 Then
'        fraSerial.BackColor = &HC0C0C0
'        fraTcp.BackColor = vbWhite
'        fraFile.BackColor = &HC0C0C0
'        chkRTS.BackColor = &HC0C0C0
'        chkDTR.BackColor = &HC0C0C0
'        optType(0).BackColor = vbWhite
'        optType(1).BackColor = vbWhite
'
'    ElseIf Index = 2 Then
'        fraSerial.BackColor = &HC0C0C0
'        fraTcp.BackColor = &HC0C0C0
'        fraFile.BackColor = vbWhite
'        chkRTS.BackColor = &HC0C0C0
'        chkDTR.BackColor = &HC0C0C0
'        optType(0).BackColor = vbWhite
'        optType(1).BackColor = vbWhite
'    End If
'
'End Sub

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
        ElseIf optUse(2).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")
        End If
        
        If optUse(0).Value = True Then
            gComm.COMPORT = Combo_Port.Text
            gComm.SPEED = Combo_BPS.Text
            gComm.DATABIT = Combo_Databit.Text
            gComm.STARTBIT = Combo_Startbit.Text
            gComm.STOPBIT = Combo_Stopbit.Text
            If Combo_Parity.ListIndex = 0 Then
               gComm.Parity = "N"
            ElseIf Combo_Parity.ListIndex = 1 Then
               gComm.Parity = "E"
            ElseIf Combo_Parity.ListIndex = 2 Then
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
            If chkDTR.Value = "1" Then
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
        
        ElseIf optUse(2).Value = True Then
            Call WritePrivateProfileString("COMM", "RSTPATH", txtRstPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
        End If
        
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

Private Sub cmdFind_Click()
    
On Error GoTo ErrHandler
    
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
        .Flags = cdlOFNHideReadOnly
        .InitDir = App.PATH
        .Filter = "All Files (*.*)|*.*|"
        .FilterIndex = 1
        .Filename = "*.*"
        .ShowOpen
        txtRstPath.Text = .Filename
    End With

Exit Sub
  
ErrHandler:
  ' »ç¿ëÀÚ°¡ [Ãë¼Ò] ´ÜÃß¸¦ ´­·¶½À´Ï´Ù.

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
        '-- Get Live Port
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
            Combo_BPS.ListIndex = 4         'Default = 9600
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
            Combo_Databit.ListIndex = 1     'Default = 8
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
            Combo_Startbit.ListIndex = 0    'Default = 1
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
            Combo_Stopbit.ListIndex = 0     'Default = 1
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
            Combo_Parity.ListIndex = 0      'Default = None
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
            
    ElseIf gComm.COMTYPE = "2" Then
        optUse(1).Value = True
        
        If gComm.TCPTYPE = "SERVER" Then
            optType(0).Value = True
        Else
            optType(1).Value = True
        End If
        
        txtIP.Text = gComm.TCPIP
        txtPort.Text = gComm.TCPPORT
        
    ElseIf gComm.COMTYPE = "3" Then
        optUse(2).Value = True
        
        txtRstPath.Text = gComm.RSTPATH
    End If
    
    txtSaveDay.Text = gHOSP.SAVEDAY
    
End Sub

Private Sub optUse_Click(Index As Integer)
    
    Dim i As Integer
    
    For i = 0 To 2
        optUse(i).ForeColor = vbBlack
        optUse(i).BackColor = vbWhite
    Next
    
    optUse(Index).ForeColor = vbWhite
    optUse(Index).BackColor = &HBF8B59

End Sub
