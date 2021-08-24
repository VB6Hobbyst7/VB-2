VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmScreenSet 
   BackColor       =   &H00BF8B59&
   BorderStyle     =   3  'Å©±â °íÁ¤ ´ëÈ­ »óÀÚ
   Caption         =   "È­¸é¼³Á¤"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   Icon            =   "frmScreenSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin VB.TextBox txtBottom 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5040
      TabIndex        =   48
      Text            =   "ÇÏ´Ü»ö"
      Top             =   5160
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdBottom 
      Caption         =   "Set"
      Height          =   285
      Left            =   5880
      TabIndex        =   47
      Top             =   5190
      Visible         =   0   'False
      Width           =   465
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6420
      Top             =   4950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTop 
      Caption         =   "Set"
      Height          =   285
      Left            =   5880
      TabIndex        =   46
      Top             =   4860
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtTop 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5040
      TabIndex        =   45
      Text            =   "»ó´Ü»ö"
      Top             =   4830
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtColWidth 
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
      Left            =   2700
      TabIndex        =   43
      Top             =   7470
      Width           =   1485
   End
   Begin VB.Frame fraView 
      BackColor       =   &H00BF8B59&
      BorderStyle     =   0  '¾øÀ½
      Height          =   7155
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4095
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   20
         Left            =   2430
         TabIndex        =   42
         Top             =   6700
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFD732&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   20
         Left            =   150
         TabIndex        =   41
         Top             =   6720
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   19
         Left            =   2430
         TabIndex        =   40
         Top             =   6370
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   19
         Left            =   150
         TabIndex        =   39
         Top             =   6390
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   18
         Left            =   2430
         TabIndex        =   38
         Top             =   6040
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF064&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   18
         Left            =   150
         TabIndex        =   37
         Top             =   6060
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAD2&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   120
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   1
         Left            =   150
         TabIndex        =   35
         Top             =   450
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFAB4&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   2
         Left            =   150
         TabIndex        =   34
         Top             =   780
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   3
         Left            =   150
         TabIndex        =   33
         Top             =   1110
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FAFA96&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   4
         Left            =   150
         TabIndex        =   32
         Top             =   1440
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   5
         Left            =   150
         TabIndex        =   31
         Top             =   1770
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFA82&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   6
         Left            =   150
         TabIndex        =   30
         Top             =   2100
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   7
         Left            =   150
         TabIndex        =   29
         Top             =   2430
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF064&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   8
         Left            =   150
         TabIndex        =   28
         Top             =   2760
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   9
         Left            =   150
         TabIndex        =   27
         Top             =   3090
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF56E&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   10
         Left            =   150
         TabIndex        =   26
         Top             =   3420
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   11
         Left            =   150
         TabIndex        =   25
         Top             =   3750
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF56E&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   12
         Left            =   150
         TabIndex        =   24
         Top             =   4080
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   13
         Left            =   150
         TabIndex        =   23
         Top             =   4410
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFE650&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   14
         Left            =   150
         TabIndex        =   22
         Top             =   4740
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   15
         Left            =   150
         TabIndex        =   21
         Top             =   5070
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "¸¼Àº °íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2430
         TabIndex        =   20
         Text            =   "10"
         Top             =   100
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   1
         Left            =   2430
         TabIndex        =   19
         Top             =   430
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   2
         Left            =   2430
         TabIndex        =   18
         Top             =   760
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   3
         Left            =   2430
         TabIndex        =   17
         Top             =   1090
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   4
         Left            =   2430
         TabIndex        =   16
         Top             =   1420
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   5
         Left            =   2430
         TabIndex        =   15
         Top             =   1750
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   6
         Left            =   2430
         TabIndex        =   14
         Top             =   2080
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   7
         Left            =   2430
         TabIndex        =   13
         Top             =   2410
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   8
         Left            =   2430
         TabIndex        =   12
         Top             =   2740
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   9
         Left            =   2430
         TabIndex        =   11
         Top             =   3070
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   10
         Left            =   2430
         TabIndex        =   10
         Top             =   3400
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   11
         Left            =   2430
         TabIndex        =   9
         Top             =   3730
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   12
         Left            =   2430
         TabIndex        =   8
         Top             =   4060
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   13
         Left            =   2430
         TabIndex        =   7
         Top             =   4390
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   14
         Left            =   2430
         TabIndex        =   6
         Top             =   4720
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   15
         Left            =   2430
         TabIndex        =   5
         Top             =   5050
         Width           =   1515
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFF064&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   16
         Left            =   150
         TabIndex        =   4
         Top             =   5400
         Width           =   2235
      End
      Begin VB.CheckBox chkColumn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÀúÀå¼ø¹ø"
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
         Index           =   17
         Left            =   150
         TabIndex        =   3
         Top             =   5730
         Width           =   2235
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   16
         Left            =   2430
         TabIndex        =   2
         Top             =   5380
         Width           =   1515
      End
      Begin VB.TextBox txtColumn 
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
         Height          =   330
         Index           =   17
         Left            =   2430
         TabIndex        =   1
         Top             =   5710
         Width           =   1515
      End
   End
   Begin BHButton.BHImageButton cmdSave 
      Height          =   555
      Left            =   1260
      TabIndex        =   49
      Top             =   7980
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
      Picture         =   "frmScreenSet.frx":08A8
      TransparentPicture=   "frmScreenSet.frx":0A02
      ButtonAttrib    =   2
      ButtonTransparent=   1
      ForeColor       =   16777215
      BackColor       =   16777215
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdCancel 
      Height          =   555
      Left            =   2730
      TabIndex        =   50
      Top             =   7980
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
      Picture         =   "frmScreenSet.frx":33F4
      TransparentPicture=   "frmScreenSet.frx":354E
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
      Caption         =   "°Ë»çÇ×¸ñ ³ÐÀÌ"
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
      Height          =   225
      Index           =   7
      Left            =   1440
      TabIndex        =   44
      Top             =   7500
      Width           =   1140
   End
End
Attribute VB_Name = "frmScreenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBottom_Click()
    Dim LetColor
    Dim GetColor
    
    '¼±ÅÃÇÏ±âÀü »öÀ» °®°í ÀÖ´Â´Ù.
    LetColor = txtBottom.BackColor
    
    CommonDialog1.ShowColor
    
    '¼±ÅÃÇÑ »öÀÌ ¾ø´Ù¸é
    If CommonDialog1.Color = 0 Then
        txtBottom.BackColor = LetColor
        GetColor = LetColor
    Else
        GetColor = CommonDialog1.Color
        txtBottom.BackColor = GetColor
    End If
    
    Call WritePrivateProfileString("VIEW", "BOTTOMCOLOR", CStr(GetColor), App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i          As Integer
    Dim strSPDView As String
    Dim strSPDSize As String
    
    strSPDView = ""
    
    For i = 0 To 20
        strSPDView = strSPDView & IIf(chkColumn(i).Value = "1", "1", "0")
        strSPDSize = strSPDSize & txtColumn(i).Text & "|"
    Next
    
    Call WritePrivateProfileString("VIEW", "SPDVIEW", strSPDView, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("VIEW", "SPDSIZE", strSPDSize, App.PATH & "\INI\" & gMACH & ".ini")

    '-- ÄÃ·³º¸ÀÌ±â¼³Á¤
    Call SetColumnView(frmInterface.spdOrder)
    
    Call WritePrivateProfileString("VIEW", "COLWIDTH", txtColWidth.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    MsgBox "ÄÃ·³Á¤º¸°¡ º¯°æµÇ¾ú½À´Ï´Ù.", vbInformation + vbOKOnly, Me.Caption

End Sub

Private Sub cmdTop_Click()
    Dim LetColor
    Dim GetColor
    
    '¼±ÅÃÇÏ±âÀü »öÀ» °®°í ÀÖ´Â´Ù.
    LetColor = txtTop.BackColor
    
    CommonDialog1.ShowColor
    
    '¼±ÅÃÇÑ »öÀÌ ¾ø´Ù¸é
    If CommonDialog1.Color = 0 Then
        txtTop.BackColor = LetColor
        GetColor = LetColor
    Else
        GetColor = CommonDialog1.Color
        txtTop.BackColor = GetColor
    End If
    
    Call WritePrivateProfileString("VIEW", "TOPCOLOR", CStr(GetColor), App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub Form_Load()

    '-- È­¸é¼³Á¤
    Call SetColumnName
    
    'Call SetColumnView(frmInterface.spdOrder)
    
    txtColWidth.Text = gCOLWIDTH

    txtTop.BackColor = frmInterface.picHeader.BackColor
    txtBottom.BackColor = frmInterface.picBottom.BackColor
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub SetColumnName()
    Dim i       As Integer
    Dim varView As Variant
    Dim varSize As Variant
    
    chkColumn(0).Caption = "¼±ÅÃ"
    chkColumn(1).Caption = "°Ë»çÀÏ½Ã"
    chkColumn(2).Caption = "°Ë»ç½Ã°£"
    chkColumn(3).Caption = "°Ë»ç¼ø¹ø"
    chkColumn(4).Caption = "ER"
    chkColumn(5).Caption = "RT"
    chkColumn(6).Caption = "Á¢¼öÀÏÀÚ"
    chkColumn(7).Caption = "°ËÃ¼¹øÈ£"
    chkColumn(8).Caption = "°ËÃ¼"
    chkColumn(9).Caption = "RackNo"
    chkColumn(10).Caption = "TubePos"
    chkColumn(11).Caption = "SeqNo"
    chkColumn(12).Caption = "ÀÌ¸§"
    chkColumn(13).Caption = "Sex"
    chkColumn(14).Caption = "Age"
    chkColumn(15).Caption = "º´·Ï¹øÈ£"
    chkColumn(16).Caption = "Ã­Æ®¹øÈ£"
    chkColumn(17).Caption = "ÀÇ·Ú°ú"
    chkColumn(18).Caption = "ÀÔ/¿Ü±¸ºÐ"
    chkColumn(19).Caption = "¿À´õ°¹¼ö"
    chkColumn(20).Caption = "°á°ú°¹¼ö"
    
    
    For i = 0 To 20
        'If Mid(varViewi + 1, 1) = "1" Then
        chkColumn(i).Value = Mid(gCOLVIEW, i + 1, 1)
        txtColumn(i).BackColor = chkColumn(i).BackColor
    Next
    
    
    varSize = Split(gCOLSIZE, "|")
    
    For i = 0 To 20
        txtColumn(i).Alignment = 2
        txtColumn(i).Text = varSize(i)
        txtColumn(i).FontSize = 9 '11
        If Mid(gCOLVIEW, i + 1, 1) = "1" Then
            txtColumn(i).FontBold = False 'True
        Else
            txtColumn(i).FontBold = False
        End If
    Next

End Sub

'Private Sub SetColumnView()
'    Dim i       As Integer
'    Dim varSize As Variant
'
'    varSize = Split(gCOLSIZE, "|")
'
'    For i = 0 To UBound(varSize) - 1
'
'        If Mid(gCOLVIEW, i + 1, 1) = 1 Then
'            frmScreenSet.chkColumn(i).Value = "1"
'        Else
'            frmScreenSet.chkColumn(i).Value = "0"
'        End If
'        'spdOrder.ColWidth(i + 1) = varSize(i + 1)
'    Next
'
'
'End Sub
