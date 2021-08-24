VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmCaution 
   Caption         =   "감염관리"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   7590
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdExit 
      Caption         =   "종 료"
      Height          =   615
      Left            =   5220
      TabIndex        =   25
      Top             =   6000
      Width           =   1665
   End
   Begin VB.Frame Frame3 
      Caption         =   "특이소견"
      Height          =   2685
      Left            =   60
      TabIndex        =   17
      Top             =   3240
      Width           =   6795
      Begin RichTextLib.RichTextBox RichText 
         Height          =   1935
         Left            =   150
         TabIndex        =   24
         Top             =   570
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmCaution.frx":0000
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fungus"
         Height          =   195
         Index           =   14
         Left            =   5610
         TabIndex        =   23
         Top             =   270
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "C.difficile"
         Height          =   195
         Index           =   13
         Left            =   4200
         TabIndex        =   22
         Top             =   270
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tb"
         Height          =   195
         Index           =   12
         Left            =   3360
         TabIndex        =   21
         Top             =   270
         Width           =   885
      End
      Begin VB.CheckBox Check1 
         Caption         =   "AFB"
         Height          =   195
         Index           =   11
         Left            =   2370
         TabIndex        =   20
         Top             =   270
         Width           =   885
      End
      Begin VB.CheckBox Check1 
         Caption         =   "VRE"
         Height          =   195
         Index           =   10
         Left            =   1290
         TabIndex        =   19
         Top             =   270
         Width           =   885
      End
      Begin VB.CheckBox Check1 
         Caption         =   "MRSA"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   18
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Drug Allergy"
      Height          =   1095
      Left            =   60
      TabIndex        =   13
      Top             =   2100
      Width           =   6795
      Begin VB.TextBox txtDrug 
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   570
         Width           =   6465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "RadioContrast"
         Height          =   195
         Index           =   8
         Left            =   2640
         TabIndex        =   15
         Top             =   270
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Penicillin"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Viral Marker"
      Height          =   1335
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   6795
      Begin VB.TextBox txtVival 
         Height          =   315
         Left            =   210
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   870
         Width           =   6465
      End
      Begin VB.CheckBox Check1 
         Caption         =   "anti_HAV lgM"
         Height          =   195
         Index           =   6
         Left            =   2610
         TabIndex        =   11
         Top             =   570
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "anti_HBc lgM"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   10
         Top             =   570
         Width           =   1845
      End
      Begin VB.CheckBox Check1 
         Caption         =   "기 타"
         Height          =   195
         Index           =   4
         Left            =   5670
         TabIndex        =   9
         Top             =   330
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "anti_HCV"
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   8
         Top             =   330
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "HBsAg"
         Height          =   195
         Index           =   2
         Left            =   2610
         TabIndex        =   7
         Top             =   330
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "VDRL"
         Height          =   195
         Index           =   1
         Left            =   1290
         TabIndex        =   6
         Top             =   330
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "HIV"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   330
         Width           =   1125
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   3690
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "최종기록일"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   3690
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "최종기록자"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblWDt 
      Height          =   300
      Left            =   5010
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   529
      BackColor       =   16777215
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblWNm 
      Height          =   300
      Left            =   5010
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   529
      BackColor       =   16777215
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
End
Attribute VB_Name = "frmCaution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset
Public strCPtid         As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtVival.Text = ""
    txtDrug.Text = ""
    RichText.Text = ""
    
    If Caution(strCPtid) = False Then
        Unload Me
    Else
'        Me.Show
    End If
End Sub

Public Function Caution(ByVal strPtid As String) As Boolean
    Dim SQL As String
    Dim iCnt As Integer

    Set AdoCn_ORACLE = New ADODB.Connection

    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
'        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("Persist Security Info") = True
        
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        
'        Screen.MousePointer = vbHourglass
        .Open
    End With
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
    SQL = "SELECT HIVYN,                                                         "
    SQL = SQL + "       VDRLYN,                                                        "
    SQL = SQL + "       HBSAGYN,                                                       "
    SQL = SQL + "       HCVYN,                                                         "
    SQL = SQL + "       VMOTHRYN,                                                      "
    SQL = SQL + "       HBCYN,                                                         "
    SQL = SQL + "       HAVYN,                                                         "
    SQL = SQL + "       PENICILN,                                                      "
    SQL = SQL + "       RADCONT,                                                       "
    SQL = SQL + "       MRSAYN,                                                        "
    SQL = SQL + "       VREYN,                                                         "
    SQL = SQL + "       AFBYN,                                                         "
    SQL = SQL + "       TBYN,                                                          "
    SQL = SQL + "       CDIFFIYN,                                                      "
    SQL = SQL + "       FUNGUSYN,                                                      "
    SQL = SQL + "       VMREMARK,                                                      "
    SQL = SQL + "       OTHERRMK,                                                      "
    SQL = SQL + "       DRUGALGY,                                                      "
    SQL = SQL + "       PATNO,                                                         "
    SQL = SQL + "       SEQ,                                                           "
    SQL = SQL + "       TO_CHAR(EDITDATE,'YYYYMMDD') AS EDITDATE,                      "
    SQL = SQL + "       EDITID,                                                        "
    SQL = SQL + "       FN_USERNAME_SELECT(EDITID) AS EDITNM                           "
    SQL = SQL + "  FROM MDCAUTNT                                                       "
    SQL = SQL + " WHERE PATNO = '" & strPtid & "'                                             "
    SQL = SQL + "   AND SEQ = (SELECT MAX(SEQ) FROM MDCAUTNT WHERE PATNO = '" & strPtid & "') "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open SQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            For iCnt = 0 To 17
                If .Fields(iCnt).Value = "Y" Then
                    Check1(iCnt).Value = 1
                End If
            Next
            lblWDt.Caption = .Fields("EDITDATE").Value & ""
            lblWNm.Caption = .Fields("EDITID").Value & ""
            txtVival.Text = .Fields("VMREMARK").Value & ""
            txtDrug.Text = .Fields("DRUGALGY").Value & ""
            RichText.Text = .Fields("OTHERRMK").Value & ""
            Caution = True
        Else
            Caution = False
        End If
        .Close
    End With
    Set AdoCn_ORACLE = Nothing

End Function

