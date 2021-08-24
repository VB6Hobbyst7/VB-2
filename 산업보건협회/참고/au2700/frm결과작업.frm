VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm°á°úÀÛ¾÷ 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "°á°úÀÛ¾÷"
   ClientHeight    =   10635
   ClientLeft      =   3135
   ClientTop       =   4410
   ClientWidth     =   15075
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm°á°úÀÛ¾÷.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   15075
   Begin FPSpread.vaSpread vasTemp 
      Height          =   2565
      Left            =   17760
      TabIndex        =   0
      Top             =   9300
      Visible         =   0   'False
      Width           =   4095
      _Version        =   393216
      _ExtentX        =   7223
      _ExtentY        =   4524
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":0442
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Á¶È¸(&V)"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   28
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Data Sort"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4680
      TabIndex        =   27
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton txtResPrint 
      Appearance      =   0  'Æò¸é
      Caption         =   "°á°ú Ãâ·Â"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   26
      Top             =   60
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "[°Ë»çÀÏÀÚ]"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   24
      Top             =   60
      Width           =   1755
      Begin MSComCtl2.DTPicker dtpExamDate 
         Height          =   315
         Left            =   60
         TabIndex        =   25
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21430273
         CurrentDate     =   39699
      End
   End
   Begin FPSpread.vaSpread vasOrder 
      Height          =   1815
      Left            =   17520
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   7425
      _Version        =   393216
      _ExtentX        =   13097
      _ExtentY        =   3201
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":4949
   End
   Begin FPSpread.vaSpread vasOrder1 
      Height          =   2505
      Left            =   17700
      TabIndex        =   10
      Top             =   6300
      Visible         =   0   'False
      Width           =   7005
      _Version        =   393216
      _ExtentX        =   12356
      _ExtentY        =   4419
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":4B6A
   End
   Begin VB.CheckBox ChkAll 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   165
   End
   Begin VB.CommandButton cmd_Trans 
      Caption         =   "¼±ÅÃÀü¼Û"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      Picture         =   "frm°á°úÀÛ¾÷.frx":4D8B
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12780
      TabIndex        =   3
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Á¾·á"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   9495
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   14955
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   7050
         TabIndex        =   13
         Top             =   120
         Width           =   7635
         Begin VB.CommandButton cmd°ËÃ¼¹øÈ£»ý¼º 
            Caption         =   "°ËÃ¼¹øÈ£»ý¼º"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   23
            Top             =   120
            Width           =   1515
         End
         Begin VB.TextBox txtReceHead 
            Appearance      =   0  'Æò¸é
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   180
            TabIndex        =   20
            Top             =   150
            Width           =   1515
         End
         Begin VB.TextBox txtResN 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6810
            TabIndex        =   16
            Top             =   150
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtStartS 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5670
            TabIndex        =   15
            Top             =   150
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtStartR 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3540
            TabIndex        =   14
            Top             =   150
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "~"
            Height          =   195
            Left            =   6540
            TabIndex        =   19
            Top             =   210
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "°Ë»ç¹øÈ£ :"
            Height          =   285
            Left            =   4500
            TabIndex        =   18
            Top             =   210
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "Start Row : "
            Height          =   315
            Left            =   2310
            TabIndex        =   17
            Top             =   210
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5340
         TabIndex        =   12
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   39699
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   8565
         Left            =   6060
         TabIndex        =   29
         Top             =   660
         Width           =   8115
         _Version        =   393216
         _ExtentX        =   14314
         _ExtentY        =   15108
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   13
         Protect         =   0   'False
         SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":5A55
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8745
         Left            =   60
         TabIndex        =   30
         Top             =   660
         Width           =   5925
         _Version        =   393216
         _ExtentX        =   10451
         _ExtentY        =   15425
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   21
         Protect         =   0   'False
         SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":99A8
      End
      Begin VB.Label Label3 
         Caption         =   "Á¢¼öÀÏÀÚ"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4230
         TabIndex        =   11
         Top             =   330
         Width           =   1035
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   10260
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2011-05-02"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "¿ÀÈÄ 3:14"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "¸Þµð¸ÞÀÌÆ® ¢Ï(051)462-1751"
            TextSave        =   "¸Þµð¸ÞÀÌÆ® ¢Ï(051)462-1751"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpread.vaSpread vasRece 
      Height          =   1875
      Left            =   8220
      TabIndex        =   9
      Top             =   5610
      Visible         =   0   'False
      Width           =   5115
      _Version        =   393216
      _ExtentX        =   9022
      _ExtentY        =   3307
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":DD35
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   5505
      Left            =   900
      TabIndex        =   7
      Top             =   3030
      Width           =   10875
      _Version        =   393216
      _ExtentX        =   19182
      _ExtentY        =   9710
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":DF56
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   2715
      Left            =   17580
      TabIndex        =   22
      Top             =   3060
      Visible         =   0   'False
      Width           =   9615
      _Version        =   393216
      _ExtentX        =   16960
      _ExtentY        =   4789
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   30
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frm°á°úÀÛ¾÷.frx":E177
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°Ë »ç ÀÚ"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6450
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frm°á°úÀÛ¾÷"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBARCODE = 2
Const colSampleNo = 3
Const colRack = 4
Const colPos = 5
Const colPID = 6
Const colPName = 7
Const colJumin = 8
Const colPSex = 9
Const colPAge = 10
Const colState = 11
Const colEXAMDATE = 12
Const colSlipNo1 = 13
Const colSlipNo2 = 14
Const colReqDate = 15

Const colEquipExam = 3
Const colExamCode = 4
Const colExamName = 5
Const colResult = 6
Const colRCheck = 7
Const colPCheck = 8
Const colDCheck = 9
Const colUnit = 10
Const colRef = 11
Const colPanic = 12
Const colResult1 = 13

Dim gRType As String
Dim ConfirmData As String
Dim sBARCODE As String
Dim llRow As Long
Dim gRefFlag As String
Dim gPanicFlag As String
Dim SysDateTime As String
Dim TimerFlag As Integer
Dim SubStr(1 To 80) As String


Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If ChkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            
            If Trim(GetText(vasID, iRow, colState)) = "Result" And Trim(GetText(vasID, iRow, colBARCODE)) <> "" Then
                vasID.Row = iRow
                vasID.Col = 1
                vasID.Value = 1
            End If
        Next iRow
    ElseIf ChkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmd_Trans_Click()
'¼±ÅÃÀü¼Û
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer

    If MsgBox(" " & vbCrLf & "¼±ÅÃÀü¼ÛÀ» ÇÏ½Ã°Ú½À´Ï±î?" & vbCrLf & " ", vbInformation + vbOKCancel, "¾Ë¸²:¼±ÅÃÀü¼Û") = vbCancel Then
        Exit Sub
    End If

    If (vasID.DataRowCnt < 1) Or (vasRes.DataRowCnt < 1) Then
        MsgBox "ÀúÀåÇÒ µ¥ÀÌÅÍ°¡ ¾ø½À´Ï´Ù."
        Exit Sub
    End If
    
'    db_BeginTran gServer
'    Connect_Server
    For vasIDRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = vasIDRow
        
        If vasID.Value = 1 Then
            liRet = -1

            If Trim(GetText(vasID, vasIDRow, colBARCODE)) <> "" Then
                liRet = Insert_Data(vasIDRow)
            End If
            
            If liRet = 1 Then
'                db_Commit gServer
                
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "¿Ï·á", vasIDRow, colState
'                DeleteRow vasID, vasIDRow, vasIDRow
                
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "½ÇÆÐ", vasIDRow, colState
            End If
            
'            vasID.Row = vasIDRow
'            vasID.Col = 1
'            vasID.Value = 0
        Else
        
        End If
    Next vasIDRow
    cmdClear_Click
    'db_Commit gServer
End Sub

Function Insert_Data(argSpcRow As Integer) As Integer
'¼­¹öÀÇ µ¥ÀÌÅ¸ º£ÀÌ½º¿¡ ÀúÀå

    Dim iRow As Integer
    Dim i As Integer
    
    Dim sDate As String
    Dim sSCP41JDATE As String      'Á¢¼öÀÏÀÚ
    Dim sSCP41SID As String     'Á¢¼ö½Ç SID
    Dim sBARCODE As String
    Dim sPID As String
    
    Dim sResGubun As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sSlipNo1 As String
    Dim sSlipNo2 As String
    Dim sPAge As String
    Dim sPSex As String
    Dim sRegion As String
    Dim sEXAMDATE As String
    
    
    Insert_Data = -1
    
    sDate = ""
    sDate = Format(dtpExamDate.Value, "yyyymmdd")

    sBARCODE = Trim(GetText(vasID, argSpcRow, colBARCODE))
'    sPID = Trim(GetText(vasID, argSpcRow, colPID))
'    sPAge = Trim(GetText(vasID, argSpcRow, colPAge))
'    sPSex = Trim(GetText(vasID, argSpcRow, colPSex))
    
    'Local¿¡¼­ È¯ÀÚº°·Î °á°ú°ª °¡Á®¿À±â
    ClearSpread vasResTemp
    
    SQL = " Select EQUIPCODE, examcode, result, EXAMDATE " & vbCrLf & _
          " From PAT_RES " & vbCrLf & _
          " Where EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          " And BARCODE = '" & sBARCODE & "' and EXAMDATE = '" & sDate & "'"
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    vasResTemp.MaxRows = vasResTemp.DataRowCnt + 1
    
    '¼­¹ö·Î °á°ú°ª ÀúÀåÇÏ±â
    For i = 1 To vasResTemp.DataRowCnt
        sExamCode = ""
        sResult = ""
        
        sExamCode = Trim(GetText(vasResTemp, i, 2))
        sResult = Trim(GetText(vasResTemp, i, 3))
        sEXAMDATE = Trim(GetText(vasResTemp, i, 4))
        Save_Result_Data Mid(sBARCODE, 1, 2) & "20" & Mid(sBARCODE, 3, 8) & sEXAMDATE & Mid(sBARCODE, 11, 4) & sExamCode & sResult

    Next i
    
'    SQL = "UPDATE tl_workhead set bogoyn = 'Y' where sample = '" & sBARCODE & "'"
'    res = SendQuery(gServer, SQL)
    
    SQL = "UPDATE PAT_RES set sendflag = '2' where  EXAMDATE = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          " And EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          " And BARCODE = '" & sBARCODE & "'"
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gServer
        Exit Function
    End If
         
         
    Insert_Data = 1
End Function

Function CheckValue(asResult As String, asExamCode As String, asAge As String, asSex As String, asRegion As String, asDate As String)
    Dim sRefHigh As String
    Dim sRefLow As String
    Dim sPanicHigh As String
    Dim sPanicLow As String
    Dim i As Integer
    SQL = "select sclvalue, schvalue, plvalue, phvalue from tl_standard " & vbCrLf & _
          "where workname = '" & asExamCode & "' and region = '" & asRegion & "' " & vbCrLf & _
          "and f_age <= '" & asAge & "' and t_age >= '" & asAge & "'" & vbCrLf & _
          "and  f_date <= '" & asDate & "' and t_date >= '" & asDate & "' "
    res = db_select_Col(gServer, SQL)
    
    If gReadBuf(0) = "" And gReadBuf(1) = "" Then
        If asSex = "M" Then
            SQL = "select smlvalue, smhvalue, plvalue, phvalue from tl_standard " & vbCrLf & _
                  "where workname = '" & asExamCode & "' and region = '" & asRegion & "' " & vbCrLf & _
                  "and f_age <= '" & asAge & "' and t_age >= '" & asAge & "' " & vbCrLf & _
                  "and  f_date <= '" & asDate & "' and t_date >= '" & asDate & "' "
            res = db_select_Col(gServer, SQL)
        Else
            SQL = "select sflvalue, sfhvalue, plvalue, phvalue from tl_standard " & vbCrLf & _
                  "where workname = '" & asExamCode & "' and region = '" & asRegion & "' " & vbCrLf & _
                  "and f_age <= '" & asAge & "' and t_age >= '" & asAge & "' " & vbCrLf & _
                  "and  f_date <= '" & asDate & "' and t_date >= '" & asDate & "' "
            res = db_select_Col(gServer, SQL)
        End If
        For i = 0 To 3
            If gReadBuf(i) = "" Then
                gReadBuf(i) = "0"
            End If
        Next
        sRefHigh = gReadBuf(1)
        sRefLow = gReadBuf(0)
        sPanicLow = gReadBuf(2)
        sPanicHigh = gReadBuf(3)
    Else
        For i = 0 To 3
            If gReadBuf(i) = "" Then
                gReadBuf(i) = "0"
            End If
        Next
        sRefHigh = gReadBuf(1)
        sRefLow = gReadBuf(0)
        sPanicLow = gReadBuf(2)
        sPanicHigh = gReadBuf(3)
    End If
    gRefFlag = ""
    gPanicFlag = ""
    If IsNumeric(asResult) = False Or IsNumeric(sRefLow) = False Or IsNumeric(sRefHigh) = False Or IsNumeric(sPanicLow) = False Or IsNumeric(sPanicHigh) = False Then
        Exit Function
    End If
    
    If CCur(asResult) < CCur(sRefLow) Then
        gRefFlag = "L"
    End If
    If CCur(asResult) > CCur(sRefHigh) Then
        gRefFlag = "H"
    End If
    If CCur(asResult) < CCur(sPanicLow) Then
        gPanicFlag = "L"
    End If
    If CCur(asResult) > CCur(sPanicHigh) Then
        gPanicFlag = "H"
    End If
    
End Function

Private Sub cmd°ËÃ¼¹øÈ£»ý¼º_Click()
    Dim intX        As Integer
    Dim lsBARCODE   As String
    
    For intX = 1 To vasID.MaxRows
        If Trim(GetText(vasID, intX, colSampleNo)) <> "" And Trim(GetText(vasID, intX, colBARCODE)) = "" Then
            lsBARCODE = txtReceHead & Trim(GetText(vasID, intX, colSampleNo))
            
            SetText vasID, lsBARCODE, intX, colBARCODE
    
            SQL = "UPDATE PAT_RES SET "
            SQL = SQL & vbCrLf & "       BARCODE   = '" & lsBARCODE & "' "
            SQL = SQL & vbCrLf & " WHERE EXAMDATE  = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' "
            SQL = SQL & vbCrLf & "   AND EQUIPNO   = '" & gEquip & "' "
            SQL = SQL & vbCrLf & "   AND SEQNO     = '" & Trim(GetText(vasID, intX, colSampleNo)) & "' "
            res = SendQuery(gLocal, SQL)
        End If
    Next intX
End Sub

Private Sub cmdView_Click()
    Dim i As Integer
    
    ClearSpread vasID

    SQL = " Select '', BARCODE, seqno, diskno, posno, pid, pname, pjumin, psex, page, sendflag From PAT_RES " & CR & _
          " Where EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & CR & _
          " And EQUIPNO = '" & gEquip & "' " & CR & _
          " Group By BARCODE, seqno, diskno, posno, pid, pname, pjumin, psex, page, sendflag " & CR & _
          " Order By seqno "
    res = db_select_Vas(gLocal, SQL, vasID)

    vasID.MaxRows = vasID.DataRowCnt
    For i = 1 To vasID.DataRowCnt
'        vasID.RowHeight(i) = 13
'        gReadBuf(0) = ""

'        SQL = "select refflag from PAT_RES where BARCODE = '" & Trim(GetText(vasID, i, colBARCODE)) & "' and refflag in ('S', 'B')"
'        res = db_select_Col(gLocal, SQL)
'
'        If gReadBuf(0) = "S" Or gReadBuf(0) = "B" Then
'            SetForeColor vasID, i, i, 0, 0, 255
'        End If

        gReadBuf(0) = ""
        SQL = "select refflag from PAT_RES where BARCODE = '" & Trim(GetText(vasID, i, colBARCODE)) & "' and refflag = 'R'"
        res = db_select_Col(gLocal, SQL)

        If gReadBuf(0) = "R" Then
            SetForeColor vasID, i, i, 255, 0, 0
        End If

        If GetText(vasID, i, colState) = "1" Then
            SetText vasID, "Result", i, colState
            SetBackColor vasID, i, i, 1, 1, 255, 250, 205
        ElseIf GetText(vasID, i, colState) = "2" Then
            SetText vasID, "¿Ï·á", i, colState
            SetBackColor vasID, i, i, colCheckBox, colCheckBox, 202, 255, 112
        End If
    Next
End Sub

Private Sub cmdClear_Click()
    Dim iRow As Integer

'    ClearSpread vasID, 1, 1
'    vasID.MaxRows = 0

    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
        
        If vasID.Value = 1 Then
            vasDeleteRow vasID, iRow
            
            iRow = iRow - 1
        End If
    Next iRow
    
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    vasSort vasID, 3, 4
End Sub

'Private Sub DTPicker1_Change()
'Dim i As Integer
'    SysDateTime = Format(DTPicker1, "yyyy/mm/dd") & " 00:00:00"
'    Command2_Click
'End Sub

Private Sub DTPicker1_CloseUp()
Dim i As Integer
    SysDateTime = Format(DTPicker1, "yyyy/mm/dd") & " 00:00:00"
    txtReceHead.Text = "09" & Format(DTPicker1, "yyyymmdd")
'    Command2_Click
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim FindFile As String
    
    Me.Left = 0
    Me.Top = 0
    
    gAllExam = ""
    
    'cmdClear_Click
    
    ClearSpread vasID, 1, 1
    'vasID.MaxRows = 1
    
    GetSetup    'ini¿¡¼­ DBÁ¤º¸ ºÒ·¯¿À±â
       
    '·ÎÄÃ¿¡ Á¢¼Ó
    If Not Connect_Local Then
        MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
        Exit Sub
    End If
    
    raw_data = ""
    DTPicker1 = Format(Date, "yyyy-mm-dd")
    dtpExamDate = Date
    
    '°Ë»çÄÚµå °¡Á®¿À±â
    GetExamCode
      
    'MultiSelect Mode
    vasRes.OperationMode = 1
    SysDateTime = Format(Date, "yyyy/mm/dd") & " 00:00:00"
    TimerFlag = 1
    txtReceHead.Text = "09" & Format(Date, "yyyymmdd")
    txtStartR.Text = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    DisConnect_Server
    DisConnect_Local
End Sub

Sub GetExamCode()
'°Ë»çÄÚµå¸¦ array¿¡ ÀúÀå
    Dim i As Integer
    Dim j As Integer
    
    gAllExam = ""
    
    ClearSpread vasTemp
    
    SQL = "Select EQUIPCODE, examcode, examname From equipexam " & CR & _
          " where EQUIPNO = '" & gEquip & "' " & CR & _
          " Order by seqno"
          
    res = db_select_Vas(gLocal, SQL, vasTemp)

    If res > 0 Then
        ReDim gArr_ExamCode(1 To vasTemp.DataRowCnt, 1 To 3)
    Else
        SaveQuery SQL
        Exit Sub
    End If
        
    For i = 1 To vasTemp.DataRowCnt
        gArr_ExamCode(i, 1) = i
        For j = 1 To 2
            gArr_ExamCode(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        If gAllExam = "" Then
            gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
        End If
    Next i
    
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuResult_Click()

End Sub

Private Sub mnuSetSub_Click(Index As Integer)
    Select Case Index
        Case 0: '/ÇÁ·ÎÆÄÀÏ¼³Á¤
            MsgBox "°ø»çÁß...", vbCritical, "È®ÀÎ"
        
        Case 1: '/°Ë»çÄÚµå¼³Á¤
            frmEquipExam.SSPanel1.Caption = "  AU2700 Àåºñ ÄÚµå ¼³Á¤"
            frmEquipExam.Show 1
            GetExamCode
        
        Case 2: '/Åë½Å¼³Á¤
            frmConfig.SSPanel_machine.Caption = "AU2700"
            frmConfig.Show 1

    End Select
End Sub

Private Sub mnuWorkList_Click()
    frmPatSear.Show vbModal
End Sub

Private Sub MSComm1_OnComm()
'    Dim lsChar As String
'    Dim SendMsg As String
'    Dim x, Y As Long
'
'
'    lsChar = MSComm1.Input
'
'    Select Case lsChar
'    Case chrSOH
'        Save_Raw_Data "[RX]" & lsChar
'        MSComm1.Output = chrENQ
'        Save_Raw_Data "[TX]" & chrENQ
'
'    Case chrSTX
'        txtBuff.Text = ""
'        txtBuff.Text = chrSTX
'
''    Case chrENQ
''        Save_Raw_Data "[RX]" & chrENQ
''        Save_Raw_Data "[TX]" & "ÿ FILE     " & Chr(13) & ""
''        MSComm1.Output = "ÿ END     " & Chr(13) & ""
'        Case chrACK, chrENQ
'        Save_Raw_Data "[RX]" & lsChar
'        SendMsg = ""
'
'        If vasOrder.DataRowCnt < 1 Then
'            Exit Sub
'        End If
'
'        If Trim(GetText(vasOrder, 1, 2)) = "END" Then
'            SendMsg = Trim(GetText(vasOrder, 1, 1))
'            ClearSpread vasOrder
'            MsgBox "Order Àü¼Û ¿Ï·á"
'        Else
'            'Àåºñ¹øÈ£,¹ÙÄÚµå¹øÈ£,È¯ÀÚÀÌ¸§,³ªÀÌ,¼ºº°,¿À´õ,È¯ÀÚ¹øÈ£
'
'            SendMsg = Chr(&HFF) & " FILE    " & Chr(13)
''            SendMsg = SendMsg & Chr(&H70) & " " & Format(Trim(GetText(vasOrder1, 1, 1)), "0#") & Chr(13)
'            SendMsg = SendMsg & Chr(&H75) & " " & SetSpace(Trim(GetText(vasOrder, 1, 2)), 16, 2) & Chr(13)
'
'            SendMsg = SendMsg & Chr(&H72) & " " & SetSpace(Trim(GetText(vasOrder, 1, 1)), 16, 2) & Chr(13)
''            SendMsg = SendMsg & Chr(&H76) & " " & SetSpace(Trim(GetText(vasOrder1, 1, 7)), 30, 2) & Chr(13)
''            SendMsg = SendMsg & Chr(&H78) & " " & Chr(13)
''            If Trim(GetText(vasOrder1, 1, 5)) = "M" Then
''                SendMsg = SendMsg & Chr(&H79) & " 1" & Chr(13)
''            Else
''                SendMsg = SendMsg & Chr(&H79) & " 2" & Chr(13)
''            End If
'            SendMsg = SendMsg & Chr(&H80) & " " & Trim(GetText(vasOrder, 1, 6)) & Chr(13)
''            SendMsg = SendMsg & Chr(&H8B) & " " & SetSpace(Trim(GetText(vasOrder1, 1, 7)), 30, 2) & Chr(13)
'            SendMsg = SendMsg
'            x = Len(SendMsg)
'            For Y = 1 To Len(SendMsg)
'                If Mid(SendMsg, Y, 1) = "?" Then
'                    x = x + 1
'                End If
'            Next
'            SendMsg = Format(x + 6, "0000#") & Chr(13) & SendMsg  '& Chr(&HFE) & "V5.0 " & Chr(13)
'           ' SendMsg = SendMsg & Chr(&HFD) & " " & CheckSum(SendMsg) & Chr(13)
'            SendMsg = chrSTX & SendMsg & chrETX
'
'            DeleteRow vasOrder, 1, 1
'
'        End If
'
'        If SendMsg = "" Then
'        Else
'            Save_Raw_Data "[TX]" & SendMsg
'            MSComm1.Output = SendMsg
'        End If
'
'
'    Case chrETX
'
'        txtBuff.Text = txtBuff.Text & lsChar
'        Save_Raw_Data "[RX]" & txtBuff.Text
'        AllPentra txtBuff
'        MSComm1.Output = chrACK
'        Save_Raw_Data "[TX]" & chrACK
'        txtBuff.Text = ""
'
'    Case Else
'        txtBuff.Text = txtBuff.Text & lsChar
'    End Select

    Dim S As String
    Dim i   As Integer
    Dim sGubun As String
    Dim sMode   As Integer
    Dim LineData As String
    
    S = MSComm1.Input
    
    Select Case S
    Case chrSTX 'Chr(2)
        If Right(txtBuff, 1) = "" Then
            sMode = 1
            i = InStr(1, txtBuff, "")
            txtBuff = Mid(txtBuff, 1, i - 1)
        Else
            sMode = 0
            txtBuff = ""
        End If
        
    Case chrETX
        Save_Raw_Data "[RX" & CDate(Time) & "]" & txtBuff.Text & S
        
        sGubun = Mid(txtBuff, 1, 2)
        LineData = Mid(txtBuff.Text, 3)
        Call AU2700(sGubun, LineData)
        
        If sGubun = "R " Then
            SendOrder
        End If
        
    Case Else
        If sMode = 1 Then
            If S = "E" Then
                sMode = 0
            End If
        Else
            txtBuff = txtBuff & S
        End If
    End Select
End Sub

Sub SendOrder()
    If gOrderMessage <> "" Then
        gPreData = gOrderMessage
        gOrderMessage = ""
        
        MSComm1.Output = gPreData
        Save_Raw_Data "[TX" & CDate(Time) & "]" & gPreData
    End If
End Sub

Private Sub AU2700(asGubun As String, asData As String)
    Dim i As Integer
    Dim j As Integer
    
    Dim iRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim lResRow As Long         '°á°ú°ü·Ã Row
    
    Dim lsRackNo As String
    Dim lsPos As String
    
    Dim lsSampleType As String
    Dim lsSampleNo As String
    Dim lsSampleID As String
    Dim lsID As String
    Dim lsPID As String
    
    Dim sExamCode As String
    Dim sSubCode As String
    Dim sExamName As String
    
    Dim lsCode As String
    Dim lsRt As String
    Dim lsFlag As String
    
    Dim lsSeqNo As String
    
    Dim sEXAMDATE As String
    Dim sExamTime As String
    Dim sDate As String

    Dim lsData As String
    
    Dim iCnt As Integer
    Dim iExamCnt As Integer
    Dim sAllResult As String
    
    Dim iLen As String
    
    Dim lsControlNo As String
    Dim lsLotNo As String
    Dim lsLevel As String
    Dim lsLevelName As String
    
    sEXAMDATE = Format(DTPicker1.Value, "yyyymmdd")
    sExamTime = Format(Time, "hhmmss")
    
    sDate = sEXAMDATE & " " & sExamTime
    
    Select Case asGubun
    Case "R "    'Inquery
'        lsRackNo = Mid(asData, 1, 4)
'        lsPos = Mid(asData, 5, 2)
'
'        lsSampleType = Mid(asData, 7, 1)
'
'        lsSampleNo = Trim(Mid(asData, 8, 4))
'
'        '¹ÙÄÚµå¹øÈ£
'        lsSampleID = Trim(Mid(asData, 15, 20))
'        lsID = CStr(lsSampleID)
'
'        '°°Àº ¹ÙÄÚµå¹øÈ£ÀÇ °ËÃ¼´Â µð½ºÇÃ·¹ÀÌµÇÁö ¾ÊÀ½
'        glRow = -1
'        For iRow = 1 To vasID.DataRowCnt
'            If Trim(GetText(vasID, iRow, colBARCODE)) = lsID Then
'                glRow = iRow
'
'                If IsNumeric(gOrdCnt) = True Then
'                    If gOrdCnt > 0 Then
'                        gOrdCnt = CStr(CInt(gOrdCnt) - 1)
'
'                        vasID.Row = glRow
'                        vasID.Col = 1
'                        vasID.Value = 0
'                    End If
'                End If
'
'                Exit For
'            End If
'        Next iRow
'
'        If glRow = -1 Then      'vasID¿¡ ¾ø´Â °ËÃ¼ÀÇ °á°ú°¡ ³ª¿Ã ¶§ µ¥ÀÌÅÍ Ãß°¡
'            glRow = vasID.DataRowCnt + 1
'            If glRow > vasID.MaxRows Then
'                vasID.MaxRows = glRow + 1
'            End If
'
'            SetText vasID, lsID, glRow, colBARCODE
'
'            SetText vasID, lsRackNo, glRow, colRack
'            SetText vasID, lsPos, glRow, colPos
'
'            vasActiveCell vasID, glRow, colBARCODE
'        End If
'
'        'È¯ÀÚÁ¤º¸
''        If Trim(GetText(vasID, glRow, colPID)) = "" Then
''            Get_Sample_Info lsID, glRow
''        End If
'
'        'Order »ý¼º
'        lsData = Make_Order(lsID, glRow)
'
'        iLen = 20 - Len(lsID)
'
'        If lsSampleType = "" Then
'            lsSampleType = Space(1)
'        End If
'
'        If lsSampleNo = "" Then
'            lsSampleType = Space(4)
'        End If
'
'        gOrderMessage = chrSTX & _
'                        "S " & _
'                        lsRackNo & lsPos & lsSampleType & lsSampleNo & Space(iLen) & lsID & "    " & "E" & _
'                        lsData & _
'                        chrETX
                        
                        
    Case "D "    'Result
        lsRackNo = Mid(asData, 1, 4)
        lsPos = Mid(asData, 5, 2)
        
        lsSampleType = Mid(asData, 7, 1)
        lsSampleNo = Trim(Mid(asData, 8, 4))
        
        lsSampleID = Trim(Mid(asData, 15, 20))
        lsID = CStr(lsSampleID)
        
        '°°Àº ¹ÙÄÚµå¹øÈ£ÀÇ °ËÃ¼´Â µð½ºÇÃ·¹ÀÌµÇÁö ¾ÊÀ½
        glRow = -1
        For iRow = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, iRow, colBARCODE)) = lsID And Trim(GetText(vasID, iRow, colSampleNo)) = lsSampleNo Then
                glRow = iRow
                Exit For
            End If
        Next iRow

        If glRow = -1 Then      'vasID¿¡ ¾ø´Â °ËÃ¼ÀÇ °á°ú°¡ ³ª¿Ã ¶§ µ¥ÀÌÅÍ Ãß°¡
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow + 1
            End If
            vasActiveCell vasID, glRow, colBARCODE
        End If
        
        SetText vasID, lsID, glRow, colBARCODE
         
        'È¯ÀÚÁ¤º¸
'        If Trim(GetText(vasID, glRow, colPID)) = "" Then
'            Get_Sample_Info lsID, glRow
'        End If
        
        '°á°ú µð½ºÇÃ·¹ÀÌ
        vasActiveCell vasID, glRow, colBARCODE
        
        ClearSpread vasRes, 1, 1
        
        SetText vasID, lsSampleID, glRow, colBARCODE
        SetText vasID, lsSampleNo, glRow, colSampleNo
        SetText vasID, lsRackNo, glRow, colRack
        SetText vasID, lsPos, glRow, colPos
        
        '¼ö½ÅÁß========================================================
        SetText vasID, "¼ö½ÅÁß", glRow, colState
        SetBackColor vasID, glRow, glRow, 1, 1, 255, 250, 205
        '==============================================================
        
        vasRes.MaxRows = 0
        
'''        '°Ë»çÄÚµå¸¸Å­ RowÀÇ °¹¼ö¸¦ ¼³Á¤
'''        gReadBuf(0) = "0"
'''
'''        SQL = "Select count(EXAMCODE) From  SLX747CT" & vbCrLf & _
'''              " Where EQIPCD = '" & gEquip & "' "
'''        res = db_select_Col(gServer, SQL)
'''
'''        If gReadBuf(0) = "" Then
'''            vasRes.MaxRows = 50
'''        Else
'''            vasRes.MaxRows = Trim(gReadBuf(0))
'''        End If
        
        '°á°ú Àß¶ó ³Ö±â
        j = 0
                                    
        If Trim(Mid(asData, 36, 2)) = "E0" Or Trim(Mid(asData, 36, 2)) = "00" Then
            lsData = Trim(Mid(asData, 43)) '/°Ë»ç°á°úÇ×¸ñÀÌ ±æ¸é µÎ¹øÂ° °á°ú ¹ÞÀ»¶§ E·Î Ç¥±âµÊ.
        Else
            lsData = Trim(Mid(asData, 37))
        End If
        
        Do While Len(lsData) >= 5
            lsCode = Trim(Left(lsData, 2))
            lsRt = Trim(Mid(lsData, 3, 9)) '/Parameter->Online->Protocal Tab(Text Format->Data Format->9ÀÚ¸®)
'''            lsFlag = Trim(Mid(lsData, 12, 2))
'''            If lsFlag = "%?" Then
'''                lsRt = ""
'''                lsFlag = ""
'''            Else
'''                i = InStr(1, lsFlag, "r")
'''                If i = 0 Then
'''
'''                Else
'''                    lsFlag = Left(lsFlag, i - 1)
'''                End If
'''
'''            End If
            
            '°á°ú µð½ºÇÃ·¹ÀÌ
            SQL = "select examcode, examname, seqno from equipexam where EQUIPCODE = '" & Trim(lsCode) & "'"
            res = db_select_Col(gLocal, SQL)
            If res > 0 Then
                vasRes.MaxRows = vasRes.MaxRows + 1
                lResRow = vasRes.MaxRows
                
                SetText vasRes, lsID, lResRow, colBARCODE
                SetText vasRes, lsCode, lResRow, colEquipExam           'ÀåºñÄÚµå
                SetText vasRes, Trim(gReadBuf(0)), lResRow, colExamCode '°Ë»çÄÚµå
                SetText vasRes, Trim(gReadBuf(1)), lResRow, colExamName '°Ë»ç¸í
                SetText vasRes, lsRt, lResRow, colResult                '°Ë»ç°á°ú
                SetText vasRes, lsFlag, lResRow, colRCheck              'ÆÇÁ¤
    
                Save_Local_One glRow, lResRow, "1"
                
                j = j + 1
                

                SetText vasID, "¼ö½Å¿Ï·á", glRow, colState
                SetBackColor vasID, glRow, glRow, 1, 1, 0, 128, 64
            End If
                   
            lsData = Mid(lsData, 14)
            
            If Mid(lsData, 1, 1) = "D" Then     'ETB
                'lsData = Mid(lsData, 37)
                lsData = Mid(lsData, 39)
            End If
        Loop

    
    Case "DQ"   'QC Result
'2007/12/21 ÀÌ»óÀº***************************************************
        lsSampleNo = Trim(Mid(asData, 8, 4))
        lsControlNo = Trim(Mid(asData, 34, 2))
        
        lsID = Trim(Mid(asData, 12, 20))        'Control ID
        
        If lsID = "" Then
            If lsControlNo = "01" Then
                lsLevel = "1"
            ElseIf lsControlNo = "02" Then
                lsLevel = "2"
            ElseIf lsControlNo = "03" Then  'CRP
                lsLevel = "1"
            ElseIf lsControlNo = "04" Then  'CRP
                lsLevel = "2"
            End If
        End If
    
'        'LevelÁ¤º¸
'        lsLevelName = ""
'
'        SQL = " Select LEVELCODE, LEVELNAME From SLXQCMST " & CR & _
'              " Where WORKCODE = '" & gEquip & "' " & CR & _
'              " And LEVELCODE = '" & lsLevel & "' "
'        res = db_select_Col(gServer, SQL)
'
'        lsLevel = Trim(gReadBuf(0))
'        lsLevelName = Trim(gReadBuf(1))
'
'        'ÇØ´ç LevelÀÇ LotNo Á¤º¸
'        lsLotNo = ""
'        SQL = " Select LotNo From SLXQCMCT " & CR & _
'              " Where WORKCODE = '" & gEquip & "' " & CR & _
'              " And LEVELCODE = '" & lsLevel & "' " & CR & _
'              " And LOTNOS <= '" & Format(txtToday.Text, "yyyymmdd") & "' " & CR & _
'              " And LOTNOE >= '" & Format(txtToday.Text, "yyyymmdd") & "' " & CR & _
'              " Order By LOTNOS desc "
'        res = db_select_Col(gServer, SQL)
'        lsLotNo = Trim(gReadBuf(0))
        
        '°°Àº ¹ÙÄÚµå¹øÈ£ÀÇ °ËÃ¼´Â µð½ºÇÃ·¹ÀÌµÇÁö ¾ÊÀ½
        glRow = -1
        For iRow = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, iRow, colPID)) = lsLevel Then
                glRow = iRow
                
                SetText vasID, lsID, glRow, colBARCODE
                Exit For
            End If
        Next iRow

        If glRow = -1 Then      'vasID¿¡ ¾ø´Â °ËÃ¼ÀÇ °á°ú°¡ ³ª¿Ã ¶§ µ¥ÀÌÅÍ Ãß°¡
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow + 1
            End If
            
            SetText vasID, lsID, glRow, colBARCODE
            vasActiveCell vasID, glRow, colBARCODE
        End If
        
        SetText vasID, lsLotNo, glRow, colBARCODE       'Lot NO
        SetText vasID, lsLevel, glRow, colPID           'LevelÄÚµå
        SetText vasID, lsLevelName, glRow, colPName     'Level¸í
        
        SetText vasID, "QC", glRow, colState
        
        
        '°á°ú µð½ºÇÃ·¹ÀÌ
        vasActiveCell vasID, glRow, colBARCODE
        
        ClearSpread vasRes, 1, 1
        
        '°Ë»çÄÚµå¸¸Å­ RowÀÇ °¹¼ö¸¦ ¼³Á¤
        gReadBuf(0) = "0"
        
        SQL = "Select count(EXAMCODE) From  EquipExam" & vbCrLf & _
              " Where EQIPNO = '" & gEquip & "' "
        res = db_select_Col(gLocal, SQL)
        
        If gReadBuf(0) = "" Then
            vasRes.MaxRows = 50
        Else
            vasRes.MaxRows = Trim(gReadBuf(0))
        End If
        
        '°á°ú Àß¶ó ³Ö±â
        j = 0
                            
        lsData = Trim(Mid(asData, 39))
        
        Do While Len(lsData) >= 5
            lsCode = Trim(Left(lsData, 2))
            lsRt = Trim(Mid(lsData, 3, 9))
            lsFlag = Trim(Mid(lsData, 12, 2))
            If lsFlag = "%?" Then
                lsRt = ""
                lsFlag = ""
            Else
                i = InStr(1, lsFlag, "r")
                If i = 0 Then
                    
                Else
                    lsFlag = Left(lsFlag, i - 1)
                End If
                
            End If
            
            '°á°ú µð½ºÇÃ·¹ÀÌ
            SQL = "select examcode, examname, seqno from equipexam where EQUIPCODE = '" & Trim(lsCode) & "'"
            res = db_select_Col(gLocal, SQL)
            If res > 0 Then
                lResRow = vasRes.DataRowCnt + 1
                If lResRow > vasRes.MaxRows Then
                    vasRes.MaxRows = lResRow
                End If
                
                SetText vasRes, lsID, lResRow, colBARCODE
                SetText vasRes, lsCode, lResRow, colEquipExam           'ÀåºñÄÚµå
                SetText vasRes, Trim(gReadBuf(0)), lResRow, colExamCode '°Ë»çÄÚµå
                SetText vasRes, Trim(gReadBuf(1)), lResRow, colExamName '°Ë»ç¸í
                SetText vasRes, lsRt, lResRow, colResult                '°Ë»ç°á°ú
                SetText vasRes, lsFlag, lResRow, colRCheck              'ÆÇÁ¤
                
                Save_Local_One glRow, lResRow, "1"
                
                j = j + 1
                
                SetText vasID, "¼ö½Å¿Ï·á", glRow, colState
                SetBackColor vasID, glRow, glRow, 1, 1, 0, 128, 64
            End If
                    
            lsData = Mid(lsData, 14)
            
            If Mid(lsData, 1, 1) = "D" Then     'ETB
                lsData = Mid(lsData, 39)
            End If
        Loop
        
    Case "DB"   'Result Start

    Case "DE"   'Result End
    End Select

End Sub

Sub AllPentra(asSig As String)
    Dim AllStr As String
'    Dim SubStr(1 To 80) As String
    Dim Subint As Long
    Dim i As Long
    Dim ResFlag As Integer
    
    ResFlag = 0
    
    For i = 1 To 80
        SubStr(i) = ""
    Next
    AllStr = asSig
    Subint = 1
    For i = 1 To Len(AllStr)
        If Mid(AllStr, i, 1) = Chr(13) Then
            If Subint = 2 Then
                '2010.03.17 ÀÌ»óÀº - Àç°Ë °á°ú°¡ ÀÎÅÍÆäÀÌ½º·Î Àü¼Û ¾È µÊ
                'Trim(Mid(SubStr(Subint), 3)) = "RES-RR"  Ãß°¡
                If Trim(Mid(SubStr(Subint), 3)) = "RESULT" Or Trim(Mid(SubStr(Subint), 3)) = "RES-RR" Then
                    ResFlag = 1
                Else
                    Exit Sub
                End If
            End If
            Subint = Subint + 1
            SubStr(Subint) = ""
            
        Else
            SubStr(Subint) = SubStr(Subint) & Mid(AllStr, i, 1)
        End If
    Next
    
    If ResFlag = 1 Then
        Pentra120
    End If
End Sub



Sub Pentra120()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Resi As Integer

    Dim ResultTbl(1 To 40) As String        'Array¿¡ ´ã±â
    Dim TablePtr As Integer
    Dim sTmp As String

    Dim sCnt As String

    Dim sDate As String
    Dim sRefFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sReceNo As String
    Dim sReceDate  As String
    Dim sPID As String
    Dim sPName As String
    Dim sJumin As String
    Dim sPSex As String
    Dim sPAge As String
    Dim sTestID As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sResult1 As String
    Dim sEXAMDATE As String
    Dim sExamName As String
    Dim sCount As Integer
    Dim sSeqno As String
    Dim liRet  As Integer
    Dim sSpace As Integer
    Dim resFormat As String
    Dim sRack As String

    Dim sSampleType As String

    Dim sLevelNo As String
    Dim ResErr As String

    Dim lsTemp1 As String
    Dim jRow As Integer
    Dim X, Y As Integer
    
'    Connect_Server
    ClearSpread vasRes
    
    For Resi = 1 To 80
        
        If Mid(SubStr(Resi), 1, 1) = "u" Then
            sBARCODE = Trim(Mid(SubStr(Resi), 3, 16))
            
            Exit For
        End If
    Next
    
    llRow = -1
    
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBARCODE)) = sBARCODE Then
            llRow = i
            Exit For
        End If
    Next
    If llRow = -1 Then
        llRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < llRow Then
            vasID.MaxRows = llRow
        End If
    End If
    SetText vasID, sBARCODE, llRow, colBARCODE
    For Resi = 1 To 80
        If Mid(SubStr(Resi), 1, 1) = "r" Then
            sRack = Trim(Mid(SubStr(Resi), 3, 10))
            SetText vasID, sRack, llRow, colPos
            
        End If
        
        If Mid(SubStr(Resi), 1, 1) = "s" Then
           
            SetText vasID, Trim(Mid(SubStr(Resi), 3, 4)), llRow, colRack
            
        End If
    Next
    
'    SetText vasID, llRow, llRow, colRack
    
    For Resi = 1 To 80

        
        sTestID = Mid(SubStr(Resi), 1, 1)
        SQL = "select examcode, examname, seqno from equipexam where EQUIPCODE = '" & Trim(sTestID) & "'"
        res = db_select_Col(gLocal, SQL)
        If res > 0 Then
            j = vasRes.DataRowCnt + 1
            If j > vasRes.MaxRows Then
                vasRes.MaxRows = j
            End If
            
            X = InStr(1, Mid(SubStr(Resi), 3, 5), ".")
            
            If X > 0 Then
                X = 5 - X
                For Y = 1 To X
                    If Y = 1 Then
                        resFormat = "0.0"
                    Else
                        resFormat = resFormat & "0"
                    End If
                Next
            Else
                resFormat = "0"
            End If
            
            If Trim(Mid(SubStr(Resi), 8, 2)) = "R" Then
                ResErr = "R"
                SetForeColor vasID, llRow, llRow, 255, 0, 0
            ElseIf Trim(Mid(SubStr(Resi), 8, 2)) = "B" Then
                ResErr = "B"
                SetForeColor vasID, llRow, llRow, 0, 0, 255
            ElseIf Trim(Mid(SubStr(Resi), 8, 2)) = "S" Then
                ResErr = "S"
                SetForeColor vasID, llRow, llRow, 0, 0, 255
            Else
                ResErr = ""
                
            
            End If
            
            
            If IsNumeric(Mid(SubStr(Resi), 3, 5)) = False Then
                sResult = "0.0"
                SetForeColor vasID, llRow, llRow, 0, 0, 255
            Else
                sResult = Format(CCur(Trim(Mid(SubStr(Resi), 3, 5))), resFormat)
                If Trim(sTestID) = "!" Then
                    sResult = Format(sResult * 1000, "#0")
                ElseIf Trim(sTestID) = "2" Then
                    sResult = Format(sResult * 100, "#0")
                ElseIf Trim(sTestID) = "@" Then
                    sResult = Format(sResult / 10, "#0")
                    
                '2010.03.17 ÀÌ»óÀº - ¹éºÐÀ²(Diff) ¼Ò¼ö 1ÀÚ¸®·Î
'                ElseIf Trim(sTestID) = "+" Then
'                    sResult = CLng(sResult)
'                ElseIf Trim(sTestID) = "-" Then
'                    sResult = CLng(sResult)
'                ElseIf Trim(sTestID) = "%" Then
'                    sResult = CLng(sResult)
'                ElseIf Trim(sTestID) = "#" Then
'                    sResult = CLng(sResult)
'                ElseIf Trim(sTestID) = ")" Then
'                    sResult = CLng(sResult)

                ElseIf Trim(sTestID) = "+" Then
                    sResult = Format(sResult, "#0.0")
                ElseIf Trim(sTestID) = "-" Then
                    sResult = Format(sResult, "#0.0")
                ElseIf Trim(sTestID) = "%" Then
                    sResult = Format(sResult, "#0.0")
                ElseIf Trim(sTestID) = "#" Then
                    sResult = Format(sResult, "#0.0")
                ElseIf Trim(sTestID) = ")" Then
                    sResult = Format(sResult, "#0.0")
                End If
                
            End If
            SetText vasRes, Trim(sTestID), j, colEquipExam      'ÀåºñÄÚµå
            SetText vasRes, gReadBuf(0), j, colExamCode         '°Ë»çÄÚµå
            SetText vasRes, gReadBuf(1), j, colExamName         '°Ë»ç¸í
            SetText vasRes, sResult, j, colResult               '°Ë»ç°á°ú
            SetText vasRes, sResult, j, colResult1              '°Ë»ç°á°ú
            SetText vasRes, ResErr, j, colRCheck
            
'                SetText vasRes, CInt(gReadBuf(2)), j, 13


            Save_Local_One llRow, j, "1"
        End If
'        If Mid(SubStr(Resi), 1, 1) = "r" Then
'            sRack = Trim(Mid(SubStr(Resi), 3, 10))
'            SetText vasID, sRack, llRow, colPos
'
''            Get_Sample_Info llRow
'
''            For i = 1 To vasRes.DataRowCnt
''                Save_Local_One llRow, i, "1"
''            Next
''            SQL = "delete from PAT_RES where BARCODE = '' and seqno = '" & Trim(GetText(vasID, llRow, colRack)) & "'"
''            res = SendQuery(gLocal, SQL)
''            Exit For
'        End If
'
'        If Mid(SubStr(Resi), 1, 1) = "s" Then
'            sBARCODE = Trim(Mid(SubStr(Resi), 3, 4))
'            SetText vasID, sBARCODE, llRow, colRack
'
''            Get_Sample_Info llRow
'
'            For i = 1 To vasRes.DataRowCnt
'                Save_Local_One llRow, i, "1"
'            Next
'            SQL = "delete from PAT_RES where diskno = '' and seqno = ''"
'            res = SendQuery(gLocal, SQL)
'
'        End If
'
'        If Mid(SubStr(Resi), 1, 1) = "u" Then
'            sBARCODE = Trim(Mid(SubStr(Resi), 3, 16))
'            SetText vasID, sBARCODE, llRow, colBARCODE
'
'            SQL = "UPDATE PAT_RES set BARCODE = '" & sBARCODE & "' " & vbCrLf & _
'                  "where diskno = '" & Trim(GetText(vasID, llRow, colPos)) & "' " & vbCrLf & _
'                  "and seqno = '" & Trim(GetText(vasID, llRow, colRack)) & "' and BARCODE = ''"
'            res = SendQuery(gLocal, SQL)
'
''            Get_Sample_Info llRow
'
''            For i = 1 To vasRes.DataRowCnt
''                Save_Local_One llRow, i, "1"
''            Next
''            SQL = "delete from PAT_RES where diskno = '' and seqno = ''"
''            res = SendQuery(gLocal, SQL)
'            Exit For
'        End If
        
               
    Next
    
    SetText vasID, "Result", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205

End Sub

Sub Save_Result_Data(ArgSQL As String)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir("C:\AU2700Result", vbDirectory) <> "AU2700Result" Then
        MkDir ("C:\AU2700Result")
    End If
    
    sFileName = Format(CDate(dtpExamDate.Value), "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open "C:\AU2700Result\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, ArgSQL
    Close FilNum
End Sub

Sub Save_Raw_Data(ArgSQL As String)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
        MkDir (App.Path & "\Result")
    End If
    
    sFileName = Format(CDate(dtpExamDate.Value), "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open App.Path & "\Result\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, ArgSQL
    Close FilNum
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim lsBARCODE As String
    Dim lsSeqNo As String
    Dim lsDate As String
    
    'Á¢¼öÀÏÀÚ,Á¢¼ö¹øÈ£·Î »ùÇÃ È¯ÀÚ Á¤º¸ °¡Á®¿À±â
    lsBARCODE = Trim(GetText(vasID, asRow, colBARCODE))   '»ùÇÃ ¹ÙÄÚµå ¹øÈ£
    lsDate = ""
    lsDate = Format(Trim(dtpExamDate.Value), "yyyymmdd")
    
    lsSeqNo = ""
    lsSeqNo = Trim(GetText(vasID, asRow, colRack))
    

'    SQL = "select a.ptno, b.sname, a.sex, a.ageyy, a.JEOBSUDT, a.slipno1, a.slipno2 from twexam_general_sub a, tw_mis_pmpa.twbas_patient b " & vbCrLf & _
'          "where a.ptno = b.ptno and a.ptno = '" & lsBARCODE & "' and jeobsudt = to_date('" & Format(Date, "yyyymmdd") & "', 'yyyy/mm/dd hh24/mi/ss') and itemcd in (" & gAllExam & ") order by slipno2"
          
          
    SQL = "select a.hospno, b.name, b.sex, a.requestdate, b.jumin" & vbCrLf & _
          "from tl_workhead a, tb_idmast b " & vbCrLf & _
          "where a.sample = '" & lsBARCODE & "' and a.hospno = b.hospno"
    res = db_select_Col(gServer, SQL)
    
    If res = 1 Then
        SetText vasID, Trim(gReadBuf(0)), asRow, colPID
        
        SetText vasID, Trim(gReadBuf(1)), asRow, colPName
'        SetText vasID, Trim(gReadBuf(4)), asRow, colJumin
        
        CalAgeSex Trim(gReadBuf(4)), dtpExamDate.Value
        SetText vasID, gPatGen.Age, asRow, colPAge
        SetText vasID, Trim(gReadBuf(2)), asRow, colPSex
        
        
        SetText vasID, Format(Trim(gReadBuf(3)), "yyyymmdd"), asRow, colReqDate
'        SetText vasID, Trim(gReadBuf(5)), asRow, colSlipNo1
'        SetText vasID, Trim(gReadBuf(6)), asRow, colSlipNo2
'
'        SetText vasID, Trim(gReadBuf(6)), asRow, colBARCODE
    Else
        SetText vasID, "", asRow, colPID
        SetText vasID, "", asRow, colPName
'        SetText vasID, "", asRow, colJumin
        SetText vasID, "", asRow, colPSex
        SetText vasID, "0", asRow, colPAge
        SetText vasID, "", asRow, colReqDate
'        SetText vasID, "", asRow, colBARCODE
    End If
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
    
End Function

Function SetResult(asResult As String, aiItem As Integer) As String
'DB¿¡¼­ ºÒ·¯¿À±â
    Dim iFloat As Integer
    
    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    Select Case aiItem
    Case 7, 16
        iFloat = 2
    Case 14
        iFloat = 0
    Case Else
        iFloat = 1
    End Select

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
    End If
 
End Function

Private Sub subDel_Click()
    Dim i As Long
    
    i = vasID.ActiveRow
    
    vasID.DeleteRows i, 1
    If i > vasID.DataRowCnt Then
        i = vasID.DataRowCnt
    End If
    vasID.MaxRows = vasID.DataRowCnt
    vasActiveCell vasID, i, colBARCODE
    vasID.SetFocus
End Sub

Private Sub subUp_Click()
Dim sValue As String
Dim sTmp As String
Dim i As Integer
Dim j As Integer

    sTmp = ""
    
    vasID.Row = vasID.ActiveRow
    vasID.Col = vasID.ActiveCol
    
    sTmp = vasID.Text
    
    sValue = InputBox("º¯°æÇÒ °ËÃ¼¹øÈ£¸¦ ÀÔ·ÂÇÏ¼¼¿ä")
        
    If Trim(sValue) <> "" Then
        If MsgBox("" & sTmp & "¸¦ " & sValue & "·Î ¼öÁ¤ÇÏ½Ã°Ú½À´Ï±î?", vbYesNo, "È®ÀÎ") = vbYes Then
            SetText vasID, sValue, vasID.Row, vasID.Col
            
            If Trim(GetText(vasID, vasID.Row, colBARCODE)) <> "" Then
                Get_Sample_Info vasID.Row
                            
                For i = 1 To vasRes.DataRowCnt
                    Save_Local_One vasID.Row, i, "A"
                Next
            End If
        End If
    End If

End Sub

Private Sub txtPS_GotFocus()
    SelectFocus txtPS
End Sub

Private Sub txtPS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtPS = "" Then
            txtPS.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtPS) = False Then
            txtPS.SetFocus
            Exit Sub
        End If
        
        txtPS.Text = Format(Trim(txtPS.Text), "000#")
        
        txtPE.SetFocus
    End If
End Sub

Private Sub txtPE_GotFocus()
    SelectFocus txtPE
End Sub

Private Sub txtPE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim lsBARCODE As String
    Dim lRow As Long
    
    If KeyCode = vbKeyReturn Then
        If txtPE = "" Then
            txtPE.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtPE) = False Then
            txtPE.SetFocus
            Exit Sub
        End If
        txtPE.Text = Format(Trim(txtPE.Text), "000#")
        txtResPrint.SetFocus
    End If
End Sub

Private Sub txtResPrint_Click()

frmPrint.Show 1
    
End Sub

Private Sub txtStartR_GotFocus()
    SelectFocus txtStartR
End Sub

Private Sub txtStartR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtStartR = "" Then
            txtStartR.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtStartR) = False Then
            txtStartR.SetFocus
            Exit Sub
        End If
        txtStartS.SetFocus
    End If
End Sub

Private Sub txtStartS_GotFocus()
    SelectFocus txtStartS
End Sub

Private Sub txtStartS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtStartS = "" Then
            txtStartS.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtStartS) = False Then
            txtStartS.SetFocus
            Exit Sub
        End If
        
        txtResN.SetFocus
    End If
End Sub

Private Sub txtResN_GotFocus()
    SelectFocus txtResN
End Sub

Private Sub txtResN_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim lsBARCODE As String
    Dim lRow As Long
    
    If KeyCode = vbKeyReturn Then
        If txtResN = "" Then
            txtResN.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtResN) = False Then
            txtResN.SetFocus
            Exit Sub
        End If
    'vasid ¿¡ Á¢¼ö¹øÈ£ ÀÔ·ÂÈÄ ÀúÀå
    
        For i = 1 To CInt(txtResN.Text) - CInt(txtStartS.Text) + 1
            lsBARCODE = txtReceHead & Format(CInt(txtStartS) + i - 1, "000#")
            lRow = CInt(txtStartR.Text) + i - 1
'            txtStartS.Text = CInt(txtStartS.Text) + 1
            
            If Trim(GetText(vasID, lRow, colRack)) = "" Then
            Else
                 SetText vasID, lsBARCODE, lRow, colBARCODE
    
    '            Get_Sample_Info llRow
    
                SQL = "UPDATE PAT_RES SET "
                SQL = SQL & vbCrLf & "       BARCODE   = '" & lsBARCODE & "' "
                SQL = SQL & vbCrLf & " WHERE EXAMDATE  = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPNO   = '" & gEquip & "' "
                SQL = SQL & vbCrLf & "   AND SEQNO     = '" & Trim(GetText(vasID, lRow, colSampleNo)) & "' "
                res = SendQuery(gLocal, SQL)
            End If
           
            
        Next
        
        txtStartR.SetFocus
    End If
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim lsTmpID As String
    
    Dim i As Integer
    
    '»ùÇÃ¹øÈ£¿¡ ÇØ´ç ÇÏ´Â °Ë»ç°á°ú Local Databse¿¡¼­ °¡Á®¿À±â
    If Row = 0 Then
        vasSort vasID, Col
    End If
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBARCODE))

    ClearSpread vasRes
    vasRes.MaxRows = 0
    
'''    SQL = "select '', a.BARCODE, a.EQUIPCODE,  a.examcode, a.examname, a.result, a.refflag, a.panicflag, a.deltaflag, a.unit, a.refvalue, a.panicvalue, a.result " & vbCrLf & _
          "FROM PAT_RES a, equipexam b" & vbCrLf & _
          "WHERE a.EXAMDATE = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND a.EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND a.BARCODE = '" & Trim(GetText(vasID, vasID.Row, colBARCODE)) & "' " & vbCrLf & _
          "  AND a.seqno = '" & Trim(GetText(vasID, vasID.Row, colSampleNo)) & "' " & vbCrLf & _
          "  AND a.diskno = '" & Trim(GetText(vasID, vasID.Row, colPos)) & "' " & vbCrLf & _
          "  AND a.examcode = b.examcode and a.EQUIPCODE = b.EQUIPCODE " & vbCrLf & _
          "  order by b.seqno"
    
    SQL = "select '', a.BARCODE, a.EQUIPCODE,  a.examcode, a.examname, a.result, a.refflag, a.panicflag, a.deltaflag, a.unit, a.refvalue, a.panicvalue, a.result " & vbCrLf & _
          "FROM PAT_RES a, equipexam b" & vbCrLf & _
          "WHERE a.EXAMDATE = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND a.EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND a.BARCODE = '" & Trim(GetText(vasID, vasID.Row, colBARCODE)) & "' " & vbCrLf & _
          "  AND a.seqno = '" & Trim(GetText(vasID, vasID.Row, colSampleNo)) & "' " & vbCrLf & _
          "  AND a.examcode = b.examcode and a.EQUIPCODE = b.EQUIPCODE " & vbCrLf & _
          "  order by b.seqno"
          
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
'    For i = 1 To vasRes.DataRowCnt
'        vasRes.RowHeight(i) = 13
'    Next
    
'    For i = 1 To vasRes.DataRowCnt
'        'ÂüÁ¶Ä¡
'        Select Case Trim(GetText(vasRes, i, colRCheck))
'        Case "H"
'            vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
'
'        'Panic
'        Select Case Trim(GetText(vasRes, i, 8))
'        Case "H"
'            vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 8
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
'
'        'Delta
'        Select Case Trim(GetText(vasRes, i, 9))
'        Case "D"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End Select
'    Next i

End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sEXAMDATE As String
    
'    If Trim(GetText(vasID, asRow1, colSeqNo)) = "QC" Then
'        sEXAMDATE = Trim(GetText(vasID, asRow1, colEXAMDATE))
'
'        '2004/05/28 ÀÌ»óÀº
'        'sEXAMDATE = Left(sEXAMDATE, 4) & "-" & Mid(sEXAMDATE, 5, 2) & "-" & Mid(sEXAMDATE, 7, 2) & " " & Mid(sEXAMDATE, 9, 2) & ":" & Mid(sEXAMDATE, 11, 2) & ":00"
'    Else
'        sEXAMDATE = GetDateFull
'    End If
    
    sCnt = ""
    SQL = "DELETE FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMDATE = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "' "
    SQL = SQL & vbCrLf & "   AND BARCODE   = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' "
    SQL = SQL & vbCrLf & "   AND SEQNO     = '" & Trim(GetText(vasID, asRow1, colSampleNo)) & "' "
'    SaveQuery SQL
    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colEXAMDATE))) Then
'        SetText vasExam, "1900-01-01", asRow, colEXAMDATE
'    End If
    
    SQL = "INSERT INTO PAT_RES "
    SQL = SQL & vbCrLf & " (EXAMDATE,   EQUIPNO,    BARCODE,    receno,     pid, "
    SQL = SQL & vbCrLf & "  pname,      pjumin,     page,       psex,       resdate, "
    SQL = SQL & vbCrLf & "  EQUIPCODE,  examcode,   examtype,   result,     sendflag, "
    SQL = SQL & vbCrLf & "  examname,   refflag,    panicflag,  deltaflag,  unit, "
    SQL = SQL & vbCrLf & "  refvalue,   panicvalue, seqno,      diskno,     posno) "
    SQL = SQL & vbCrLf & " VALUES "
    SQL = SQL & vbCrLf & " ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', " '/EXAMDATE
    SQL = SQL & vbCrLf & "  '" & Trim(gEquip) & "', " '/EQUIPNO
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "', " '/BARCODE
    SQL = SQL & vbCrLf & "  '', " '/receno
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPID)) & "', " '/pid
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPName)) & "', " '/pname
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colJumin)) & "', " '/pjumin
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPAge)) & "', " '/page
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " '/psex
    SQL = SQL & vbCrLf & "  '" & sEXAMDATE & "', " '/resdate
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', " '/EQUIPCODE
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " '/examcode
    SQL = SQL & vbCrLf & "  '', " '/examtype
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " '/result
    SQL = SQL & vbCrLf & "  '" & asSend & "', " '/sendflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " '/examname
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', " '/refflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colPCheck)) & "', " '/panicflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colDCheck)) & "', " '/deltaflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colUnit)) & "', "     '/unit
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colRef)) & "', "      '/refvalue
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colPanic)) & "', "    '/panicvalue
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colSampleNo)) & "', "      '/seqno
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colRack)) & "', "       '/diskno
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPos)) & "') "       '/posno
    
'    SaveQuery SQL
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


'Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim j As Integer
Dim iStrNum As Integer

Dim iRow As Integer
Dim lRow As Long

    iRow = vasID.ActiveRow
    lRow = iRow
    
    If KeyCode = vbKeyReturn Then
'        If Trim(GetText(vasID, iRow, colSeqNo)) <> "" Then
'                iStrNum = Trim(GetText(vasID, iRow, colSeqNo))
'            For j = iRow To vasID.DataRowCnt
'                If j = iRow Then
'                Else
'                    iStrNum = iStrNum + 1
'                    SetText vasID, iStrNum, j, colSeqNo
'                End If
'            Next j
'        ElseIf Trim(GetText(vasID, iRow, colBARCODE)) <> "" Then
            SetText vasID, Trim(GetText(vasID, lRow, colBARCODE)), lRow, colBARCODE
            Get_Sample_Info lRow

            '2004/03/10 ÀÌ»óÀº
            For i = 1 To vasRes.DataRowCnt
                Save_Local_One lRow, i, "1"
            Next
            SQL = "delete from PAT_RES where BARCODE = '' and seqno = '" & Trim(GetText(vasID, lRow, colRack)) & "'"
            res = SendQuery(gLocal, SQL)
'        End If
    End If
End Sub
'End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop
End Sub

Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
   vasRes.Row = vasRes.ActiveRow
   vasRes.Col = vasRes.ActiveCol
   ConfirmData = vasRes.Value
    
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Response, Help
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    If KeyCode = vbKeyReturn Then
        vasIDRow = vasID.ActiveRow
        If vasResCol = colResult And _
           Trim(GetText(vasRes, vasResRow, colResult)) <> Trim(GetText(vasRes, vasResRow, colResult1)) Then
            
            Response = MsgBox("ÀúÀåÇÏ½Ã°Ú½À´Ï±î?", vbYesNo + vbCritical + vbDefaultButton2, "ÁÖÀÇ!!!  È®ÀÎ!!!", Help, 100)
            If Response = vbYes Then
                'ÆÇÁ¤, µ¨Å¸, ÆÐ´Ð ¼öÁ¤
                Check_Result Trim(GetText(vasID, vasIDRow, colBARCODE)), _
                             Trim(GetText(vasID, vasIDRow, colPID)), _
                             Trim(GetText(vasRes, vasResRow, colExamCode)), _
                             Trim(GetText(vasRes, vasResRow, colResult)), _
                             vasResRow, Trim(GetText(vasID, vasIDRow, colPSex))

                SQL = " UPDATE PAT_RES " & vbCrLf & _
                      " Set result = '" & Trim(GetText(vasRes, vasResRow, colResult)) & "', " & vbCrLf & _
                      " refFlag = '" & Trim(GetText(vasRes, vasResRow, colRCheck)) & "', " & vbCrLf & _
                      " panicFlag = '" & Trim(GetText(vasRes, vasResRow, colPCheck)) & "', " & vbCrLf & _
                      " deltaFlag = '" & Trim(GetText(vasRes, vasResRow, colDCheck)) & "' " & vbCrLf & _
                      " WHERE EXAMDATE = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      "  AND EQUIPCODE = '" & Trim(GetText(vasRes, vasResRow, colEquipExam)) & "'" & vbCrLf & _
                      "  AND BARCODE = '" & Trim(GetText(vasID, vasIDRow, colBARCODE)) & "' "
                res = SendQuery(gLocal, SQL)
                
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                
            End If
        End If
        
    End If
End Sub

'Public Function QC_Result(argBARCODE As String, argExamCode As String, _
'                            argResult As String, ByVal argRow As Integer) As Integer
'    Dim sDiffRet, sDiffRet1 As String
'    Dim PreResult   As String
'
'    Dim sResClassCode As String     '°á°úÁ¾·ù
'    Dim sLow        As String       'ÂüÁ¶Ä¡
'    Dim sHigh       As String
'    Dim RefRet      As String
'
'    Dim sPart       As String
'    Dim sEquip      As String
'    Dim sLevel      As String
'    Dim sLotNo      As String
'
'    Dim sTmpRece1, sTmpRet1 As String
'    Dim sTmpRece2, sTmpRet2 As String
'    Dim i           As Integer
'    Dim sReceNo     As String
'    Dim sPID        As String
'
'    Dim sTmpStr As String
'
'    QC_Result = -1
'
'    If argBARCODE = "" Then
'        Exit Function
'    End If
'
'    If argExamCode = "" Then
'        Exit Function
'    End If
'
'
'    RefRet = ""
'
'    sDiffRet = argResult
'    If sDiffRet = "" Then
'        QC_Result = -1
'        Exit Function
'    End If
'    sPart = Trim(GetText(vasID, argRow, colJumin))
'    sEquip = gEquip
'    sLevel = Trim(GetText(vasID, argRow, colPName))
'    sLotNo = Trim(GetText(vasID, argRow, colPID))
'
'    SQL = "Select Max(q.AppDate), e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh   " & vbCrLf & _
'          "From QCInItem q, ExamMaster e " & vbCrLf & _
'          "Where q.LabCode = '" & sPart & "' " & vbCrLf & _
'          "  and q.EQUIPCODE = '" & sEquip & "' " & vbCrLf & _
'          "  and q.QCInLevel = '" & sLevel & "' " & vbCrLf & _
'          "  and q.LotNo = '" & sLotNo & "' " & vbCrLf & _
'          "  and q.QCBARCODE = '" & argBARCODE & "' " & vbCrLf & _
'          "  and q.ExamCode = '" & argExamCode & "' " & vbCrLf & _
'          "  and q.AppDate >= '1900-01-01' " & vbCrLf & _
'          "  and e.AppDate = (select Max(c.AppDate) from ExamMaster c Where c.AppDate >= '1900-01-01' and c.ExamCode = q.ExamCode)" & vbCrLf & _
'          "  and e.ExamCode = q.ExamCode " & vbCrLf & _
'          "Group by e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh"
'    res = db_select_Col(gServer, SQL)
'    sResClassCode = Trim(gReadBuf(1))
'
'    If sResClassCode = "1" Then '¼ýÀÚ
''ÂüÁ¶Ä¡ Ã¼Å©
'        sLow = ""
'        sHigh = ""
'
'        '¼ýÀÚÀÎÁö ¾Æ´ÑÁö È®ÀÎ
'        If IsNumeric(sDiffRet) = False Then
'           MsgBox "°á°úÇü½ÄÀÌ ÀÏÄ¡ÇÏÁö ¾Ê½À´Ï´Ù.", vbInformation, "¾Ë¸²"
'           QC_Result = -1
'           Exit Function
'        End If
'
'        If IsNumeric(gReadBuf(2)) Then
'            If CInt(gReadBuf(2)) > 0 Then
'                sTmpStr = "#0."
'                For i = 1 To CInt(gReadBuf(2))
'                    sTmpStr = sTmpStr & "0"
'                Next i
'            Else
'                sTmpStr = "#0"
'            End If
'            sDiffRet = Format(sDiffRet, sTmpStr)
'            SetText vasRes, sDiffRet, argRow, colResult
'            SetText vasRes, sDiffRet, argRow, colResult1
'        End If
'
'        sLow = Trim(gReadBuf(3))
'        sHigh = Trim(gReadBuf(4))
'
'        If sLow = "" And sHigh = "" Then
'            RefRet = ""
'        ElseIf sLow = "" And sHigh <> "" Then   'ÀÌ»ó
'            If CCur(sHigh) < CCur(sDiffRet) Then
'                RefRet = "H"
'            End If
'        ElseIf sLow <> "" And sHigh = "" Then   'ÀÌÇÏ
'            If CCur(sLow) > CCur(sDiffRet) Then
'                RefRet = "L"
'            End If
'        Else
'            If CCur(sLow) > CCur(sDiffRet) Then
'                RefRet = "L"
'            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
'                RefRet = "H"
'            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
'                RefRet = ""
'            End If
'        End If
'
'
'
'    ElseIf sResClassCode = "2" Then '¹®ÀÚ
''        Dim sRefValue As String
''        Dim sPanicValue As String
''        Dim sResult As String
''
''        sLow = ""
''        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
''
''        '2003/03/17 ÀÌ»óÀº ¼öÁ¤
''        '°Ë»ç Ç×¸ñ °á°ú ÂüÁ¶ ÄÚµå Ã¼Å©¿¡¼­ 1 ÀÌ»óÀÏ °æ¿ì¸¸ ÆÇÁ¤µÇ°Ô
''        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''            Exit Function
''        End If
''
''        '2002³â 3¿ù 12ÀÏ +-¿¡¼­ +/-·Î ¼öÁ¤
''        '2002³â 5¿ù 13ÀÏ NON-REACTIVE ÆÇÁ¤ ¾ÈµÅ¼­ Ãß°¡
''        '2003³â 2¿ù 4ÀÏ ÀÌ»óÀº ¼öÁ¤ - 0-1·Î ÂüÁ¶Ä¡´Â 1ÀÌ³ª ÆÇÁ¤µÊ
''        '=================================================================================
''        '2002³â 5¿ù 13ÀÏ 1 : 40 ¹Ì¸¸ ÆÇÁ¤ ¾ÈµÊ
''        '2002³â 6¿ù 11ÀÏ (°á°úÂüÁ¶°¡ 1:·Î ½ÃÀÛÇÏ¸é ÆÇÁ¤Ã¼Å© ¾ÈÇÏ°Ô ¼öÁ¤)
''        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
''            Exit Function
''        End If
''        '=================================================================================
''
''        Select Case UCase(sDiffRet)
''        Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-1"
''            sResult = 1
''        Case "+/-", "2", "+-", "2-5"
''            sResult = 2
''        Case "+", "POSITIVE", "¾ç¼º", "3", "6-10"
''            sResult = 3
''        Case "++", "4", "11-20"
''            sResult = 4
''        Case "+++", "5", "21-30"
''            sResult = 5
''        Case "++++", "6"
''            sResult = 6
''        Case "+++++", "7"
''            sResult = 7
''        Case "++++++", "8"
''            sResult = 8
''        Case Else
''            sResult = sDiffRet
''        End Select
''        'sLow = "0-2"
''        If Trim(sLow) <> "" Then
''            Select Case UCase(Trim(sLow))
''            Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-2"
''                sRefValue = 1
''            Case "+/-", "2", "+-"
''                sRefValue = 2
''            Case "+", "POSITIVE", "¾ç¼º", "3"
''                sRefValue = 3
''            Case "++", "4"
''                sRefValue = 4
''            Case "+++", "5"
''                sRefValue = 5
''            Case "++++", "6"
''                sRefValue = 6
''            Case "+++++", "7"
''                sRefValue = 7
''            Case "++++++", "8"
''                sRefValue = 8
''            Case Else
''                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
''                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
''                    RefRet = sDiffRet
''                End If
''            End Select
''            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''
''            ElseIf sRefValue < sResult Then
'''                RefRet = "H"
''                RefRet = sDiffRet
''
'''                argTable.Row = argRow
'''                argTable.Col = iresDecision
'''                argTable.ForeColor = RGB(205, 55, 0)
''
''
''            End If
''        End If
''        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''            RefRet = Trim(GetText(argTable, argRow, iresDecision))
''        End If
'    End If
'
'    SetText vasRes, RefRet, argRow, colRCheck
'
'    If RefRet = "L" Then
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(65, 105, 225)
'    Else
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(205, 55, 0)
'    End If
'
'    QC_Result = 1
'
'End Function

Public Function Check_Result(argBARCODE As String, argPID As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Integer, asSex As String) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '°á°úÁ¾·ù
    Dim sLow        As String       'ÂüÁ¶Ä¡
    Dim sHigh       As String
    Dim RefRet      As String
    Dim sPanicGubun As String
    Dim sPanicLow   As String       'Panic
    Dim sPanicHigh  As String
    Dim PanicRet    As String
    Dim sDeltaGubun As String
    Dim sDeltaLow   As String       'Delta
    Dim sDeltaHigh  As String
    Dim DeltaRet    As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim sMax_ReceNo As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    Check_Result = -1
    
    If argBARCODE = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    PanicRet = ""
    DeltaRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        Check_Result = -1
        Exit Function
    End If
    
    SQL = " Select ResClassCode, Res_M_Low, Res_M_High, Res_F_Low, Res_F_High, " & CR & _
          "        PanicValueGubun, Panic_M_Low, Panic_M_High, Panic_F_Low, Panic_F_High, " & CR & _
          "        DeltaValueGubun, DeltaLow, DeltaHigh, Point " & CR & _
          "From ExamMaster " & CR & _
          " Where HID = '117' " & CR & _
          " And ExamCode = '" & Trim(argExamCode) & "' "
    res = db_select_Col(gServer, SQL)
    
    sResClassCode = Trim(gReadBuf(0))
    
    If sResClassCode = "1" Then '¼ýÀÚ
'ÂüÁ¶Ä¡ Ã¼Å©
        sLow = ""
        sHigh = ""
        
        '¼ýÀÚÀÎÁö ¾Æ´ÑÁö È®ÀÎ
        If IsNumeric(sDiffRet) = False Then
           'MsgBox "°á°úÇü½ÄÀÌ ÀÏÄ¡ÇÏÁö ¾Ê½À´Ï´Ù.", vbInformation, "¾Ë¸²"
           Check_Result = -1
           Exit Function
        End If
        
        If IsNumeric(gReadBuf(13)) Then
            If CInt(gReadBuf(13)) > 0 Then
                sTmpStr = "#0."
                For i = 1 To CInt(gReadBuf(13))
                    sTmpStr = sTmpStr & "0"
                Next i
            Else
                sTmpStr = "#0"
            End If
            sDiffRet = Format(sDiffRet, sTmpStr)
            SetText vasRes, sDiffRet, argRow, colResult
            SetText vasRes, sDiffRet, argRow, colResult1
        End If
        
        Select Case asSex
        Case "M", ""
            sLow = Trim(gReadBuf(1))
            sHigh = Trim(gReadBuf(2))
        Case "F"
            sLow = Trim(gReadBuf(3))
            sHigh = Trim(gReadBuf(4))
        End Select
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   'ÀÌ»ó
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   'ÀÌÇÏ
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


'Panic Ã¼Å©
        sPanicLow = ""
        sPanicHigh = ""
        
        sPanicGubun = Trim(gReadBuf(5))
        
        Select Case asSex
        Case "M", ""
            sPanicLow = Trim(gReadBuf(6))
            sPanicHigh = Trim(gReadBuf(7))
        Case "F"
            sPanicLow = Trim(gReadBuf(8))
            sPanicHigh = Trim(gReadBuf(9))
        End Select
        
        If sPanicGubun = "0" Then '»óÇÑ/ÇÏÇÑ
            If sPanicLow = "" Or sPanicHigh = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) > CCur(sDiffRet) Then
                    PanicRet = "L"
                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
                    PanicRet = "H"
                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
                    PanicRet = ""
                End If
            End If
        ElseIf sPanicGubun = "1" Then 'percent
            If sPanicLow = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "L"
                    Else
                        PanicRet = ""
                    End If
                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "H"
                    Else
                        PanicRet = ""
                    End If
                Else
                    PanicRet = ""
                End If
            End If
        End If
        

'Delta Ã¼Å©
        sDeltaLow = ""
        sDeltaHigh = ""
                
        sTmpRece1 = ""
        sTmpRet1 = ""
        sTmpRece2 = ""
        sTmpRet2 = ""
        PreResult = ""
        
        sMax_ReceNo = ""
'        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
        sReceNo = argBARCODE
        
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where HID = '117' " & CR & _
'              " And PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBARCODE & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " Group By Result"
              
'2004/12/30 ÀÌ»óÀº - Á¤·ÄºÎºÐ Ãß°¡
        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
              " Where HID = '117' " & CR & _
              " And PID = '" & Trim(argPID) & "' " & CR & _
              " And ReceNo < '" & argBARCODE & "' " & CR & _
              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
              " Group By Result" & CR & _
              " Order by 2 desc "
        res = db_select_Col(gServer, SQL)
              
        If res > 0 And gReadBuf(0) <> "" Then
            PreResult = gReadBuf(0)
        Else
            PreResult = ""
        End If
        
        'ÀÌÀü°á°ú°¡ °ø¹éÀÌ ¾Æ´Ï°í, ¼ýÀÚÀÎ °æ¿ì¸¸
        If PreResult <> "" And IsNumeric(PreResult) Then
          'PreResult = Trim(gReadBuf(0))
          sDeltaGubun = Trim(gReadBuf(10))
          
          sDeltaLow = Trim(gReadBuf(11))
          sDeltaHigh = Trim(gReadBuf(12))
          
            'ÀÌÀü°á°ú¿¡¼­ ÇöÀç°á°ú »«°ªÀÌ sDiffRetÀÓ (2002³â 3¿ù 15ÀÏ ¼öÁ¤)
'            sDiffRet = PreResult - sDiffRet
            sDiffRet1 = sDiffRet - PreResult
            If sDeltaGubun = "0" Then '»óÇÑ/ÇÏÇÑ
                If sDeltaLow = "" Or sDeltaHigh = "" Then
                    DeltaRet = ""
                Else
                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
                        DeltaRet = "L"
                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
                        DeltaRet = "H"
                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
                        DeltaRet = ""
                    End If
                End If
              
            ElseIf sDeltaGubun = "1" Then 'percent
               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
                  DeltaRet = ""
               Else
                   If sDeltaLow = "" Then
                        DeltaRet = ""
                    Else
                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
                            DeltaRet = "D"
                        Else
                            DeltaRet = ""
                        End If
                    End If
               End If
            End If
        End If
        
    ElseIf sResClassCode = "2" Then '¹®ÀÚ
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 ÀÌ»óÀº ¼öÁ¤
'        '°Ë»ç Ç×¸ñ °á°ú ÂüÁ¶ ÄÚµå Ã¼Å©¿¡¼­ 1 ÀÌ»óÀÏ °æ¿ì¸¸ ÆÇÁ¤µÇ°Ô
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002³â 3¿ù 12ÀÏ +-¿¡¼­ +/-·Î ¼öÁ¤
'        '2002³â 5¿ù 13ÀÏ NON-REACTIVE ÆÇÁ¤ ¾ÈµÅ¼­ Ãß°¡
'        '2003³â 2¿ù 4ÀÏ ÀÌ»óÀº ¼öÁ¤ - 0-1·Î ÂüÁ¶Ä¡´Â 1ÀÌ³ª ÆÇÁ¤µÊ
'        '=================================================================================
'        '2002³â 5¿ù 13ÀÏ 1 : 40 ¹Ì¸¸ ÆÇÁ¤ ¾ÈµÊ
'        '2002³â 6¿ù 11ÀÏ (°á°úÂüÁ¶°¡ 1:·Î ½ÃÀÛÇÏ¸é ÆÇÁ¤Ã¼Å© ¾ÈÇÏ°Ô ¼öÁ¤)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        Select Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "¾ç¼º", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End Select
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "¾ç¼º", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End Select
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
'        sLow = ""
'        sLow = Trim(GetText(argTable, argRow, iresPanicValue))
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "À½¼º"
'                sPanicValue = 1
'            Case "+/-"
'                sPanicValue = 2
'            Case "+", "POSITIVE", "¾ç¼º"
'                sPanicValue = 3
'            Case "++"
'                sPanicValue = 4
'            Case "+++"
'                sPanicValue = 5
'            Case "++++"
'                sPanicValue = 6
'            Case "+++++"
'                sPanicValue = 7
'            Case "++++++"
'                sPanicValue = 8
'            Case Else
'                If UCase(sDiffRet) > UCase(sLow) Then
'                    PanicRet = sDiffRet
'                End If
'            End Select
'            If sPanicValue < sResult Then
'                'PanicRet = "H"
'                PanicRet = sDiffRet
'            End If
'        End If
'
'        'Delta Check
'        sMax_ReceNo = ""
'        DeltaRet = ""
'        sReceNo = Trim(GetText(argForm.vasPatient, 1, 1))
'        sPID = Trim(GetText(argForm.vasPatient, 1, 3))
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & sPID & "' " & CR & _
'              " And ReceNo < '" & sReceNo & "' " & CR & _
'              " And ExamCode = '" & Trim(GetText(argTable, argRow, iresExamCode)) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'               If sDiffRet <> gReadBuf(0) Then
'                  DeltaRet = "D"
'               End If
'        Else
'            DeltaRet = ""
'        End If
    End If
    
    SetText vasRes, RefRet, argRow, colRCheck
    SetText vasRes, PanicRet, argRow, colPCheck
    SetText vasRes, DeltaRet, argRow, colDCheck
    

    '2002³â 2¿ù 15ÀÏ ¼öÁ¤ (ÆÇÁ¤½Ã H, L ÀÏ¶§ ±ÛÀÚ »ö±ò º¯È­)
    '2002³â 3¿ù 14ÀÏ ¼öÁ¤ (ÆÇÁ¤½Ã LÀÏ¶§´Â ÆÄ¶õ»ö ±× ¿Ü´Â »¡°£»ö)
    If RefRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If PanicRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If DeltaRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    ElseIf DeltaRet = "D" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    Check_Result = 1

End Function

