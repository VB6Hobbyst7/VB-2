VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmTBReport 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "항산성균 감수성 결과"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "frmTBReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   6885
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   240
      Left            =   5100
      TabIndex        =   6
      Top             =   6690
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   423
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "("
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8175
      Left            =   30
      TabIndex        =   2
      Top             =   -30
      Width           =   6810
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   60
         ScaleHeight     =   420
         ScaleWidth      =   6660
         TabIndex        =   17
         Top             =   660
         Width           =   6660
         Begin VB.OptionButton optDILaw 
            BackColor       =   &H00DBE6E6&
            Caption         =   " 2. 간 접 법"
            Height          =   275
            Index           =   1
            Left            =   5175
            TabIndex        =   19
            Top             =   75
            Width           =   1275
         End
         Begin VB.OptionButton optDILaw 
            BackColor       =   &H00DBE6E6&
            Caption         =   " 1. 직접법 ( 배양결과 :"
            Height          =   275
            Index           =   0
            Left            =   30
            TabIndex        =   18
            Top             =   45
            Width           =   2160
         End
         Begin MedControls1.LisLabel LisLabel1 
            Height          =   330
            Left            =   4740
            TabIndex        =   20
            Top             =   30
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   582
            BackColor       =   14411494
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   ")"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblDGrow 
            Height          =   270
            Left            =   2235
            TabIndex        =   21
            Top             =   45
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   476
            BackColor       =   13752531
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   ""
            Appearance      =   0
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00F1F5F4&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            Height          =   360
            Left            =   2190
            Shape           =   4  '둥근 사각형
            Top             =   0
            Width           =   2550
         End
      End
      Begin VB.OptionButton optRALaw 
         BackColor       =   &H00DBE6E6&
         Caption         =   " 절대농도법 ( 대조배지균발육 : "
         Height          =   275
         Index           =   1
         Left            =   390
         TabIndex        =   12
         Top             =   1350
         Width           =   2865
      End
      Begin VB.OptionButton optRALaw 
         BackColor       =   &H00DBE6E6&
         Caption         =   " 내성비율법"
         Height          =   275
         Index           =   0
         Left            =   390
         TabIndex        =   11
         Top             =   1050
         Width           =   2550
      End
      Begin MedControls1.LisLabel lblPZAValue 
         Height          =   270
         Left            =   5235
         TabIndex        =   8
         Top             =   6720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   16777215
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   330
         Left            =   6195
         TabIndex        =   7
         Top             =   6675
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   582
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ")"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPZA 
         Height          =   270
         Left            =   2205
         TabIndex        =   5
         Top             =   6705
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   476
         BackColor       =   16777215
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblRst 
         Height          =   4920
         Left            =   300
         TabIndex        =   3
         Top             =   2130
         Width           =   6300
         _Version        =   196608
         _ExtentX        =   11112
         _ExtentY        =   8678
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   15857140
         MaxCols         =   6
         MaxRows         =   11
         MoveActiveOnFocus=   0   'False
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frmTBReport.frx":144A
         UserResize      =   0
      End
      Begin MedControls1.LisLabel lblRemark 
         Height          =   630
         Left            =   75
         TabIndex        =   9
         Top             =   7425
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   1111
         BackColor       =   13752531
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBacRstNm 
         Height          =   270
         Left            =   60
         TabIndex        =   13
         Top             =   390
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   476
         BackColor       =   13752531
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   330
         Left            =   4395
         TabIndex        =   14
         Top             =   1335
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ")"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRGrow 
         Height          =   270
         Left            =   3330
         TabIndex        =   15
         Top             =   1335
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   476
         BackColor       =   13752531
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00F1F5F4&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         Height          =   360
         Left            =   3270
         Shape           =   4  '둥근 사각형
         Top             =   1290
         Width           =   1125
      End
      Begin VB.Label Label9 
         BackColor       =   &H00DBE6E6&
         Caption         =   "1. 결핵균 배양검사 결과 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   16
         Tag             =   "25601"
         Top             =   165
         Width           =   2190
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "2. 특 기 사 항"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   10
         Tag             =   "25601"
         Top             =   7200
         Width           =   2190
      End
      Begin VB.Label Label10 
         BackColor       =   &H00DBE6E6&
         Caption         =   "◈ 약제 감수성 결과"
         Height          =   270
         Left            =   300
         TabIndex        =   4
         Tag             =   "25607"
         Top             =   1815
         Width           =   2070
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D1D8D3&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   5505
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8205
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D1D8D3&
      Caption         =   "출 력(&P)"
      Height          =   510
      Left            =   4155
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8205
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "frmTBReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------
'항산성균 감수성 결과 조회
'--------------------------

Public Event Click()
Private objSQL As New clsLISSqlTubercle
Private mvarWorkarea As String
Private mvarAccdt    As String
Private mvarAccSeq   As String

Private Sub cmdExit_Click()
    RaiseEvent Click
    Unload Me
End Sub


Private Sub Clear()
    lblBacRstNm.Caption = ""
    optDILaw(0).Value = 0
    optDILaw(1).Value = 0
    lblDGrow.Caption = ""
    optRALaw(0).Value = 0
    optRALaw(1).Value = 0
    lblRGrow.Caption = ""
    lblRemark.Caption = ""
    
    With tblRst
        .BlockMode = True
        .Row = 1: .Col = 1
        .Row2 = .MaxRows: .Col2 = .MaxCols
        .Value = ""
        .BlockMode = False
    End With
    
End Sub

Public Sub GetTBResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                       ByVal pAccSeq As String)
    
    Call AFPCultureRstVauleLoad(pWorkArea, pAccDt, pAccSeq)
    
    Call AFPSensBody(pWorkArea, pAccDt, pAccSeq)
    
    Call AFPSensHeader(pWorkArea, pAccDt, pAccSeq)
    
End Sub

Private Sub AFPCultureRstVauleLoad(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
    Dim RS As New Recordset
    Dim strSQL As String
    
    strSQL = objSQL.SQLAFPCultureRstLoad(pWorkArea, pAccDt, pAccSeq)
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        lblBacRstNm.Caption = RS.Fields("rstnm").Value & ""
    End If
    
NoData:
    Set RS = Nothing

End Sub

Private Sub AFPSensBody(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
    Dim RS      As New Recordset
    Dim strSQL  As String
    Dim ii      As Integer
    Dim jj      As Integer
    
    strSQL = objSQL.SQLAFPSensBodyLoad(pWorkArea, pAccDt, pAccSeq)
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        ii = 1
        jj = 1
        With tblRst
            Do Until RS.EOF
                
                .Row = ii
                If ii <> 11 Then
                    .Col = RS.Fields("seq").Value & ""
                        .Value = RS.Fields("rstvalue").Value & ""
                Else
                    If RS.Fields("seq").Value & "" = 1 Then
                        .Col = RS.Fields("seq").Value & ""
                        .Value = RS.Fields("rstvalue").Value & ""
                    ElseIf RS.Fields("seq").Value & "" = 2 Then
                        lblPZA.Caption = RS.Fields("rstvalue").Value & ""
                    Else
                        lblPZAValue.Caption = RS.Fields("rstvalue").Value & ""
                    End If
                End If
                If jj = 5 Then
                    jj = 1
                    ii = ii + 1
                Else
                    jj = jj + 1
                End If
                RS.MoveNext
            Loop
        End With
    End If
    RS.Close
    
NoData:
    Set RS = Nothing
End Sub

Private Sub AFPSensHeader(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
    Dim RS As New Recordset
    Dim strSQL As String
    
    strSQL = objSQL.SQLAFPSensHeaderLoad(pWorkArea, pAccDt, pAccSeq)
    
    RS.Open strSQL, dbconn
    If RS.RecordCount > 0 Then
'        txtTBNo1.Text = Mid(RS.Fields("tbno").Value & "", 1, 4)
'        txtTBNo2.Text = Mid(RS.Fields("tbno").Value & "", 5, 2)
'        txtTBNo3.Text = Mid(RS.Fields("tbno").Value & "", 7, 2)
'        txtTBNo4.Text = Mid(RS.Fields("tbno").Value & "", 9)
'
'        txtBacNo1.Text = Mid(RS.Fields("bacno").Value & "", 1, 4)
'        txtBacNo2.Text = Mid(RS.Fields("bacno").Value & "", 5, 2)
'        txtBacNo3.Text = Mid(RS.Fields("bacno").Value & "", 7, 2)
'        txtBacNo4.Text = Mid(RS.Fields("bacno").Value & "", 9)
       
        
        If RS.Fields("dilaw").Value & "" = "0" Then
            optDILaw(0).Value = True
            lblDGrow.Caption = RS.Fields("dgrow").Value & ""
        Else
            optDILaw(1).Value = True
        End If
        
        
        If RS.Fields("ralaw").Value & "" = "0" Then
            optRALaw(0).Value = True
        Else
            optRALaw(1).Value = True
            lblRGrow.Caption = RS.Fields("rgrow").Value & ""
        End If
        
        lblRemark.Caption = RS.Fields("remark").Value & ""
       
    End If
    RS.Close
    
NoData:
    Set RS = Nothing
End Sub

Private Sub Form_Load()
    Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSQL = Nothing
End Sub
