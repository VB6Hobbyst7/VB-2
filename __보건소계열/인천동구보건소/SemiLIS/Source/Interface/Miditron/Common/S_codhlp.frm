VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmS0codhlp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   1800
   ClientTop       =   2145
   ClientWidth     =   4170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   4170
   Begin Threed.SSPanel pnlbottom 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   3540
      Width           =   4170
      _Version        =   65536
      _ExtentX        =   7355
      _ExtentY        =   714
      _StockProps     =   15
      ForeColor       =   16576
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Font3D          =   3
      Alignment       =   1
      Begin Threed.SSPanel pnlMsg 
         Height          =   330
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   8388608
         BackColor       =   14737632
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
         BevelOuter      =   1
      End
   End
   Begin Threed.SSPanel pnlmain 
      Align           =   1  'Align Top
      Height          =   3525
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4170
      _Version        =   65536
      _ExtentX        =   7355
      _ExtentY        =   6218
      _StockProps     =   15
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      Begin FPSpread.vaSpread SpdCode 
         Height          =   3000
         Left            =   90
         OleObjectBlob   =   "S_codhlp.frx":0000
         TabIndex        =   5
         Top             =   90
         Width           =   3990
      End
      Begin VB.TextBox txtCd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   690
         TabIndex        =   0
         Top             =   3150
         Width           =   2235
      End
      Begin Threed.SSCommand CmdClk 
         Height          =   285
         Left            =   2970
         TabIndex        =   3
         Top             =   3150
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "View"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand CmdEsc 
         Cancel          =   -1  'True
         Height          =   285
         Left            =   3510
         TabIndex        =   2
         Top             =   3150
         Width           =   555
         _Version        =   65536
         _ExtentX        =   970
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Esc"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "코드명"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   3195
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmS0codhlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sKeyCod As String
Dim SelectCount As Integer

Dim ChangeKeyFlag As Integer

Private Const SUCCEED% = 1
Private Const FAIL% = 0

Dim sStr$, tData() As String


Dim SqlConn As Integer
Private Sub CmdClk_Click()

    MousePointer = HOURGLASS

    If ChangeKeyFlag = True Then
        sKeyCod = Trim(txtCd)
        ChangeKeyFlag = False
    
        SqlStr = "SELECT " & S0COM_code_col & ", " & S0COM_name_col & " FROM " & S0COM_table
        SqlStr = SqlStr & " where " & S0COM_name_col & " LIKE '" & Trim(sKeyCod) & "%' "
'        SQlStr = SQlStr & " WHERE " & S0COM_NAME_COL & " LIKE '%" & Trim(sKeyCod) & "%' "
'        SQlStr = SQlStr & " OR " & S0COM_NAME_COL & " LIKE '" & Trim(sKeyCod) & "%' "
'        SQlStr = SQlStr & " OR " & S0COM_NAME_COL & " LIKE '%" & Trim(sKeyCod) & "' "
        If S0COM_cd_gbn <> "" Then
            SqlStr = SqlStr & " AND " & S0COM_cd_gbn
        End If
        
        SpdCode.Col = 1
        SpdCode.Col2 = SpdCode.MaxCols
        SpdCode.Row = 1
        SpdCode.Row2 = SpdCode.MaxRows
        SpdCode.BlockMode = True
        SpdCode.Action = SS_ACTION_CLEAR_TEXT
        SpdCode.BlockMode = False
        
        S0SUB_CDNAME_HLP
    
    End If
    MousePointer = DEFAULT

End Sub

Private Sub CmdEsc_Click()
    
    S0COM_ret = -1
    S0COM_cd_gbn = ""
    Unload Me

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_ESCAPE Then
        CmdEsc_Click
    End If

End Sub

Private Sub Form_Load()

    Dim ret%
    Dim pw As Integer
    
    sKeyCod = ""
    SelectCount = 0
    
    Me.Width = Screen.Width
    
    If S0COM_length >= 30 Then
        SpdCode.ColWidth(1) = S0COM_length - 10
    Else
        SpdCode.ColWidth(1) = 18
    End If
    SpdCode.AutoSize = True
    ret = DoEvents()
    SpdCode.AutoSize = False
    
    Me.Width = SpdCode.Width + 260
    pw = pnlmain.Width
    pnlMsg.Width = pw - 90
    CmdClk.Left = pw - 1245
    CmdEsc.Left = pw - 705
    txtCd.Width = pw - 1950
    
    If InStr(S0COM_FORMNAME, "DSN") <> 0 Then
         '본 화면에서 사용할 Index Open
        ret = S0SUB_Open(S0COM_FORMNAME, Me.hWnd, SqlConn)
        If ret <> QSQL_SUCCESS Then
            Exit Sub
        End If
    Else
        SqlConn = QsqlCode
    End If
    
    SqlStr = "SELECT DISTINCT " & S0COM_code_col & ", " & S0COM_name_col & " FROM " & S0COM_table
    If S0COM_cd_gbn <> "" Then
        SqlStr = SqlStr & " WHERE " & S0COM_cd_gbn
    End If
    SqlStr = SqlStr & " ORDer by " & S0COM_code_col

    S0SUB_CDNAME_HLP

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim ret%

    If InStr(S0COM_FORMNAME, "DSN") <> 0 Then
        ret = Qsqlclose(SqlConn, ONECLOSE)
    End If

End Sub



'-------------------------------------------------------
'CODE HELP를 위한 프로시져
'X=0 : 코드로 SORT
'X=1 : 내용으로 SORT
'--------------------------------------------------------
Private Sub S0SUB_CDNAME_HLP()
    
    Dim ret%
    Dim sqldata() As String
            
 '   On Error GoTo S0SUB_CDNAME_HLP_ERROR
                
    SelectCount = 0
    
    ret = QSqlDBExec(SqlStr, SqlConn%)
    
    If ret = QSQL_SUCCESS Then
        Do Until QSqlGetRow(RECORD, SqlConn%) <> QSQL_SUCCESS
    
            QSqlGetField 2, RECORD, sqldata()
            
            If SelectCount = SpdCode.MaxRows Then
                SpdCode.MaxRows = SpdCode.MaxRows + 1
                SelectCount = SpdCode.MaxRows
            Else
                SelectCount = SelectCount + 1
            End If
            
            SpdCode.SetText 1, SelectCount, sqldata(2)
            SpdCode.SetText 2, SelectCount, sqldata(1)
            
        Loop
    End If
    ret = QSqlSelectFree(SqlConn)
    
    If SelectCount >= 10 Then SpdCode.MaxRows = SelectCount

    If SelectCount = 0 Then
        Beep
        S0COM_ret = ret
        pnlMsg = "데이타가 없습니다."
    Else
        pnlMsg = "Table : " & S0COM_table & "   Count : " & SelectCount
    End If
    
'S0SUB_CDNAME_HLP_ERROR:
 '   If Err <> 0 Then
'        MsgBox "Error :" & Err & Chr$(13) & Error(Err)
'    End If
End Sub

Private Sub SpdCode_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim ret As Integer
    Dim Nam As Variant
    Dim Cod As Variant
    
    If Row <> 0 Then
        ret = SpdCode.GetText(1, Row, Nam)
        ret = SpdCode.GetText(2, Row, Cod)
        S0COM_name = Nam
        S0COM_code = Cod
        S0COM_cd_gbn = ""
        S0COM_ret = SUCCEED
        Unload Me
    End If

End Sub

Private Sub SpdCode_GotFocus()

    With SpdCode
        .Row = SpdCode.ActiveRow
        .Row2 = SpdCode.ActiveRow
        .Col = 1
        .Col2 = SpdCode.MaxCols
    
        .BlockMode = True
        .BackColor = &HEEEFFF
        .BlockMode = False
    End With
    
End Sub

Private Sub SpdCode_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call SpdCode_DblClick(SpdCode.ActiveCol, SpdCode.ActiveRow)
    
End Sub

Private Sub SpdCode_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If Row <> NewRow Then
    
            With SpdCode
            .Row = Row:      .Row2 = Row
            .Col = 1:            .Col2 = 1
        
            .BlockMode = True
            .BackColor = 16318448
            .BlockMode = False
        
            .Col = 2:           .Col2 = 2
        
            .BlockMode = True
            .BackColor = 16776686
            .BlockMode = False
            
            If NewRow <> -1 Then
                .Row = NewRow: .Row2 = NewRow
                .Col = 1: .Col2 = .MaxCols
            
                .BlockMode = True
                .BackColor = &HEEEFFF
                .BlockMode = False
                
            End If
        End With
        
    End If

End Sub

Private Sub txtCD_Change()

    ChangeKeyFlag = True

End Sub

Private Sub txtCD_GotFocus()

    txtCd.SelStart = 0

End Sub

Private Sub txtCD_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then CmdClk_Click

End Sub




