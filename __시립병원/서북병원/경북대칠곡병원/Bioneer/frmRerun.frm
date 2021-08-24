VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmRerun 
   Caption         =   "재검 건수 입력"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5250
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdClose 
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3960
      TabIndex        =   4
      Top             =   180
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "확인"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2940
      TabIndex        =   3
      Top             =   180
      Width           =   885
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   1485
      Left            =   210
      TabIndex        =   0
      Top             =   750
      Width           =   4785
      _Version        =   196613
      _ExtentX        =   8440
      _ExtentY        =   2619
      _StockProps     =   64
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   4
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmRerun.frx":0000
   End
   Begin MSComCtl2.DTPicker dtpSch 
      Height          =   330
      Left            =   1200
      TabIndex        =   1
      Top             =   210
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   64290817
      CurrentDate     =   37174
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   270
      Width           =   900
   End
End
Attribute VB_Name = "frmRerun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Long
    Dim sCnt As String
    Dim a As String
    
    For i = 1 To vasList.DataRowCnt
        sCnt = ""
        SQL = "Select count(*) from rerun_cnt where examdate = '" & SeperatorCls(dtpSch.Value) & "' and equipcode = '" & Trim(GetText(vasList, i, 1)) & "' "
        res = db_select_Var(gLocal, SQL, sCnt)
        If Not IsNumeric(sCnt) Then
            sCnt = "0"
        End If
        a = Trim(GetText(vasList, i, 3))
        If Not IsNumeric(a) Then
            a = "0"
        End If
        
        If CInt(sCnt) > 0 Then
            SQL = "Update rerun_cnt set r_cnt = " & a & " where examdate = '" & SeperatorCls(dtpSch.Value) & "'  and equipcode = '" & Trim(GetText(vasList, i, 1)) & "' "
        Else
            SQL = "Insert into rerun_cnt (examdate, equipcode, r_cnt) values ('" & SeperatorCls(dtpSch.Value) & "', '" & Trim(GetText(vasList, i, 1)) & "', " & a & " ) "
        End If
        res = SendQuery(gLocal, SQL)
    Next i
End Sub

Private Sub Form_Activate()
    vasActiveCell vasList, 1, 3
    vasList.SetFocus

End Sub

Private Sub Form_Load()
    Dim i, j As Long
    
    dtpSch.Value = Form_Main.Text_Today.Text
    
    SQL = "Select equipcode, examname from equipexam order by 1"
    db_select_Vas gLocal, SQL, vasList
    
    ClearSpread Form_Main.vasTemp
    
    SQL = "Select equipcode, r_cnt from rerun_cnt where examdate = '" & SeperatorCls(dtpSch.Value) & "' "
    res = db_select_Vas(gLocal, SQL, Form_Main.vasTemp)
    
    For i = 1 To Form_Main.vasTemp.DataRowCnt
        For j = 1 To vasList.DataRowCnt
            If Trim(GetText(vasList, j, 1)) = Trim(GetText(Form_Main.vasTemp, i, 1)) Then
                SetText vasList, Trim(GetText(Form_Main.vasTemp, i, 2)), j, 3
                Exit For
            End If
        Next j
    Next i
    
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i, j As Long
    
    i = vasList.ActiveRow
    j = vasList.ActiveCol
    If KeyCode = vbKeyReturn Then
        If j = 3 Then
            If i = vasList.DataRowCnt Then
                cmdSave.SetFocus
            Else
                vasActiveCell vasList, i + 1, 3
            End If
        End If
    End If
End Sub
