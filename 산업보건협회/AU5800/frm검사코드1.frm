VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frm검사코드1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "프로파일"
   ClientHeight    =   8535
   ClientLeft      =   3840
   ClientTop       =   1725
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9315
   Begin VB.Frame Frame1 
      Caption         =   "[검사코드내역]"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Index           =   0
      Left            =   5640
      TabIndex        =   12
      Top             =   780
      Width           =   3615
      Begin VB.CommandButton cmdCancel 
         Caption         =   "취소(&C)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2460
         TabIndex        =   20
         Top             =   3000
         Width           =   1035
      End
      Begin VB.TextBox txtpanichigh 
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2520
         Width           =   1155
      End
      Begin VB.TextBox txtpaniclow 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2160
         Width           =   1155
      End
      Begin VB.TextBox txtseqno 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1740
         Width           =   2355
      End
      Begin VB.TextBox txtexamname 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1260
         Width           =   2355
      End
      Begin VB.TextBox txtexamcode 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   780
         Width           =   2355
      End
      Begin VB.TextBox txtequipcode 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   300
         Width           =   2355
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제(&D)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1260
         TabIndex        =   8
         Top             =   3000
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장(&S)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "이상"
         Height          =   180
         Index           =   6
         Left            =   2400
         TabIndex        =   19
         Top             =   2580
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "이하"
         Height          =   180
         Index           =   5
         Left            =   2400
         TabIndex        =   18
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Panic Ref"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   2220
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    번"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   1800
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[검사코드현황]"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Index           =   2
      Left            =   60
      TabIndex        =   11
      Top             =   780
      Width           =   5535
      Begin FPSpread.vaSpread vasList 
         Height          =   7395
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   5385
         _Version        =   393216
         _ExtentX        =   9499
         _ExtentY        =   13044
         _StockProps     =   64
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         OperationMode   =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frm검사코드1.frx":0000
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8220
      TabIndex        =   9
      Top             =   60
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   60
      X2              =   9240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '단색
      Height          =   435
      Index           =   1
      Left            =   3960
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   435
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "프로파일 설정"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   9
      Left            =   480
      TabIndex        =   10
      Top             =   180
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '단색
      Height          =   615
      Index           =   1
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4035
   End
End
Attribute VB_Name = "frm검사코드1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height의 초기값
Dim lngMeWidth      As Long '/Me.Width의 초기값

Private Type ConWhere   ' 사용자 정의 형식을 만듭니다.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Public Sub MM_CANCEL()
    If vasList.MaxRows > 0 Then vasList.MaxRows = 0
    txtequipcode = ""
    txtexamcode = ""
    txtexamname = ""
    txtseqno = ""
    txtpaniclow = ""
    txtpanichigh = ""
End Sub

Public Sub MM_INITIAL()
    '/Resize를 위한 초기 Size Setting----------------------------------------------------------------------------------------------------/
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Menu
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is MSComm
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Size Setting
    lngMeHeight = 9015
    lngMeWidth = 9435
    
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    '/Resize를 위한 초기 Size Setting----------------------------------------------------------------------------------------------------/
    
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    
    '/변동 컨트롤 초기화----------------------------------------------------------------------------------------------------/
    Call MM_CANCEL
    '/변동 컨트롤 초기화----------------------------------------------------------------------------------------------------/
    
'    Call SET_CBO_DT_L(GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_PG_WORKLIST), cboPG_WORKLIST, 1)
'    txtWaitTime = "10"
'    Call SET_CBO_DT_L(GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_PG_QC), cboPG_QC, 1)
    
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:

Return
End Sub

Public Function MM_SAVE() As Boolean
    If vas검체결과.MaxRows = 0 Then
        MsgBox "저장할 검체별 결과가 없습니다.", vbCritical, "저장 불가": Exit Sub
    End If
    If MsgBox("저장 하겠습니까?", vbQuestion + vbOKCancel, "저장 질의") = vbCancel Then Exit Sub
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub
    
    ADC.BeginTrans
    
    For intY = 1 To vas검체결과.MaxRows
        gstrQuy = "UPDATE PAT_RES SET "
        gstrQuy = gstrQuy & vbCrLf & "       result    = '" & Trim(GET_CELL(vas검체결과, colResult, intY)) & "', "
        gstrQuy = gstrQuy & vbCrLf & "       sendflag  = '1' " '/1.결과저장, 2.결과전송
        gstrQuy = gstrQuy & vbCrLf & " WHERE examdate  = '" & Trim(GET_CELL(vas검체리스트, colEXAMDATE, vas검체리스트.ActiveRow)) & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND equipno   = '" & Trim(gtypREG_INFO.EQUIPCD) & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND barcode   = '" & Trim(GET_CELL(vas검체리스트, colBARCODE, vas검체리스트.ActiveRow)) & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND equipcode = '" & Trim(GET_CELL(vas검체결과, colEquipExam, intY)) & "' "
        If RunSQL(gstrQuy) = False Then
            ADC.RollbackTrans
            Call CloseDB
            Call ErrQuery(gstrQuy, 0)
            Exit Sub
        End If
    Next intY
    
    ADC.CommitTrans
    
    Call CloseDB
    
    Call cmdView_Click '/상태를 다시 보이기 위해...
    
    MsgBox "저장 되었습니다.", vbInformation, "확인"
End Function

Public Sub MM_VIEW()
    Dim lngRow  As Long
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM equipexam "
    gstrQuy = gstrQuy & vbCrLf & " WHERE equipno  = '" & gtypREG_INFO.EQUIPCD & "' "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY seqno "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If
    
    If Not ADR Is Nothing Then
        vasList.MaxRows = ARC
        
        Do Until ADR.EOF
            lngRow = lngRow + 1
        
            Call SET_CELL(vasList, 1, lngRow, Trim(ADR!equipcode & ""))     '/장비코드
            Call SET_CELL(vasList, 2, lngRow, Trim(ADR!ExamCode & ""))      '/검사코드
            Call SET_CELL(vasList, 3, lngRow, Trim(ADR!examname & ""))      '/검사명
            Call SET_CELL(vasList, 4, lngRow, Trim(ADR!seqno & ""))         '/순서
            Call SET_CELL(vasList, 5, lngRow, Trim(ADR!paniclow & ""))      '/Panic하한치
            Call SET_CELL(vasList, 6, lngRow, Trim(ADR!panichigh & ""))     '/Panic상한치
            
            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
        
        Call vasList.SetActiveCell(1, 1)
        Call vasList_DblClick(1, 1)
    End If
    Call CloseDB
End Sub

Private Sub cmdCancel_Click()
    txtequipcode = ""
    txtexamcode = ""
    txtexamname = ""
    txtseqno = ""
    txtpaniclow = ""
    txtpanichigh = ""
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtequipcode) = "" Then MsgBox "삭제할 장비코드를 (재)입력하십시오!", vbCritical, "삭제 불가": txtequipcode.SetFocus: Exit Sub
    If Trim(txtexamcode) = "" Then MsgBox "삭제할 검사코드를 (재)입력하십시오!", vbCritical, "삭제 불가": txtexamcode.SetFocus: Exit Sub
    
    If MsgBox("삭제하겠습니까?", vbQuestion + vbOKCancel, "삭제 질의") = vbCancel Then Exit Sub
              
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub
    
    ADC.BeginTrans

    gstrQuy = "DELETE FROM equipexam "
    gstrQuy = gstrQuy & vbCrLf & " WHERE equipno   = '" & gtypREG_INFO.EQUIPCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND equipcode = '" & txtequipcode & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND examcode  = '" & txtexamcode & "' "
    If RunSQL(gstrQuy) = False Then
        ADC.RollbackTrans
        Call CloseDB
        Call ErrQuery(gstrQuy, 0)
        Exit Sub
    End If
    
    ADC.CommitTrans
    
    Call CloseDB
    
    Call MM_VIEW '/상태를 다시 보이기 위해...
    
    MsgBox "삭제 되었습니다.", vbInformation, "확인"
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    '/입력항목Check
    If Trim(txtequipcode) = "" Then MsgBox "장비코드를 (재)입력하십시오!", vbCritical, "저장 불가": txtequipcode.SetFocus: Exit Sub
    If Trim(txtexamcode) = "" Then MsgBox "검사코드를 (재)입력하십시오!", vbCritical, "저장 불가": txtexamcode.SetFocus: Exit Sub
    If IsNumeric(txtseqno) = False Then MsgBox "순번을 (재)입력하십시오!", vbCritical, "저장 불가": txtseqno.SetFocus: Exit Sub
    
    If MsgBox("저장하겠습니까?", vbQuestion + vbOKCancel, "저장 질의") = vbCancel Then Exit Sub
              
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM equipexam "
    gstrQuy = gstrQuy & vbCrLf & " WHERE equipno   = '" & gtypREG_INFO.EQUIPCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND equipcode = '" & txtequipcode & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND examcode  = '" & txtexamcode & "' "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If

    If Not ADR Is Nothing Then
        ADC.BeginTrans
    
        gstrQuy = "UPDATE equipexam SET "
        gstrQuy = gstrQuy & vbCrLf & "       examname  = '" & Trim(txtexamname) & "', "
        gstrQuy = gstrQuy & vbCrLf & "       seqno     =  " & Val(txtseqno) & ", "
        gstrQuy = gstrQuy & vbCrLf & "       paniclow  = '" & Trim(txtpaniclow) & "', "
        gstrQuy = gstrQuy & vbCrLf & "       panichigh = '" & Trim(txtpanichigh) & "' "
        gstrQuy = gstrQuy & vbCrLf & " WHERE equipno   = '" & gtypREG_INFO.EQUIPCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND equipcode = '" & txtequipcode & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND examcode  = '" & txtexamcode & "' "
        If RunSQL(gstrQuy) = False Then
            ADC.RollbackTrans
            Call CloseDB
            Call ErrQuery(gstrQuy, 0)
            Exit Sub
        End If
        
        ADC.CommitTrans
    Else
        ADC.BeginTrans
        
        gstrQuy = "INSERT INTO equipexam "
        gstrQuy = gstrQuy & vbCrLf & " (equipno,    equipcode,  examcode,   examtype,   examname, "
        gstrQuy = gstrQuy & vbCrLf & "  resprec,    seqno,      reflow,     refhigh,    paniclow, "
        gstrQuy = gstrQuy & vbCrLf & "  panichigh,  deltavalue) "
        gstrQuy = gstrQuy & vbCrLf & " VALUES "
        gstrQuy = gstrQuy & vbCrLf & " ('" & gtypREG_INFO.EQUIPCD & "', "
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtequipcode) & "', "
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtexamcode) & "', "
        gstrQuy = gstrQuy & vbCrLf & "  '', "
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtexamname) & "', "
        gstrQuy = gstrQuy & vbCrLf & "   0, "
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(txtseqno) & ", "
        gstrQuy = gstrQuy & vbCrLf & "  '', "
        gstrQuy = gstrQuy & vbCrLf & "  '', "
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtpaniclow) & "', "
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtpanichigh) & "', "
        gstrQuy = gstrQuy & vbCrLf & "  '') "
        If RunSQL(gstrQuy) = False Then
            ADC.RollbackTrans
            Call CloseDB
            Call ErrQuery(gstrQuy, 0)
            Exit Sub
        End If
        
        ADC.CommitTrans
    End If
    
    Call CloseDB
    
    Call MM_VIEW '/상태를 다시 보이기 위해...
    
    MsgBox "저장 되었습니다.", vbInformation, "확인"
End Sub

Private Sub Form_Load()
    Call MM_INITIAL
    Call MM_VIEW
End Sub

Private Sub txtequipcode_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtequipcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtequipcode_LostFocus()
    txtequipcode = UCase(txtequipcode)
End Sub

Private Sub txtexamcode_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtexamcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtexamcode_LostFocus()
    txtexamcode = UCase(txtexamcode)
End Sub

Private Sub txtexamname_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtexamname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtpanichigh_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtpanichigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtpaniclow_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtpaniclow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtseqno_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtseqno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    txtequipcode = Trim(GET_CELL(vasList, 1, Row))
    txtexamcode = Trim(GET_CELL(vasList, 2, Row))
    txtexamname = Trim(GET_CELL(vasList, 3, Row))
    txtseqno = Trim(GET_CELL(vasList, 4, Row))
    txtpaniclow = Trim(GET_CELL(vasList, 5, Row))
    txtpanichigh = Trim(GET_CELL(vasList, 6, Row))
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vasList_DblClick(1, vasList.ActiveRow)
    End If
End Sub
