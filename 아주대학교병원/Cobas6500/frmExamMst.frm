VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmExamMst 
   BorderStyle     =   1  '단일 고정
   Caption         =   "검사오더"
   ClientHeight    =   7110
   ClientLeft      =   10125
   ClientTop       =   2355
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdInstRow 
      Caption         =   "줄삽입"
      Height          =   375
      Left            =   1050
      TabIndex        =   6
      Top             =   630
      Width           =   855
   End
   Begin VB.CommandButton cmdExamLoad 
      Caption         =   "검사불러오기"
      Height          =   375
      Left            =   1380
      TabIndex        =   5
      Top             =   150
      Width           =   1425
   End
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "줄삭제"
      Height          =   375
      Left            =   150
      TabIndex        =   4
      Top             =   630
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "화면정리"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   150
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "종료"
      Height          =   375
      Left            =   6630
      TabIndex        =   2
      Top             =   150
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   1155
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   6015
      Left            =   150
      TabIndex        =   0
      Top             =   1050
      Width           =   7665
      _Version        =   393216
      _ExtentX        =   13520
      _ExtentY        =   10610
      _StockProps     =   64
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmExamMst.frx":0000
   End
End
Attribute VB_Name = "frmExamMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    Display_Equip
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDelRow_Click()
    DeleteRow vasExam, vasExam.ActiveRow, vasExam.ActiveRow
End Sub

Private Sub cmdExamLoad_Click()
    ClearSpread vasExam
    
    SQL = "select EQUIPCODE, MAX(EXAMNAME) , '','','',MIN(SEQNO) from EQUIPEXAM "
    SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE"
    SQL = SQL & vbCrLf & " order by MIN(SEQNO)"
    res = db_select_Vas(gLocal, SQL, vasExam)
End Sub

Private Sub cmdInstRow_Click()
    InsertRow vasExam, vasExam.ActiveRow + 1
    SetText vasExam, GetText(vasExam, vasExam.ActiveRow, 1), vasExam.ActiveRow + 1, 1
    SetText vasExam, GetText(vasExam, vasExam.ActiveRow, 2), vasExam.ActiveRow + 1, 2
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    
    SQL = "DELETE FROM EXAMMST"
    res = SendQuery(gLocal, SQL)
    
    For i = 1 To vasExam.DataRowCnt
        If Trim(GetText(vasExam, i, 1)) <> "" And Trim(GetText(vasExam, i, 2)) <> "" Then
            SQL = "INSERT INTO EXAMMST(EQUIPCODE, EXAMNAME, LOW, HIGH, EQUAL, VALSTRING) "
            SQL = SQL & " VALUES('" & Trim(GetText(vasExam, i, 1)) & "', '" & Trim(GetText(vasExam, i, 2)) & "', "
            SQL = SQL & " '" & Trim(GetText(vasExam, i, 3)) & "','" & Trim(GetText(vasExam, i, 4)) & "',"
            SQL = SQL & " '" & Trim(GetText(vasExam, i, 5)) & "',"
            SQL = SQL & " '" & Trim(GetText(vasExam, i, 6)) & "')"
            res = SendQuery(gLocal, SQL)
        End If
    Next
    
    Display_Equip
End Sub


Private Sub Form_Load()
    
    Display_Equip
    
End Sub

Private Sub Display_Equip()
    ClearSpread vasExam
    
    SQL = "select EQUIPCODE, EXAMNAME, LOW, HIGH, EQUAL, VALSTRING from EXAMMST "
    SQL = SQL & vbCrLf & " WHERE LOW <> ''"
    SQL = SQL & vbCrLf & "order by EQUIPCODE, INT(LOW)"
    res = db_select_Vas(gLocal, SQL, vasExam)
    
    SQL = "select EQUIPCODE, EXAMNAME, LOW, HIGH, EQUAL, VALSTRING from EXAMMST "
    SQL = SQL & vbCrLf & " WHERE LOW = ''"
    SQL = SQL & vbCrLf & "order by EQUIPCODE, EQUAL"
    res = db_select_Vas(gLocal, SQL, vasExam, vasExam.DataRowCnt + 1)
    
    
End Sub

