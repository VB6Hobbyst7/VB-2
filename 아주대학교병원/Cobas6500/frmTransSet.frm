VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTransSet 
   Caption         =   "자동전송 설정"
   ClientHeight    =   7980
   ClientLeft      =   1875
   ClientTop       =   2295
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "종료"
      Height          =   375
      Left            =   4980
      TabIndex        =   4
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "화면정리"
      Height          =   375
      Left            =   2790
      TabIndex        =   3
      Top             =   180
      Width           =   1065
   End
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "줄삭제"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdExamLoad 
      Caption         =   "검사불러오기"
      Height          =   375
      Left            =   1290
      TabIndex        =   1
      Top             =   180
      Width           =   1425
   End
   Begin VB.CommandButton cmdInstRow 
      Caption         =   "줄삽입"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   6015
      Left            =   60
      TabIndex        =   6
      Top             =   1860
      Width           =   6075
      _Version        =   393216
      _ExtentX        =   10716
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
      SpreadDesigner  =   "frmTransSet.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "※NEMD(신장내과)의 환자결과는 자동전송 하지 않습니다."
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1920
      TabIndex        =   8
      Top             =   1410
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "※아래 항목의 결과에 해당하지 않을때 검경 결과가 자동으로 전송됩니다."
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   4215
   End
End
Attribute VB_Name = "frmTransSet"
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
    
    SQL = "DELETE FROM TRANSMST"
    res = SendQuery(gLocal, SQL)
    
    For i = 1 To vasExam.DataRowCnt
        If Trim(GetText(vasExam, i, 1)) <> "" And Trim(GetText(vasExam, i, 2)) <> "" Then
            SQL = "INSERT INTO TRANSMST(EQUIPCODE, EXAMNAME, LOW, HIGH, EQUAL) "
            SQL = SQL & " VALUES('" & Trim(GetText(vasExam, i, 1)) & "', '" & Trim(GetText(vasExam, i, 2)) & "', "
            SQL = SQL & " '" & Trim(GetText(vasExam, i, 3)) & "','" & Trim(GetText(vasExam, i, 4)) & "',"
            SQL = SQL & " '" & Trim(GetText(vasExam, i, 5)) & "')"
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
    
    SQL = "select EQUIPCODE, EXAMNAME, LOW, HIGH, EQUAL from TRANSMST "
    SQL = SQL & vbCrLf & " WHERE LOW <> ''"
    SQL = SQL & vbCrLf & "order by EQUIPCODE, INT(LOW)"
    res = db_select_Vas(gLocal, SQL, vasExam)
    
    SQL = "select EQUIPCODE, EXAMNAME, LOW, HIGH, EQUAL from TRANSMST "
    SQL = SQL & vbCrLf & " WHERE LOW = ''"
    SQL = SQL & vbCrLf & "order by EQUIPCODE, EQUAL"
    res = db_select_Vas(gLocal, SQL, vasExam, vasExam.DataRowCnt + 1)
    
    
End Sub


