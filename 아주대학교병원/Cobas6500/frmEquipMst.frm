VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmEquipMst 
   BorderStyle     =   1  '단일 고정
   Caption         =   "검사오더"
   ClientHeight    =   6885
   ClientLeft      =   4260
   ClientTop       =   2505
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6120
   Begin VB.CommandButton cmdClear 
      Caption         =   "화면정리"
      Height          =   375
      Left            =   1350
      TabIndex        =   3
      Top             =   150
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "종료"
      Height          =   375
      Left            =   2550
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
   Begin FPSpread.vaSpread vasEquip 
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5865
      _Version        =   393216
      _ExtentX        =   10345
      _ExtentY        =   10874
      _StockProps     =   64
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frmEquipMst.frx":0000
   End
End
Attribute VB_Name = "frmEquipMst"
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

Private Sub cmdSave_Click()
    Dim i As Integer
    
    SQL = "DELETE FROM EQUIPMST"
    res = SendQuery(gLocal, SQL)
    
    For i = 1 To vasEquip.DataRowCnt
        If Trim(GetText(vasEquip, i, 1)) <> "" And Trim(GetText(vasEquip, i, 2)) <> "" Then
            SQL = "INSERT INTO EQUIPMST(EQUIPCODE, EXAMCODE, EXAMNAME) "
            SQL = SQL & " VALUES('" & Trim(GetText(vasEquip, i, 1)) & "', '" & Trim(GetText(vasEquip, i, 2)) & "', "
            SQL = SQL & " '" & Trim(GetText(vasEquip, i, 3)) & "')"
            res = SendQuery(gLocal, SQL)
        End If
    Next
    
    Display_Equip
End Sub

Private Sub Form_Load()
    
    Display_Equip
    
End Sub

Private Sub Display_Equip()
    ClearSpread vasEquip
    
    SQL = "select EQUIPCODE, EXAMCODE, EXAMNAME from equipmst order by EQUIPCODE"
    res = db_select_Vas(gLocal, SQL, vasEquip)
    
End Sub
