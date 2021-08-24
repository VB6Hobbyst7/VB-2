VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmData 
   Caption         =   "data"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4290
      TabIndex        =   1
      Top             =   5190
      Width           =   1155
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5445
      _Version        =   393216
      _ExtentX        =   9604
      _ExtentY        =   8916
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
      MaxRows         =   1000
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmData.frx":0000
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    Dim lsCode, lsOrdcode As String
    Dim lRow As Long
    
    For lRow = 1 To vasTemp.DataRowCnt
        lsOrdcode = Trim(GetText(vasTemp, lRow, 1))
        lsCode = Mid(lsOrdcode, 2)
        
        vasTemp.SetText 2, lRow, lsCode
        
        SQL = "Update equipexam set ordcode = '" & lsOrdcode & "' where examcode = '" & lsCode & "' "
        res = SendQuery(gLocal, SQL)
        
        
    Next lRow
    
    MsgBox "완료"
End Sub
