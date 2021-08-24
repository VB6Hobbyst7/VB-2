VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmDetail 
   Caption         =   "상세정보"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18570
   Icon            =   "frmDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   18570
   StartUpPosition =   2  '화면 가운데
   Begin FPSpreadADO.fpSpread spdDetail 
      Height          =   8325
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   18495
      _Version        =   458752
      _ExtentX        =   32623
      _ExtentY        =   14684
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
      SpreadDesigner  =   "frmDetail.frx":058A
   End
End
Attribute VB_Name = "frmDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Spd_Initializing()
    
    With spdDetail
    
        '.Reset
        .MaxRows = 0
        .MaxRows = 500
        
        .OperationMode = OperationModeRow
        .GridSolid = False
        
        .Appearance = Appearance3D
                
        'Hide row header
        .RowHeadersShow = False
        
        'Turn off font bold
        .Col = -1
        .Row = -1
        .FontBold = False
        
        'Change the amount of data each cell will hold
        .Col = -1
        .Row = -1
        .TypeEditLen = 200
        
        'Set column display type
        .ColHeaderDisplay = DispBlank
        .AllowCellOverflow = True
        .ReDraw = True
        
        .ShowScrollTips = ShowScrollTipsVertical
        .GrayAreaBackColor = &HFFFFFF
        
        .TextTip = TextTipFloating
        
        .MaxCols = 2
        .MaxRows = 0
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "관측소"
        '.ColWidth(1) = 7
        .SetText 2, 0, "관측시간"
        '.ColWidth(2) = 15
    End With
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Escape(27)
    If KeyCode = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    Call Spd_Initializing

End Sub
