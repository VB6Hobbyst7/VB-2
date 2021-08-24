VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PrintDlg 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Print"
   ClientHeight    =   2355
   ClientLeft      =   2160
   ClientTop       =   2025
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Printer Setup"
      Height          =   375
      Left            =   3135
      TabIndex        =   19
      Top             =   1860
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Options"
      Height          =   1515
      Left            =   2820
      TabIndex        =   9
      Top             =   180
      Width           =   3750
      Begin VB.CheckBox Check1 
         Caption         =   "Shadows"
         Height          =   255
         Index           =   6
         Left            =   2025
         TabIndex        =   18
         Top             =   1200
         Value           =   1  '확인
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Color"
         Height          =   255
         Index           =   5
         Left            =   2025
         TabIndex        =   16
         Top             =   900
         Value           =   1  '확인
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Data Cells Only"
         Height          =   255
         Index           =   4
         Left            =   2025
         TabIndex        =   15
         Top             =   600
         Value           =   1  '확인
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Border"
         Height          =   255
         Index           =   3
         Left            =   2025
         TabIndex        =   13
         Top             =   300
         Value           =   1  '확인
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Grid Lines"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   900
         Value           =   1  '확인
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Row Headers"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Value           =   1  '확인
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Column Headers"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Value           =   1  '확인
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page Range"
      Height          =   1515
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   2715
      Begin VB.CommandButton Command1 
         Caption         =   "Setup"
         Height          =   315
         Left            =   1740
         TabIndex        =   17
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1740
         TabIndex        =   8
         Text            =   "1"
         Top             =   1020
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Text            =   "1"
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pages"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1020
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Current Page"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Selected Cells"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   255
         Left            =   1500
         TabIndex        =   7
         Top             =   1080
         Width           =   435
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   375
      Left            =   1590
      TabIndex        =   0
      Top             =   1860
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   450
      Top             =   1815
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "PrintDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
'    PageSetup.Show
End Sub

Private Sub Command2_Click()
    PrintSpread
    Unload frmSpPreview
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Sub PrintSpread()
    '** 출력폼 선택(폼 이름)==============================================
    Select Case Gbl_FormName
        Case "frmStatistics"
            With frmStatistics
                'Set printing options for spreadsheet
                .spdResult1.PrintColHeaders = Check1(0).Value
                .spdResult1.PrintRowHeaders = Check1(1).Value
                .spdResult1.PrintBorder = Check1(3).Value
                .spdResult1.PrintColor = Check1(5).Value
                .spdResult1.PrintGrid = Check1(2).Value
                .spdResult1.PrintShadows = Check1(6).Value
                .spdResult1.PrintUseDataMax = Check1(4).Value
                
                'Page Range
                'All
                If Option1(0).Value = True Then
                    .spdResult1.PrintType = PrintTypeAll
                    
                'Selected cells
                ElseIf Option1(1).Value = True Then
                    .spdResult1.Col = .spdResult1.SelBlockCol
                    .spdResult1.Col2 = .spdResult1.SelBlockCol2
                    .spdResult1.Row = .spdResult1.SelBlockRow
                    .spdResult1.Row2 = .spdResult1.SelBlockRow2
                    .spdResult1.PrintType = PrintTypeCellRange
                    
                'Current Page
                ElseIf Option1(2).Value = True Then
                    .spdResult1.PrintType = PrintTypeCurrentPage
                    
                'Pages
                Else
                    .spdResult1.PrintPageStart = CInt(Text1(0).Text)
                    .spdResult1.PrintPageEnd = CInt(Text1(1).Text)
                    .spdResult1.PrintType = PrintTypePageRange
                End If
                
                'Print control
                Screen.MousePointer = 11
                .spdResult1.Action = ActionPrint
                Screen.MousePointer = 0
            End With
            
        Case Else
            
    End Select
    '=====================================================================
End Sub

Private Sub Command4_Click()
    CommonDialog1.ShowPrinter
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 3 Then
        Text1(0).Enabled = True
        Text1(1).Enabled = True
        Text1(0).SetFocus
    Else
        Text1(0).Enabled = False
        Text1(1).Enabled = False
    End If
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Verify if a numeric number
    
    If Not IsNumeric(Text1(Index)) Then
        Text1(Index).Text = "1"
    End If
End Sub
