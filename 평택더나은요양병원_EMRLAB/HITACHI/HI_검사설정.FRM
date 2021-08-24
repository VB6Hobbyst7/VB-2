VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm검사명설정 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "검사항목 설정 및 수정"
   ClientHeight    =   5985
   ClientLeft      =   1800
   ClientTop       =   1035
   ClientWidth     =   5280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread spdOrdCd 
      Height          =   4185
      Left            =   420
      OleObjectBlob   =   "Hi_검사설정.frx":0000
      TabIndex        =   0
      Top             =   420
      Width           =   4470
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   735
      Left            =   780
      TabIndex        =   1
      Top             =   4920
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3351
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "등      록"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   2
      RoundedCorners  =   0   'False
      MouseIcon       =   "Hi_검사설정.frx":0327
      Picture         =   "Hi_검사설정.frx":0343
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   735
      Left            =   2670
      TabIndex        =   2
      Top             =   4920
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3351
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "취      소"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   2
      RoundedCorners  =   0   'False
      Picture         =   "Hi_검사설정.frx":0C1D
   End
End
Attribute VB_Name = "frm검사명설정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *=====================================================*
' *    Test Item Name&Code Display                      *
' *=====================================================*
Sub Disp_Test()
    
    Dim i   As Integer
            
    With spdORDCD
        .Col = -1: .row = -1
        .Action = SS_ACTION_CLEAR_TEXT    'Spread 초기화
            
        For i = 1 To MAX_NUM
            Call .SetText(1, i, TestNameTable(i).EqCd)  ' 장비코드
            Call .SetText(2, i, TestNameTable(i).Name)  ' 검사명
            Call .SetText(3, i, TestNameTable(i).code)  ' 검사코드
        Next i
    End With
        
End Sub

Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub cmdSave_Click()

    Dim i   As Integer
    Dim EqipCd  As String
    
    Dim temp    As Variant
    Dim FileName    As String
        
    MousePointer = 11
    
    '----- DataBase Open
    FileName = strLAB_PATH & "System\Config.mdb"
    Set Db = OpenDatabase(FileName)
    Set TableCd = Db.OpenRecordset("Test_Item")
    
    '----- Spread의 내용을 구조체에 저장
    For i = 1 To MAX_NUM
        With spdORDCD
            Call .GetText(1, i, temp)
            TestNameTable(i).EqCd = Trim(temp)
            Call .GetText(2, i, temp)
            TestNameTable(i).Name = Trim(temp)
            Call .GetText(3, i, temp)
            TestNameTable(i).code = Trim(temp)
        End With
    Next i

    '----- 검사항목Record가 존재하는지 여부 검사하여 Table에 저장
    For i = 1 To MAX_NUM
        If RecordExist(TableCd, "PrimaryKey", Format(i, "00")) Then
            TableCd.Edit
        Else
            TableCd.AddNew
            TableCd!seqno = Format(i, "00")
        End If
        TableCd!Eq_Code = Trim(TestNameTable(i).EqCd)
        TableCd!ordnm = Trim(TestNameTable(i).Name)
        TableCd!ordcd = Trim(TestNameTable(i).code)
                        
        TableCd.Update
        TableCd.MoveFirst
    Next i

    '----- DataBase Close
    TableCd.Close
    Db.Close
    
    MousePointer = 0
    
    cmdExit.Caption = "닫      기"
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Call cmdExit.DoClick
    End Select
    
End Sub

Private Sub Form_Load()

    Me.Top = (Hi_Main.Height - Hi_Main.pnlBack.Height - Me.Height) / 3
    Me.Left = (Hi_Main.Width - Me.Width) / 2
    
    Screen.MousePointer = 11
    
    '검사항목 조회/표시
    Call Disp_Test
        
    Screen.MousePointer = 0
    
End Sub

