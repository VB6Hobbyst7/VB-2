VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form INTtname30 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사명 설정 및 수정"
   ClientHeight    =   6060
   ClientLeft      =   2625
   ClientTop       =   1530
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread spdcode 
      Height          =   4620
      Left            =   360
      OleObjectBlob   =   "INFACE30.frx":0000
      TabIndex        =   0
      Top             =   1110
      Width           =   4425
   End
   Begin Threed.SSCommand cmdexit 
      Height          =   870
      Left            =   3975
      TabIndex        =   2
      Top             =   120
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "종   료"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE30.frx":0334
   End
   Begin Threed.SSCommand cmdenrole 
      Height          =   870
      Left            =   3150
      TabIndex        =   1
      Top             =   120
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "저   장"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      MouseIcon       =   "INFACE30.frx":1CD6
      Picture         =   "INFACE30.frx":2128
   End
End
Attribute VB_Name = "INTtname30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdenrole_Click()

   Dim maxrow%, i%, eqno$, cname, ccode
   Dim RtnCd As Boolean
          
    Screen.MousePointer = 11
    maxrow = spdcode.MaxRows
    
    For i = 1 To maxrow
        eqno = Format$(i, "00")
        RtnCd = spdcode.GetText(1, i, cname)
        RtnCd = spdcode.GetText(2, i, ccode)
        
        tbcode.Index = "primarykey"
        tbcode.Seek "=", eqno
        If tbcode.NoMatch = False Then
            With tbcode
                .Edit
                !Name = cname & " "
                !code = Trim$(ccode) & " "
                .Update
            End With
        Else
            tbcode.AddNew
            tbcode!EQIPNO = eqno
            tbcode!Name = cname & " "
            tbcode!code = ccode & " "
            tbcode.Update
        End If
    Next i
    'Screen.MousePointer = 0
    'MsgBox "등록이 되었습니다.  확인을 누르신 후 다른 곳으로 이동하려면 닫기를 누르세요!!"
    Unload Me
    Screen.MousePointer = 0
    
End Sub


Private Sub cmdExit_Click()

    Unload Me
    FrmFlag = 0
End Sub

Private Sub Form_Load()

    Dim i%
    
    'form을 가운데에 위치
    Me.Top = (INTmain00.Height - INTmain00.pnlMain.Height - Me.Height) / 3
    Me.Left = (INTmain00.Width - Me.Width) / 2
    
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable", dbOpenTable)

    MaxTestItem = 99 ' tbcode.RecordCount
    
    If MaxTestItem = 0 Then Exit Sub
    
    spdcode.MaxRows = MaxTestItem
    i = 1
    If tbcode.RecordCount > 0 Then tbcode.MoveFirst
    Do Until tbcode.EOF
    
        i = Val(tbcode!EQIPNO)
        
        Call spdsettext(spdcode, 1, i, tbcode!Name)
        Call spdsettext(spdcode, 2, i, tbcode!code)
        
        tbcode.MoveNext
    
    Loop
    
    FrmFlag = 30
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    tbcode.Close
    dbcode.Close
    
End Sub


Private Sub spdcode_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If spdcode.ActiveCol = 2 Then
            spdcode.Col = 0
            spdcode.Row = spdcode.ActiveRow + 1
            spdcode.Action = SS_ACTION_ACTIVE_CELL
        End If
        KeyAscii = 0
    End If
    
End Sub


