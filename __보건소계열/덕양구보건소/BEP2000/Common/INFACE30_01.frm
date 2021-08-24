VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form INTtname30 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사명 설정 및 수정"
   ClientHeight    =   6060
   ClientLeft      =   2625
   ClientTop       =   1530
   ClientWidth     =   7230
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
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread spdcode 
      Height          =   4620
      Left            =   360
      TabIndex        =   0
      Top             =   1110
      Width           =   6735
      _Version        =   196608
      _ExtentX        =   11880
      _ExtentY        =   8149
      _StockProps     =   64
      ColsFrozen      =   1
      EditEnterAction =   2
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   50
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "INFACE30_01.frx":0000
      UserResize      =   0
      VisibleCols     =   500
      VisibleRows     =   500
   End
   Begin Threed.SSCommand cmdexit 
      Height          =   870
      Left            =   6255
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
      Picture         =   "INFACE30_01.frx":031E
   End
   Begin Threed.SSCommand cmdenrole 
      Height          =   870
      Left            =   5400
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
      MouseIcon       =   "INFACE30_01.frx":1CC0
      Picture         =   "INFACE30_01.frx":2112
   End
End
Attribute VB_Name = "INTtname30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdenrole_Click()

   Dim maxrow%, i%, seqno$, sName$, sCode$, sEtc$
   Dim RtnCd As Boolean
   Dim vTmp     As Variant
   
    Me.MousePointer = 11
    maxrow = spdcode.MaxRows
    
    For i = 1 To maxrow
        
        spdcode.GetText 1, i, vTmp: seqno = Format$(vTmp, "00")
        spdcode.GetText 2, i, vTmp: sName = Trim$(vTmp)
        spdcode.GetText 3, i, vTmp: sCode = Trim$(vTmp)
        spdcode.GetText 4, i, vTmp: sEtc = Trim$(vTmp)
        
        If seqno = "000" Or seqno = "" Then Exit For
        
        tbcode.Index = "primarykey"
        tbcode.Seek "=", seqno
        If tbcode.NoMatch = False Then
            With tbcode
                .Edit
                !Name = Trim(sName) & " "
                !code = Trim$(sCode) & " "
                !etc = Trim(sEtc) & " "
                
                .Update
            End With
        Else
            tbcode.AddNew
            tbcode!EQIPNO = seqno
            tbcode!Name = Trim(sName) & " "
            tbcode!code = Trim(sCode) & " "
            tbcode!etc = Trim(sEtc) & " "
            tbcode.Update
        End If
    Next i
    'Screen.MousePointer = 0
    'MsgBox "등록이 되었습니다.  확인을 누르신 후 다른 곳으로 이동하려면 닫기를 누르세요!!"
    Unload Me
    Me.MousePointer = 0
    
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
    i = 0
    If tbcode.RecordCount > 0 Then tbcode.MoveFirst
    Do Until tbcode.EOF
    
        i = i + 1
        Call spdsettext(spdcode, 1, i, tbcode!EQIPNO)
        Call spdsettext(spdcode, 2, i, tbcode!Name)
        Call spdsettext(spdcode, 3, i, tbcode!code)
        Call spdsettext(spdcode, 4, i, tbcode!etc)
        
        tbcode.MoveNext
    
    Loop
    
    FrmFlag = 30
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    tbcode.Close
    dbcode.Close
    
End Sub


Private Sub spdcode_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        If spdcode.ActiveCol = 2 Then
'            spdcode.Col = 0
'            spdcode.Row = spdcode.ActiveRow + 1
'            spdcode.Action = SS_ACTION_ACTIVE_CELL
'        End If
'        KeyAscii = 0
'    End If
    
End Sub


