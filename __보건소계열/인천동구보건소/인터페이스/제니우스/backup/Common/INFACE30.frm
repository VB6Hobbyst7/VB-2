VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form INTtname30 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검사명 설정 및 수정"
   ClientHeight    =   6060
   ClientLeft      =   2625
   ClientTop       =   1530
   ClientWidth     =   5205
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
   ScaleWidth      =   5205
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
      Picture         =   "INFACE30.frx":033E
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
      MouseIcon       =   "INFACE30.frx":1CE0
      Picture         =   "INFACE30.frx":2132
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
   Dim ccode2 As Variant
   Dim RtnCd As Boolean
   Dim P As Integer
          
    Screen.MousePointer = 11
    maxrow = spdcode.MaxRows
    
    With tbcode
        For i = 1 To maxrow
            eqno = Format$(i, "00")
            With spdcode
                .GetText 1, i, cname
                .GetText 2, i, ccode
            End With
            '---------------------
            '[수정] 1999-08-25  김희정
            '   코드가 ','로 구분되어 있을 경우
            '   Current Row+30 으로 저장
            '---------------------
            P = InStr(ccode, ",")
            If P > 0 Then
                ccode2 = Mid(ccode, P + 1)
                .Index = "primarykey"
                .Seek "=", eqno + 30
                If .NoMatch = False Then
                    .Edit
                Else
                    .AddNew
                    !EQIPNO = eqno + 30
                End If
                !name = cname
                !Code = ccode2
                .Update
                ccode = Left(ccode, P - 1)
            Else
                .Index = "primarykey"
                .Seek "=", eqno + 30
                If .NoMatch = False Then .Delete
            End If
            
            .Index = "primarykey"
            .Seek "=", eqno
            If .NoMatch = False Then
                .Edit
            Else
                .AddNew
                !EQIPNO = eqno
            End If
            !name = cname
            !Code = ccode
            .Update
            
        Next i
    End With
    
    'Screen.MousePointer = 0
    'MsgBox "등록이 되었습니다.  확인을 누르신 후 다른 곳으로 이동하려면 닫기를 누르세요!!"
    Unload Me
    Screen.MousePointer = 0
End Sub


Private Sub cmdexit_Click()

    Unload Me
    FrmFlag = 0
End Sub

Private Sub Form_Load()

    Dim i%
    Dim Code As Variant
    
    'form을 가운데에 위치
    Me.Top = (INTmain00.Height - INTmain00.pnlMain.Height - Me.Height) / 3
    Me.Left = (INTmain00.Width - Me.Width) / 2
    
    Set dbcode = OpenDatabase(FileName & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable", dbOpenTable)

    MaxTestItem = tbcode.RecordCount
    
    If MaxTestItem = 0 Then Exit Sub
    
'    spdcode.MaxRows = 30
    i = 1
    tbcode.MoveFirst
    With tbcode
        Do Until .EOF
            If !EQIPNO <= spdcode.MaxRows Then
                Call spdsettext(spdcode, 1, i, !name)
                Call spdsettext(spdcode, 2, i, !Code)
                .MoveNext
                i = i + 1
            Else
                spdcode.GetText 2, !EQIPNO - 30, Code
                spdcode.SetText 2, !EQIPNO - 30, Code & "," & !Code
                .MoveNext
            End If
        Loop
    End With
    
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


