VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmDisplayLot 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Lot No"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows 기본값
   Begin Threed.SSPanel SSPanel1 
      Height          =   525
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   926
      _StockProps     =   15
      Caption         =   "SSPanel1"
      ForeColor       =   0
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
   End
   Begin FPSpread.vaSpread vasLotNo 
      Height          =   4635
      Left            =   90
      TabIndex        =   0
      Top             =   660
      Width           =   5265
      _Version        =   196608
      _ExtentX        =   9287
      _ExtentY        =   8176
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   50
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "DisplayLot.frx":0000
   End
End
Attribute VB_Name = "frmDisplayLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sPart   As String
Dim sEquip  As String
Dim sLevel(0 To 4)  As String
Dim sLevels As String

    sPart = ""
    IsolateCode frmQCResSch.cboPart.Text
    sPart = Trim(gCode)
    
    SSPanel1.Caption = Trim(gName)
    
    For i = 0 To 4
        sLevel(i) = ""
    Next i
    
    For i = 0 To frmQCResSch.lstLevel(1).ListCount - 1
        SetText vasLotNo, frmQCResSch.lstLevel(1).List(i), 0, i + 2
    Next i

    vasLotNo.MaxCols = frmQCResSch.lstLevel(1).ListCount + 1

    If sPart <> "" Then
        SQL = "Select distinct(validstart) From qcexam " & CR & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              "   and equipcode = '" & sPart & "' " & vbCrLf & _
              " Order By validstart desc "
              
        res = db_select_Vas(gLocal, SQL, vasLotNo, 1, 1)
        
        ClearSpread Form_Main.vasTemp
    
        SQL = "Select Max(validstart), levelname, lotno " & CR & _
              "From qcexam " & CR & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              "   and equipcode = '" & sPart & "' " & vbCrLf & _
              "Group By levelname, lotno "
    
        res = db_select_Vas(gLocal, SQL, Form_Main.vasTemp)
              
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        For iRow = 1 To vasLotNo.DataRowCnt
            sAppDate = GetText(vasLotNo, iRow, 1)
            For iCol = 2 To vasLotNo.MaxCols
                sLevels = GetText(vasLotNo, 0, iCol)
                IsolateCode sLevels
                sLevels = Trim(gCode)
                For jRow = 1 To Form_Main.vasTemp.DataRowCnt
                    If sAppDate = GetText(Form_Main.vasTemp, jRow, 1) Then
                        If sLevels = GetText(Form_Main.vasTemp, jRow, 2) Then
                            SetText vasLotNo, GetText(Form_Main.vasTemp, jRow, 3), iRow, iCol
                            Exit For
                        End If
                    End If
                Next jRow
            Next iCol
        Next iRow
        
        '너비 설정
        If vasLotNo.MaxCols = 3 Then
            For iCol = 2 To vasLotNo.MaxCols
                With vasLotNo
                    .ColWidth(iCol) = 12
                    .Row = -1
                    .Col = i
                End With
            Next iCol
        Else
            For iCol = 2 To vasLotNo.MaxCols
                With vasLotNo
                    .ColWidth(iCol) = 9
                    .Row = -1
                    .Col = i
                End With
            Next iCol
        End If
    Else
        MsgBox "검사코드,레벨을 선택하세요.", vbInformation, "확인"
        Unload Me
    End If


End Sub

Private Sub vasLotNo_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim iCol As Integer

If GetText(vasLotNo, Row, 1) <> "" Then
    For iCol = 2 To vasLotNo.MaxCols
        If frmQCResSch.lstLevel(1).Selected(iCol - 2) = True Then
            frmQCResSch.txtLotNo(iCol - 2).Text = GetText(vasLotNo, Row, iCol)
        End If
    Next iCol
    Unload Me
End If


End Sub
