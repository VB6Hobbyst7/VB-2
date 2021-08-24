VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS924 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "월별 혈액입출고 현황"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS924.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   4
      Left            =   1860
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "용량 "
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   6
      Left            =   90
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "혈액형"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   270
      Left            =   3885
      TabIndex        =   6
      Top             =   1605
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   476
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "4월 혈액 입고 현황"
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15101"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker dtpMonth 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "gg yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2670
      TabIndex        =   1
      Top             =   60
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy-MM"
      Format          =   25034755
      CurrentDate     =   36799
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "월별혈액입출고현황(월선택)"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6300
      Left            =   75
      TabIndex        =   5
      Top             =   1965
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   11113
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   35
      MaxRows         =   20
      OperationMode   =   1
      RowsFrozen      =   1
      ShadowColor     =   14737632
      ShadowDark      =   13818331
      SpreadDesigner  =   "frmBBS924.frx":076A
      TextTip         =   4
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1305
      Left            =   75
      TabIndex        =   10
      Top             =   285
      Width           =   10770
      Begin VB.CheckBox chkALL 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체혈액형"
         Height          =   255
         Index           =   0
         Left            =   4770
         TabIndex        =   23
         Top             =   210
         Width           =   1215
      End
      Begin VB.ComboBox cboCenter 
         Height          =   300
         ItemData        =   "frmBBS924.frx":8C8B
         Left            =   1170
         List            =   "frmBBS924.frx":8C8D
         Style           =   2  '드롭다운 목록
         TabIndex        =   22
         Top             =   135
         Width           =   2415
      End
      Begin VB.ComboBox cboVolume 
         Height          =   300
         ItemData        =   "frmBBS924.frx":8C8F
         Left            =   1170
         List            =   "frmBBS924.frx":8C9F
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   495
         Width           =   2415
      End
      Begin VB.ComboBox cboDiv 
         Height          =   300
         ItemData        =   "frmBBS924.frx":8CBC
         Left            =   1170
         List            =   "frmBBS924.frx":8CCC
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   855
         Width           =   2415
      End
      Begin VB.CheckBox chkALL 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Rh 전체"
         Height          =   255
         Index           =   1
         Left            =   6315
         TabIndex        =   19
         Top             =   915
         Width           =   1050
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9195
         Style           =   1  '그래픽
         TabIndex        =   18
         Tag             =   "15101"
         Top             =   645
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   495
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "조회용량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "조회건물"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   3690
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   510
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "혈 액 형"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   855
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "조회구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   3690
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   870
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Rh "
         Appearance      =   0
      End
      Begin VB.Frame fraSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   420
         Left            =   4785
         TabIndex        =   24
         Tag             =   "136"
         Top             =   420
         Width           =   2460
         Begin VB.OptionButton optABO 
            BackColor       =   &H00DBE6E6&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   28
            Tag             =   "15304"
            Top             =   135
            Width           =   495
         End
         Begin VB.OptionButton optABO 
            BackColor       =   &H00DBE6E6&
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   585
            TabIndex        =   27
            Tag             =   "15305"
            Top             =   135
            Width           =   450
         End
         Begin VB.OptionButton optABO 
            BackColor       =   &H00DBE6E6&
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1140
            TabIndex        =   26
            Tag             =   "15304"
            Top             =   135
            Width           =   495
         End
         Begin VB.OptionButton optABO 
            BackColor       =   &H00DBE6E6&
            Caption         =   "AB"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   1725
            TabIndex        =   25
            Tag             =   "15305"
            Top             =   135
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   420
         Left            =   4785
         TabIndex        =   29
         Tag             =   "136"
         Top             =   780
         Width           =   1020
         Begin VB.OptionButton optRh 
            BackColor       =   &H00DBE6E6&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   525
            TabIndex        =   31
            Tag             =   "15305"
            Top             =   120
            Width           =   390
         End
         Begin VB.OptionButton optRh 
            BackColor       =   &H00DBE6E6&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   30
            Tag             =   "15304"
            Top             =   120
            Width           =   435
         End
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   7
      Left            =   8415
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "센터"
      Appearance      =   0
   End
   Begin VB.Label lblCenter 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label7"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9495
      TabIndex        =   9
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label lblVo 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label7"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2925
      TabIndex        =   8
      Top             =   1620
      Width           =   690
   End
   Begin VB.Label lblabo 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label7"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1155
      TabIndex        =   7
      Top             =   1620
      Width           =   690
   End
End
Attribute VB_Name = "frmBBS924"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mode As Long  '1:입고,2:출고,3:반환,4:폐기
Private QueryFg As Boolean
Private lngCompCnt As Long

Private Sub cboCenter_Click()
    lblCenter.Caption = cboCenter.Text
    
    Call cmdClear_Click
End Sub

Private Sub cboDiv_Click()
    lblTitle.Caption = Format(dtpMonth.Value, "mm") & "월 혈액 " & cboDiv.Text & " 현황"
    With tblList
        Call .SetText(2, .MaxRows, cboDiv.Text & "계")
    End With
    
    Call cmdClear_Click
End Sub

Private Sub cboVolume_Click()
    lblVo.Caption = cboVolume.Text
    
    Call cmdClear_Click
End Sub

Private Sub chkALL_Click(Index As Integer)
    Call cmdClear_Click
End Sub

Private Sub cmdClear_Click()
    Dim ii As Integer
    Dim jj As Integer
    
    With tblList
        For ii = 2 To .MaxRows
            .Row = ii
            For jj = 3 To .MaxCols - 1
                .Col = jj
                .Value = ""
            Next
        Next
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Call cmdClear_Click
    Call DirectQuery

    Exit Sub
    
    
    QueryFg = True
    Call cmdClear_Click
    Call cmdPrint_Click
    QueryFg = False
End Sub

Private Sub dtpMonth_Change()
    lblTitle.Caption = Format(dtpMonth.Value, "mm") & "월 혈액 " & cboDiv.Text & " 현황"
    
    Call cmdClear_Click
End Sub

Private Sub Form_Load()

    Dim objcom003 As clsCom003
    Dim objSql    As clsStatics
    Dim RS        As Recordset
    Dim ii        As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    QueryFg = False
    
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter, True)
    
    optABO(0).Value = True
    optRh(0).Value = True
    dtpMonth.Value = Format(GetSystemDate, "yyyy-mm")
    cboCenter.ListIndex = 0
    cboVolume.ListIndex = 0
    cboDiv.ListIndex = 0
    chkALL(0).Value = 0
    chkALL(1).Value = 0
    
    Set objSql = New clsStatics
    Set RS = objSql.Get_GroupCompo
    
'    Call medClearTable(tblList)
    
    If Not RS.EOF Then
        lngCompCnt = RS.RecordCount
        
        '센터 제목지우기
        With tblList
            .MaxRows = (lngCompCnt * 3) + 2
            .RowHeight(-1) = 13.3
            .Row = 2: .Row2 = .MaxRows
            .Col = 1: .Col2 = 1
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
        End With
        'border line 없애기
        With tblList
            .Row = 2: .Row2 = .MaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderType = 1 Or 2 Or 4 Or 8
            .CellBorderColor = vbWhite
            .Action = 16
            .BlockMode = False
        End With
        
        For i = 0 To 2 * lngCompCnt Step lngCompCnt
            j = 0
            RS.MoveFirst
            Do Until RS.EOF
                With tblList
                    .Row = 2 + i + j: .Col = 2: .Value = RS.Fields("field1").Value & ""
                                      .Col = 35: .Value = RS.Fields("cdval1").Value & ""
'                                      .Col = 1: .Value = RS.Fields("cdval1").Value & ""
                    j = j + 1
                End With
                RS.MoveNext
            Loop
            
'            센터 블럭별 색깔 변경
            With tblList
                .Row = 2 + i: .Row2 = 2 + i + lngCompCnt
                .Col = 1: .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = Choose(i / lngCompCnt + 1, RGB(255, 255, 202), RGB(238, 255, 255), RGB(255, 208, 255))
                .BlockMode = False
                
'            센터 블럭 border line 그리기
                .Row = 2 + i: .Row2 = 2 + i + lngCompCnt
                .Col = 1: .Col2 = 1
                .BlockMode = True
                .CellBorderStyle = CellBorderStyleSolid
                .CellBorderType = 16 '4 Or 8
                .CellBorderColor = vbBlack
                .Action = 16
                .BlockMode = False
'            센터명 등록하기
                Call .SetText(1, 2 + i, Choose(i / lngCompCnt + 1, "혈", "헌", "외"))
                Call .SetText(1, 2 + i + 1, Choose(i / lngCompCnt + 1, "액", "", ""))
                Call .SetText(1, 2 + i + 2, Choose(i / lngCompCnt + 1, "원", "혈", "부"))
                For k = 2 + i + 3 To lngCompCnt
                    Call .SetText(1, k, "")
                Next
            End With
        Next i
        
        With tblList
            '입고계 그리기
            Call .SetText(2, .MaxRows, "입고계")
            .Row = .MaxRows: .Row2 = .MaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = RGB(213, 170, 255)
            .BlockMode = False
            
            'border line 그리기
            .Row = 2: .Row2 = .MaxRows
            .Col = 2: .Col2 = .MaxCols
            .BlockMode = True
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderType = 1 Or 2 Or 4 Or 8
            .CellBorderColor = vbBlack
            .Action = 16
            .BlockMode = False
            
            '입고계 보더라인 그리기
            .Row = .MaxRows: .Row2 = .MaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderType = 1 Or 2 Or 4 Or 8
            .CellBorderColor = vbBlack
            .Action = 16
            .BlockMode = False
        End With
    End If
    
'    If Not RS.EOF Then
'        Do Until RS.EOF
'            RS.MoveFirst
'            With tblList
'                For ii = 2 To 7
'
'                    .Row = ii:      .Col = 2: .Value = RS.Fields("field1").Value & ""
'                                    .Col = 35: .Value = RS.Fields("cdval1").Value & ""
'                    .Row = ii + 6:  .Col = 2: .Value = RS.Fields("field1").Value & ""
'                                    .Col = 35: .Value = RS.Fields("cdval1").Value & ""
'                    .Row = ii + 12: .Col = 2: .Value = RS.Fields("field1").Value & ""
'                                    .Col = 35: .Value = RS.Fields("cdval1").Value & ""
'                    RS.MoveNext
'                Next
'            End With
'        Loop
'    End If
    
    Call cmdClear_Click
    
    optRh(0).Value = True
    optABO(0).Value = True
    
    lblabo.Caption = optABO(0).Caption & optRh(0).Caption
    lblCenter.Caption = cboCenter.Text
    lblVo.Caption = cboVolume.Text
    lblTitle.Caption = Format(dtpMonth.Value, "mm") & "월 혈액 " & cboDiv.Text & " 현황"
    
    Set RS = Nothing
    Set objcom003 = Nothing
    Set objSql = Nothing
End Sub





Private Sub cmdPrint_Click()
    Dim objStatics As New clsStatics
    Dim FMonth     As String
    Dim TMonth     As String
    Dim Centercd   As String
    Dim Centernm   As String
    Dim Volume     As String
    Dim ABO        As String
    Dim Rh         As String
    
    If cboCenter.Text = "(ALL)" Then
        Centercd = "ALL": Centernm = "ALL"
    Else
        Centercd = medGetP(cboCenter.Text, 1, " ")
        Centernm = medGetP(cboCenter.Text, 2, " ")
    End If
    
    Select Case cboVolume.ListIndex
        Case 0: Volume = "ALL"
        Case 1: Volume = "320"
        Case 2: Volume = "400"
        Case 3: Volume = "Etc"
    End Select
    
    Select Case cboDiv.ListIndex
        Case 0: mode = 1
        Case 1: mode = 2
        Case 2: mode = 3
        Case 3: mode = 4
    End Select
    
    If chkALL(0).Value = 0 Then
        If optABO(0).Value = True Then ABO = "A"
        If optABO(1).Value = True Then ABO = "B"
        If optABO(2).Value = True Then ABO = "O"
        If optABO(3).Value = True Then ABO = "AB"
    Else
        ABO = "ALL"
    End If
    
    If chkALL(1).Value = 0 Then
        If optRh(0).Value = True Then Rh = "+"
        If optRh(1).Value = True Then Rh = "-"
    Else
        Rh = "ALL"
    End If
    
    
    FMonth = Format(dtpMonth.Value, "yyyymm") & "01"
    TMonth = Format(dtpMonth.Value, "yyyymm") & "31"
    
'    objStatics.setDbConn DBConn
    
    If objStatics.bloodcnt(FMonth, TMonth, mode) = True Then
        Call Query(FMonth, TMonth, Centercd, Centernm, ABO, Rh, Volume, mode)
    Else
        MsgBox "해당자료가 없습니다", vbInformation + vbOKOnly, "혈액현황출력"
    End If
    Set objStatics = Nothing
    
End Sub
Private Sub Query(ByVal FMonth As String, ByVal TMonth As String, ByVal Centercd As String, ByVal Centernm As String, _
                  ByVal ABO As String, ByVal Rh As String, ByVal Volume As String, ByVal mode As String)

    Dim objStatics As New clsStatics
    Dim RS         As New Recordset
    Dim objDic     As clsDictionary
    Dim objPrint   As clsInOutStatus
    Dim GRs        As Recordset
    
'    objStatics.setDbConn DBConn
    Set RS = objStatics.BloodCondition(FMonth, TMonth, Centercd, ABO, Rh, Volume, "", mode)
    If RS.RecordCount > 1 Then
        Set objDic = New clsDictionary
        Set objDic = objStatics.Bld_Dic(RS)
    Else
        Set RS = Nothing
        Set objStatics = Nothing
        Exit Sub
    End If
    
    Set GRs = New Recordset
    Set objPrint = New clsInOutStatus
    Set GRs = objStatics.Get_GroupCompo
    
    
    If QueryFg = True Then
        Call Display(objDic)
    Else
        objPrint.Hearder_Line GRs, Centernm, ABO, Rh, Volume, Format(Mid(FMonth, 5, 2), "##"), mode, objDic
    End If
    Set RS = Nothing
    Set GRs = Nothing
    Set objDic = Nothing
    Set objPrint = Nothing
    Set objStatics = Nothing

End Sub
Private Sub Display(ByVal objDic As clsDictionary)

'Bld_Dic.FieldInialize "seq,gorupcd", "day1,day2,day3,day4,day5,day6,day7,day8,day9,day10,day11," & _
'                                    "day12,day13,day14,day15,day16,day17,day18,day19,day20,day21," & _
'                                    "day22,day23,day24,day25,day26,day27,day28,day29,day30,day31,tot"
    Dim ii As Integer
    Dim jj As Integer
    
    objDic.MoveFirst
    With tblList
        Do Until objDic.EOF
            Select Case objDic.Fields("seq")
                Case "1"
                    For ii = 2 To 7
                        .Row = ii: .Col = 35
                        If .Value = objDic.Fields("gorupcd") Then
                            For jj = 3 To 34
                                .Row = ii: .Col = jj
                                .Value = objDic.Fields("day" & jj - 2)
                                .Value = IIf(.Value = "0", "", .Value)
                            Next
                            .Col = 34: .Value = objDic.Fields("tot"): .Value = IIf(.Value = "0", "", .Value)
                        End If
                    Next
                Case "2"
                    For ii = 8 To 13
                        .Row = ii: .Col = 35
                        If .Value = objDic.Fields("gorupcd") Then
                            For jj = 3 To 34
                                .Row = ii: .Col = jj
                                .Value = objDic.Fields("day" & jj - 2)
                                .Value = IIf(.Value = "0", "", .Value)
                            Next
                            .Col = 34: .Value = objDic.Fields("tot"): .Value = IIf(.Value = "0", "", .Value)
                        End If
                    Next
                Case "3"
                    For ii = 14 To 19
                        .Row = ii: .Col = 35
                        If .Value = objDic.Fields("gorupcd") Then
                            For jj = 3 To 34
                                .Row = ii: .Col = jj
                                .Value = objDic.Fields("day" & jj - 2)
                                .Value = IIf(.Value = "0", "", .Value)
                            Next
                            .Col = 34: .Value = objDic.Fields("tot"): .Value = IIf(.Value = "0", "", .Value)
                        End If
                    Next
            End Select
            objDic.MoveNext
        Loop
        Dim lngTot As Long
        
'        For ii = 2 To 19
'            .Row = ii
'            For jj = 3 To 34
'                .Col = jj: lngTot = lngTot + Val(.Value)
'            Next
'            .Col = 34: .Value = IIf(lngTot = 0, "", lngTot)
'            lngTot = 0
'        Next
        lngTot = 0
        For ii = 3 To 34
            .Col = ii
            For jj = 2 To 19
                .Row = jj: lngTot = lngTot + Val(.Value)
            Next
            .Row = 20: .Value = IIf(lngTot = 0, "", lngTot)
            lngTot = 0
        Next
                
    End With

End Sub

Private Sub optABO_Click(Index As Integer)
    lblabo.Caption = optABO(Index).Caption
    If optRh(0).Value Then
        lblabo.Caption = lblabo.Caption & optRh(0).Caption
    Else
        lblabo.Caption = lblabo.Caption & optRh(1).Caption
    End If
    
    If chkALL(0).Value = 1 Then chkALL(0).Value = 0
    
    Call cmdClear_Click
End Sub

Private Sub optRh_Click(Index As Integer)
    If optABO(0).Value Then
        lblabo.Caption = optABO(0).Caption & optRh(Index).Caption
    ElseIf optABO(1).Value Then
        lblabo.Caption = optABO(1).Caption & optRh(Index).Caption
    ElseIf optABO(2).Value Then
        lblabo.Caption = optABO(2).Caption & optRh(Index).Caption
    ElseIf optABO(3).Value Then
        lblabo.Caption = optABO(3).Caption & optRh(Index).Caption
    End If
    
    If chkALL(1).Value = 1 Then chkALL(1).Value = 0
    
    Call cmdClear_Click
End Sub


'혈액마감 없이 Direct조회시 사용하기 위해서 새로 작성

Private Sub DirectQuery()
    Dim objStatics As New clsStatics
    Dim FMonth     As String
    Dim TMonth     As String
    Dim Centercd   As String
    Dim Centernm   As String
    Dim Volume     As String
    Dim ABO        As String
    Dim Rh         As String
    
    If cboCenter.Text = "(ALL)" Then
        Centercd = "": Centernm = ""
    Else
        Centercd = medGetP(cboCenter.Text, 1, " ")
        Centernm = medGetP(cboCenter.Text, 2, " ")
    End If
    
    Select Case cboVolume.ListIndex
        Case 0: Volume = ""
        Case 1: Volume = "320"
        Case 2: Volume = "400"
        Case 3: Volume = "Etc"
    End Select
    
    Select Case cboDiv.ListIndex
        Case 0: mode = 1
        Case 1: mode = 2
        Case 2: mode = 3
        Case 3: mode = 4
    End Select
    
    If chkALL(0).Value = 0 Then
        If optABO(0).Value = True Then ABO = "A"
        If optABO(1).Value = True Then ABO = "B"
        If optABO(2).Value = True Then ABO = "O"
        If optABO(3).Value = True Then ABO = "AB"
    Else
        ABO = ""
    End If
    
    If chkALL(1).Value = 0 Then
        If optRh(0).Value = True Then Rh = "+"
        If optRh(1).Value = True Then Rh = "-"
    Else
        Rh = ""
    End If
    
    FMonth = Format(dtpMonth.Value, "yyyymm") & "01"
    TMonth = Format(dtpMonth.Value, "yyyymm") & "31"
    
    Select Case mode
        Case "1": Call DirectStorage(FMonth, TMonth, Centercd, ABO, Rh, Volume)
        Case "2": Call DirectDelivery(FMonth, TMonth, Centercd, ABO, Rh, Volume)
        Case "3": Call DirectReturn(FMonth, TMonth, Centercd, ABO, Rh, Volume)
        Case "4": Call DirectExpire(FMonth, TMonth, Centercd, ABO, Rh, Volume)
    End Select

    
'    If objStatics.bloodcnt(FMonth, TMonth, mode) = True Then
'        Call Query(FMonth, TMonth, Centercd, Centernm, ABO, Rh, Volume, mode)
'    Else
'        MsgBox "해당자료가 없습니다", vbInformation + vbOKOnly, "혈액현황출력"
'    End If
'    Set objStatics = Nothing
End Sub

'입고
Private Sub DirectStorage(ByVal FMonth As String, ByVal TMonth As String, ByVal Centercd As String, _
                          ByVal ABO As String, ByVal Rh As String, ByVal Volume As String)
    Dim RS   As Recordset
    Dim SSQL As String
    
    
    SSQL = " SELECT a.entdt as dt ,b.groupcd,'M1' as div,count(*) as cnt FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & DBW("a.entdt>=", FMonth) & " AND " & DBW("a.entdt<=", TMonth) & " AND " & DBW("a.hosfg<>", "1") & _
           " AND (a.localcd='' or a.localcd is null) AND a.compocd=b.compocd"
         
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY a.entdt,a.compocd,b.groupcd"
    
    SSQL = SSQL & " UNION ALL SELECT a.entdt as dt,b.groupcd,'M2' as div,count(*) as cnt FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & DBW("a.entdt>=", FMonth) & " AND " & DBW("a.entdt<=", TMonth) & " AND " & DBW("a.hosfg=", "1") & _
           " AND a.compocd=b.compocd"
         
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY a.entdt,a.compocd,b.groupcd"
    
    SSQL = SSQL & " UNION ALL SELECT a.entdt as dt,b.groupcd,'M3' as div,count(*) as cnt FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & DBW("a.entdt>=", FMonth) & " AND " & DBW("a.entdt<=", TMonth) & " AND " & DBW("a.hosfg<>", "1") & _
           " AND a.localcd is not null AND a.compocd=b.compocd"
         
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY a.entdt,a.compocd,b.groupcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Call DirectDisplay(RS)
    Else
        MsgBox "조회할 자료가 없습니다.", vbExclamation
    End If
    Set RS = Nothing
    
End Sub
'출고
Private Sub DirectDelivery(ByVal FMonth As String, ByVal TMonth As String, ByVal Centercd As String, _
                          ByVal ABO As String, ByVal Rh As String, ByVal Volume As String)

    Dim RS   As Recordset
    Dim SSQL As String
    
    
    SSQL = " SELECT b.deliverydt as dt,c.groupcd,count(*) as cnt, 'M1' as div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b " & _
           " WHERE  " & DBW("b.deliverydt>=", FMonth) & _
           " AND " & DBW("b.deliverydt <=", TMonth) & _
           " AND b.bldsrc=a.bldsrc AND b.bldyy=a.bldyy AND b.bldno=a.bldno AND b.compocd=a.compocd" & _
           " AND a.hosfg<>'1' AND (a.localcd='' or a.localcd is null)" & _
           " AND b.compocd=c.compocd"
    
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY b.deliverydt,c.groupcd"
    
    SSQL = SSQL & " UNION ALL  SELECT b.deliverydt as dt,c.groupcd,count(*) as cnt, 'M2' as div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b " & _
           " WHERE  " & DBW("b.deliverydt>=", FMonth) & _
           " AND " & DBW("b.deliverydt <=", TMonth) & _
           " AND b.bldsrc=a.bldsrc AND b.bldyy=a.bldyy AND b.bldno=a.bldno AND b.compocd=a.compocd" & _
           " AND a.hosfg='1'" & _
           " AND b.compocd=c.compocd"
    
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY b.deliverydt,c.groupcd"
    
    SSQL = SSQL & " UNION ALL  SELECT b.deliverydt as dt,c.groupcd,count(*) as cnt, 'M3' as div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b " & _
           " WHERE  " & DBW("b.deliverydt>=", FMonth) & _
           " AND " & DBW("b.deliverydt <=", TMonth) & _
           " AND b.bldsrc=a.bldsrc AND b.bldyy=a.bldyy AND b.bldno=a.bldno AND b.compocd=a.compocd" & _
           " AND a.hosfg<>'1' AND a.localcd is not null" & _
           " AND b.compocd=c.compocd"
    
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY b.deliverydt,c.groupcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Call DirectDisplay(RS)
    Else
        MsgBox "조회할 자료가 없습니다.", vbExclamation
    End If
    Set RS = Nothing
End Sub

'반환
Private Sub DirectReturn(ByVal FMonth As String, ByVal TMonth As String, ByVal Centercd As String, _
                          ByVal ABO As String, ByVal Rh As String, ByVal Volume As String)
    Dim RS   As Recordset
    Dim SSQL As String
    
    
    SSQL = " SELECT b.retdt as dt,c.groupcd,count(*) as cnt, 'M1' as div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b " & _
           " WHERE  " & DBW("b.retdt>=", FMonth) & _
           " AND " & DBW("b.retdt <=", TMonth) & _
           " AND b.bldsrc=a.bldsrc AND b.bldyy=a.bldyy AND b.bldno=a.bldno AND b.compocd=a.compocd" & _
           " AND a.hosfg<>'1' AND (a.localcd='' or a.localcd is null)" & _
           " AND b.compocd=c.compocd"
    
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY b.retdt,c.groupcd"
    
    SSQL = SSQL & " UNION ALL  SELECT b.retdt as dt,c.groupcd,count(*) as cnt, 'M2' as div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b " & _
           " WHERE  " & DBW("b.retdt>=", FMonth) & _
           " AND " & DBW("b.retdt <=", TMonth) & _
           " AND b.bldsrc=a.bldsrc AND b.bldyy=a.bldyy AND b.bldno=a.bldno AND b.compocd=a.compocd" & _
           " AND a.hosfg='1'" & _
           " AND b.compocd=c.compocd"
    
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY b.retdt,c.groupcd"
    
    SSQL = SSQL & " UNION ALL  SELECT b.retdt as dt,c.groupcd,count(*) as cnt, 'M3' as div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a," & T_BBS402 & " b " & _
           " WHERE  " & DBW("b.retdt>=", FMonth) & _
           " AND " & DBW("b.retdt <=", TMonth) & _
           " AND b.bldsrc=a.bldsrc AND b.bldyy=a.bldyy AND b.bldno=a.bldno AND b.compocd=a.compocd" & _
           " AND a.hosfg<>'1' AND a.localcd is not null" & _
           " AND b.compocd=c.compocd"
    
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY b.retdt,c.groupcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Call DirectDisplay(RS)
    Else
        MsgBox "조회할 자료가 없습니다.", vbExclamation
    End If
    Set RS = Nothing
End Sub

'폐기
Private Sub DirectExpire(ByVal FMonth As String, ByVal TMonth As String, ByVal Centercd As String, _
                          ByVal ABO As String, ByVal Rh As String, ByVal Volume As String)

    Dim RS   As Recordset
    Dim SSQL As String
    
    
    SSQL = " SELECT a.realexpdt as dt,b.groupcd,'M1' as div,count(*) as cnt FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & DBW("a.realexpdt>=", FMonth) & " AND " & DBW("a.realexpdt<=", TMonth) & " AND " & DBW("a.hosfg<>", "1") & _
           " AND (a.localcd='' or a.localcd is null) AND a.compocd=b.compocd"
         
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY a.realexpdt,a.compocd,b.groupcd"
    
    SSQL = SSQL & " UNION ALL  SELECT a.realexpdt as dt,b.groupcd,'M2' as div,count(*) as cnt FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & DBW("a.realexpdt>=", FMonth) & " AND " & DBW("a.realexpdt<=", TMonth) & " AND " & DBW("a.hosfg=", "1") & _
           " AND a.compocd=b.compocd"
         
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    SSQL = SSQL & " GROUP BY a.realexpdt,a.compocd,b.groupcd"
    
    SSQL = SSQL & " UNION ALL  SELECT a.realexpdt as dt,b.groupcd,'M3' as div,count(*) as cnt FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & DBW("a.realexpdt>=", FMonth) & " AND " & DBW("a.realexpdt<=", TMonth) & " AND " & DBW("a.hosfg<>", "1") & _
           " AND a.localcd is not null AND a.compocd=b.compocd"
         
    SSQL = SSQL & QueryString(Centercd, ABO, Rh, Volume)
    
    SSQL = SSQL & " GROUP BY a.realexpdt,a.compocd,b.groupcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Call DirectDisplay(RS)
    Else
        MsgBox "조회할 자료가 없습니다.", vbExclamation
    End If
    Set RS = Nothing
End Sub
Private Function QueryString(ByVal Centercd As String, ByVal ABO As String, ByVal Rh As String, ByVal Volume As String) As String
    If Centercd <> "" Then QueryString = QueryString & " AND " & DBW("a.centercd=", Centercd)
    If ABO <> "" Then QueryString = QueryString & " AND " & DBW("a.abo=", ABO)
    If Rh <> "" Then QueryString = QueryString & " AND " & DBW("a.rh=", Rh)
    If Volume <> "" And Volume <> "Etc" Then QueryString = QueryString & " AND " & DBW("a.volumn=", Volume)
    If Volume = "Etc" Then QueryString = QueryString & " AND (" & DBW("a.volumn<>", Volume) & " AND " & DBW("a.volumn<>", Volume) & ")"
End Function

Private Sub DirectDisplay(ByVal RS As Recordset)
    Dim ii As Integer
    Dim jj As Integer
    Dim lngTot As Long
'M1:혈액원, M2:헌혈, M3:외부
    
    With tblList
        .ReDraw = False
        Do Until RS.EOF
            Select Case RS.Fields("div").Value & ""
                Case "M1" 'lngCompCnt
'                    For ii = 2 To 7
                    For ii = 2 To lngCompCnt + 1
                        .Row = ii: .Col = 35
                        If .Value = RS.Fields("groupcd").Value & "" Then
                            .Col = Val(Mid(RS.Fields("dt").Value & "", 7)) + 2
                            .Value = RS.Fields("cnt").Value & ""
                        End If
                    Next
                Case "M2"
'                    For ii = 8 To 13
                    For ii = lngCompCnt + 2 To lngCompCnt * 2 + 1
                        .Row = ii: .Col = 35
                        If .Value = RS.Fields("groupcd").Value & "" Then
                            .Col = Val(Mid(RS.Fields("dt").Value & "", 7)) + 2
                            .Value = RS.Fields("cnt").Value & ""
                        End If
                    Next
                Case "M3"
'                    For ii = 14 To 19
                    For ii = lngCompCnt * 2 + 2 To lngCompCnt * 3 + 1
                        .Row = ii: .Col = 35
                        If .Value = RS.Fields("groupcd").Value & "" Then
                            .Col = Val(Mid(RS.Fields("dt").Value & "", 7)) + 2
                            .Value = RS.Fields("cnt").Value & ""
                        End If
                    Next
            End Select
            RS.MoveNext
        Loop
        
        '가로합(제재별계)
'        For ii = 2 To 19
'            lngTot = 0
'            .Row = ii
'            For jj = 3 To 33
'                .Col = jj
'                lngTot = lngTot + Val(.Value)
'            Next
'            .Col = 34: .Value = IIf(lngTot = 0, "", lngTot)
'        Next
        For ii = 2 To .MaxRows - 1
            lngTot = 0
            .Row = ii
            For jj = 3 To 33
                .Col = jj
                lngTot = lngTot + Val(.Value)
            Next
            .Col = 34: .Value = IIf(lngTot = 0, "", lngTot)
        Next
        
        '세로합(일별계)
'        For ii = 3 To 34
'            lngTot = 0
'            .Col = ii
'            For jj = 2 To 19
'                .Row = jj
'                lngTot = lngTot + Val(.Value)
'            Next
'            .Row = 20: .Value = IIf(lngTot = 0, "", lngTot)
'        Next
        For ii = 3 To 34
            lngTot = 0
            .Col = ii
            For jj = 2 To .MaxRows - 1
                .Row = jj
                lngTot = lngTot + Val(.Value)
            Next
            .Row = .MaxRows: .Value = IIf(lngTot = 0, "", lngTot)
        Next
        
        .ReDraw = True
    End With
    
End Sub
