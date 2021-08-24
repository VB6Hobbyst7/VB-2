VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS307 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액재고조회"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14550
   Icon            =   "frmBBS307.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14550
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8265
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "128"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   510
      Left            =   9585
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   9930
      _ExtentX        =   17515
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
      Caption         =   "혈액재고조회"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblAbo 
      Height          =   6225
      Left            =   2280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Width           =   9930
      _Version        =   196608
      _ExtentX        =   17515
      _ExtentY        =   10980
      _StockProps     =   64
      BackColorStyle  =   1
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
      GrayAreaBackColor=   14411494
      MaxCols         =   10
      MaxRows         =   22
      OperationMode   =   1
      RowsFrozen      =   3
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS307.frx":076A
      TextTip         =   4
      ScrollBarTrack  =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   540
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   9930
      Begin VB.ComboBox cboCenter 
         Height          =   300
         Left            =   1080
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   2715
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   18
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   635
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
         Caption         =   "Center"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2001,02,12 by kjg
'혈액조회(현시점의 혈액 조회기능을 가진다.)
'센터별로 조회를 하며, 센터가 전체 선택일경우 혈액입고내역의 모든 혈액이 조회대상이다.)

Private Enum TblColumn
    TcCOMP = 1
    TcAP
    TcBP
    TcOP
    TcABP
    
    TcAM
    TcBM
    TcOM
    TcABM
    TcTOT
End Enum

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub tableClear()
    Dim ii As Integer
    Dim jj As Integer
    
    With tblAbo
        For ii = 3 To .MaxRows
            .Row = ii
            For jj = 2 To .MaxCols
                .Col = jj: .value = ""
            Next
        Next
    End With
        
End Sub

Private Function Query()
'조회(전체병동일 경우는 입고내역의 모든 혈액의 조회기능을 가진다.)
    
    If cboCenter.ListIndex < 0 Then Exit Function
    
    Dim objbld    As New clsDictionary
    Dim objGetSql As New clsGetSqlStatement
    Dim RS        As Recordset
    Dim center    As String
    Dim strTmp    As String
    Dim ii        As Integer
    
'    objGetSql.setDbConn DBConn
    Set RS = objGetSql.Get_CompoRecordSet
        
    With objbld
        .FieldInialize "compocd", "componm,ap,bp,op,abp,am,bm,om,abm,other,tot"
        .Sort = False
        Do Until RS.EOF
            If RS.Fields("compocd").value & "" <> "" Then
                .AddNew RS.Fields("compocd").value & "", Join(Array(RS.Fields("abbrnm").value & "", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"), COL_DIV)
            End If
            RS.MoveNext
        Loop
        .Sort = True
        
    End With
    Set RS = Nothing
    center = cboCenter.List(cboCenter.ListIndex)
    If center = "전체센터" Then
        center = ""
        Set RS = objGetSql.Get_BldEnter(center)
    Else
        center = medGetP(center, 1, " ")
        Set RS = objGetSql.Get_BldEnter(center)
    End If
    
    If RS.EOF Then
        MsgBox "해당조건의 자료가 없습니다.", vbInformation + vbOKOnly, "혈액조회"
        tableClear
        Set objbld = Nothing
        Set objGetSql = Nothing
        Exit Function
    Else
        With objbld
            For ii = 1 To .RecordCount
                Do Until RS.EOF
                    If .Exists(RS.Fields("compocd").value & "") Then
                        .KeyChange RS.Fields("compocd").value & ""
                        strTmp = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
                        Select Case strTmp
                            Case "A+"
                                .Fields("ap") = Val(.Fields("ap")) + Val(RS.Fields("cnt").value & "")
                            Case "B+"
                                .Fields("bp") = Val(.Fields("bp")) + Val(RS.Fields("cnt").value & "")
                            Case "O+"
                                .Fields("op") = Val(.Fields("op")) + Val(RS.Fields("cnt").value & "")
                            Case "AB+"
                                .Fields("abp") = Val(.Fields("abp")) + Val(RS.Fields("cnt").value & "")
                            Case "A-"
                                .Fields("am") = Val(.Fields("am")) + Val(RS.Fields("cnt").value & "")
                            Case "B-"
                                .Fields("bm") = Val(.Fields("bm")) + Val(RS.Fields("cnt").value & "")
                            Case "O-"
                                .Fields("om") = Val(.Fields("om")) + Val(RS.Fields("cnt").value & "")
                            Case "AB-"
                                .Fields("abm") = Val(.Fields("abm")) + Val(RS.Fields("cnt").value & "")
                            Case Else
                                .Fields("other") = Val(.Fields("other")) + Val(RS.Fields("cnt").value & "")
                        End Select
                        .Fields("tot") = Val(.Fields("tot")) + Val(RS.Fields("cnt").value & "")
                    End If
                    
                    RS.MoveNext
                Loop
            Next ii
        End With
    End If
    
    Dim objPro As New clsProgress
    
    objPro.Container = MainFrm.stsBar
    objPro.Max = 100
    
    For ii = 1 To 50
        objPro.value = ii
    Next ii
    TblDisplay objbld
    For ii = 51 To 100
        objPro.value = ii
    Next ii
    QueryAssign center            'Assign 된 혈액조회
    
    With tblAbo
        .ReDraw = False
        
    'border line 없애기
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .CellBorderColor = vbWhite
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 1 Or 2 Or 4 Or 8
        .Action = 16
        .BlockMode = False
    
    'border line 그리기
        .Row = 3: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 1 Or 2 Or 4 Or 8
        .CellBorderColor = vbBlack
        .Action = 16
        .BlockMode = False
        
    'row=1,row2=2,col=1,col2=1 component 라인긋기
        .Row = 1: .Row2 = 2
        .Col = 1: .COL2 = 1
        .BlockMode = True
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 1 Or 2 Or 4 Or 8
        .CellBorderColor = vbBlack
        .Action = 16
        .BlockMode = False
        
    'row=1,row2=1,col=2,col2=5 Rh + 라인긋기
        .Row = 1: .Row2 = 1
        .Col = 2: .COL2 = 5
        .BlockMode = True
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 16 '1 Or 2 Or 4 Or 8
        .CellBorderColor = vbBlack
        .Action = 16
        .BlockMode = False
    
    'row=1,row2=1,col=6,col2=9 Rh - 라인긋기
        .Row = 1: .Row2 = 1
        .Col = 6: .COL2 = 9
        .BlockMode = True
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 16 '1 Or 2 Or 4 Or 8
        .CellBorderColor = vbBlack
        .Action = 16
        .BlockMode = False
    
    'row=2,row2=2,col=2,col2=9 ABO 라인긋기
        .Row = 2: .Row2 = 2
        .Col = 2: .COL2 = 9
        .BlockMode = True
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 1 Or 2 Or 4 Or 8
        .CellBorderColor = vbBlack
        .Action = 16
        .BlockMode = False
    
    'row=1,row2=2,col=10,col2=10 합계 라인긋기
        .Row = 1: .Row2 = 2
        .Col = 10: .COL2 = 10
        .BlockMode = True
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 16 '1 Or 2 Or 4 Or 8
        .CellBorderColor = vbBlack
        .Action = 16
        .BlockMode = False
        
        .ReDraw = True
    End With
    
    
    Set RS = Nothing
    Set objbld = Nothing
    Set objGetSql = Nothing
    Set objPro = Nothing
End Function

Private Sub TblDisplay(ByVal objbld As clsDictionary)
    Dim ii As Integer
    Dim TOTAP As Long
    Dim TOTBP As Long
    Dim TOTOP As Long
    Dim TOTABP As Long
    Dim TOTAM As Long
    Dim TOTBM As Long
    Dim TOTOM As Long
    Dim TOTABM As Long
    Dim TOTAL As Long
    
    ii = 3
    With tblAbo
        .ReDraw = False
        
        .MaxRows = objbld.RecordCount + ii
        .RowHeight(-1) = 12
        .RowHeight(1) = 18
        .RowHeight(2) = 18
        
        objbld.MoveFirst
        Do Until objbld.EOF
            ii = ii + 1
            .Row = ii
            .Col = TblColumn.TcCOMP: .value = objbld.Fields("componm")
            
            .Col = TblColumn.TcAP: .value = IIf(objbld.Fields("ap") = 0, "", objbld.Fields("ap"))
            TOTAP = TOTAP + Val(.value)
            
            .Col = TblColumn.TcBP: .value = IIf(objbld.Fields("bp") = 0, "", objbld.Fields("bp"))
            TOTBP = TOTBP + Val(.value)
            
            .Col = TblColumn.TcOP: .value = IIf(objbld.Fields("op") = 0, "", objbld.Fields("op"))
            TOTOP = TOTOP + Val(.value)
            
            .Col = TblColumn.TcABP: .value = IIf(objbld.Fields("abp") = 0, "", objbld.Fields("abp"))
            TOTABP = TOTABP + Val(.value)
            
            .Col = TblColumn.TcAM: .value = IIf(objbld.Fields("am") = 0, "", objbld.Fields("am"))
            TOTAM = TOTAM + Val(.value)
            
            .Col = TblColumn.TcBM: .value = IIf(objbld.Fields("bm") = 0, "", objbld.Fields("bm"))
            TOTBM = TOTBM + Val(.value)
            
            .Col = TblColumn.TcOM: .value = IIf(objbld.Fields("om") = 0, "", objbld.Fields("om"))
            TOTOM = TOTOM + Val(.value)
            
            .Col = TblColumn.TcABM: .value = IIf(objbld.Fields("abm") = 0, "", objbld.Fields("abm"))
            TOTABM = TOTABM + Val(.value)
            
            .Col = TblColumn.TcTOT: .value = IIf(objbld.Fields("tot") = 0, "", objbld.Fields("tot"))
            TOTAL = TOTAL + Val(.value)
            
            objbld.MoveNext
        Loop
        
        .Row = 3
        .Col = TblColumn.TcAP: .value = IIf(TOTAP = 0, "", TOTAP)
        .Col = TblColumn.TcBP: .value = IIf(TOTBP = 0, "", TOTBP)
        .Col = TblColumn.TcOP: .value = IIf(TOTOP = 0, "", TOTOP)
        .Col = TblColumn.TcABP: .value = IIf(TOTABP = 0, "", TOTABP)
        .Col = TblColumn.TcAM: .value = IIf(TOTAM = 0, "", TOTAM)
        .Col = TblColumn.TcBM: .value = IIf(TOTBM = 0, "", TOTBM)
        .Col = TblColumn.TcOM: .value = IIf(TOTOM = 0, "", TOTOM)
        .Col = TblColumn.TcABM: .value = IIf(TOTABM = 0, "", TOTABM)
        .Col = TblColumn.TcTOT: .value = IIf(TOTAL = 0, "", TOTAL)
        
        .ReDraw = True
    End With
End Sub

Private Sub QueryAssign(ByVal CenterCd As String)
    Dim objGetSql As New clsGetSqlStatement
    Dim objAssign As New clsDictionary
    Dim RS        As Recordset
    Dim strTmp    As String
    Dim ii        As Integer
    
    With objAssign
        Set RS = objGetSql.Get_CompoRecordSet
        .FieldInialize "compocd", "ap,bp,op,abp,am,bm,om,abm,other,tot"
        .Sort = False
        Do Until RS.EOF = True
            If RS.Fields("compocd").value & "" <> "" Then
                .AddNew RS.Fields("compocd").value & "", Join(Array("0", "0", "0", "0", "0", "0", "0", "0", "0", "0"), COL_DIV)
            End If
            RS.MoveNext
        Loop
        
        .Sort = True
        Set RS = Nothing
    End With
    
    Set RS = objGetSql.Get_AssignCnt(CenterCd)
    If RS.EOF Then
    Else
        With objAssign
            For ii = 1 To .RecordCount
                Do Until RS.EOF
                    If .Exists(RS.Fields("compocd").value & "") Then
                        .KeyChange RS.Fields("compocd").value & ""
                        strTmp = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
                        Select Case strTmp
                            Case "A+"
                                .Fields("ap") = Val(.Fields("ap")) + Val(RS.Fields("cnt").value & "")
                            Case "B+"
                                .Fields("bp") = Val(.Fields("bp")) + Val(RS.Fields("cnt").value & "")
                            Case "O+"
                                .Fields("op") = Val(.Fields("op")) + Val(RS.Fields("cnt").value & "")
                            Case "AB+"
                                .Fields("abp") = Val(.Fields("abp")) + Val(RS.Fields("cnt").value & "")
                            Case "A-"
                                .Fields("am") = Val(.Fields("am")) + Val(RS.Fields("cnt").value & "")
                            Case "B-"
                                .Fields("bm") = Val(.Fields("bm")) + Val(RS.Fields("cnt").value & "")
                            Case "O-"
                                .Fields("om") = Val(.Fields("om")) + Val(RS.Fields("cnt").value & "")
                            Case "AB-"
                                .Fields("abm") = Val(.Fields("abm")) + Val(RS.Fields("cnt").value & "")
                            Case Else
                                .Fields("other") = Val(.Fields("other")) + Val(RS.Fields("cnt").value & "")
                        End Select
                        .Fields("tot") = Val(.Fields("tot")) + Val(RS.Fields("cnt").value)
                    End If
                    
                    RS.MoveNext
                Loop
           Next ii
        End With
    End If
    'Assign 된 혈액갯수를 보여준다.
    TblDisplay_Assign objAssign
    
    Set RS = Nothing
    Set objAssign = Nothing
    Set objGetSql = Nothing
End Sub

Private Sub TblDisplay_Assign(ByVal objAssign As clsDictionary)
    Dim TOTAP     As Long
    Dim TOTBP     As Long
    Dim TOTOP     As Long
    Dim TOTABP    As Long
    Dim TOTAM     As Long
    Dim TOTBM     As Long
    Dim TOTOM     As Long
    Dim TOTABM    As Long
    Dim TOTETC    As Long
    Dim TOTAL     As Long
    Dim ii As Integer
    
    
    ii = 3
    With tblAbo
        .ReDraw = False
        
        .MaxRows = objAssign.RecordCount + ii
        objAssign.MoveFirst
        Do Until objAssign.EOF
            ii = ii + 1
            .Row = ii
            
            .Col = TblColumn.TcAP: .value = .value & IIf(objAssign.Fields("ap") = 0, "", "(" & objAssign.Fields("ap") & ")")
            TOTAP = TOTAP + Val(objAssign.Fields("ap"))
            
            .Col = TblColumn.TcBP: .value = .value & IIf(objAssign.Fields("bp") = 0, "", "(" & objAssign.Fields("bp") & ")")
            TOTBP = TOTBP + Val(objAssign.Fields("bp"))
            
            .Col = TblColumn.TcOP: .value = .value & IIf(objAssign.Fields("op") = 0, "", "(" & objAssign.Fields("op") & ")")
            TOTOP = TOTOP + Val(objAssign.Fields("op"))
            
            .Col = TblColumn.TcABP: .value = .value & IIf(objAssign.Fields("abp") = 0, "", "(" & objAssign.Fields("abp") & ")")
            TOTABP = TOTABP + Val(objAssign.Fields("abp"))
            
            .Col = TblColumn.TcAM: .value = .value & IIf(objAssign.Fields("am") = 0, "", "(" & objAssign.Fields("am") & ")")
            TOTAM = TOTAM + Val(objAssign.Fields("am"))
            
            .Col = TblColumn.TcBM: .value = .value & IIf(objAssign.Fields("bm") = 0, "", "(" & objAssign.Fields("bm") & ")")
            TOTBM = TOTBM + Val(objAssign.Fields("bm"))
            
            .Col = TblColumn.TcOM: .value = .value & IIf(objAssign.Fields("om") = 0, "", "(" & objAssign.Fields("om") & ")")
            TOTOM = TOTOM + Val(objAssign.Fields("om"))
            
            .Col = TblColumn.TcABM: .value = .value & IIf(objAssign.Fields("abm") = 0, "", "(" & objAssign.Fields("abm") & ")")
            TOTABM = TOTABM + Val(objAssign.Fields("abm"))
            
            .Col = TblColumn.TcTOT: .value = .value & IIf(objAssign.Fields("tot") = 0, "", "(" & objAssign.Fields("tot") & ")")
            TOTAL = TOTAL + Val(objAssign.Fields("tot"))
            
            objAssign.MoveNext
        Loop
        
        .Row = 3
        .Col = TblColumn.TcAP: .value = .value & IIf(TOTAP = 0, "", "(" & TOTAP & ")")
        .Col = TblColumn.TcBP: .value = .value & IIf(TOTBP = 0, "", "(" & TOTBP & ")")
        .Col = TblColumn.TcOP: .value = .value & IIf(TOTOP = 0, "", "(" & TOTOP & ")")
        .Col = TblColumn.TcABP: .value = .value & IIf(TOTABP = 0, "", "(" & TOTABP & ")")
        .Col = TblColumn.TcAM: .value = .value & IIf(TOTAM = 0, "", "(" & TOTAM & ")")
        .Col = TblColumn.TcBM: .value = .value & IIf(TOTBM = 0, "", "(" & TOTBM & ")")
        .Col = TblColumn.TcOM: .value = .value & IIf(TOTOM = 0, "", "(" & TOTOM & ")")
        .Col = TblColumn.TcABM: .value = .value & IIf(TOTABM = 0, "", "(" & TOTABM & ")")
        .Col = TblColumn.TcTOT: .value = .value & IIf(TOTAL = 0, "", "(" & TOTAL & ")")
                
        .ReDraw = True
    End With
End Sub

Private Sub cmdPrint_Click()
    Dim strFont   As String
    Dim strYear   As String
    Dim strMonth  As String
    Dim strDay As String
    
    strYear = Format(GetSystemDate, "yyyy")
    strMonth = Format(Format(GetSystemDate, "mm"), "##")
    
    With tblAbo
        .PrintJobName = "혈액재고 조회"
        .PrintAbortMsg = "혈액재고 내역을 출력중입니다..."
        .PrintColor = False
        .PrintFirstPageNumber = 1
         strFont = "/fn""굴림체""/fz""12"""
        .PrintHeader = strFont & "/n/n/l/fb1 " & "【 혈액 재고 내역 】" & "/n/n/n/n"
        .PrintFooter = "/l" & String(125, Chr(6)) & "/n" & _
                       "/l" & HOSPITAL_NAME & "/c/p/fb0" & "/r" & " 출력일시 : " & GetSystemDate
        
        .PrintMarginBottom = 100
        .PrintMarginLeft = 900
        .PrintMarginRight = 100
        .PrintMarginTop = 300
        
        .PrintShadows = True
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = True
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        
'        .GridSolid = True
        .PrintType = PrintTypeAll
        .Action = ActionPrint
'        .GridSolid = False
    End With
End Sub

Private Sub cmdQuery_Click()
    Me.MousePointer = 11
    Query
    Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim objGetSql As New clsGetSqlStatement
    Dim RS        As Recordset
    
    Set RS = objGetSql.Get_CenterRecordSet
    With RS
        If Not .EOF Then
            cboCenter.AddItem "전체센터"
            .MoveFirst
            Do Until .EOF
                cboCenter.AddItem .Fields("cdval1").value & "" & " " & .Fields("field1").value & ""
                .MoveNext
            Loop
            If cboCenter.ListCount > 2 Then
                cboCenter.ListIndex = 0
            Else
                cboCenter.ListIndex = 1
            End If
        End If
    End With

    Set RS = Nothing
    Set objGetSql = Nothing
End Sub
