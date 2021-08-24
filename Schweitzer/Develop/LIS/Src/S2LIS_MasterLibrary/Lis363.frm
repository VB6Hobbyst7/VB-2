VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm363MicWsKind 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGcAcAsFcFs 
      BackColor       =   &H00DBE6E6&
      Height          =   7515
      Left            =   195
      TabIndex        =   1
      Top             =   1065
      Width           =   10470
      Begin VB.ListBox lstTest 
         Appearance      =   0  '평면
         BackColor       =   &H00F7FFF7&
         Height          =   3990
         Left            =   240
         TabIndex        =   5
         Top             =   585
         Width           =   3045
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   8850
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   6900
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   7515
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   6900
         Width           =   1335
      End
      Begin FPSpread.vaSpread ssBRTmp 
         Height          =   6255
         Left            =   3420
         TabIndex        =   10
         Top             =   585
         Width           =   3405
         _Version        =   196608
         _ExtentX        =   6006
         _ExtentY        =   11033
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   -2147483633
         MaxCols         =   2
         MaxRows         =   0
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis363.frx":0000
         UserResize      =   0
      End
      Begin VB.ListBox lstRstType 
         Appearance      =   0  '평면
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   7
         Top             =   6615
         Width           =   3045
      End
      Begin VB.ListBox lstWS 
         Appearance      =   0  '평면
         BackColor       =   &H00F7FFF7&
         Height          =   1290
         Left            =   240
         TabIndex        =   6
         Top             =   4995
         Width           =   3045
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   4680
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "적용 WorkSheet"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검사항목"
         Appearance      =   0
      End
      Begin FPSpread.vaSpread ssGRTmp 
         Height          =   6255
         Left            =   6900
         TabIndex        =   11
         Top             =   585
         Width           =   3405
         _Version        =   196608
         _ExtentX        =   6006
         _ExtentY        =   11033
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   0
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis363.frx":027A
         UserResize      =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   6300
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "결과Type"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   300
         Left            =   3420
         TabIndex        =   25
         Top             =   270
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "미생물 배치 결과코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Left            =   6900
         TabIndex        =   26
         Top             =   270
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검사항목별 결과코드"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.TabStrip tabGroup 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   690
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   582
      Style           =   1
      TabMinWidth     =   88
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "일반 Culture"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gram Stain"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AFB Culture"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AFB Stain"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fungus Culture"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fungus Stain"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraGs 
      BackColor       =   &H00DBE6E6&
      Height          =   7515
      Left            =   195
      TabIndex        =   14
      Top             =   1065
      Width           =   10470
      Begin VB.ListBox lstGSTest 
         BackColor       =   &H00F7FFF7&
         Height          =   960
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3045
      End
      Begin VB.ListBox lstGSDetailTest 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Height          =   3090
         Left            =   240
         TabIndex        =   19
         Top             =   1575
         Width           =   3045
      End
      Begin VB.ListBox lstGSWS 
         Appearance      =   0  '평면
         BackColor       =   &H00F7FFFF&
         Height          =   1290
         Left            =   240
         TabIndex        =   18
         Top             =   4995
         Width           =   3045
      End
      Begin VB.ListBox lstGSRstType 
         Appearance      =   0  '평면
         BackColor       =   &H00F7FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   17
         Top             =   6615
         Width           =   3045
      End
      Begin VB.CommandButton cmdGsSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   7530
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   6915
         Width           =   1320
      End
      Begin VB.CommandButton cmdGsExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   8850
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   6915
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   4680
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "적용 WorkSheet"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   270
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검사항목"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   6300
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "결과Type"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   300
         Left            =   3450
         TabIndex        =   8
         Top             =   270
         Width           =   6780
         _ExtentX        =   11959
         _ExtentY        =   529
         BackColor       =   16703181
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    General Result Template"
         Appearance      =   0
      End
      Begin FPSpread.vaSpread ssGsGRtmp 
         Height          =   6255
         Left            =   3450
         TabIndex        =   9
         Top             =   600
         Width           =   6780
         _Version        =   196608
         _ExtentX        =   11959
         _ExtentY        =   11033
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   0
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis363.frx":04E6
         UserResize      =   0
      End
   End
   Begin VB.Label lblRName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "미생물 배치 결과등록"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00613636&
      Height          =   315
      Left            =   450
      TabIndex        =   23
      Top             =   210
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0085A3A3&
      BorderWidth     =   2
      Height          =   390
      Left            =   210
      Shape           =   4  '둥근 사각형
      Top             =   660
      Width           =   10440
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00D191A2&
      BorderWidth     =   3
      FillColor       =   &H00F7F0F0&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   210
      Shape           =   4  '둥근 사각형
      Top             =   105
      Width           =   4845
   End
End
Attribute VB_Name = "frm363MicWsKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1

Private iBRTmpCurRow As Integer
Private iGRTmpCurRow As Integer
Private iGsGrTmpCurRow  As Integer

Private objSpcDic As New clsDictionary
Private objSql As New clsLISSqlStatement
Private objMicSql As New clsLISSqlMasters


Private Sub SetSGroup()

    Dim sRstTp1 As String, sRstTp2 As String
    Dim tmpRstTp As Variant
    Dim i As Integer, j As Long
    Dim objRs As Recordset
    Dim sSGCnt As Integer, sWsName As String, sWsGrp As String, sWorkArea As String
    
    MouseRunning
    
    objSpcDic.Clear
    objSpcDic.FieldInialize "wsgrp", "grpnm,workarea,rsttp1,rsttp2"
    objSpcDic.Sort = False
    
    tabGroup.Tabs.Clear
    
    'Prt As Integer          ' 검체군별 프린터 여부 지정시 사용 (현재 기능 막았슴)
    Set objRs = New Recordset
    objRs.Open objMicSql.SqlGetWsGroup, DBConn
       
    If objRs.EOF Then
        MsgBox "등록되어 있는 검체군이 없습니다.", vbExclamation, "검체군"
        Set objRs = Nothing
        Exit Sub
    End If

    sSGCnt = objRs.RecordCount
    
    objRs.MoveFirst
    
    For i = 1 To sSGCnt
    
        sWsGrp = "" & objRs.Fields("wscd").Value
        sWsName = "" & objRs.Fields("wsnm").Value
        sWorkArea = medGetP("" & objRs.Fields("wa").Value, 1, ";")
        sRstTp1 = "" & objRs.Fields("rsttp").Value
        
        tmpRstTp = Split(sRstTp1, ",")
        For j = LBound(tmpRstTp) To UBound(tmpRstTp)
            tmpRstTp(j) = "'" & tmpRstTp(j) & "'"
        Next
        sRstTp2 = Join(tmpRstTp, ",")
        
        If Not objSpcDic.Exists(sWsGrp) Then
            objSpcDic.AddNew sWsGrp, medConcatString(COL_DIV, sWsName, sWorkArea, sRstTp1, sRstTp2)
            tabGroup.Tabs.Add , sWsGrp, sWsName
        End If
        objRs.MoveNext
        
    Next i
    
    Set objRs = Nothing
    
    MouseDefault
    
End Sub

Private Function TmpDataChk() As Boolean
    Dim i%, j%
    Dim StandTmpCd As String, CompareTmpCd As String, StandTmpNm As String
    
    With ssBRTmp
        For i = 1 To .MaxRows
            .Col = 1: .Row = i: StandTmpCd = .Text
            .Col = 2: .Row = i: StandTmpNm = .Text
            If Len(Trim(StandTmpCd)) < 1 And Len(Trim(StandTmpNm)) > 1 Then
                TmpDataChk = False
                Exit Function
            End If
            For j = 1 To .MaxRows
                .Col = 1: .Row = j: CompareTmpCd = .Text
                If StandTmpCd = CompareTmpCd And i <> j Then
                    TmpDataChk = False
                    Exit Function
                End If
            Next j
        Next i
    End With
        
    With ssGRTmp
        For i = 1 To .MaxRows
            .Col = 1: .Row = i: StandTmpCd = .Text
            .Col = 2: .Row = i: StandTmpNm = .Text
            If Len(Trim(StandTmpCd)) < 1 And Len(Trim(StandTmpNm)) > 1 Then
                TmpDataChk = False
                Exit Function
            End If
            
            For j = 1 To .MaxRows
                .Col = 1: .Row = j: CompareTmpCd = .Text
                If StandTmpCd = CompareTmpCd And i <> j Then
                    TmpDataChk = False
                    Exit Function
                End If
            Next j
        Next i
    End With
    
    TmpDataChk = True
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGsExit_Click()
    Unload Me
End Sub

Private Function GsGrTmpDataChk() As Boolean
    Dim StandTmpCd As String, CompareTmpCd As String, StandTmpNm As String
    Dim i%, j%
    
    With ssGsGRtmp
        For i = 1 To .MaxRows
            .Col = 1: .Row = i: StandTmpCd = .Text
            .Col = 2: .Row = i: StandTmpNm = .Text
            If Len(Trim(StandTmpCd)) < 1 And Len(Trim(StandTmpNm)) > 1 Then
                GsGrTmpDataChk = False
                Exit Function
            End If
            
            For j = 1 To .MaxRows
                .Col = 1: .Row = j: CompareTmpCd = .Text
                If StandTmpCd = CompareTmpCd And i <> j Then
                    GsGrTmpDataChk = False
                    Exit Function
                End If
            Next j
        Next i
    End With
    GsGrTmpDataChk = True
End Function

Private Sub cmdGsSave_Click()
    
    Dim sSqlDelC110 As String
    Dim sSqlInsC110 As String
    Dim sTestCd As String
    Dim sGRstCd As String, sGRstNm As String
    Dim i%, j%
            
    If GsGrTmpDataChk = False Then
        MsgBox " template 코드의 입력이 올바르지 않습니다."
        Exit Sub
    End If
    
On Error GoTo DBExecError
    DBConn.BeginTrans
    
    ' Delete C110 ------------------------------------------------------------
    sTestCd = Trim(Mid(lstGSDetailTest.List(lstGSDetailTest.ListIndex), 1, _
                  InStr(1, lstGSDetailTest.List(lstGSDetailTest.ListIndex), vbTab) - 1))
    sSqlDelC110 = objSql.SqlDeleteLAB031(LC2_ItemResult, sTestCd)
                  
    DBConn.Execute (sSqlDelC110)

    ' Insert C110 ------------------------------------------------------------
    For j = 1 To ssGsGRtmp.MaxRows
        ssGsGRtmp.Col = 1: ssGsGRtmp.Row = j: sGRstCd = ssGsGRtmp.Text
        ssGsGRtmp.Col = 2: ssGsGRtmp.Row = j: sGRstNm = ssGsGRtmp.Text
        
        If Len(Trim(sGRstCd)) < 1 Or Len(Trim(sGRstNm)) < 1 Then Exit For
        
        sSqlInsC110 = objSql.SqlSaveLAB031(LC2_ItemResult, sTestCd, sGRstCd, sGRstNm, "", "", "", "", "", "", 1)
                             
        DBConn.Execute (sSqlInsC110)
    Next j
        
    DBConn.CommitTrans
    ClearssGSGRTmp
    Exit Sub
    
DBExecError:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdSave_Click()
    Dim sSqlDelC110 As String
    Dim sSqlInsC110 As String
    Dim sSqlDelC113 As String
    Dim sSqlInsC113 As String
    Dim sTestCd As String
    Dim sBRstCd As String, sBRstNm As String
    Dim sGRstCd As String, sGRstNm As String
    Dim sGroupCd As String
    Dim i%, j%, K%
            
    If TmpDataChk = False Then
'        MsgBox " template 코드의 입력이 올바르지 않습니다."
        MsgBox " 결과 코드의 입력이  중복되었거나 올바르지 않습니다."
        Exit Sub
    End If
    
On Error GoTo DBExecError
    DBConn.BeginTrans
    
    For i = 0 To lstTest.ListCount - 1
        sTestCd = Trim(Mid(lstTest.List(i), 1, _
                  InStr(1, lstTest.List(i), vbTab) - 1))
        sSqlDelC110 = objSql.SqlDeleteLAB031(LC2_ItemResult, sTestCd)
                      
        DBConn.Execute (sSqlDelC110)
    Next i
    
    sGroupCd = tabGroup.SelectedItem.Key
    sSqlDelC113 = objSql.SqlDeleteLAB031(LC2_MBatchRst, sGroupCd)
    DBConn.Execute (sSqlDelC113)
    
    '## 미생물 배치결과코드 입력
    For i = 1 To ssBRTmp.DataRowCnt
        ssBRTmp.Col = 1: ssBRTmp.Row = i: sBRstCd = ssBRTmp.Text
        ssBRTmp.Col = 2: ssBRTmp.Row = i: sBRstNm = ssBRTmp.Text
        If Len(Trim(sBRstCd)) < 1 Or Len(Trim(sBRstNm)) < 1 Then Exit For
        sSqlInsC113 = objSql.SqlSaveLAB031(LC2_MBatchRst, sGroupCd, sBRstCd, sBRstNm, "", "", "", "", "", "", 1)
                          
        DBConn.Execute (sSqlInsC113)
    Next i
            
    For i = 0 To lstTest.ListCount - 1
        sTestCd = Trim(Mid(lstTest.List(i), 1, _
                  InStr(1, lstTest.List(i), vbTab) - 1))
        
        '## 미생물 배치결과코드 입력
        For j = 1 To ssBRTmp.MaxRows
            ssBRTmp.Col = 1: ssBRTmp.Row = j: sBRstCd = ssBRTmp.Text
            ssBRTmp.Col = 2: ssBRTmp.Row = j: sBRstNm = ssBRTmp.Text
            If Len(Trim(sBRstCd)) < 1 Or Len(Trim(sBRstNm)) < 1 Then Exit For
            sSqlInsC110 = objSql.SqlSaveLAB031(LC2_ItemResult, sTestCd, sBRstCd, sBRstNm, "", "B", "", "", "", "", 1)

            DBConn.Execute (sSqlInsC110)
        Next j
        
        '## 검사항목별 결과코드 입력
        For j = 1 To ssGRTmp.MaxRows
            ssGRTmp.Col = 1: ssGRTmp.Row = j: sGRstCd = ssGRTmp.Text
            ssGRTmp.Col = 2: ssGRTmp.Row = j: sGRstNm = ssGRTmp.Text
            For K = 1 To ssBRTmp.DataRowCnt
                ssBRTmp.Row = K
                ssBRTmp.Col = 1
                If ssBRTmp.Value = sGRstCd Then GoTo Skip
            Next
            If Len(Trim(sGRstCd)) < 1 Or Len(Trim(sGRstNm)) < 1 Then Exit For
            sSqlInsC110 = objSql.SqlSaveLAB031(LC2_ItemResult, sTestCd, sGRstCd, sGRstNm, "", "", "", "", "", "", 1)
                          
            DBConn.Execute (sSqlInsC110)
Skip:
        Next j
    Next i
    
    DBConn.CommitTrans
    
    tabGroup.DeselectAll
    ClearfraGcAcAsFcFsAll
    Exit Sub
    
DBExecError:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub ClearfraGcAcAsFcFsAll()
    lstTest.Clear
    lstWS.Clear
    lstRstType.Clear
    ClearssBRTmp
    ClearssGRTmp
End Sub

Private Sub Form_Load()
    objSpcDic.Clear
    objSpcDic.FieldInialize "grpcd", "grpnm,media,workarea,fseq,tseq,rptseq,wsgrp,excfg," & _
                                     "wsunit,startdt,starttm,fnshdt,fnshtm,count,worksheet,extable,excount"
            
    Call SetSGroup
    
    iBRTmpCurRow = -1
    iGRTmpCurRow = -1
    iGsGrTmpCurRow = -1
    
    tabGroup.Tabs(1).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing

End Sub

Private Sub lstGSDetailTest_Click()
    Dim RS          As ADODB.Recordset
    Dim strTestCd   As String   '검사코드
    Dim SQL         As String
    Dim i           As Long
    
    '## 수정:이상대(2004-11-12)
    '##     1.에러처리 추가
    '##     2.디자인 변경에 따른 Spread 표현방법 변경, 소스정리..
    strTestCd = Trim(Mid(lstGSDetailTest.List(lstGSDetailTest.ListIndex), 1, _
                InStr(1, lstGSDetailTest.List(lstGSDetailTest.ListIndex), vbTab) - 1))
    SQL = objSql.SqlLAB031CodeList(LC2_ItemResult, "cdval2, field1", strTestCd)
                     
On Error GoTo Errors
    Set RS = New ADODB.Recordset
    RS.Open SQL, DBConn
        
    Call ClearssGSGRTmp
    If Not (RS.BOF Or RS.EOF) Then
        With ssGsGRtmp
            For i = 1 To RS.RecordCount
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = 1: .Text = "" & RS.Fields("cdval2").Value & ""
                .Col = 2: .Text = "" & RS.Fields("field1").Value & ""
                
                RS.MoveNext
            Next i
            .RowHeight(-1) = 11
            .Col = 1: .Col2 = 2
            .Row = 1: .Row2 = .DataRowCnt
            .BlockMode = True
            .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
        End With
    End If
    ssGsGRtmp.MaxRows = ssGsGRtmp.MaxRows + 1
    RS.Close
    Set RS = Nothing
    Exit Sub
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub

Private Sub ClearssGSGRTmp()
    With ssGsGRtmp
        .Col = -1
        .Row = -1
        .Action = ActionClearText
        .MaxRows = 0
    End With
End Sub

Private Sub lstGSTest_Click()
    
    Dim sSqlGetDetailTest As String
    Dim rsGetDetailTest As Recordset
    Dim i%
    Dim sTestNm As String
    Dim objWsSql As New clsLISSqlMasters
        
    ClearssGSGRTmp
    lstGSDetailTest.Clear
    
    sTestNm = Trim(Mid(lstGSTest.List(lstGSTest.ListIndex), 1, _
              InStr(1, lstGSTest.List(lstGSTest.ListIndex), vbTab) - 1))

    If Not objSpcDic.Exists("GS") Then Exit Sub
    
    objSpcDic.KeyChange "GS"    'Gram Stain
    Set rsGetDetailTest = New Recordset
    rsGetDetailTest.Open objWsSql.SqlGetDetailTest(sTestNm, objSpcDic.Fields("workarea")), DBConn
    
    If rsGetDetailTest.EOF = True Then
        Set rsGetDetailTest = Nothing
        Set objWsSql = Nothing
        lstGSDetailTest.AddItem lstGSTest.List(lstGSTest.ListIndex)
        lstGSDetailTest.ListIndex = 0
        Call lstGSDetailTest_Click
        Exit Sub
    End If
    
    For i = 1 To rsGetDetailTest.RecordCount
        lstGSDetailTest.AddItem "" & rsGetDetailTest.Fields("testcd").Value & vbTab & _
                                "" & rsGetDetailTest.Fields("testnm").Value
        rsGetDetailTest.MoveNext
    Next i
    
    Set rsGetDetailTest = Nothing
    Set objWsSql = Nothing
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            Dim sTmpCd As String, sTmpNm As String
            
            
            If iBRTmpCurRow <> -1 Then       ' ssBrTmp 선택시
                With ssBRTmp
                .Row = iBRTmpCurRow
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                End With
                iBRTmpCurRow = -1
            ElseIf iGRTmpCurRow <> -1 Then     'ssGRtmp 선택시
                With ssGRTmp
                .Row = iGRTmpCurRow
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                End With
                iGRTmpCurRow = -1
            ElseIf iGsGrTmpCurRow <> -1 Then
                With ssGsGRtmp
                    .Row = iGsGrTmpCurRow
                    .Action = ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                End With
                iGsGrTmpCurRow = -1
            End If
    End Select
End Sub

Private Sub ssBRTmp_Advance(ByVal AdvanceNext As Boolean)
    Dim sTmpCd As String, sTmpNm As String
    Dim iMaxRow As Integer
    
    With ssBRTmp
        .Col = 1: .Row = .MaxRows: sTmpCd = Trim(.Text)
        .Col = 2: .Row = .MaxRows: sTmpNm = Trim(.Text)
    End With
    
    If sTmpCd = "" Or sTmpNm = "" Then Exit Sub
    
    If DuplicateBRTmp(ssBRTmp.ActiveRow) = True Then
        ssBRTmp.Row = ssBRTmp.ActiveRow
        ssBRTmp.Col = -1
        ssBRTmp.BackColor = &H490FF9
        Exit Sub
    End If
    
    If AdvanceNext Then
        iMaxRow = ssBRTmp.MaxRows
        ssBRTmp.MaxRows = iMaxRow + 1
        ssBRTmp.RowHeight(iMaxRow + 1) = 13
        ssBRTmp.Col = 1: ssBRTmp.Row = iMaxRow + 1
        ssBRTmp.Action = ActionActiveCell
    End If
    
End Sub

Private Function DuplicateBRTmp(iCurRow As Integer) As Boolean
    
    Dim i%
    Dim sCmpBRcd As String, sCurBRcd As String
    With ssBRTmp
        For i = 1 To .MaxRows
            .Row = i: .Col = 1: sCmpBRcd = Trim(.Text)
            .Row = iCurRow: .Col = 1: sCurBRcd = Trim(.Text)
            If (sCmpBRcd = sCurBRcd) And (i <> iCurRow) Then
                DuplicateBRTmp = True
                Exit Function
            End If
        Next i
    End With
    DuplicateBRTmp = False
    
End Function

Private Function DuplicateGRTmp(iCurRow As Integer) As Boolean
    Dim i%
    Dim sCmpGRcd As String, sCurGRcd As String
    With ssGRTmp
        For i = 1 To .MaxRows
            .Row = i: .Col = 1: sCmpGRcd = Trim(.Text)
            .Row = iCurRow: .Col = 1: sCurGRcd = Trim(.Text)
            If (sCmpGRcd = sCurGRcd) And (i <> iCurRow) Then
                DuplicateGRTmp = True
                Exit Function
            End If
        Next i
    End With
    DuplicateGRTmp = False
End Function

Private Sub ssBRTmp_Click(ByVal Col As Long, ByVal Row As Long)
    With ssBRTmp
        If .BackColor = &H490FF9 Then
            .Col = -1
            .Row = Row
            .BackColor = &HFFFFFF
        End If
    End With
End Sub

Private Sub ssBRTmp_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim sTmpCd As String, sTmpNm As String
    
    If Row <= 0 Then Exit Sub
    With ssBRTmp
        .Col = 1: .Row = Row:  sTmpCd = .Text
        .Col = 2: .Row = Row:  sTmpNm = .Text
    End With
    If Len(Trim(sTmpCd)) < 1 And Len(Trim(sTmpNm)) < 1 Then Exit Sub
    
    Set objPop = Nothing
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE ROW"
    
        iBRTmpCurRow = Row
    
        .PopupMenus Me.hWnd
    End With
    
    Set objPop = Nothing
    
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Row Delete"
   
'        iBRTmpCurRow = Row
        
'    PopupMenu mnuPopup
End Sub

Private Sub ssGRTmp_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim sTmpCd As String, sTmpNm As String
    
    If Row <= 0 Then Exit Sub
    With ssGRTmp
        .Col = 1: .Row = Row:  sTmpCd = .Text
        .Col = 2: .Row = Row:  sTmpNm = .Text
    End With
    If Len(Trim(sTmpCd)) < 1 And Len(Trim(sTmpNm)) < 1 Then Exit Sub
    
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE ROW"
        iGRTmpCurRow = Row
        
        .PopupMenus Me.hWnd
    End With
    
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Row Delete"
'
'    iGRTmpCurRow = Row
'
'    PopupMenu mnuPopup

End Sub

'Private Sub mnuDelete_Click()
'    Dim sTmpCd As String, sTmpNm As String
'
'
'    If iBRTmpCurRow <> -1 Then       ' ssBrTmp 선택시
'        With ssBRTmp
'        .Row = iBRTmpCurRow
'        .Action = ActionDeleteRow
'        .MaxRows = .MaxRows - 1
'        End With
'        iBRTmpCurRow = -1
'    ElseIf iGRTmpCurRow <> -1 Then     'ssGRtmp 선택시
'        With ssGRTmp
'        .Row = iGRTmpCurRow
'        .Action = ActionDeleteRow
'        .MaxRows = .MaxRows - 1
'        End With
'        iGRTmpCurRow = -1
'    ElseIf iGsGrTmpCurRow <> -1 Then
'        With ssGsGRtmp
'            .Row = iGsGrTmpCurRow
'            .Action = ActionDeleteRow
'            .MaxRows = .MaxRows - 1
'        End With
'        iGsGrTmpCurRow = -1
'    End If
'End Sub

Private Sub ssGRTmp_Advance(ByVal AdvanceNext As Boolean)

    Dim sTmpCd As String, sTmpNm As String
    Dim iMaxRow As Integer
    
    With ssGRTmp
        .Col = 1: .Row = .MaxRows: sTmpCd = Trim(.Text)
        .Col = 2: .Row = .MaxRows: sTmpNm = Trim(.Text)
    End With
    
    If sTmpCd = "" Or sTmpNm = "" Then Exit Sub
    
    If DuplicateGRTmp(ssGRTmp.ActiveRow) = True Then
        ssGRTmp.Row = ssGRTmp.ActiveRow
        ssGRTmp.Col = -1
        ssGRTmp.BackColor = &H490FF9
        Exit Sub
    End If
    
    If AdvanceNext Then
        iMaxRow = ssGRTmp.MaxRows
        ssGRTmp.MaxRows = iMaxRow + 1
        ssGRTmp.RowHeight(iMaxRow + 1) = 13
        ssGRTmp.Col = 1: ssGRTmp.Row = iMaxRow + 1
        ssGRTmp.Action = ActionActiveCell
    End If
End Sub

Private Sub ssGRTmp_Click(ByVal Col As Long, ByVal Row As Long)
    With ssGRTmp
        If .BackColor = &H490FF9 Then
            .Col = -1
            .Row = Row
            .BackColor = &HFFFFFF
        End If
    End With
End Sub


Private Sub ssGsGRtmp_Advance(ByVal AdvanceNext As Boolean)
    Dim sTmpCd As String, sTmpNm As String
    Dim iMaxRow As Integer
    
    With ssGsGRtmp
        .Col = 1: .Row = .MaxRows: sTmpCd = Trim(.Text)
        .Col = 2: .Row = .MaxRows: sTmpNm = Trim(.Text)
    End With
    
    If sTmpCd = "" Or sTmpNm = "" Then Exit Sub
    
    If DuplicateGsGRTmp(ssGsGRtmp.ActiveRow) = True Then
        ssGsGRtmp.Row = ssGsGRtmp.ActiveRow
        ssGsGRtmp.Col = -1
        ssGsGRtmp.BackColor = &H490FF9
        Exit Sub
    End If
    
    If AdvanceNext Then
        iMaxRow = ssGsGRtmp.MaxRows
        ssGsGRtmp.MaxRows = iMaxRow + 1
        ssGsGRtmp.RowHeight(iMaxRow + 1) = 13
        ssGsGRtmp.Col = 1: ssGsGRtmp.Row = iMaxRow + 1
        ssGsGRtmp.Action = ActionActiveCell
    End If

End Sub

Private Function DuplicateGsGRTmp(iCurRow As Integer) As Boolean
    
    Dim i%
    Dim sCmpGsGrcd As String, sCurGsGRcd As String
    With ssGsGRtmp
        For i = 1 To .MaxRows
            .Row = i: .Col = 1: sCmpGsGrcd = Trim(.Text)
            .Row = iCurRow: .Col = 1: sCurGsGRcd = Trim(.Text)
            If (sCmpGsGrcd = sCurGsGRcd) And (i <> iCurRow) Then
                DuplicateGsGRTmp = True
                Exit Function
            End If
        Next i
    End With
    DuplicateGsGRTmp = False
    
End Function

Private Sub ssGsGRtmp_Click(ByVal Col As Long, ByVal Row As Long)
    With ssGsGRtmp
        If .BackColor = &H490FF9 Then
            .Col = -1
            .Row = Row
            .BackColor = &HFFFFFF
        End If
    End With

End Sub

Private Sub ssGsGRtmp_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim sTmpCd As String, sTmpNm As String
    
    If Row <= 0 Then Exit Sub
    With ssGsGRtmp
        .Col = 1: .Row = Row:  sTmpCd = .Text
        .Col = 2: .Row = Row:  sTmpNm = .Text
    End With
    If Len(Trim(sTmpCd)) < 1 And Len(Trim(sTmpNm)) < 1 Then Exit Sub
        
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE ROW"
        iGsGrTmpCurRow = Row
        
        .PopupMenus Me.hWnd
    End With
    
    Set objPop = Nothing
        
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Row Delete"
'
'    iGsGrTmpCurRow = Row
'
'    PopupMenu mnuPopup
End Sub

Private Sub tabGroup_Click()
    objSpcDic.KeyChange Trim(tabGroup.SelectedItem.Key)
    '## Gram Stain는 검사코드별로 관리
    Select Case objSpcDic.Fields("wsgrp")
        Case "GS"
            LoadGSInitData
        Case Else
            LoadInitData (objSpcDic.Fields("wsgrp"))
    End Select
End Sub

Private Sub LoadGSInitData()
    ClearfraGSAll
    LoadFraGS
    LoadlstGSTest
    LoadGSWS
    LoadGSRstType
End Sub

Private Sub ClearfraGSAll()
    lstGSTest.Clear
    lstGSDetailTest.Clear
    lstGSWS.Clear
    lstGSRstType.Clear
    ClearssGSGRTmp
End Sub

Private Sub LoadFraGS()
    fraGcAcAsFcFs.Visible = False
    fraGs.Visible = True
    fraGs.Top = fraGcAcAsFcFs.Top
    fraGs.Left = fraGcAcAsFcFs.Left
    
End Sub

Private Sub LoadGSRstType()
    Dim sSqlGetRstType As String
    Dim rsGetRstType As Recordset
    
    sSqlGetRstType = objSql.SqlLAB032CodeList(LC3_MWSKinds, "field2", "GS")
                           
    Set rsGetRstType = New Recordset
    rsGetRstType.Open sSqlGetRstType, DBConn
    
    If rsGetRstType.EOF = True Then Exit Sub
    
    lstGSRstType.Clear
        
    lstGSRstType.AddItem "" & rsGetRstType.Fields("field2").Value
    
    lstGSRstType.Enabled = False
    
    Set rsGetRstType = Nothing

End Sub

Private Sub LoadGSWS()
    Dim sSqlGetWS As String
    Dim rsGetWS  As Recordset
    Dim i%
    
    sSqlGetWS = objSql.SqlLAB032CodeList(LC3_SGroup, "cdval1, field1", "GS")
    Set rsGetWS = New Recordset
    rsGetWS.Open sSqlGetWS, DBConn
    
    If rsGetWS.EOF = True Then Exit Sub
    
    lstGSWS.Clear
    
    For i = 1 To rsGetWS.RecordCount
        lstGSWS.AddItem "" & rsGetWS.Fields("cdval1").Value & vbTab & _
                        "" & rsGetWS.Fields("field1").Value
        rsGetWS.MoveNext
    Next i
    
    lstGSWS.Enabled = False
    Set rsGetWS = Nothing
End Sub

Private Sub LoadlstGSTest()
    Dim ssqlGetGSTest As String
    Dim rsGetGSTest As Recordset
    Dim objWsSql As New clsLISSqlMasters
    Dim i%
    
    If Not objSpcDic.Exists("GS") Then Exit Sub
    
    objSpcDic.KeyChange ("GS")
    ssqlGetGSTest = objWsSql.SqlGetGSTest(objSpcDic.Fields("workarea"))
                    
    Set rsGetGSTest = New Recordset
    rsGetGSTest.Open ssqlGetGSTest, DBConn
    
    If rsGetGSTest.EOF = True Then
        Set rsGetGSTest = Nothing
        Exit Sub
    End If
    
    lstGSTest.Clear
    
    For i = 1 To rsGetGSTest.RecordCount
        lstGSTest.AddItem "" & rsGetGSTest.Fields("testcd").Value & vbTab & _
                          "" & rsGetGSTest.Fields("testnm").Value
        rsGetGSTest.MoveNext
    Next i
    
    Set rsGetGSTest = Nothing
    Set objWsSql = Nothing
    
End Sub

Private Sub LoadInitData(sGroupCd As String)
    LoadfraGcAcAsFcFs
    LoadLstTest (sGroupCd)
    LoadWS (sGroupCd)
    LoadRstType (sGroupCd)
End Sub

Private Sub LoadfraGcAcAsFcFs()
    fraGs.Visible = False
    fraGcAcAsFcFs.Visible = True
End Sub

Private Sub LoadGRTmpFromC110(Optional FirstRstType As String, _
                              Optional SecondRstType As String, _
                              Optional sWA As String)
    Dim RS  As ADODB.Recordset
    Dim SQL As String
    Dim i   As Long
    
    '## 수정:이상대(2004-11-12)
    '##     1.에러처리 추가
    '##     2.디자인 변경에 따른 Spread 표현방법 변경, 소스정리..
    If Len(Trim(SecondRstType)) < 1 Then
        '## General Culture가 아닌경우
        SQL = " select distinct b.cdval2 as tmpcd, " & _
              "                 b.field1 as tmpNm " & _
              " from " & T_LAB001 & " a, " & T_LAB031 & " b " & _
              " where " & _
                    DBW("a.rsttype=", FirstRstType) & _
                    " and " & DBW("a.workarea=", sWA) & _
                    " and " & DBW("a.testdiv=", "2") & _
                    " and " & DBW("b.cdindex=", LC2_ItemResult) & _
                    " and ( b.field3 = '' or b.field3 is null )" & _
                    " and a.testcd = b.cdval1 " & _
              " order by tmpcd"
                       
    Else
        '## General Culture인 경우
        SQL = " select distinct b.cdval2 as tmpcd, " & _
              "                 b.field1 as tmpNm " & _
            " from " & T_LAB001 & " a, " & T_LAB031 & " b " & _
            " where " & _
                "(" & DBW("a.rsttype=", FirstRstType) & " or " & _
                DBW("a.rsttype=", SecondRstType) & ")" & _
                " and " & DBW("a.workarea=", sWA) & _
                " and " & DBW("a.testdiv=", "2") & _
                " and " & DBW("b.cdindex=", LC2_ItemResult) & _
                " and ( b.field3 = '' or b.field3 is null ) " & _
                " and a.testcd = b.cdval1 " & _
            " order by tmpcd"
    End If
    
On Error GoTo Errors
    Set RS = New ADODB.Recordset
    RS.Open SQL, DBConn
    
    Call ClearssGRTmp
    If Not (RS.BOF Or RS.EOF) Then
        With ssGRTmp
            For i = 1 To RS.RecordCount
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = 1: .Text = RS.Fields("tmpCd").Value & ""
                .Col = 2: .Text = RS.Fields("tmpNm").Value & ""
                RS.MoveNext
            Next i
            .RowHeight(-1) = 11
            .Col = 1: .Col2 = 2
            .Row = 1: .Row2 = .DataRowCnt
            .BlockMode = True
            .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
        End With
    End If
    ssGRTmp.MaxRows = ssGRTmp.MaxRows + 1
    RS.Close
    Set RS = Nothing
    Exit Sub
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub
Private Sub ClearssGRTmp()
    With ssGRTmp
        .Col = -1
        .Row = -1
        .Action = ActionClearText
        .MaxRows = 0
    End With
End Sub

Private Sub LoadBRTmpFromC110(Optional FirstRstType As String, _
                              Optional SecondRstType As String, _
                              Optional sWA As String)
    Dim RS  As ADODB.Recordset
    Dim SQL As String
    Dim i   As Long
    
    '## 수정:이상대(2004-11-12)
    '##     1.에러처리 추가
    '##     2.디자인 변경에 따른 Spread 표현방법 변경, 소스정리..
    If Len(Trim(SecondRstType)) < 1 Then
        '## General Culture가 아닌경우
        SQL = " select distinct b.cdval2 as tmpcd, " & _
                       "                 b.field1 as tmpNm " & _
                       " from " & T_LAB001 & " a, " & T_LAB031 & " b " & _
                       " where " & _
                                 DBW("a.rsttype  = ", FirstRstType) & _
                       " and " & DBW("a.workarea = ", sWA) & _
                       " and " & DBW("a.testdiv  = ", enTestDiv.TST_MicTest) & _
                       " and " & DBW("b.cdindex  = ", LC2_ItemResult) & _
                       " and   b.field3 = 'B' " & _
                       " and   a.testcd = b.cdval1 " & _
                       " order by tmpcd"
                       
    Else
        '## General Culture인 경우
        SQL = " select distinct b.cdval2 as tmpcd, " & _
                       "                 b.field1 as tmpNm " & _
                       " from " & T_LAB001 & " a, " & T_LAB031 & " b " & _
                       " where (" & _
                                 DBW("a.rsttype=", FirstRstType) & " or " & _
                                 DBW("a.rsttype=", SecondRstType) & _
                              ") " & _
                              " and " & DBW("a.workarea=", sWA) & _
                              " and " & DBW("a.testdiv=", enTestDiv.TST_MicTest) & _
                              " and " & DBW("b.cdindex=", LC2_ItemResult) & _
                              " and " & DBW("b.field3=", "B") & _
                              " and a.testcd = b.cdval1 " & _
                       " order by tmpcd"
    End If
' LC2_MBatchRst   LC2_ItemResult
On Error GoTo Errors
    Set RS = New ADODB.Recordset
    RS.Open SQL, DBConn
    
    Call ClearssBRTmp
    If Not (RS.BOF Or RS.EOF) Then
        With ssBRTmp
            For i = 1 To RS.RecordCount
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = 1: .Text = "" & RS.Fields("tmpCd").Value
                .Col = 2: .Text = "" & RS.Fields("tmpNm").Value
                RS.MoveNext
            Next i
            .RowHeight(-1) = 11
            .Col = 1: .Col2 = 2
            .Row = 1: .Row2 = .DataRowCnt
            .BlockMode = True
            .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
        End With
    End If
    ssBRTmp.MaxRows = ssBRTmp.MaxRows + 1
    RS.Close
    Set RS = Nothing
    Exit Sub
    
Errors:
    Set RS = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub

Private Sub ClearssBRTmp()
    With ssBRTmp
        .Col = -1
        .Row = -1
        .Action = ActionClearText
        .MaxRows = 0
    End With
End Sub

Private Sub LoadRstType(sGroupCd As String)
    Dim sSqlGetRstType As String
    Dim rsGetRstType As Recordset
    
    sSqlGetRstType = " select field2 as rsttype " & _
                     " from " & T_LAB032 & _
                     " where " & DBW("cdindex=", LC3_MWSKinds) & _
                     " and " & DBW("cdval1=", sGroupCd)
                     
                     
    Set rsGetRstType = New Recordset
    rsGetRstType.Open sSqlGetRstType, DBConn
    
    lstRstType.Clear
        
    lstRstType.AddItem "" & rsGetRstType.Fields("rsttype").Value
    
    lstRstType.Enabled = False
    
    Set rsGetRstType = Nothing
    
End Sub

Private Sub LoadWS(sGroupCd As String)
    Dim sSqlGetWS As String
    Dim rsGetWS As Recordset
    Dim i%, j%
    
    ' 적용 worksheet(적용검체군) 로드
    sSqlGetWS = " select cdval1 as speCd, field1 as speNm " & _
                " from  " & T_LAB032 & _
                " where " & DBW("cdindex=", LC3_SGroup) & _
                " and   " & DBW("field4=", sGroupCd)
                
    Set rsGetWS = New Recordset
    rsGetWS.Open sSqlGetWS, DBConn
    
    If rsGetWS.EOF = True Then Exit Sub
    
    lstWS.Clear
    For i = 1 To rsGetWS.RecordCount
        
       lstWS.AddItem "" & rsGetWS.Fields("specd").Value & vbTab & _
                     "" & rsGetWS.Fields("speNm").Value
       rsGetWS.MoveNext
    Next i
    
    lstWS.Enabled = False

    Set rsGetWS = Nothing
End Sub

Private Sub LoadLstTest(sGroupCd As String)
    Dim sSqlGetTest As String
    Dim rsGetTest As Recordset
    Dim i%
    Dim strRstTp As String, strRstTp1 As String
    Dim strWorkArea As String
    
    lstTest.Clear
    objSpcDic.KeyChange sGroupCd
    strRstTp = objSpcDic.Fields("rsttp2")
    strRstTp1 = objSpcDic.Fields("rsttp1")
    strWorkArea = objSpcDic.Fields("workarea")
    
    sSqlGetTest = " select testcd, testnm " & _
                  " from   " & T_LAB001 & _
                  " where  rsttype in (" & strRstTp & ")" & _
                  " and " & DBW("workarea=", strWorkArea) & _
                  " and " & DBW("testdiv =", TST_MicTest)
                  
    Select Case sGroupCd
        Case "GC"
            Call LoadBRTmpFromC110("S", "C", strWorkArea)
            Call LoadGRTmpFromC110("S", "C", strWorkArea)
        Case Else
            Call LoadBRTmpFromC110(strRstTp1, , strWorkArea)
            Call LoadGRTmpFromC110(strRstTp1, , strWorkArea)
    End Select
    
    Set rsGetTest = New Recordset
    rsGetTest.Open sSqlGetTest, DBConn
    
    If rsGetTest.EOF = True Then
        Set rsGetTest = Nothing
        Exit Sub
    End If
    
    
    
    For i = 1 To rsGetTest.RecordCount
        lstTest.AddItem "" & rsGetTest.Fields("testcd").Value & vbTab & _
                        "" & rsGetTest.Fields("testnm").Value
        rsGetTest.MoveNext
    Next i
    lstTest.Enabled = False
    
    Set rsGetTest = Nothing
End Sub
