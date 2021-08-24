VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmChange 
   BorderStyle     =   1  '단일 고정
   Caption         =   "수작업 변경"
   ClientHeight    =   6600
   ClientLeft      =   3315
   ClientTop       =   2475
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10065
   Begin VB.PictureBox picResult 
      Height          =   6555
      Left            =   4830
      ScaleHeight     =   6495
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   0
      Width           =   5235
      Begin FPSpread.vaSpread SSR 
         Height          =   6480
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   5175
         _Version        =   196608
         _ExtentX        =   9128
         _ExtentY        =   11430
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   50
         SpreadDesigner  =   "FrmChange.frx":0000
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      _Version        =   65536
      _ExtentX        =   8440
      _ExtentY        =   2725
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   90
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "접 수 일 자"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand CmdSave 
         Height          =   915
         Left            =   90
         TabIndex        =   3
         Top             =   570
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "저장 (S)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Font3D          =   2
         Picture         =   "FrmChange.frx":0BB3
      End
      Begin MSComCtl2.DTPicker GeomDate 
         Height          =   375
         Left            =   2010
         TabIndex        =   1
         Top             =   120
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24510465
         CurrentDate     =   36892
      End
      Begin Threed.SSCommand CmdDelete 
         Height          =   915
         Left            =   1650
         TabIndex        =   4
         Top             =   570
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "삭제(&D)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Font3D          =   2
         Picture         =   "FrmChange.frx":0ECD
      End
      Begin Threed.SSCommand CmdExit 
         Height          =   915
         Left            =   3210
         TabIndex        =   6
         Top             =   570
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "E&xit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
         Picture         =   "FrmChange.frx":131F
      End
      Begin Threed.SSCommand CmdSearch 
         Height          =   495
         Left            =   3810
         TabIndex        =   7
         Top             =   60
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "조회(&Q)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin FPSpread.vaSpread SS 
      Height          =   4980
      Left            =   30
      TabIndex        =   9
      Top             =   1560
      Width           =   4755
      _Version        =   196608
      _ExtentX        =   8387
      _ExtentY        =   8784
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   4
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421504
      MaxCols         =   4
      SpreadDesigner  =   "FrmChange.frx":1BB1
      UserResize      =   1
      VisibleCols     =   4
      VisibleRows     =   120
   End
End
Attribute VB_Name = "FrmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StrBdate                As String
Dim StrSlipNo1              As String
Dim StrSlipNo2              As String

Private Sub CmdDelete_Click()
    
    
    If MsgBox("해당 결과를 삭제하시겠습니까?" & vbNewLine & vbNewLine, _
                            vbQuestion + vbYesNo, "결과 삭제") = vbNo Then Exit Sub
                            
    strSQL = ""
    strSQL = strSQL & " UPDATE  TWEXAM_GENERAL_SUB                                  " & vbLf
    strSQL = strSQL & " SET RESULT1 = '',                                           " & vbLf
    strSQL = strSQL & "     VERIFY  = 'N'                                           " & vbLf
    strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & StrBdate & "','YYYY-MM-DD')  " & vbLf    '입력된 날자로 검색
    strSQL = strSQL & "   AND SLIPNO1  =   " & StrSlipNo1 & "                       " & vbLf                       '구분
    strSQL = strSQL & "   AND SLIPNO2  =   " & StrSlipNo2 & "                       " & vbLf                       '구분

    adoConnect.BeginTrans
    
    Result = adoSQL(strSQL)
    If Result <> 0 Or rowindicator = 0 Then
        adoConnect.RollbackTrans
        MsgBox "데이타 삭제중 오류 발생!!, 전산실 문의 바람", vbInformation + vbOKOnly, "오류"
        Exit Sub
    End If
    
    '/==================================================================================================
' TWEXAM_GENERAL 테이블 변경 ... 검사구분... 'E'로 셋팅... 결과 입력 완료...
    
    strSQL = ""
    strSQL = strSQL & "Update Twexam_General           Set                          " & vbLf
    strSQL = strSQL & "         STATUS = 'R'                                        " & vbLf
    strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & StrBdate & "','YYYY-MM-DD')  " & vbLf    '입력된 날자로 검색
    strSQL = strSQL & "   AND SLIPNO1  =   " & StrSlipNo1 & "                       " & vbLf                       '구분
    strSQL = strSQL & "   AND SLIPNO2  =   " & StrSlipNo2 & "                       " & vbLf                       '구분
    
    Result = adoSQL(strSQL)
    
    If Result <> 0 Then
        adoConnect.RollbackTrans
        MsgBox "데이타 삭제중 오류 발생!!, 전산실 문의 바람", vbInformation + vbOKOnly, "오류"
        Exit Sub
    End If
'/==================================================================================================
    
    adoConnect.CommitTrans
    
    Call vaSpread_Clear(SSR, 1, 1, SSR.MaxCols, SSR.MaxRows)
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim i                   As Integer
    Dim ErrChk              As Boolean
    Dim TempItemCd          As String
    Dim TempItemValue       As String
    
    
    If MsgBox("변경된 데이타를 저장하시겠습니까?", vbQuestion + vbYesNo, "저장확인") = vbNo Then Exit Sub
    
    StrBdate = GeomDate.Value
    
    If StrBdate = "" Or StrSlipNo1 = "" Or StrSlipNo2 = "" Then
        MsgBox "변경 저장하려는 데이타를 정확히 선택해 주세요", vbInformation + vbOKOnly, "정보": Exit Sub
    End If
    
    adoConnect.BeginTrans
    
    GoSub GeneralSub_Update
    
    If ErrChk = False Then GoSub General_Update
    
    If ErrChk = False Then
        adoConnect.CommitTrans
    Else
        MsgBox "데이타 저장중 오류 발생!!", vbInformation + vbOKOnly, "정보": Exit Sub
    End If
    Exit Sub
'/==================================================================================================
GeneralSub_Update:

    With SSR
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 2:       TempItemValue = Trim(.Text)
            .Col = 3:       TempItemCd = Trim(.Text)
            If TempItemValue <> "" Or TempItemValue <> "0" Then
                strSQL = ""
                strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB                                    " & vbLf
                strSQL = strSQL & "   SET RESULT1  =   '" & TempItemValue & "',                 " & vbLf
                strSQL = strSQL & "       VERIFY   =  'Y'                                       " & vbLf                                    ' 접수결과에서 VERIFY OK한경우에는 UPDATE하지않음
                strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & StrBdate & "','YYYY-MM-DD')  " & vbLf    '입력된 날자로 검색
                strSQL = strSQL & "   AND SLIPNO1  =   " & StrSlipNo1 & "                       " & vbLf                       '구분
                strSQL = strSQL & "   AND SLIPNO2  =   " & StrSlipNo2 & "                       " & vbLf                       '구분
                strSQL = strSQL & "   AND ITEMCD   =    '" & TempItemCd & "'                    " & vbLf                    'ITEMCODE
                
                Result = adoSQL(strSQL)
                
                If Result <> 0 Or rowindicator = 0 Then
                    adoConnect.RollbackTrans
                    ErrChk = True               '에러발생
                    .Row = i
                    .ReDraw = True
                    .Col = i:       .Col2 = .MaxCols
                    .Row = i:       .Row2 = i
                    .CellBorderColor = &HC0C0&
                    
                    .Row = i:           .Col = 2
                    .Action = ActionActiveCell
                    .ReDraw = False
                    Return
                Else
                    .Row = i
                    .ReDraw = True
                    .Col = i:       .Col2 = .MaxCols
                    .Row = i:       .Row2 = i
                    .CellBorderColor = &HFFFFFF
                    .ReDraw = False
                End If
            End If
        Next i
    End With
    
''/==================================================================================================
' TWEXAM_GENERAL 테이블 변경 ... 검사구분... 'E'로 셋팅... 결과 입력 완료...
General_Update:

    strSQL = ""
    strSQL = strSQL & "Update Twexam_General           Set                         " & vbLf
    strSQL = strSQL & "         Status = 'C'                                       " & vbLf
    strSQL = strSQL & "WHERE JEOBSUDT = TO_DATE('" & StrBdate & "','YYYY-MM-DD')   " & vbLf
    strSQL = strSQL & "    AND SLIPNO1 =   " & Val(StrSlipNo1) & "                 " & vbLf        ' 일련번호
    strSQL = strSQL & "    AND SLIPNO2 =   " & Val(StrSlipNo2) & "                 " & vbLf        ' 일련번호
    
    Result = adoSQL(strSQL)
    
    If Result <> 0 Then
        adoConnect.RollbackTrans
        MsgBox "데이타 갱신중 에러 발생 !!", vbExclamation + vbOKOnly, "Error !!"
    End If

    Return
'/===================================================================================================


End Sub

Private Sub CmdSearch_Click()
    Dim i                       As Integer
    Dim TempPtno                As String
    
    Call vaSpread_Clear(SS, 1, 1, SSR.MaxCols, SSR.MaxRows)
    Call vaSpread_Clear(SSR, 1, 1, SSR.MaxCols, SSR.MaxRows)
    
    strSQL = ""             '/ Verify 된 데이타는 뷰 에서 제외된다.
    strSQL = strSQL & "   SELECT DISTINCT PT.SNAME, TG.PTNO, TG.SLIPNO1, TG.SLIPNO2         " & vbLf
    strSQL = strSQL & "   FROM TWEXAM_GENERAL TG,                                           " & vbLf
    strSQL = strSQL & "        TW_MIS_PMPA.TWBAS_PATIENT PT,                                " & vbLf
    strSQL = strSQL & "     TWEXAM_GENERAL_SUB GS,                                          " & vbLf
    strSQL = strSQL & "     (SELECT CODEKY                                                  " & vbLf
    strSQL = strSQL & "      FROM TWEXAM_ITEMML                                             " & vbLf
    strSQL = strSQL & "      WHERE GEOMJAN1 = '" & GGCODE & "'                              " & vbLf
    strSQL = strSQL & "            AND GEOMJAN3 <>'99') IT                                  " & vbLf
    strSQL = strSQL & "   WHERE TG.JEOBSUDT = TO_DATE('" & GeomDate & "', 'YYYY-MM-DD')     " & vbLf
    strSQL = strSQL & "       AND PT.PTNO = TG.PTNO                                         " & vbLf
    strSQL = strSQL & "       AND TG.SLIPNO1 = GS.SLIPNO1                                   " & vbLf
    strSQL = strSQL & "    AND IT.CODEKY = GS.ITEMCD                                        " & vbLf
    strSQL = strSQL & "    AND GS.PTNO = TG.PTNO                                            " & vbLf
    strSQL = strSQL & "    AND GS.JEOBSUDT = TG.JEOBSUDT                                    " & vbLf
    strSQL = strSQL & "       AND TG.STATUS <> 'C'                                          " & vbLf
    strSQL = strSQL & "   ORDER BY TG.SLIPNO2,TG.PTNO                                       " & vbLf
    '해당장비의 슬립넘버를 기초로... 장비코드를 기준으로 한다
    
'    strSQL = ""             '/ Verify 된 데이타는 뷰 에서 제외된다.
'    strSQL = strSQL & "  SELECT DISTINCT PT.SNAME, TG.PTNO, TG.SLIPNO1, TG.SLIPNO2          " & vbLf
'    strSQL = strSQL & "  FROM TWEXAM_GENERAL TG,                                            " & vbLf
'    strSQL = strSQL & "       TW_MIS_PMPA.TWBAS_PATIENT PT                                 " & vbLf
'    strSQL = strSQL & "  WHERE TG.JEOBSUDT = TO_DATE('" & GeomDate & "', 'YYYY-MM-DD')      " & vbLf
'    strSQL = strSQL & "      AND PT.PTNO = TG.PTNO                                          " & vbLf
'    strSQL = strSQL & "      AND TG.SLIPNO1 IN ( 21)                                            " & vbLf
'    strSQL = strSQL & "      AND TG.STATUS <> 'C'                                           " & vbLf
'    strSQL = strSQL & "  ORDER BY TG.SLIPNO2,TG.PTNO                                        " & vbLf
    Result = adoSQL(strSQL)
    
    If Result <> 0 Or rowindicator = 0 Then Exit Sub
    With SS
        .MaxRows = 15
        .ReDraw = False
        For i = 0 To rowindicator - 1
            If .MaxRows = i Then .MaxRows = i + 2
            .Row = i + 1
            If TempPtno <> AdoGetString(Rs, "PTNO", i) Then
                .Col = 1:           .Text = AdoGetString(Rs, "PTNO", i)
                .Col = 2:           .Text = AdoGetString(Rs, "SNAME", i)
            End If
            .Col = 3:           .Text = AdoGetString(Rs, "SLIPNO1", i)
            .Col = 4:           .Text = AdoGetString(Rs, "SLIPNO2", i)
            TempPtno = AdoGetString(Rs, "PTNO", i)
        Next i
        .ReDraw = True
    End With
End Sub

Private Sub Form_Load()
    GeomDate.Value = GstrSysDate
    Me.Top = CT500.Top + (CT500.ScaleHeight - Me.ScaleHeight)
    Me.Left = CT500.Left + (CT500.ScaleWidth - Me.ScaleWidth)
End Sub

Private Sub SS_DblClick(ByVal Col As Long, ByVal Row As Long)
  
    If Row > 0 And (Col = 3 Or Col = 4) Then
        With SS
            .Row = Row:
            .Col = 1:       StrBdate = Format(GeomDate.Value, "YYYYMMDD")
            .Col = 3:       StrSlipNo1 = .Text
            .Col = 4:       StrSlipNo2 = .Text
            If StrBdate <> "" And StrSlipNo1 <> "" And StrSlipNo2 <> "" Then
                Call Result_View(StrBdate, StrSlipNo1, StrSlipNo2)
                Button_Setting ("C")
            End If
            
        End With
    Else
        Button_Setting
    End If
End Sub

Sub Button_Setting(Optional ChkOpt As String)
    If ChkOpt = "" Then
        CmdSave.Enabled = False
        CmdDelete.Enabled = False
    Else
        CmdSave.Enabled = True
        CmdDelete.Enabled = True
    End If
End Sub

Private Sub Result_View(strJDate As String, strSlno1 As String, strSlno2 As String)
    Dim i               As Integer
    
    Call vaSpread_Clear(SSR, 1, 1, SSR.MaxCols, SSR.MaxRows)
    
    strSQL = ""
    strSQL = strSQL & " SELECT A.ITEMNM , A.CODEKY, B.RESULT1                           " & vbLf
    strSQL = strSQL & " FROM   TWEXAM_ITEMML A,                                         " & vbLf
    strSQL = strSQL & "     TWEXAM_GENERAL_SUB B                                        " & vbLf
    strSQL = strSQL & " WHERE B.JEOBSUDT  = TO_DATE('" & strJDate & "', 'YYYY-MM-DD')   " & vbLf
    strSQL = strSQL & "     AND B.SLIPNO1 = " & strSlno1 & vbLf
    strSQL = strSQL & "     AND B.SLIPNO2 = " & strSlno2 & vbLf
    strSQL = strSQL & "     AND A.CODEKY  = B.ITEMCD                                    " & vbLf
    strSQL = strSQL & " ORDER BY A.CODEKY                                               " & vbLf

    
    Result = adoSQL(strSQL)
    
    If Result <> 0 Or rowindicator = 0 Then
        MsgBox "결과 입력작업이 이루어지지 않은 데이타입니다", vbOKOnly + vbInformation, "결과"
        Exit Sub
    End If
    With SSR
        For i = 0 To rowindicator - 1
            .Row = i + 1
            .Col = 1:           .Text = AdoGetString(Rs, "ITEMNM", i)
            .Col = 2:           .Text = AdoGetString(Rs, "RESULT1", i)
            .Col = 3:           .Text = AdoGetString(Rs, "CODEKY", i)
        Next i
    End With
End Sub


