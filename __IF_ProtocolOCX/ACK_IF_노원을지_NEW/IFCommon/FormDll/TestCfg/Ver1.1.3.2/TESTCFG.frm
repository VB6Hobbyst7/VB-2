VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmTestCfg 
   Caption         =   "인터페이스 검사항목 설정"
   ClientHeight    =   10140
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15285
   Icon            =   "TESTCFG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   15285
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread spdIFItem 
      Height          =   7035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   12409
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      ColsFrozen      =   2
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
      MaxCols         =   22
      MaxRows         =   200
      NoBeep          =   -1  'True
      SpreadDesigner  =   "TESTCFG.frx":014A
      UserResize      =   0
      VisibleCols     =   11
      VisibleRows     =   100
      TextTip         =   1
   End
   Begin FPSpread.vaSpread spdCalItem 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   8340
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   2990
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      ColsFrozen      =   2
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
      MaxCols         =   16
      MaxRows         =   10
      NoBeep          =   -1  'True
      SpreadDesigner  =   "TESTCFG.frx":1629
      UserResize      =   0
      VisibleCols     =   9
      VisibleRows     =   10
      TextTip         =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   2145
      _Version        =   65536
      _ExtentX        =   3784
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "   계산항목"
      ForeColor       =   8454143
      BackColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Alignment       =   1
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   900
      Left            =   13170
      TabIndex        =   3
      Top             =   7290
      Width           =   1020
      _Version        =   65536
      _ExtentX        =   1799
      _ExtentY        =   1587
      _StockProps     =   78
      Caption         =   "저 장"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "TESTCFG.frx":204F
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   900
      Left            =   14175
      TabIndex        =   4
      Top             =   7290
      Width           =   1020
      _Version        =   65536
      _ExtentX        =   1799
      _ExtentY        =   1587
      _StockProps     =   78
      Caption         =   "닫 기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "TESTCFG.frx":30F1
   End
   Begin Threed.SSCommand cmdFlag 
      Height          =   375
      Left            =   135
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7305
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Flag 설정"
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "frmTestCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CFValidationChk() As Boolean
    Dim i%, j%
    Dim vNm, vCF, vRtnVal
    Dim sCF$
    
    CFValidationChk = True
    
    With spdCalItem
        For i = 1 To .MaxRows
            Call .GetText(2, i, vNm)
            Call .GetText(4, i, vCF)
            
            If vNm = "" And vCF = "" Then
            Else
                sCF = CStr(vCF)
                
                For j = 1 To MAXIFITEM
                    sCF = Replace(sCF, "I", "10")
                Next
                
                Call CFCompute(sCF, vRtnVal)
                
                If vRtnVal = False Then
                    MsgBox "계산식에 오류가 있습니다!!" & vbCrLf & "()를 넣어 보시기 바랍니다..."
                    CFValidationChk = False
                    Exit Function
                Else
                End If
            End If
        Next
    End With
End Function

Private Sub DisplayIFItem(ByVal iIFItemCnt As Integer)
    Dim i%
    Dim vSeq
    
    With spdIFItem
        For i = 1 To iIFItemCnt
            Call .SetText(2, Val(gIFItem(i).s01), CVar(gIFItem(i).s02 & ""))
            Call .SetText(3, Val(gIFItem(i).s01), CVar(gIFItem(i).s03 & ""))
            Call .SetText(4, Val(gIFItem(i).s01), CVar(gIFItem(i).s04 & ""))
            Call .SetText(5, Val(gIFItem(i).s01), CVar(gIFItem(i).s05 & ""))
            Call .SetText(6, Val(gIFItem(i).s01), CVar(gIFItem(i).s06 & ""))
            Call .SetText(7, Val(gIFItem(i).s01), CVar(gIFItem(i).s07 & ""))
            Call .SetText(8, Val(gIFItem(i).s01), CVar(gIFItem(i).s08 & ""))
                 
            'ComboBox Text
            .Row = Val(gIFItem(i).s01)
            .Col = 9
            .TypeComboBoxCurSel = CInt(Val(gIFItem(i).s09))
                 
            Call .SetText(10, Val(gIFItem(i).s01), CVar(gIFItem(i).s10 & ""))
            Call .SetText(11, Val(gIFItem(i).s01), CVar(gIFItem(i).s11 & ""))
            Call .SetText(12, Val(gIFItem(i).s01), CVar(gIFItem(i).s12 & ""))
            Call .SetText(13, Val(gIFItem(i).s01), CVar(gIFItem(i).s13 & ""))
            Call .SetText(14, Val(gIFItem(i).s01), CVar(gIFItem(i).s14 & ""))
            Call .SetText(15, Val(gIFItem(i).s01), CVar(gIFItem(i).s15 & ""))
            
            'ComboBox Text
            .Row = Val(gIFItem(i).s01)
            .Col = 16
            .TypeComboBoxCurSel = CInt(Val(gIFItem(i).s16))
            
            Call .SetText(17, Val(gIFItem(i).s01), CVar(gIFItem(i).s17 & ""))
            Call .SetText(18, Val(gIFItem(i).s01), CVar(gIFItem(i).s18 & ""))
            
            'ComboBox Text
            .Row = Val(gIFItem(i).s01)
            .Col = 19
            
            If gIFItem(i).s19 = "" Then
                .TypeComboBoxCurSel = 0
            Else
                .TypeComboBoxCurSel = Val(gIFItem(i).s19) + 1
            End If
            
            Call .SetText(20, Val(gIFItem(i).s01), CVar(gIFItem(i).s20 & ""))
            
            'ComboBox Text
            .Row = Val(gIFItem(i).s01)
            .Col = 21
            
            If gIFItem(i).s21 = "" Then
                .TypeComboBoxCurSel = 0
            Else
                .TypeComboBoxCurSel = Val(gIFItem(i).s21) + 1
            End If
            
            Call .SetText(22, Val(gIFItem(i).s01), CVar(gIFItem(i).s22 & ""))
        Next
    End With
End Sub

Private Sub DisplayCalItem(ByVal iCalItemCnt As Integer)
    Dim i%
    Dim vSeq
        
    With spdCalItem
        For i = 1 To iCalItemCnt
            Call .SetText(2, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s02 & "")
            Call .SetText(3, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s03 & "")
            Call .SetText(4, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s04 & "")
            Call .SetText(5, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s05 & "")
            Call .SetText(6, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s06 & "")
                 
            'ComboBox Text
            .Row = CInt(Right(gCalItem(i).s01, 1)) + 1
            .Col = 7
            .TypeComboBoxCurSel = CInt(Val(gCalItem(i).s07))
            
            Call .SetText(8, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s08 & "")
            Call .SetText(9, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s09 & "")
            Call .SetText(10, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s10 & "")
            Call .SetText(11, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s11 & "")
            Call .SetText(12, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s12 & "")
            Call .SetText(13, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s13 & "")
            
            'ComboBox Text
            .Row = CInt(Right(gCalItem(i).s01, 1)) + 1
            .Col = 14
            .TypeComboBoxCurSel = CInt(Val(gCalItem(i).s14))
            
            Call .SetText(15, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s15 & "")
            Call .SetText(16, CInt(Right(gCalItem(i).s01, 1)) + 1, gCalItem(i).s16 & "")
        Next
    End With
End Sub

Private Sub GetItem()
    On Error GoTo ErrHandler
    
    Dim objDB As Object
    Dim sRetVal1$, sRetVal2$
    Dim iItemCnt%
    
    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    sRetVal1 = objDB.Get_IFTestItem(gsMachineCd, 0)
    
    sRetVal2 = objDB.Get_CalTestItem(gsMachineCd, 0)
        
    If sRetVal1 = "NONE" Then
    Else
        iItemCnt = GetByOneUserSymbol(sRetVal1, sRetVal1, Chr(3))
        Call MakeIFItemStruct(sRetVal1, iItemCnt)
        Call DisplayIFItem(iItemCnt)
    End If
    
    If sRetVal2 = "NONE" Then
    Else
        iItemCnt = GetByOneUserSymbol(sRetVal2, sRetVal2, Chr(3))
        Call MakeCalItemStruct(sRetVal2, iItemCnt)
        Call DisplayCalItem(iItemCnt)
    End If
    
    Exit Sub
    
ErrHandler:
    Set objDB = Nothing
    MsgBox "GetItem - Local DB 연결 실패!!"
End Sub

Private Sub SpdInit()
    Dim i%
    
    With spdIFItem
        .MaxRows = MAXIFITEM
        
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = RGB(255, 255, 255)
        .EditModeReplace = True
        .NoBeep = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
        
        For i = 1 To MAXIFITEM
            Call .SetText(1, i, Format(i, "000") & "")
        Next
        
        .BlockMode = True
        .Col = 1
        .Col2 = 1
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
    End With

    With spdCalItem
        .MaxRows = MAXCALITEM
        
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = RGB(255, 255, 255)
        .EditModeReplace = True
        .NoBeep = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
        
        For i = 1 To MAXCALITEM
            Call .SetText(1, i, "C" & CStr(i - 1) & "")
        Next
        
        .BlockMode = True
        .Col = 1
        .Col2 = 1
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFlag_Click()

    Load frmFlagCfg
    frmFlagCfg.Show 1
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    Dim objDB As Object
    Dim sRtnVal$
    Dim i%, iExist1%, iExist2%
    Dim vSeq, vNm, vOcd, vRcd, vSPCcd, vSVRcd, vDot, vLHU, vJudge, vMRef1, vMRef2, vFRef1, vFRef2, vPanL, vPanH, vDelGbn, vDelL, vDelH, vCF
    Dim sTSeq$, sTNm$, sTOcd$, sTRcd$, sTSPCcd$, sTSVRcd$, sTDot$, sTLHU$, sTJudge$, sTMRef1$, sTMRef2$, sTFRef1$, sTFRef2$, sTCF$
    Dim sTPanL$, sTPanH$, sTDelGbn$, sTDelL$, sTDelH$
    Dim sTLimit1Gbn$, sTLimit1$, sTLimit2Gbn$, sTLimit2$
    Dim vLimit1Gbn, vLimit1, vLimit2Gbn, vLimit2
    
    If CFValidationChk = False Then
        Exit Sub
    End If
        
    For i = 1 To MAXIFITEM
        With spdIFItem
            Call .GetText(1, i, vSeq)
            Call .GetText(2, i, vNm)
            Call .GetText(3, i, vOcd)
            Call .GetText(4, i, vRcd)
            Call .GetText(5, i, vSPCcd)
            Call .GetText(6, i, vSVRcd)
            Call .GetText(7, i, vDot)
            Call .GetText(8, i, vLHU)
            
            If vLHU = "" And vDot <> "" Then
                vLHU = "H"
            End If
               
            .Col = 9
            .Row = i
        
            If .Value = "" Or .Value = "0" Then
                vJudge = "0"
            Else
                vJudge = .Value
            End If
            
            Call .GetText(10, i, vMRef1)
            Call .GetText(11, i, vMRef2)
            Call .GetText(12, i, vFRef1)
            Call .GetText(13, i, vFRef2)
            Call .GetText(14, i, vPanL)
            Call .GetText(15, i, vPanH)
            
            .Col = 16
            .Row = i
            
            If .Value = "" Or .Value = "0" Then
                vDelGbn = "0"
            Else
                vDelGbn = .Value
            End If
            
            Call .GetText(17, i, vDelL)
            Call .GetText(18, i, vDelH)
            
            vCF = ""
            
            .Col = 19
            .Row = i
            
            If .Value = "" Or .Value = "0" Then
                vLimit1Gbn = ""
            Else
                vLimit1Gbn = CStr(Val(.Value) - 1)
            End If
            
            Call .GetText(20, i, vLimit1)
            
            .Col = 21
            .Row = i
            
            If .Value = "" Or .Value = "0" Then
                vLimit2Gbn = ""
            Else
                vLimit2Gbn = CStr(Val(.Value) - 1)
            End If
            
            Call .GetText(22, i, vLimit2)
            
            If vNm = "" Then
                iExist1 = iExist1 + 0
            Else
                iExist1 = iExist1 + 1
                sTSeq = sTSeq & CStr(vSeq) & Chr(124)
                sTNm = sTNm & CStr(vNm) & Chr(124)
                sTOcd = sTOcd & CStr(vOcd) & Chr(124)
                sTRcd = sTRcd & CStr(vRcd) & Chr(124)
                sTSPCcd = sTSPCcd & CStr(vSPCcd) & Chr(124)
                sTSVRcd = sTSVRcd & CStr(vSVRcd) & Chr(124)
                sTDot = sTDot & CStr(vDot) & Chr(124)
                sTLHU = sTLHU & CStr(vLHU) & Chr(124)
                sTJudge = sTJudge & CStr(vJudge) & Chr(124)
                sTMRef1 = sTMRef1 & CStr(vMRef1) & Chr(124)
                sTMRef2 = sTMRef2 & CStr(vMRef2) & Chr(124)
                sTFRef1 = sTFRef1 & CStr(vFRef1) & Chr(124)
                sTFRef2 = sTFRef2 & CStr(vFRef2) & Chr(124)
                sTPanL = sTPanL & CStr(vPanL) & Chr(124)
                sTPanH = sTPanH & CStr(vPanH) & Chr(124)
                sTDelGbn = sTDelGbn & CStr(vDelGbn) & Chr(124)
                sTDelL = sTDelL & CStr(vDelL) & Chr(124)
                sTDelH = sTDelH & CStr(vDelH) & Chr(124)
                sTCF = sTCF & CStr(vCF) & Chr(124)
                sTLimit1Gbn = sTLimit1Gbn & CStr(vLimit1Gbn) & Chr(124)
                sTLimit1 = sTLimit1 & CStr(vLimit1) & Chr(124)
                sTLimit2Gbn = sTLimit2Gbn & CStr(vLimit2Gbn) & Chr(124)
                sTLimit2 = sTLimit2 & CStr(vLimit2) & Chr(124)
            End If
        End With
    Next
    
    For i = 1 To MAXCALITEM
        With spdCalItem
            Call .GetText(1, i, vSeq)
            Call .GetText(2, i, vNm)
            
            vOcd = ""
            vRcd = ""
            vSPCcd = ""
            
            Call .GetText(3, i, vSVRcd)
            Call .GetText(4, i, vCF)
            Call .GetText(5, i, vDot)
            Call .GetText(6, i, vLHU)
            
            If vLHU = "" And vDot <> "" Then
                vLHU = "H"
            End If
            
            .Col = 7
            .Row = i
        
            If .Value = "" Or .Value = "0" Then
                vJudge = "0"
            Else
                vJudge = .Value
            End If
            
            Call .GetText(8, i, vMRef1)
            Call .GetText(9, i, vMRef2)
            Call .GetText(10, i, vFRef1)
            Call .GetText(11, i, vFRef2)
            Call .GetText(12, i, vPanL)
            Call .GetText(13, i, vPanH)
            
            .Col = 14
            .Row = i
            
            If .Value = "" Or .Value = "0" Then
                vDelGbn = "0"
            Else
                vDelGbn = .Value
            End If
            
            Call .GetText(15, i, vDelL)
            Call .GetText(16, i, vDelH)
            
            If vNm = "" Or vCF = "" Then
                '계산식이 없는 항목은 제외
                iExist2 = iExist2 + 0
            Else
                iExist2 = iExist2 + 1
                sTSeq = sTSeq & CStr(vSeq) & Chr(124)
                sTNm = sTNm & CStr(vNm) & Chr(124)
                sTOcd = sTOcd & CStr(vOcd) & Chr(124)
                sTRcd = sTRcd & CStr(vRcd) & Chr(124)
                sTSPCcd = sTSPCcd & CStr(vSPCcd) & Chr(124)
                sTSVRcd = sTSVRcd & CStr(vSVRcd) & Chr(124)
                sTDot = sTDot & CStr(vDot) & Chr(124)
                sTLHU = sTLHU & CStr(vLHU) & Chr(124)
                sTJudge = sTJudge & CStr(vJudge) & Chr(124)
                sTMRef1 = sTMRef1 & CStr(vMRef1) & Chr(124)
                sTMRef2 = sTMRef2 & CStr(vMRef2) & Chr(124)
                sTFRef1 = sTFRef1 & CStr(vFRef1) & Chr(124)
                sTFRef2 = sTFRef2 & CStr(vFRef2) & Chr(124)
                sTPanL = sTPanL & CStr(vPanL) & Chr(124)
                sTPanH = sTPanH & CStr(vPanH) & Chr(124)
                sTDelGbn = sTDelGbn & CStr(vDelGbn) & Chr(124)
                sTDelL = sTDelL & CStr(vDelL) & Chr(124)
                sTDelH = sTDelH & CStr(vDelH) & Chr(124)
                sTCF = sTCF & CStr(vCF) & Chr(124)
            End If
        End With
    Next
    
    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))
    
    If iExist1 = 0 And iExist2 = 0 Then
        MsgBox "검사항목 설정 - 저장할 내용이 없습니다!!"
        
        Exit Sub
    Else
        sRtnVal = objDB.Trans_SaveTestItem(gsMachineCd, 0, sTSeq, _
                            sTNm, sTOcd, sTRcd, sTSPCcd, sTSVRcd, sTDot, sTLHU, sTJudge, sTMRef1, sTMRef2, sTFRef1, sTFRef2, _
                            sTPanL, sTPanH, sTDelGbn, sTDelL, sTDelH, sTCF, sTLimit1Gbn, sTLimit1, sTLimit2Gbn, sTLimit2)
    End If
    
    If IsNumeric(sRtnVal) = True Then
        MsgBox "검사항목 설정 - 저장 실패!!", vbExclamation, Me.Caption
    Else
        MsgBox "저장작업이 성공적으로 수행되었습니다!!", vbInformation, Me.Caption
    End If
    
    Exit Sub
ErrHandler:
    Set objDB = Nothing
End Sub

Private Sub Form_Load()
    '화면 초기화
    Call SpdInit
    
    '인터페이스 검사항목 가져와 화면에 뿌리기
    Call GetItem
    
    Me.Caption = "[" & gsMachineNm & "]  " & Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RegEditCurFrmTitle "TestCfg", ""
    ViewMsg ""
End Sub

Private Sub spdCalItem_KeyPress(KeyAscii As Integer)
    With spdCalItem
        If .ActiveCol = 4 Then
            Select Case KeyAscii
                Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8, 40, 41, 73, 46, 42, 47, 43, 45
                    '0~9, BS, (, ), I, ., *, /, +, - 만 입력가능하도록
                Case Else
                    KeyAscii = 0
            End Select
        ElseIf .ActiveCol = 5 Then
            Select Case KeyAscii
                Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
                    '숫자만 입력가능하도록
                Case Else
                    KeyAscii = 0
            End Select
        ElseIf .ActiveCol = 6 Then
            Select Case Chr(KeyAscii)
                Case "L", "H", "U", Chr(8)
                    'L, H, U 만 입력가능하도록
                Case Else
                    KeyAscii = 0
            End Select
        Else
        End If
    End With
End Sub

Private Sub spdIFItem_KeyPress(KeyAscii As Integer)
    With spdIFItem
        If .ActiveCol = 7 Or .ActiveCol = 12 Or .ActiveCol = 13 Or .ActiveCol = 14 Or .ActiveCol = 15 Then
            Select Case KeyAscii
                Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8, 46
                    '숫자만 입력가능하도록
                Case Else
                    KeyAscii = 0
            End Select
        ElseIf .ActiveCol = 8 Then
            Select Case Chr(KeyAscii)
                Case "L", "H", "U", Chr(8)
                    'L, H, U 만 입력가능하도록
                Case Else
                    KeyAscii = 0
            End Select
        Else
        End If
    End With
End Sub
