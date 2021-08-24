VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWBatchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsResultPt"
Option Explicit

Private Enum tblCol
    tcNO = 1
    tcLABNO
    tcPTID
    tcSpc
    tcStat
    tcHold
End Enum

Private Const P_FreeRstLength = 30

'/* Event Information & ETC */
Dim WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Dim WithEvents mnuPopup     As Menu
Attribute mnuPopup.VB_VarHelpID = -1
Dim WithEvents mnuUtil1     As Menu
Attribute mnuUtil1.VB_VarHelpID = -1
Dim WithEvents mnuUtil2     As Menu
Attribute mnuUtil2.VB_VarHelpID = -1

Private ssRst               As Object
Private objdic              As New clsDictionary

'/* Help List Popup 위치
Private mvarSsLeft      As Long
Private mvarSsTop       As Long

Private mvarAGE         As String
Private mvarSEX         As String
Private mvarPtNm        As String

Private mvarApply       As String
Private mvarRefFrVal    As String
Private mvarRefToVAl    As String

Private mvarRefCd       As String

Private mvarARefFrVal   As String
Private mvarARefToVal   As String

Private mvarPanicFrVal  As String
Private mvarPanicToVal  As String

Private mvarArletFrVal  As String
Private mvarArletToVal  As String

' AMR 판정추가
Private mvarAMRFrVal  As String
Private mvarAMRToVal  As String

Private mvarWsSheet9    As Boolean

Public Property Get Age() As String
    Age = mvarAGE
End Property
Public Property Get Sex() As String
    Sex = mvarSEX
End Property
Public Property Get Ptnm() As String
    Ptnm = mvarPtNm
End Property
Public Property Get Apply() As String
    Apply = mvarApply
End Property
Public Property Get RefFrVal() As String
    RefFrVal = mvarRefFrVal
End Property
Public Property Get RefToVal() As String
    RefToVal = mvarRefToVAl
End Property
Public Property Get RefCd() As String
    RefCd = mvarRefCd
End Property
Public Property Get ARefFrVal() As String
    ARefFrVal = mvarARefFrVal
End Property
Public Property Get ARefToVal() As String
    ARefToVal = mvarARefToVal
End Property
Public Property Get PanicFrVal() As String
    PanicFrVal = mvarPanicFrVal
End Property
Public Property Let PanicFrVal(ByVal vData As String)
    mvarPanicFrVal = vData
End Property

Public Property Get PanicToVal() As String
    PanicToVal = mvarPanicToVal
End Property
Public Property Let PanicToVal(ByVal vData As String)
    mvarPanicToVal = vData
End Property

Public Property Get ArletFrVal() As String
    ArletFrVal = mvarArletFrVal
End Property

Public Property Let ArletFrVal(ByVal vData As String)
    mvarArletFrVal = vData
End Property

Public Property Get ArletToVal() As String
    ArletToVal = mvarArletToVal
End Property

Public Property Let ArletToVal(ByVal vData As String)
    mvarArletToVal = vData
End Property

Public Property Let SsTop(ByVal vData As Long)
    mvarSsTop = vData
End Property
'
Public Property Get SsTop() As Long
    SsTop = mvarSsTop
End Property
'
Public Property Let SsLeft(ByVal vData As Long)
    mvarSsLeft = vData
End Property

Public Property Get SsLeft() As Long
    SsLeft = mvarSsLeft
End Property


Private Sub Class_Initialize()
    mvarWsSheet9 = False
End Sub

Private Sub Class_Terminate()
    Set objdic = Nothing
    Set objCodeList = Nothing
    Set mnuPopup = Nothing
    Set mnuUtil1 = Nothing
End Sub


Public Sub LoadPic(ByVal strKey As String, ByVal iCol As Long, ByVal strType As String, ByVal Row As Long)
    Dim strValue    As String
    Dim Col         As Long

    If Row < 1 Then Exit Sub

    objdic.KeyChange strKey

    
    Select Case UCase(strType)
            Case "*":
                strValue = "*"
                objdic.Fields("hldiv") = ""

                objdic.Fields("hldiv") = strValue

            Case "H":
                strValue = "H"
                objdic.Fields("hldiv") = ""
                
                objdic.Fields("hldiv") = strValue

            Case "L":
                strValue = "L"
                objdic.Fields("hldiv") = ""
                objdic.Fields("hldiv") = strValue
            Case "A":
                strValue = "A"
                objdic.Fields("dpdiv") = ""

                objdic.Fields("dpdiv") = strValue
            Case "D":
                strValue = "D"
                objdic.Fields("dpdiv") = ""

                objdic.Fields("dpdiv") = strValue
            Case "P":
                strValue = "P"
                objdic.Fields("dpdiv") = ""
                
                objdic.Fields("dpdiv") = strValue
            Case "DP":
                strValue = "DP"
                objdic.Fields("dpdiv") = ""
                                
                objdic.Fields("dpdiv") = strValue
            Case "C":
                strValue = ""
                objdic.Fields("hldiv") = "": objdic.Fields("dpdiv") = ""
                objdic.Fields("hldiv") = strValue
            Case "T":
            Case "G":
            Case "X":
    End Select

    '
    With ssRst
        .Row = Row
        .Col = iCol
        .CellType = 9                                   'CellTypePicture
        objdic.Fields("saverst") = .Value
        If strType = "C" Then
            .CellType = 1                               'CellTypeEdit
'            objdic.Fields("rstcd") = Trim(.Value)
            .ForeColor = vbBlack
            .TypeHAlign = 2                             'TypeHAlignCenter
            .TypeVAlign = 2                             'TypeVAlignCenter

        ElseIf strType = "*" Then
            .CellType = 1                               'CellTypeEdit
            .TypeHAlign = 2                             'TypeHAlignCenter
            .TypeVAlign = 2                             'TypeVAlignCenter
'            objdic.Fields("rstcd") = Trim(.Value)
            .Value = .Value & strType
            .ForeColor = DCM_LightRed                   '약간 붉은색
            Exit Sub
        ElseIf strType = "H" Then
            .CellType = 1                               'CellTypeEdit
            .TypeHAlign = 2                             'TypeHAlignCenter
            .TypeVAlign = 2                             'TypeVAlignCenter
'            objdic.Fields("rstcd") = Trim(.Value)
            .Value = .Value & "▲"
            .ForeColor = DCM_LightRed                   '약간 붉은색
            Exit Sub
        ElseIf strType = "L" Then
            .CellType = 1                               'CellTypeEdit
            .TypeHAlign = 2                             'TypeHAlignCenter
            .TypeVAlign = 2                             'TypeVAlignCenter
'            objdic.Fields("rstcd") = Trim(.Value)
            .Value = .Value & "▼"
            .ForeColor = DCM_LightBlue                  '약간 파란색
            Exit Sub
        Else
            .CellType = 1                               'CellTypeEdit
            .TypeHAlign = 2                             'TypeHAlignCenter
            .TypeVAlign = 2                             'TypeVAlignCenter
            .ForeColor = DCM_LightRed
'            objdic.Fields("rstcd") = Trim(.Value)
            .Value = .Value & strType
        End If

    End With
End Sub

Public Sub ResultCheck(ByVal iCol As Integer, Optional ByVal Row As Long = -1)
    Dim Value   As String
    Dim sLabNo  As String
    Dim sTestCD As String

    If Row = -1 Then
        Row = ssRst.ActiveRow
    End If

    ssRst.Row = Row
    ssRst.Col = tblCol.tcLABNO:  sLabNo = ssRst.Value
    
    If Not mvarWsSheet9 Then
        ssRst.Col = iCol - 1:   sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
    Else
        ssRst.Row = 1
        ssRst.Col = iCol:       sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
        ssRst.Row = Row
    End If
    ssRst.Col = iCol:            Value = Trim(ssRst.Value)

    objdic.KeyChange sLabNo & COL_DIV & sTestCD

    With objdic
        '결과가 널이면 초기화.
        If Value = "" Then
            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
            If .Fields("rsttype") = "N" Or .Fields("rsttype") = "A" Then
                .Fields("rsttype") = ""
            ElseIf .Fields("rsttype") = "F" Or .Fields("rsttype") = "R" Then
                .Fields("rsttype") = .Fields("rsttype")
            End If
            .Fields("rstval") = ""
            .Fields("rstcd") = ""
            .Fields("dpdiv") = ""
            .Fields("hldiv") = ""
            Exit Sub
        End If
        '결과가 들어왔고 결과유형이 널이 아니면 결과유형에 NUMERIC 혹은 ALPHA 셋팅.
        If .Fields("rsttype") = "N" Or .Fields("rsttype") = "A" Or .Fields("rsttype") = "" Then
            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
            If Not (Value Like "*+" Or Value Like "*-") And IsNumeric(Value) Then
                .Fields("rsttype") = "N"
            Else
                .Fields("rsttype") = "A"
            End If
        End If
    End With

    ssRst.Row = ssRst.ActiveRow
    ssRst.Col = iCol
End Sub



Public Sub NumValCheck(ByVal iCol As Long, Optional ByVal Row As Long = -1)
    Dim blnPanic    As Boolean
    Dim blnDelta    As Boolean
    Dim blnArlet    As Boolean
    
    Dim Value       As String
    Dim sLabNo      As String
    Dim sTestCD     As String

    Dim dblDeltaNum As Double

    If Row = -1 Then
       Row = ssRst.ActiveRow
    End If

    ssRst.Row = Row
    ssRst.Col = iCol:           Value = Trim(ssRst.Value)
    
    ssRst.Col = tblCol.tcLABNO: sLabNo = Trim(ssRst.Value)
    
    If Not mvarWsSheet9 Then
        ssRst.Col = iCol - 1:   sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
    Else
        ssRst.Row = 1
        ssRst.Col = iCol:       sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
        ssRst.Row = Row
    End If

    objdic.KeyChange sLabNo & COL_DIV & sTestCD

    If P_ESRResultChk = True Then
        If sTestCD = P_ESRCode Then
            Dim sSEX        As String
            Dim strTmp      As String
            Dim sEsrValue   As String
            Dim sHctValue   As String
            Dim sSQL        As String
            Dim sPTID       As String
            Dim sDefault    As String
            Dim sESRCorrRe  As String
            Dim sRealHctVal As String
            Dim sVfydt      As String
            
            Dim ii          As Long
            
            Dim RS          As Recordset
            
            objdic.MoveFirst
            Do Until objdic.EOF
                If objdic.Fields("testcd") = P_HctCode Then
                    sRealHctVal = objdic.Fields("rstcd")
                    Exit Do
                End If
                objdic.MoveNext
            Loop
            
            strTmp = MsgBox("ESR Correct Check 결과등록을 하시겠습니까?", vbInformation + vbYesNo)
            objdic.KeyChange sLabNo & COL_DIV & sTestCD
            If strTmp = vbYes Then
                sPTID = Trim(objdic.Fields("ptid"))
                sSEX = Trim(medGetP(objdic.Fields("sexage"), 1, "/"))
                
                sSQL = " SELECT rstcd,vfydt FROM " & T_LAB302 & _
                       " WHERE " & _
                                 DBW("ptid=", sPTID) & _
                       " AND " & DBW("testcd=", P_HctCode) & _
                       " AND (vfydt<>' ' or vfydt is not null) " & _
                       " AND " & DBW("vfydt>=", Format(DateAdd("d", -30, GetSystemDate), "YYYYMMDD")) & _
                       " ORDER BY vfydt DESC"
                     
                Set RS = Nothing
                Set RS = New Recordset
                RS.Open sSQL, DBConn
                
                If Not RS.EOF Then
                    sDefault = RS.Fields("rstcd").Value & ""
                    sVfydt = RS.Fields("vfydt").Value & ""
                End If
                
                If sRealHctVal <> "" Then
                    sHctValue = sRealHctVal
                Else
                    If sVfydt <> "" Then
                        sHctValue = CStr(InputBox("Hct Result" & vbCrLf & " 보고일시 : " & _
                                         Mid(sVfydt, 1, 4) & " 년 " & Mid(sVfydt, 5, 2) & " 월 " & Mid(sVfydt, 7) & " 일", "Hct Result", sDefault))
                    Else
                        sHctValue = CStr(InputBox("Hct Result", "Hct Result", sDefault))
                    End If
                End If
                
                If sHctValue = "" Then
                    sHctValue = "0"
                Else
                    sHctValue = CLng(sHctValue)
                End If
                
                If sHctValue < 9 Then sHctValue = "9"
                If sHctValue > 60 Then sHctValue = "60"
                
                sEsrValue = CStr(CLng(medGetP(Value, 1, "/")))
                
                If sEsrValue Mod 2 = 1 Then sEsrValue = sEsrValue - 1
                If sEsrValue < 2 Then sEsrValue = "2"
                If sEsrValue > 60 Then sEsrValue = "60"
                
                sSQL = " SELECT svalue" & sHctValue & " as value FROM S2LAB801 " & _
                       " WHERE " & _
                                 DBW("sex=", sSEX) & _
                       " AND evalue='" & sEsrValue & "'"
                Set RS = Nothing
                Set RS = New Recordset
                RS.Open sSQL, DBConn
                
                If Not RS.EOF Then
                    sESRCorrRe = RS.Fields("value").Value & ""
                End If
                
                If sESRCorrRe <> "" Then
                    Value = sESRCorrRe
                    ssRst.Row = Row
                    ssRst.Col = iCol
                    ssRst.Value = sEsrValue & " / " & sESRCorrRe
                End If
            End If
            Set RS = Nothing
        End If
    End If
    
    
    With objdic
        objdic.KeyChange sTestCD & COL_DIV & sLabNo

        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row

        If Trim(Value) = "" Then Exit Sub
        If Not IsNumeric(Value) Then Exit Sub

        '/* HIGH,LOW CHECK */
        If Val(.Fields("refvalfrom")) <> 0 Or Val(.Fields("refvalto")) <> 0 Then
            If Val(Value) < .Fields("refvalfrom") Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "L", Row
            End If
            If Val(Value) > Val(.Fields("refvalto")) Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "H", Row
            End If
        End If
        
        '/* DELTA CHECK */
        If .Fields("deltafg") = "1" And .Fields("lastrst") <> "" And IsNumeric(.Fields("lastrst")) Then
            If Val(.Fields("lastrst")) = 0 Then
                dblDeltaNum = 0
            Else
                dblDeltaNum = ((Val(.Fields("lastrst")) - Val(Value)) / Val(.Fields("lastrst"))) * 100
            End If
            If dblDeltaNum < CDbl(.Fields("deltaval1")) * (-1) Or dblDeltaNum > CDbl(.Fields("deltaval2")) Then
                blnDelta = True
            End If
        End If
        
        If P_Arlet_PanicChk = True Then
            '/* ARLET CHECK */
            If Val(.Fields("arletfrval")) <> 0 Or Val(.Fields("arlettoval")) <> 0 Then
                If Val(Value) < Val(.Fields("arletfrval")) Or Val(Value) > Val(.Fields("arlettoval")) Then
                    blnArlet = True
                End If
            End If
            '/* PANIC CHECK */: 성별,나이별로 체크한다.
            If Val(.Fields("panicfrval")) <> 0 Or Val(.Fields("panictoval")) <> 0 Then
                If Val(Value) < Val(.Fields("panicfrval")) Or Val(Value) > Val(.Fields("panictoval")) Then
                    blnPanic = True
                End If
            End If
        Else
            '성별/나이로 체크를 않하고, 검사항목(지정검체)으로 체크할경우는 lab004의 Value  값을 가지고 체크한다.
            '2004-02-08 : 김정규
            If .Fields("arletfg") = "1" Then
                If Val(Value) < Val(.Fields("arletfrval")) Or Val(Value) > Val(.Fields("arlettoval")) Then
                    blnArlet = True
                End If
            End If
            If .Fields("panicfg") = "1" Then
                If Val(Value) < Val(.Fields("panicfrval")) Or Val(Value) > Val(.Fields("panictoval")) Then
                    blnPanic = True
                End If
            End If
        End If
        
        '/* PANIC,DELTA PICTURE LOAD */
        '델타/패닉은 같이 들어갈수 있으며, 알터이면서 패닉인 경우는 패닉으로 간주한다.
        
        If blnPanic = True And blnDelta = True Then
            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "DP", Row
        Else
            If blnArlet = True And blnPanic = True Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "P", Row
            ElseIf blnArlet = True Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "A", Row
            ElseIf blnPanic = True Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "P", Row
            End If

            If blnDelta = True Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "D", Row
            End If
        End If
    End With
    
    ssRst.Row = ssRst.ActiveRow
    ssRst.Col = iCol
End Sub

Public Function IsAvalVal(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
    Dim Value   As String
    Dim sTestCD As String
    Dim sLabNo  As String

    IsAvalVal = False
    If Row = -1 Then
        Row = ssRst.ActiveRow
    End If

    ssRst.Row = Row: ssRst.Col = iCol
    Value = Trim(ssRst.Value)

    If Value = "" Then
        IsAvalVal = True
        Exit Function
    End If

    With ssRst
        .Row = .ActiveRow
        .Col = tblCol.tcLABNO:      sLabNo = .Value
        
        If Not mvarWsSheet9 Then
            .Col = iCol - 1: sTestCD = medGetP(.Value, 1, COL_DIV)
        Else
            .Row = 1
            .Col = iCol: sTestCD = medGetP(.Value, 1, COL_DIV)
            .Row = Row
        End If
        
        If objdic.Exists(sLabNo & COL_DIV & sTestCD) Then
            objdic.KeyChange sLabNo & COL_DIV & sTestCD
        End If
        If IsNumeric(Value) = False Then
            LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
            Exit Function
        Else
'            If objdic.Fields("avalval") = "9" Then
                objdic.Fields("rstcd") = Value
                objdic.Fields("rstval") = Value
'            End If
        End If
        
        If objdic.Fields("avalval") = "9" Then
            '유효숫자 관리를 않함. 숫자형이면 다 허용 -- by 김미경
            IsAvalVal = True
            Exit Function
        ElseIf objdic.Fields("avalval") = "0" Then
        '정수형만 관리
            Value = CStr(Val(Value))
            If Value Like "*." Or Value Like "*.*" Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
                Exit Function
            End If
            objdic.Fields("rstcd") = Value
            objdic.Fields("rstval") = Value
        Else
            '유효숫자 관리
            If Len(medGetP(Value, 2, ".")) <> objdic.Fields("avalval") Then
                LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
                Exit Function
            End If
        End If
    End With
    '
    IsAvalVal = True
    '
End Function

Public Function IsAlphaCd(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
    Dim RS      As Recordset
    Dim Value   As String
    Dim sLabNo  As String
    Dim sTestCD As String
    Dim Col     As Long

    IsAlphaCd = False
    If Row = -1 Then Row = ssRst.ActiveRow

    ssRst.Row = Row
    ssRst.Col = tblCol.tcLABNO: sLabNo = ssRst.Value
    
    If Not mvarWsSheet9 Then
        ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
    Else
        ssRst.Row = 1
        ssRst.Col = iCol: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
        ssRst.Row = Row
    End If
        
    objdic.KeyChange sLabNo & COL_DIV & sTestCD

    If ssRst.Value = "" Then
        ssRst.Row = Row:  ssRst.Col = iCol
    End If
    ssRst.Col = iCol

    Value = objdic.Fields("rstcd")
    
    '/* 아이템과 결과코드 연결 체크 */
    Set RS = GetRstCdSQL(LC2_ItemResult, sTestCD, Value)

    If RS.EOF Then
        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
        objdic.Fields("rsttype") = ""
        Set RS = Nothing
        Exit Function
    End If


    '/* 이상결과코드 체크 */
    If Trim("" & RS.Fields("Field2").Value = "") Then
        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
    Else
        LoadPic sLabNo & COL_DIV & sTestCD, iCol, "*", Row
    End If

    IsAlphaCd = True

    Set RS = Nothing
    '
End Function

Public Function IsRateCd(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
    Dim sTestCD As String
    Dim sLabNo  As String
    Dim Value   As String
    Dim Col     As Long

    IsRateCd = False
    If Row = -1 Then Row = ssRst.ActiveRow

    ssRst.Row = Row:
    ssRst.Col = tblCol.tcLABNO: sLabNo = ssRst.Value
    
    If Not mvarWsSheet9 Then
        ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
    Else
        ssRst.Row = 1
        ssRst.Col = iCol: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
        ssRst.Row = Row
    End If

    objdic.KeyChange sLabNo & COL_DIV & sTestCD

    ssRst.Col = iCol
    Value = UCase(Trim(ssRst.Value))
    '
    '/* 비율결과 체크 : 반드시 숫자형으로 들어감, DB는 1:VALUE 형태*/
    LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
    '
    If CStr(Value) = "" Then
        '결과가 들어가지 않으면 Validation체크 하지 않음.
        IsRateCd = True
        Exit Function
    End If
    '
    If IsNumeric(Value) = True Then
        '숫자형이 들어올 경우 정수타입인지 체크
        If Value Like "*.*" Then
            Exit Function
        Else
            IsRateCd = True
        End If
    Else
        '문자가 들어올 경우 아이템과 비율결과코드 연결 체크
        If UpdatableLAB031(LC2_ItemResult, sTestCD, Value) = False Then
            Exit Function
        Else
            IsRateCd = True
        End If
    End If
End Function

Public Function IsFreeResult(ByVal iCol As Long, Optional ByVal Row As Long = -1) As Boolean
    Dim sLabNo  As String
    Dim sTestCD As String
    Dim Value   As String
    Dim Col     As Long

    IsFreeResult = False
    If Row = -1 Then Row = ssRst.ActiveRow

    ssRst.Row = Row
    ssRst.Col = tblCol.tcLABNO: sLabNo = ssRst.Value
    
    If Not mvarWsSheet9 Then
        ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
    Else
        ssRst.Row = 1
        ssRst.Col = iCol:     sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
        ssRst.Row = Row
    End If
    
    objdic.KeyChange sLabNo & COL_DIV & sTestCD

    ssRst.Col = iCol
    Value = Trim(ssRst.Value)
    '
    '/* Free형태의 결과는 (P_FreeRstLength)자리 이내 */
    LoadPic sLabNo & COL_DIV & sTestCD, iCol, "C", Row
    '
    If CStr(Value) = "" Then
        '결과가 들어가지 않으면 Validation체크 하지 않음.
        IsFreeResult = True
        Exit Function
    End If
    '
    If Len(Value) <= P_FreeRstLength Then
        IsFreeResult = True
    End If
    '
End Function



Private Sub ItemRstCodePop(ByVal rCol As Integer, ByVal sTestCD As String)

    Dim strSQL  As String
    Dim PosX    As Long, PosY As Long
    Dim Wid     As Long, Hgt As Long

    Call ssRst.GetCellPos(rCol, ssRst.ActiveRow, PosX, PosY, Wid, Hgt)
    Set objCodeList = New clsPopUpList
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Result Code List"
        .ColumnHeaderText = "Code;Name"
        .ColumnHeaderWidth = "840.189;2310.236"
        .Tag = "ResultCode"
        .HideSearchTool = True
'        .HideColumnHeaders = True
        .SelectByClick = True
        .FormHeight = 2490 '2250
        .LoadPopUp SqlStmtRstCode(sTestCD)

'        .ListCaption = "Result Code List"
'        .ListColHeader = "Name" & vbTab & "Code"
'        .Width = 4050
'        .Height = 3000
'        .Top = SsTop + PosY + 1700
'        .Left = SsLeft + PosX - 810
'        .Tag = "ResultCode"
'        .CaptionOn = True
'        .MultiSel = False
'        strSQL = SqlStmtRstCode(sTestCD)
'        .PopupList strSQL, 2
    End With
End Sub

'Private Sub objCodeList_ListClick(ByVal SelList As String)
'    Dim sTestCD As String
'    Dim sLabNo  As String
'
'
'    If Not IsNull(SelList) Then
'        ssRst.Row = ssRst.ActiveRow
'        ssRst.Col = ssRst.ActiveCol
'        ssRst.Value = medGetP(SelList, 2, vbTab)
'        ssRst.Col = tblCol.tcLABNO:  sLabNo = Trim(ssRst.Value)
'        ssRst.Col = ssRst.ActiveCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
'        If objdic.Exists(sLabNo & COL_DIV & sTestCD) Then
'            objdic.KeyChange sLabNo & COL_DIV & sTestCD
'            objdic.Fields("rstcd") = medGetP(SelList, 1, vbTab)
'        End If
'    End If
'    Set objCodeList = Nothing
'End Sub


Public Function GetRstCd(ByVal TestCd As String, ByVal RstCd As String) As String
    Dim RS As Recordset
    Dim sSQL As String

    sSQL = " SELECT a.cdval2,a.field1" & _
                     " FROM   " & T_LAB031 & " a " & _
                     " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
                     " AND    " & DBW("a.cdval1", TestCd, 2) & _
                     " AND    " & DBW("a.cdval2", RstCd, 2) & _
                     " ORDER BY a.cdval1"

    Set RS = New Recordset
    RS.Open sSQL, DBConn

    If Not RS.EOF Then
        GetRstCd = RS.Fields("field1").Value & ""
    Else
        GetRstCd = RstCd
    End If
    Set RS = Nothing


End Function

Public Sub PopUp(Optional ColorSet As Boolean = True, Optional ByVal iCol As Integer = 0)

    Dim lngBkBackColor  As Long
    Dim blnRstCd        As Boolean
    Dim blnEqpCd        As Boolean
    Dim sTestCD         As String

    If ColorSet = True Then
        With ssRst
            .Row = ssRst.ActiveRow
            If Not mvarWsSheet9 Then
                ssRst.Col = iCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
            Else
                ssRst.Row = 1
                ssRst.Col = iCol: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
                ssRst.Row = ssRst.ActiveRow
            End If

            .Col = -1
'            lngBkBackColor = .BackColor
'            .BackColor = vb3DLight
            DoEvents
        End With
    End If

    Set mnuPopup = frmMenu.mnuPopup

    blnRstCd = IsRstCode(sTestCD)
    Set mnuUtil1 = frmMenu.mnuUtil1
    mnuUtil1.Caption = "Select Result Code"

    If blnRstCd = False Then
        mnuUtil1.Enabled = False
    Else
        mnuUtil1.Enabled = True
    End If

    If mnuUtil1.Enabled Then
        Call ItemRstCodePop(iCol, sTestCD)
    End If

    Set mnuPopup = Nothing
    Set mnuUtil1 = Nothing
End Sub

Public Function SetRST(ByRef objRST As Object, ByRef objDic1 As clsDictionary)
    Set ssRst = objRST
    Set objdic = objDic1
    mvarWsSheet9 = False
End Function

Public Function SetWSRST(ByRef objRST As Object, ByRef objDic1 As clsDictionary, ByVal blnWS9 As Boolean)
    Set ssRst = objRST
    Set objdic = objDic1
    mvarWsSheet9 = blnWS9
End Function


Private Function IsRstCode(ByVal Item As String) As Boolean
    Dim strSQL  As String
    Dim RS      As Recordset

    IsRstCode = False

    strSQL = " SELECT COUNT(a.cdval2) AS ITEMCNT" & _
             " FROM   " & T_LAB031 & " a " & _
             " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
             " AND    " & DBW("a.cdval1", Item, 2)

    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If Not RS.EOF Then
        If Val("" & RS.Fields("ITEMCNT").Value) > 0 Then
           IsRstCode = True
        End If
    End If
    Set RS = Nothing

End Function

Private Function SqlStmtRstCode(ByVal Item As String) As String
    SqlStmtRstCode = " SELECT a.cdval2,a.field1" & _
                     " FROM   " & T_LAB031 & " a " & _
                     " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
                     " AND    " & DBW("a.cdval1", Item, 2) & _
                     " ORDER BY a.cdval1"
End Function

Private Function GetRstCdSQL(ByVal CdIndex As String, ByVal CdVal1 As String, _
                         ByVal CdVal2 As String) As Recordset
    Dim SqlStmt As String
   '
    SqlStmt = " SELECT * FROM " & T_LAB031 & _
              " WHERE " & DBW("cdindex = ", CdIndex) & _
              " AND   " & DBW("cdval1  = ", CdVal1) & _
              " AND   " & DBW("cdval2  = ", CdVal2)
    Set GetRstCdSQL = New Recordset
    GetRstCdSQL.Open SqlStmt, DBConn

   '
End Function
Private Function UpdatableLAB031(ByVal CdIndex As String, ByVal CdVal1 As String, _
                                ByVal CdVal2 As String) As Boolean
    Dim SqlStmt As String
    Dim RS      As Recordset

    SqlStmt = " SELECT * FROM " & T_LAB031 & _
              " WHERE " & DBW("cdindex = ", CdIndex) & _
              " AND   " & DBW("cdval1  = ", CdVal1) & _
              " AND   " & DBW("cdval2  = ", CdVal2)

    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If RS.RecordCount > 0 Then
        UpdatableLAB031 = True
    Else
        UpdatableLAB031 = False
    End If

    Set RS = Nothing
End Function
'WorkSheetCoderk 존재하는지를 판단하는 SQL
Public Function GetIsWorkCd(ByVal WorkCd As String) As Boolean
'
'    GetIsWorkCd = " SELECT COUNT(workcd) AS CNT" & _
'                  " FROM  " & t_lab008 & " " & _
'                  " WHERE " & DBW("workcd", WorkCd, 2)
'
End Function

'WORKSHEET 일괄 등록 대상을 조회해온다.

Public Function WSBatchResultQuery(ByVal sWorkCD As String, ByVal sWorkDt As String, ByVal sFrWorkNo As String, ByVal sToWorkNo As String, ByVal sStat As String) As String
    Dim sSQL As String
    
    
    sSQL = " SELECT b.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " b.accdt " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " b.accseq as labno," & _
           " b.testcd,b.spccd,b.ptid,b.rstcd,b.hldiv,b.dpdiv,b.statfg,b.detailfg,b.ordno,b.ordseq,b.rstdiv," & _
           " a.workseq,d.field5 as spcnm,e.abbrnm5,b.rsttype,b.rstval,f.avalval,b.lastrst,b.txtfg,b.eqpcd," & _
           " f.panicfg,f.panicfrval,f.panictoval,f.deltafg,f.deltaval as deltaval1,f.deltaval2,b.ptid,c.rcvdt,c.rcvtm," & _
           " c.wardid,c.deptcd,c.coldt,c.coltm,c.orddoct,b.orddt,c.hosilid,b.lastvfydt," & _
           " f.arletfg,f.arletfrval,f.arlettoval,g.workinfo" & _
           " FROM " & T_LAB008 & " g," & T_LAB004 & " f," & T_LAB001 & " e," & T_LAB032 & " d," & T_LAB201 & " c," & _
                      T_LAB302 & " b," & T_LAB301 & " a" & _
           " WHERE " & DBW("a.workcd=", sWorkCD) & _
           " AND " & DBW("a.workdt=", sWorkDt) & _
           " AND " & DBW("a.workseq>=", sFrWorkNo) & _
           " AND " & DBW("a.workseq<=", sToWorkNo) & _
           " AND " & DBW("b.rstdiv<>", " ") & _
           " AND " & DBW("b.workcd=", sWorkCD) & _
           " AND " & DBW("d.cdindex=", LC3_Specimen) & _
           " AND (b.vfydt = ''  or  b.vfydt is null)" & _
           " AND b.spccd=d.cdval1" & _
           " AND a.workarea=b.workarea and a.accdt=b.accdt and a.accseq=b.accseq" & _
           " AND a.workarea=c.workarea and a.accdt=c.accdt and a.accseq=c.accseq" & _
           " AND b.workcd=g.workcd and b.testcd=g.testcd and b.spccd=g.spccd" & _
           " AND b.testcd=e.testcd" & _
           " AND b.testcd=f.testcd AND b.spccd=f.spccd"

    If sStat = "1" Then
        sSQL = sSQL & " ORDER BY b.statfg,a.workseq "
    Else
        sSQL = sSQL & " ORDER BY a.workseq,g.workinfo "
    End If
    WSBatchResultQuery = sSQL
End Function
'장비별 일괄 등록 대상을 조회해온다.

Public Function InstBatchResultQuery(ByVal EqpCd As String, ByVal TransDt As String, ByVal sStat As String) As String
    Dim sSQL As String
    
    
    sSQL = " SELECT b.workarea" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " b.accdt " & FUNC_CONCAT & "'-'" & FUNC_CONCAT & " b.accseq as labno," & _
           " b.testcd,b.spccd,b.ptid,b.rstcd,b.hldiv,b.dpdiv,b.statfg,b.detailfg,b.ordno,b.ordseq,b.rstdiv," & _
           " d.field5 as spcnm,e.abbrnm5,b.rsttype,b.rstval,f.avalval,b.lastrst,b.txtfg,b.eqpcd," & _
           " f.panicfg,f.panicfrval,f.panictoval,f.deltafg,f.deltaval as deltaval1,f.deltaval2,b.ptid,c.rcvdt,c.rcvtm," & _
           " c.wardid,c.deptcd,c.coldt,c.coltm,c.orddoct,b.orddt,c.hosilid,b.lastvfydt," & _
           " f.arletfg,f.arletfrval,f.arlettoval" & _
           " FROM " & T_LAB004 & " f," & T_LAB001 & " e," & T_LAB032 & " d," & T_LAB201 & " c," & _
                      T_LAB302 & " b," & T_LAB306 & " a" & _
           " WHERE " & DBW("a.eqpcd=", EqpCd) & _
           " AND " & DBW("a.transdt=", TransDt) & _
           " AND " & DBW("b.rstdiv<>", " ") & _
           " AND " & DBW("b.eqpcd=", EqpCd) & _
           " AND " & DBW("b.autofg=", "1") & _
           " AND " & DBW("d.cdindex=", LC3_Specimen) & _
           " AND (b.vfydt = ''  or  b.vfydt is null)" & _
           " AND b.spccd=d.cdval1" & _
           " AND a.spcyy=c.spcyy and a.spcno=c.spcno" & _
           " AND b.workarea=c.workarea and b.accdt=c.accdt and b.accseq=c.accseq" & _
           " AND b.testcd=e.testcd" & _
           " AND b.testcd=f.testcd AND b.spccd=f.spccd"

    If sStat = "1" Then
        sSQL = sSQL & " ORDER BY b.statfg desc,b.workarea,b.accdt,b.accseq "
    Else
        sSQL = sSQL & " ORDER BY b.workarea,b.accdt,b.accseq "
    End If
    InstBatchResultQuery = sSQL
End Function

'WorkSheet 작성내역의 일자별 마지막 번호를 가지고 온다.
Public Function MaxWorkSeq(ByVal sWorkCD As String, ByVal sWorkDt As String) As String
    MaxWorkSeq = " SELECT max(workseq) as MaxSeq from " & T_LAB301 & _
                 " WHERE  " & DBW("workcd =", sWorkCD) & " " & _
                 " AND    " & DBW("workdt =", sWorkDt) & " "

End Function


'장비별 일괄등록 사용

Public Sub LoadLstEQCode(ByVal pBuildingCd As String, ByRef lstEQCode As Object)
    
    Dim RS  As Recordset
    Dim I   As Integer

    Set RS = New Recordset
    RS.Open GetEqpList(pBuildingCd), DBConn

    If RS.EOF = True Then
        MsgBox " 장비내역이 존재하지 않습니다.", vbExclamation, "메세지"
        GoTo NoData
    End If

    lstEQCode.Clear
    With lstEQCode
        RS.MoveFirst
        While Not RS.EOF
            .AddItem "" & RS.Fields("EqpCd").Value & vbTab & _
                     "" & RS.Fields("EqpNm").Value & vbTab, I
            I = I + 1
            RS.MoveNext
        Wend
    End With
    
NoData:
    Set RS = Nothing
End Sub
 Function GetEqpList(ByVal pBuildingCd As String) As String

    GetEqpList = " select eqpcd, eqpnm from   " & T_LAB006 & _
                 " where  " & DBW("location = ", pBuildingCd) & _
                 " order by eqpcd"

End Function

Public Function GetEqpNm(ByVal pEqpCd As String) As String
    Dim RS     As Recordset
    
    Set RS = New Recordset
    RS.Open " SELECT eqpnm FROM " & T_LAB006 & " WHERE  " & DBW("eqpcd", pEqpCd, 2), DBConn
    
    If Not RS.EOF Then
        GetEqpNm = "" & RS.Fields("EQPNM").Value
    Else
        GetEqpNm = ""
    End If
    Set RS = Nothing
End Function

Public Function EqpHistoryDelete(ByVal EqpCd As String, ByVal TransDt As String) As Boolean
'2001-11-07 추가 : 기존 장비전송내역 삭제 (기간 : 1개월)
    
    Dim sSQL As String
    
On Error GoTo DELETE_ERROR
    sSQL = " DELETE from " & T_LAB306 & _
           " WHERE " & DBW("eqpcd=", EqpCd) & _
           " AND " & DBW("transdt<=", TransDt)
    
    DBConn.BeginTrans   '2001-11-07 추가
    DBConn.Execute sSQL
    EqpHistoryDelete = True
    DBConn.CommitTrans
    Exit Function
    
DELETE_ERROR:
    DBConn.RollbackTrans
End Function

Public Function SqlStmtTransfer(ByVal EqpCd As String) As String
    
    Dim strSQL As String
    Dim sqlTransDtTm As String
    Dim sqlSpcNo As String
    
    sqlSpcNo = "a.spcyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_CONVERT("char", "a.spcno")
    
    SqlStmtTransfer = "SELECT a.transdt," & _
                              sqlSpcNo & " AS SpcNo FROM " & T_LAB306 & " a" & _
                      " WHERE " & DBW("a.transdt >=", Format(Date - 7, "yyyymmdd")) & _
                      " AND   exists(SELECT c.accseq FROM " & T_LAB201 & " c, " & T_LAB302 & " d " & _
                                    " WHERE c.spcyy = a.spcyy AND c.spcno = a.spcno " & _
                                    " AND   d.workarea = c.workarea AND d.accdt = c.accdt AND d.accseq = c.accseq " & _
                                    " AND  (d.vfydt = ''  or  d.vfydt is null) and d.eqpcd=a.eqpcd)" & _
                      " AND   " & DBW("a.eqpcd", EqpCd, 2)
        
    SqlStmtTransfer = SqlStmtTransfer & " ORDER BY transdt "

End Function

Public Sub GetRefVal(ByVal PtId As String, ByVal TestCd As String, ByVal SpcCd As String, ByVal RcvDt As String)


    Dim RS      As Recordset
    Dim sSQL    As String
    Dim sSEX    As String
    Dim sAge    As String

    

     mvarAGE = "":        mvarSEX = "":        mvarPtNm = "":      mvarApply = "":     mvarRefFrVal = ""
     mvarRefToVAl = "":   mvarRefCd = "":      mvarARefFrVal = "": mvarARefToVal = "": mvarPanicFrVal = ""
     mvarPanicToVal = "": mvarArletFrVal = "": mvarArletToVal = "": mvarAMRFrVal = "": mvarAMRToVal = ""
 
    If PtId = "" Then Exit Sub

    sSQL = " SELECT " & F_PTNM & " as PtNm, " & F_DOB2 & " as dob, " & F_SEX & " as sex" & _
           " FROM  " & T_HIS001 & _
           " WHERE " & DBW(F_PTID & " = ", PtId)

    Set RS = New Recordset
    RS.Open sSQL, DBConn

    If RS.RecordCount > 0 Then
        '성명,생일,성별,나이(day),접수일
        Dim sDob As String
        Dim sPtnm As String

        sSEX = Trim("" & RS.Fields("sex").Value)
        If IsNumeric(sSEX) Then
            sSEX = Choose((Val(sSEX) Mod 2) + 1, "F", "M")
        End If
        
'        sSEX = Choose((Val("" & RS.Fields("sex").Value) Mod 2) + 1, "F", "M")
        
        sAge = medFindAge(RS.Fields("dob").Value & "", "D", RcvDt)
        mvarAGE = medFindAge(RS.Fields("dob").Value & "", "Y", RcvDt)
        sDob = RS.Fields("dob").Value & ""
        sPtnm = RS.Fields("ptnm").Value & ""

        sSQL = GetRefValSQL(TestCd, SpcCd, RcvDt, sSEX, sAge)
        
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open sSQL, DBConn
        
        If RS.EOF Then
            sSQL = GetRefValSQL(TestCd, SpcCd, RcvDt, "B", sAge)
            Set RS = Nothing
            Set RS = New Recordset
            RS.Open sSQL, DBConn
        End If
        
        mvarSEX = sSEX
        
        mvarPtNm = sPtnm
        
        If Not RS.EOF Then
            Do Until RS.EOF
                Select Case RS.Fields("applysex").Value & ""
                    Case "M", "F", "B"
                        mvarApply = RS.Fields("applydt").Value & ""
                        mvarRefFrVal = RS.Fields("refvalfrom").Value & ""
                        mvarRefToVAl = RS.Fields("refvalto").Value & ""
                        mvarRefCd = RS.Fields("refcd").Value & ""
                        mvarARefFrVal = RS.Fields("arefvalfrom").Value & ""
                        mvarARefToVal = RS.Fields("arefvalto").Value & ""
                    Case "X", "Y", "Z"
                        mvarPanicFrVal = RS.Fields("panicfrval").Value & ""
                        mvarPanicToVal = RS.Fields("panictoval").Value & ""
                        mvarArletFrVal = RS.Fields("arletfrval").Value & ""
                        mvarArletToVal = RS.Fields("arlettoval").Value & ""
                End Select
                RS.MoveNext
            Loop
            
        End If
        
    End If
    Set RS = Nothing

End Sub


Private Function GetRefValSQL(ByVal TestCd As String, ByVal SpcCd As String, ByVal ApplyDt As String, _
                              ByVal ApplySex As String, ByVal AgeDay As String) As String
    Dim strSQL As String
    
    Dim strSex As String
    
    
    Select Case ApplySex
        Case "M": strSex = "X"
        Case "F": strSex = "Y"
        Case "B": strSex = "Z"
    End Select

    strSQL = " SELECT a.applysex,a.applydt,a.refvalfrom,a.refvalto,a.refcd,a.arefvalfrom,a.arefvalto,a.panicfrval,a.panictoval,a.arletfrval,a.arlettoval" & _
             " FROM   " & T_LAB005 & " a" & _
             " WHERE  " & DBW("a.testcd", TestCd, 2) & _
             " AND    " & DBW("a.spccd", SpcCd, 2) & _
             " AND    a.applydt = (SELECT MAX(b.applydt) FROM " & T_LAB005 & " b" & _
                                " WHERE b.testcd = a.testcd" & _
                                " AND   b.spccd = a.spccd AND b.applysex = a.applysex " & _
                                " AND   b.agefrom = a.agefrom AND b.ageto = a.ageto  " & _
                                " AND  (b.expdt is null or b.expdt = '' ))" & _
             " AND    " & DBW("a.applysex", ApplySex, 2) & _
             " AND    " & DBW("a.agefrom <=", AgeDay) & _
             " AND    " & DBW("a.ageto   >=", AgeDay) & _
             " AND    (a.expdt is null or a.expdt = '') "
    
    strSQL = strSQL & " UNION ALL " & _
                      " SELECT a.applysex,a.applydt,a.refvalfrom,a.refvalto,a.refcd,a.arefvalfrom,a.arefvalto,a.panicfrval,a.panictoval,a.arletfrval,a.arlettoval" & _
                      " FROM   " & T_LAB005 & " a" & _
                      " WHERE  " & DBW("a.testcd", TestCd, 2) & _
                      " AND    " & DBW("a.spccd", SpcCd, 2) & _
                      " AND    a.applydt = (SELECT MAX(b.applydt) FROM " & T_LAB005 & " b" & _
                                          " WHERE b.testcd = a.testcd" & _
                                          " AND   b.spccd = a.spccd AND b.applysex = a.applysex " & _
                                          " AND   b.agefrom = a.agefrom AND b.ageto = a.ageto  " & _
                                          " AND  (b.expdt is null or b.expdt = '' ))" & _
                      " AND    " & DBW("a.applysex", strSex, 2) & _
                      " AND    " & DBW("a.agefrom <=", AgeDay) & _
                      " AND    " & DBW("a.ageto   >=", AgeDay) & _
                      " AND    (a.expdt is null or a.expdt = '') " & _
                      " ORDER BY applydt desc "
    '-> 성은 기존의 성('M','F','-','A') 중 가져오되 없으면 다시 GetRefVal을 Call하여(ApplySex = 'B')
    '    'B'를 차선으로 택해서 가져온다.
    GetRefValSQL = strSQL

End Function

Public Function UpdatableLAB303(ByVal WorkArea As String, ByVal AccDt As String, _
                                ByVal AccSeq As String, ByVal TestCd As String) As Boolean
    Dim SqlStmt As String
    Dim RS      As Recordset

    SqlStmt = " SELECT * FROM " & T_LAB303 & _
              " WHERE  " & DBW("workarea", WorkArea, 2) & _
              " AND    " & DBW("accdt", AccDt, 2) & _
              " AND    " & DBW("accseq", AccSeq, 2) & _
              " AND    " & DBW("testcd", TestCd, 2)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If Not RS.EOF Then
        UpdatableLAB303 = True
    End If
    Set RS = Nothing
End Function

Public Function GetInputCnt(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccseq As Integer)

    Dim tmpRs As New Recordset

    GetInputCnt = 0
    
    tmpRs.Open " SELECT reqinputcnt  FROM " & T_LAB201 & _
                              " where  " & DBW("workarea", pWorkArea, 2) & _
                              " AND    " & DBW("accdt", pAccDt, 2) & _
                              " AND    " & DBW("accseq", pAccseq, 2), DBConn

    GetInputCnt = GetInputCnt + Val("" & tmpRs.Fields("reqinputcnt").Value)
    
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open " SELECT count(*) as InputCnt FROM " & T_LAB302 & _
                              " where  " & DBW("workarea", pWorkArea, 2) & _
                              " AND    " & DBW("accdt", pAccDt, 2) & _
                              " AND    " & DBW("accseq", pAccseq, 2) & _
                              " AND    vfydt <> ' '  AND  vfydt is not null " & _
                              " AND   (detailfg = ''  or  detailfg is null  or rstdiv = '*')", DBConn
    
    GetInputCnt = GetInputCnt + Val("" & tmpRs.Fields("InputCnt").Value)
        
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open " SELECT count(*) as InputCnt FROM " & T_LAB351 & _
                              " where  " & DBW("workarea", pWorkArea, 2) & _
                              " AND    " & DBW("accdt", pAccDt, 2) & _
                              " AND    " & DBW("accseq", pAccseq, 2) & _
                              " AND    vfydt <> ' '  AND  vfydt is not null", DBConn
                              
    GetInputCnt = GetInputCnt + Val("" & tmpRs.Fields("InputCnt").Value)
    
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open " SELECT count(*) as InputCnt FROM " & T_LAB404 & _
                              " where  " & DBW("workarea", pWorkArea, 2) & _
                              " AND    " & DBW("accdt", pAccDt, 2) & _
                              " AND    " & DBW("accseq", pAccseq, 2) & _
                              " AND    vfydt <> ' '  AND  vfydt is not null" & _
                              " AND   (detailfg = ''  or  detailfg is null  or  rstdiv = '*')", DBConn
                              
    GetInputCnt = GetInputCnt + Val("" & tmpRs.Fields("InputCnt").Value)
    
    Set tmpRs = Nothing
End Function

Public Function GetTotCnt(ByVal sWorkArea As String, ByVal sAccDt As String, ByVal sAccSeq As String) As Integer
    Dim sSQL As String
    Dim RS   As Recordset

    sSQL = " SELECT reqtotcnt FROM " & T_LAB201 & _
           " WHERE " & DBW("workarea=", sWorkArea) & _
           " AND " & DBW("accdt=", sAccDt) & _
           " AND " & DBW("accseq=", sAccSeq)
    Set RS = New Recordset
    RS.Open sSQL, DBConn
    
    If Not RS.EOF Then
        GetTotCnt = RS.Fields("reqtotcnt").Value & ""
    End If
    Set RS = Nothing
End Function

Public Function UpdatableLAB304(ByVal WorkArea As String, ByVal AccDt As String, _
                                ByVal AccSeq As String, ByVal Seq As String) As Boolean
    Dim SqlStmt As String
    Dim RS      As Recordset

    SqlStmt = " SELECT * FROM " & T_LAB304 & _
              " WHERE  " & DBW("workarea", WorkArea, 2) & _
              " AND    " & DBW("accdt", AccDt, 2) & _
              " AND    " & DBW("accseq", AccSeq, 2) & _
              " AND    " & DBW("seq", Seq, 2)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If Not RS.EOF Then
        UpdatableLAB304 = True
    End If
    Set RS = Nothing
End Function
Public Function UpdatableLAB205(ByVal WorkArea As String, ByVal AccDt As String, _
                                ByVal AccSeq As String, ByVal TestCd As String) As Boolean
    Dim SqlStmt As String
    Dim RS      As Recordset

    SqlStmt = " SELECT * FROM " & T_LAB205 & _
              " WHERE  " & DBW("workarea", WorkArea, 2) & _
              " AND    " & DBW("accdt", AccDt, 2) & _
              " AND    " & DBW("accseq", AccSeq, 2) & _
              " AND    " & DBW("testcd", TestCd, 2)

    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If Not RS.EOF Then
        UpdatableLAB205 = True
    End If    '
End Function
Public Function IsOrderFollowUp(ByVal PtId As String, ByVal OrdDt As String, _
                                ByVal OrdNo As String, ByVal OrdSeq As String) As Boolean
    Dim SqlStmt     As String
    Dim vfyTotCnt   As Long
    Dim vfyInputCnt As Long
    Dim RS          As Recordset

    IsOrderFollowUp = False
    vfyTotCnt = 0: vfyInputCnt = 0

    SqlStmt = " SELECT COUNT(a.testcd) AS VfyTotCnt FROM " & T_LAB302 & " a" & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    " & DBW("a.orddt", OrdDt, 2) & _
              " AND    " & DBW("a.ordno", OrdNo, 2) & _
              " AND    " & DBW("a.ordseq", OrdSeq, 2) & _
              " AND   (a.vfydt = '' or a.vfydt is null) "
       '" AND a.rstdiv <> '*'"

    Set RS = New Recordset
    RS.Open SqlStmt, DBConn

    If RS.RecordCount > 0 Then
        vfyTotCnt = Val("" & RS.Fields("VfyTotCnt").Value)
        If vfyTotCnt = 0 Then
            IsOrderFollowUp = True
        End If
    End If
    Set RS = Nothing
    '
End Function

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    Dim sTestCD As String
    Dim sLabNo  As String

'    If Not IsNull(SelList) Then
        ssRst.Row = ssRst.ActiveRow
        ssRst.Col = ssRst.ActiveCol
        ssRst.Value = medGetP(pSelectedItem, 2, ";")
        ssRst.Col = tblCol.tcLABNO:  sLabNo = Trim(ssRst.Value)
        ssRst.Col = ssRst.ActiveCol - 1: sTestCD = medGetP(ssRst.Value, 1, COL_DIV)
        If objdic.Exists(sLabNo & COL_DIV & sTestCD) Then
            objdic.KeyChange sLabNo & COL_DIV & sTestCD
            objdic.Fields("rstcd") = medGetP(pSelectedItem, 1, vbTab)
        End If
'    End If
    Set objCodeList = Nothing
End Sub
