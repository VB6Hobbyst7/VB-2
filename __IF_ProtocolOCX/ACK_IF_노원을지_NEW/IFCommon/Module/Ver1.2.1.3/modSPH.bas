Attribute VB_Name = "modSPH"
'
'   저동백병원용 모듈
'
Option Explicit

Public Sub Get_WorkListDT()
    On Error GoTo ErrRtn
    
    Dim objOrd  As Object
    Dim sWKDT   As String: sWKDT = ""
    Dim tmpData()   As String
    Dim ii%
    Dim sBuf$
    
    frmInterface.cboWKDT.Clear
    
    'Order Dll을 Call하여 서버쪽에 Order를 가져옴
    sBuf = gOrdCfg.sComponent
    
    If sBuf = "" Then
        ViewMsg "오더 Dll 파일이 존재하지 않습니다!!"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set objOrd = CreateObject(sBuf)
    Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
    sWKDT = objOrd.GetWKDT(frmInterface.cboWKGbn.Text, Format(frmInterface.dtpWKDate.Value, "YYYYMMDD"))
    Set objOrd = Nothing
    
    'sWKDT = "2005-07-04 20:00:00|2005-07-04 22:00:00|2005-07-04 23:00:00|"
    If sWKDT = "" Then
        Screen.MousePointer = vbDefault
        If frmInterface.Tag = "" Then
            MsgBox "선택한 일자에 WorkList 작성정보가 존재하지 않습니다.", vbInformation
        End If
        Exit Sub
    End If
    
    tmpData() = Split(sWKDT, Chr(124))
    
    With frmInterface.cboWKDT
        For ii = 0 To UBound(tmpData())
            If Trim(tmpData(ii)) = "" Then
                Exit For
            End If
        
            .AddItem (Trim(tmpData(ii)))
        Next ii
        
        .ListIndex = 0
    End With
        
    Screen.MousePointer = vbDefault
    
ErrRtn:
    If Err <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox Err.Description, vbExclamation
        ViewMsg "Get_WorkListDT - " & Err.Description
    End If
End Sub
Private Sub cmdWorkList_Click()
'    On Error GoTo ErrRtn
'
'    Dim ii%, i%
'    Dim sRetVal$, tmpWkNo$
'    Dim tmpRow()    As String
'    Dim tmpData()   As String
'
'    Dim sIFSeq$, sIFOrdCd$
'    Dim sNo1$, sNo2$
'
'    If Trim(frmInterface.cboWorkcd.Text) = "" Or Trim(frmInterface.cboWorkgb.Text) = "" Then
'        MsgBox "조회를 원하는 작업번호를 입력해 주십시요.", vbExclamation, Me.Caption
'        Exit Sub
'    End If
'
'    If Trim(txtWkNo(0)) = "" Or Trim(txtWkNo(1)) = "" Then
'        MsgBox "조회를 원하는 작업번호 구간을 입력해 주십시요.", vbExclamation, Me.Caption
'        txtWkNo(0).SetFocus
'        Exit Sub
'    End If
'
'    MousePointer = vbHourglass
'
'    spdIntList.MaxRows = 0
'
'    'ServiceNm/EqCode
'    Dim tmpEq() As String
'    Dim sServiceNm$
'
'    If InStr(gsMachineNm, "-") > 0 Then
'        tmpEq() = Split(gsMachineNm, "-")
'        sServiceNm = "LEQUGET" & Trim(tmpEq(UBound(tmpEq())))
'    End If
'    '----------------
'
'    If txtWkNo(0) = "" Then
'        sNo1 = "1"
'    Else
'        sNo1 = txtWkNo(0)
'    End If
'    If txtWkNo(1) = "" Then
'        sNo2 = "99999"
'    Else
'        sNo2 = txtWkNo(1)
'    End If
'
'    '--- For AMC
''    sRetVal = mObjAmcSvr.GetOrderOCSensor("20050316", "L45", "00", "I", "L4514", "00001", "00099", sServiceNm)
'    sRetVal = mObjAmcSvr.GetOrderOCSensor(Format(dtpWKDT, "YYYYMMDD"), cboWorkcd.Text, "00", cboWorkgb.Text, gIFItem(1).s06, _
'                                        sNo1, sNo2, sServiceNm)
'
'    'return: wkdt/wksl/wkit/wkio/wkno/paid/dept/ward/bacd/panm
'    '        0    1    2    3    4    5    6    7    8    9
'    'test(03/16 실제 조회데이터)
'    'sRetVal = "20050424|L45|00|I|00001|Z0045201|CP||0515217351|TEST1|20050424|L45|00|I|00002|Z0045201|CP||0515217361|TEST2|"
'
'    If IsNumeric(sRetVal) Then
'        ViewMsg "Error (" & mObjAmcSvr.ErrMsg & ")"
'        Exit Sub
'    End If
'
'    tmpRow() = Split(sRetVal, Chr(3))
'
'    For ii = 0 To UBound(tmpRow())
'        If Trim(tmpRow(ii)) = "" Then
'            Exit For
'        End If
'
'        Erase tmpData()
'        tmpData() = Split(tmpRow(ii), Chr(124))
'
'        If UBound(tmpData()) >= 10 Then
'            tmpWkNo = Trim(tmpData(0)) & "-" & Trim(tmpData(1) & tmpData(2)) & "-" & Trim(tmpData(3)) _
'                    & "-" & Trim(tmpData(4))
'
'            '화면표시
'            With gOrderTable
'                .sJDate = ""
'                .sJGbn = tmpWkNo
'                .sRegNo = Trim(tmpData(5))
'                .sName = Trim(tmpData(9))
'                .sSex = ""
'                .sOther = ""
'                .iOrdCnt = 1
'                .sOrdOpt = "S"
'                .sWDate = Format$(dtpLabDate.Value, "YYYYMMDD")
'                .sJNo = Trim(tmpData(8))
'
'                sIFSeq = "001"  'ConvertIFItemInfo(1, Trim(tmpData(5)))
'                sIFOrdCd = ConvertIFItemInfo(6, sIFSeq)     'IFOrdCd로 변환
'
'                If Trim(sIFSeq) <> "" Then
'                    .iOrdCnt = 1
'                    ReDim .sIFSeq(1)
'
'                    .sIFSeq(1) = sIFSeq
'
'                    '화면표시
'                    Call DisplayOrderOK_OCSensor
'                End If
'            End With
'            '-----------------------
'        End If
'    Next ii
'
'
'    MousePointer = vbDefault
'
'ErrRtn:
'    If Err <> 0 Then
'        MousePointer = vbDefault
'        MsgBox Err.Description, vbExclamation, Me.Caption
'    End If
End Sub
Public Sub DisplayOrderOK_OCSensor(Optional ByVal sState$)
    On Error GoTo ErrHandler
    
    Dim i%
    Dim iRowCnt%
    Dim lngTwipHeight&
    Dim vTmp
    Dim sPos$
    
    If frmInterface.listTest.ListCount > 10 Then
        frmInterface.listTest.RemoveItem (0)
    End If
    frmInterface.listTest.AddItem gOrderTable.sJNo
    
    With frmInterface.spdIntList
        '작업일자를 구함
        gOrderTable.sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")
        '작업일련번호를 구함
'        gOrderTable.sWSeq = Format(Val(GetCurLastWSeq) + 1, "0000")
        gOrderTable.sWSeq = ""
        
        '해당바코드의 오더정보를 넘김
        .MaxRows = .MaxRows + 1
        gOrderTable.iCRow = .MaxRows
        
        Call .RowHeightToTwips(1, .RowHeight(1), lngTwipHeight)
        iRowCnt = Format((.Height / lngTwipHeight) - 2, "0")
        
'        If .MaxRows > iRowCnt Then
'            .TopRow = .MaxRows - iRowCnt + 1
'        End If
        
        Call .SetText(1, gOrderTable.iCRow, gOrderTable.sWSeq & "")
        If sState = "DISPLAY" Then
            Call .SetText(2, gOrderTable.iCRow, CVar("1"))
        Else
            Call .SetText(2, gOrderTable.iCRow, CVar("0"))
        End If
'        'CheckBox
'        Call .SetText(3, gOrderTable.iCRow, vbChecked)
        
        Call .SetText(4, gOrderTable.iCRow, gOrderTable.sJGbn & "")
        Call .SetText(5, gOrderTable.iCRow, gOrderTable.sJNo & "")
'        Call .SetText(6, gOrderTable.iCRow, gOrderTable.sRack & "")
        Call .SetText(6, gOrderTable.iCRow, "1")
        
        If .MaxRows > 1 Then
            Call .GetText(7, gOrderTable.iCRow - 1, vTmp)
            sPos = Trim(Val(vTmp) + 1)
        Else
            sPos = "1"
        End If
        Call .SetText(7, gOrderTable.iCRow, Format(sPos, "0000"))       'for oc-sensor
'        Call .SetText(7, gOrderTable.iCRow, gOrderTable.sPos & "")
        
        Call .SetText(8, gOrderTable.iCRow, gOrderTable.sRegNo & "")
        Call .SetText(9, gOrderTable.iCRow, gOrderTable.sName & "")
        Call .SetText(10, gOrderTable.iCRow, gOrderTable.sSex & "")
        Call .SetText(11, gOrderTable.iCRow, gOrderTable.sEmer & "")
        Call .SetText(12, gOrderTable.iCRow, gOrderTable.sReRun & "")
        Call .SetText(13, gOrderTable.iCRow, gOrderTable.sOther & "")
        Call .SetText(14, gOrderTable.iCRow, CStr(gOrderTable.iOrdCnt) & "")
        Call .SetText(15, gOrderTable.iCRow, "N")
        Call .SetText(16, gOrderTable.iCRow, CStr(gOrderTable.iOrdCnt) & "")
        
        '검사항목 정보 숨기기
        For i = 1 To gOrderTable.iOrdCnt
            Call .SetText(16 + i, gOrderTable.iCRow, gOrderTable.sIFSeq(i) & "||||")
        Next i
        
        If sState <> "DISPLAY" Then
            Call SpdForeBack(frmInterface.spdIntList, 3, 15, gOrderTable.iCRow, gOrderTable.iCRow, _
                    RGB(0, 0, 0), 연노랑)
        
            frmInterface.lblOrder = gOrderTable.sJNo
        End If
    End With
    
'    'Order 내역 Local MDB에 Insert
'    Call RegOrder(1)
    
    'gOrderTable 초기화
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        .sSampNo = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With
    
    Exit Sub
ErrHandler:
    frmInterface.listTest.AddItem "Error"
    
    'gOrderTable 초기화
    With gOrderTable
        .iCRow = 0
        .iOrdCnt = 0
        .sEmer = ""
        Erase .sIFOrdCd
        Erase .sIFRstCd
        Erase .sIFSeq
        .sIFSpcCd = ""
        .sJDate = ""
        .sJGbn = ""
        .sJNo = ""
        .sName = ""
        .sOrdOpt = ""
        .sOther = ""
        .sPos = ""
        .sRack = ""
        .sRegNo = ""
        .sReRun = ""
        .sSampID = ""
        .sSampNo = ""
        Erase .sServerCd
        .sSex = ""
        .sWDate = ""
        .sWSeq = ""
    End With
End Sub

Public Sub Get_WorkList()
    On Error GoTo ErrRtn
    
    Dim objOrd  As Object
    Dim sWKList As String
    Dim tmpRow()    As String
    Dim ii%, kk%, iOrdCnt%
    Dim sBuf$, sOneRow$, sIFSeq$, sTmp$, sTIFSeq$
        
    With frmInterface
        If .cboWKDT.Text = "" Then
            MsgBox "조회를 원하는 WorkList 작성일시를 선택해 주십시요.", vbInformation
            Exit Sub
        End If
        
        'Order Dll을 Call하여 서버쪽에 Order를 가져옴
        sBuf = gOrdCfg.sComponent
        If sBuf = "" Then
            ViewMsg "오더 Dll 파일이 존재하지 않습니다!!"
            Exit Sub
        End If
    
        Screen.MousePointer = vbHourglass
        
        Set objOrd = CreateObject(sBuf)
        Call objOrd.SetMachineInfo(gsMachineCd, gsMachineNm)
        sWKList = objOrd.GetWKList(frmInterface.cboWKGbn.Text, Trim(.cboWKDT.Text))
        Set objOrd = Nothing
        
        'sWKList = "11111111111|111|TEST1|1|001||22222222222|111|TEST2|1|001||"
        If sWKList = "" Then
            Screen.MousePointer = vbDefault
            MsgBox "해당 정보가 존재하지 않습니다.", vbInformation
            Exit Sub
        End If
        
        tmpRow() = Split(sWKList, Chr(3))
        
        For ii = 0 To UBound(tmpRow())
            If Trim(tmpRow(ii)) = "" Then
                Exit For
            End If
            
            iOrdCnt = 0: sTIFSeq = ""
            
            sOneRow = Trim(tmpRow(ii))

            '화면표시
            With gOrderTable
                .sWDate = Format$(frmInterface.dtpLabDate.Value, "YYYYMMDD")
                .sJDate = ""
                .sJGbn = ""
                .sJNo = GetByOne(sOneRow, sOneRow)      'Barcode No
                .sRegNo = GetByOne(sOneRow, sOneRow)
                .sName = GetByOne(sOneRow, sOneRow)
                .sSex = ""
                .sOther = ""
                .iOrdCnt = Val(GetByOne(sOneRow, sOneRow))
                .sOrdOpt = ""
                                    
                For kk = 1 To gOrderTable.iOrdCnt
                    sIFSeq = GetByOne(sOneRow, sOneRow)
                    
                    sTmp = sIFSeq
                    
                    'IFOrdCd로 변환
                    sTmp = ConvertIFItemInfo(6, sTmp)
                    
                    If sTmp = "" Then
                    Else
                        iOrdCnt = iOrdCnt + 1
                        
                        'IFSeq를 합친다
                        sTIFSeq = sTIFSeq & sIFSeq & Chr(124)
                    End If
                Next kk
                
                'IFSeq 순서로 재구성
                sTIFSeq = ReOrder_IFSeq_And_RealOrdCnt(sTIFSeq, iOrdCnt)
                
                gOrderTable.iOrdCnt = iOrdCnt
                ReDim gOrderTable.sIFSeq(iOrdCnt)
                
                For kk = 1 To iOrdCnt
                    gOrderTable.sIFSeq(kk) = GetByOne(sTIFSeq, sTIFSeq)
                Next kk
            End With
            
            If iOrdCnt > 0 Then
                '화면표시
                Call DisplayOrderOK_OCSensor
            End If
        Next ii
    End With
    
    Screen.MousePointer = vbDefault
    
ErrRtn:
    If Err <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Get_WorkList - " & Err.Description, vbExclamation
    End If
End Sub
