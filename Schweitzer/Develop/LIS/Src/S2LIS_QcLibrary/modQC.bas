Attribute VB_Name = "modQC"
Option Explicit

'Public Const HWND_TOPMOST = -1
'Public Const HWND_NOTOPMOST = -2
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOACTIVATE = &H10
'Public Const SWP_SHOWWINDOW = &H40
'Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'Public LangType As Long '언어 설정 변수


Public Sub LoadControl(ByRef lstList As ListBox)

    Dim SqlStmt As String
    Dim Rs As Recordset
    
    SqlStmt = "select a.*, b.field1 as sectnm from " & T_LAB021 & " a, " & T_LAB032 & " b " & _
              "where b.cdindex = '" & LC3_Section & "' " & _
              "and   b.cdval1 = a.sectcd "
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    While (Not Rs.EOF)
        lstList.addItem Rs.Fields("CtrlCd").Value & vbTab & _
                        Rs.Fields("LevelCd").Value & vbTab & _
                        Rs.Fields("CtrlNm").Value & vbTab & _
                        Rs.Fields("SectCd").Value & vbTab & _
                        Rs.Fields("SectNm").Value
        Rs.MoveNext
    Wend
    
Nodata:
    Set Rs = Nothing
    
End Sub

Public Sub LoadEqpCd(ByRef lstList As ListBox)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim i%
    
    SqlStmt = " select eqpcd , eqpnm from " & T_LAB006
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    For i = 1 To Rs.RecordCount
        lstList.addItem Rs.Fields("eqpcd").Value & vbTab & _
                        Rs.Fields("eqpnm").Value
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    
End Sub


Public Sub LoadSection(ByRef cboCombo As ComboBox)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim i%
    
    SqlStmt = " select cdval1 as SectCd , field1 as SectNm from " & T_LAB032 & _
              " where cdindex = '" & LC3_Section & "' "
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    cboCombo.Clear
    For i = 1 To Rs.RecordCount
        cboCombo.addItem Rs.Fields("SectCd").Value & "   " & _
                         Rs.Fields("SectNm").Value
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    
End Sub

Public Sub LoadItems(ByRef lstList As ListBox, ByRef lstList1 As ListBox, _
                     Optional ByVal pWorkArea As String = "", Optional ByVal pSpcFg As Boolean = False)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim tmpStr As String
    Dim i%
    
    If Not pSpcFg Then
        SqlStmt = " select distinct a.testcd, a.testnm from " & T_LAB001 & " a " & _
                  " where a.detailfg = '' and  a.panelfg = ''"   '그룹항목/상세항목 제외...
    Else
        SqlStmt = " select distinct a.testcd, a.testnm, a.rsttype, a.rstdiv, b.rstunit, b.avalval" & _
                  " from " & T_LAB001 & " a, " & T_LAB004 & " b " & _
                  " where a.detailfg = '' and  a.panelfg = ''" & _
                  " and   b.testcd = a.testcd  and  b.expdt = '' " & _
                  " and   b.seq = (select max(seq) from  " & T_LAB004 & " where testcd = b.testcd and expdt = '') "
    End If
    
    If pWorkArea <> "" Then
        SqlStmt = SqlStmt & " and a.workarea = '" & pWorkArea & "'"
    End If
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    lstList1.Clear
    For i = 1 To Rs.RecordCount
        tmpStr = Rs.Fields("TestCd").Value & Space(9)
        tmpStr = Mid(tmpStr, 1, 10) & Rs.Fields("TestNm").Value & Space(40)
        If Not pSpcFg Then
            lstList.addItem tmpStr  ', 1, 50) & _
                            rs.Fields("TestNm").Value
        Else
            lstList.addItem Mid(tmpStr, 1, 50) & vbTab & _
                            Rs.Fields("RstType").Value & vbTab & _
                            Rs.Fields("RstDiv").Value & vbTab & Rs.Fields("RstUnit").Value & vbTab & Rs.Fields("AvalVal").Value
        End If
        lstList1.addItem Rs.Fields("TestNm").Value & vbTab & Rs.Fields("TestCd").Value  'Index List
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    
End Sub

Public Sub LoadPrinter(ByRef cboCombo As ComboBox)

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim i As Long
    Dim tmpStr As String

    SqlStmt = "select * from " & T_LAB032 & " " & _
              "where cdindex = " & DBStr(LC3_PrinterId)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    cboCombo.Clear
    For i = 1 To Rs.RecordCount
        tmpStr = Rs.Fields("CdVal1").Value & Space(9)
        cboCombo.addItem Mid(tmpStr, 1, 6) & _
                        Rs.Fields("Field1").Value
        Rs.MoveNext
    Next i
    
Nodata:
    Set Rs = Nothing
    
End Sub


Public Sub LoadCtrlForOrder(ByRef lstList As ListBox, ByVal pBuildCd As String)

    Dim SqlStmt As String
    Dim Rs As Recordset
    
    SqlStmt = "select distinct a.ctrlcd, a.levelcd, a.lotno, b.ctrlnm, b.workarea " & _
              "from " & T_LAB023 & " a, " & T_LAB021 & " b " & _
              "where a.opendt <= '" & Format(Now, CS_DateDbFormat) & "' " & _
              "and   a.expdt >= '" & Format(Now, CS_DateDbFormat) & "' " & _
              "and   b.ctrlcd = a.ctrlcd " & _
              "and   b.levelcd = a.levelcd " & _
              "and   b.buildcd = '" & pBuildCd & "'"
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then GoTo Nodata
    
    lstList.Clear
    While (Not Rs.EOF)
        lstList.addItem Rs.Fields("CtrlCd").Value & vbTab & _
                        Rs.Fields("CtrlNm").Value & vbTab & _
                        Rs.Fields("LevelCd").Value & vbTab & _
                        Rs.Fields("LotNo").Value & vbTab & _
                        Rs.Fields("WorkArea").Value
        Rs.MoveNext
    Wend
    
Nodata:
    Set Rs = Nothing
    
End Sub

Public Sub CodeHelp(iKeyAscii As Integer, lstbox As ListBox, sPreStr As String, _
                    CurCtrl As Control, NextCtrl As Control)

    Dim i%
    Dim sMadenStr As String
    
    sPreStr = Trim(sPreStr)
    '***************  BackSpace 입력시 ( 나머지 문자로 Search )
    If iKeyAscii = vbKeyBack Then
        If Len(sPreStr) < 2 Then Exit Sub
        sMadenStr = Mid(sPreStr, 1, Len(sPreStr) - 1)
        For i = 0 To lstbox.ListCount
            If sMadenStr = Mid(lstbox.List(i), 1, Len(sMadenStr)) Then
                lstbox.Selected(i) = True
                Exit For
            End If
        Next i
    '**************  방향키 입력시
    ElseIf iKeyAscii = vbKeyDown Then
        lstbox.SetFocus
        lstbox.Selected(0) = True
    
   '***************  Return 입력시 ( 현재 Cell에 입력한 그대로의 값을 실제 리스트의
'                                     항목과 비교한후 존재하면 검체코드 로드
    ElseIf iKeyAscii = vbKeyReturn Then

        For i = 0 To lstbox.ListCount - 1
            If sPreStr = Trim(Mid(lstbox.List(i), 1, _
                            InStr(1, lstbox.List(i), Chr(vbKeyTab)) - 1)) Then
                                    
                Exit For
            End If
        Next i
        
        If i > lstbox.ListCount - 1 Then
            'MsgBox " 존재하지 않는 코드 입니다."
            Exit Sub
        End If
        NextCtrl.SetFocus
    
   '***************  Space Bar 입력시( 현재Cell 에 입력한내용을 바탕으로 온전한
    '                 검사항목을 찾아 Cell에 Write
    ElseIf iKeyAscii = vbKeySpace Then
        For i = 0 To lstbox.ListCount - 1
            If sPreStr = Mid(lstbox.List(i), 1, Len(sPreStr)) Then
                Exit For
            End If
        Next i
        
        If i > lstbox.ListCount - 1 Then
            MsgBox " 존재하지 않는 코드입니다."
            Exit Sub
        End If
        CurCtrl.Text = Mid(lstbox.List(i), 1, _
                             InStr(1, lstbox.List(i), Chr(vbKeyTab)) - 1)
  '***************  기타 일반적인 문자 입력시
    Else
        sMadenStr = sPreStr & Chr(iKeyAscii)
        For i = 0 To lstbox.ListCount
            If sMadenStr = Mid(lstbox.List(i), 1, Len(sMadenStr)) Then
                lstbox.Selected(i) = True
                Exit For
            Else
                lstbox.ListIndex = -1
            End If
        Next i
    End If
End Sub

Public Function ksR(ByVal pVal As Single, ByVal pN As Integer) As String
Dim vbMask As String

    Select Case True
        Case pN < 0
            ksR = "E"
            Exit Function
        Case pN = 0
            vbMask = "########0"
        Case Else
            vbMask = "#########0." & Mid("000000000000000", 1, pN)
    End Select
    
    ksR = Format(pVal, vbMask)

End Function

