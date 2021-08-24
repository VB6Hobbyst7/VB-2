Attribute VB_Name = "modCodeSearch"
Option Explicit

Public Const gCboSplitStr$ = "\"

Public Function gfSystemDate() As String
' 시스템 날짜시간
Dim sDate As String

    gSql = "select sysdate from dual"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sDate = Format(.Fields("sysdate").Value, "yyyy-MM-dd HH:nn:SS")
            End If
            .Close
        End If
    End With
    
    gfSystemDate = sDate

End Function

Public Function gfEmpName(ByVal brCode As String) As String
' 사원명
Dim sReturn As String

    If gWorkArea Then
        gSql = "SELECT EMPNM FROM S2COM006 WHERE EMPID = '" & brCode & "'"
    Else
        gSql = "SELECT USER_NM AS EMPNM FROM " & gKahpUserTable & " WHERE USERID='" & brCode & "'"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("EMPNM").Value
            End If
            .Close
        End If
    End With
    gfEmpName = sReturn
    
End Function

Public Function gfTestName(ByVal brCode As String) As String
' 검사명
Dim sReturn As String

    If gWorkArea Then
        gSql = "SELECT TESTNM FROM S2LAB001 WHERE TESTCD = '" & brCode & "'"
    Else
        gSql = "SELECT ITEMHNM AS TESTNM FROM " & gKahpUser & "TWMED_ITEM WHERE ITEMCODE='" & brCode & "'"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("TESTNM").Value
            End If
            .Close
        End If
    End With
    gfTestName = sReturn

End Function

Public Function gfStkName(ByVal brCode As String) As String
' 품명
Dim sReturn As String

    gSql = "SELECT X.CD_ITEM AS STKCD, X.NM_ITEM AS STKNM FROM " & gTBLstk & " X " & vbNewLine & _
           " WHERE X.CD_ITEM = '" & brCode & "'" & gERPStkCondition
    If Len(gERPStkGroup) > 0 Then
        gSql = gSql & " AND GRP_ITEM IN (" & gERPStkGroup & ")"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("STKNM").Value
            End If
            .Close
        End If
    End With
    gfStkName = sReturn
    
End Function

Public Function gfMachName(ByVal brCode As String) As String
' 장비명
Dim sReturn As String

    If gWorkArea Then
        gSql = "SELECT EQPNM FROM S2LAB006 WHERE EQPCD = '" & brCode & "'"
    Else
        gSql = "SELECT MACHNAME AS EQPNM FROM " & gKahpUser & "TWMED_MACHINE WHERE MACHCODE='" & brCode & "'"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("EQPNM").Value
            End If
            .Close
        End If
    End With
    gfMachName = sReturn

End Function

Public Function gfOperName(ByVal brCode As String) As String
' 장비운영명
Dim sReturn As String

    gSql = "SELECT OPERNM FROM S2PIS005 WHERE OPERCD = '" & brCode & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("OPERNM").Value
            End If
            .Close
        End If
    End With
    gfOperName = sReturn

End Function

Public Function gfReasonName(ByVal brCode As String) As String
' 사유명
Dim sReturn As String

    gSql = "SELECT REASONNM FROM S2PIS006 WHERE REASONCD = '" & brCode & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("REASONNM").Value
            End If
            .Close
        End If
    End With
    gfReasonName = sReturn
    
End Function

Public Function gfSpcName(ByVal brCode As String) As String
' 검체명
Dim sReturn As String

    If gWorkArea Then
        gSql = "SELECT FIELD3 AS SPCNM FROM S2LAB032 WHERE CDINDEX='C215' AND CDVAL1='" & brCode & "'"
    Else
        gSql = "SELECT SPECNAME AS SPCNM FROM " & gKahpUser & "TWMED_SPEC WHERE SPECCODE='" & brCode & "'"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("SPCNM").Value
            End If
            .Close
        End If
    End With
    gfSpcName = sReturn
    
End Function

Public Function gfDepotName(ByVal brCode As String) As String
' 저장고명
Dim sReturn As String

    gSql = "SELECT DEPOTNM FROM S2PIS092 WHERE DEPOTCD = '" & brCode & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("DEPOTNM").Value
            End If
            .Close
        End If
    End With
    gfDepotName = sReturn

End Function

Public Function gfHosName(ByVal brCode As String) As String
' 의뢰처명
Dim sReturn As String

    If gWorkArea Then
        gSql = "SELECT HOSNM FROM S2FIN002 WHERE HOSCD = '" & brCode & "'"
    Else
        gSql = "SELECT CORPNAME AS HOSNM FROM " & gKahpUser & "TWMED_CORP WHERE CORPCODE = '" & brCode & "'"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("HOSNM").Value
            End If
            .Close
        End If
    End With
    gfHosName = sReturn

End Function

Public Function gfCustomName(ByVal brCode As String) As String
' 의뢰처명
Dim sReturn As String

    gSql = "SELECT CSTNM FROM S2PIS002 WHERE CSTCD = '" & brCode & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("FROM").Value
            End If
            .Close
        End If
    End With
    gfCustomName = sReturn

End Function

Public Sub gsSpcCombo(ByVal brObj As XComboBox, Optional ByVal brNull As Boolean = False)
' 검체콤보설정
Dim sRow As Long

    brObj.Clear
    If brNull Then
        brObj.AddItem "", sRow
        sRow = sRow + 1
    End If

    If gWorkArea Then
        'gSql = "SELECT CDVAL1 AS SPCCD, FIELD3 AS SPCNM, FIELD2 AS USEDAY FROM S2LAB032 WHERE CDINDEX='C215' ORDER BY FIELD3"
        brObj.AddItem "진단검체" & Space(50) & gCboSplitStr & "L" & gCboSplitStr & "7", sRow:     sRow = sRow + 1
        brObj.AddItem "세포검체" & Space(50) & gCboSplitStr & "P" & gCboSplitStr & "10", sRow:    sRow = sRow + 1
        brObj.AddItem "조직검체" & Space(50) & gCboSplitStr & "S" & gCboSplitStr & "10", sRow:    sRow = sRow + 1
        brObj.AddItem "일반세포" & Space(50) & gCboSplitStr & "C" & gCboSplitStr & "10", sRow:    sRow = sRow + 1
    Else
        gSql = "SELECT SPECCODE AS SPCCD, SPECNAME AS SPCNM, '10' AS USEDAY FROM " & gKahpUser & "TWMED_SPEC ORDER BY SPECNAME"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    While (Not .EOF)
                        brObj.AddItem .Fields("SPCNM").Value & Space(50 - HLen(.Fields("SPCNM").Value)) & gCboSplitStr _
                                      & .Fields("SPCCD").Value & gCboSplitStr & .Fields("USEDAY").Value, sRow
                        sRow = sRow + 1
                    
                        .MoveNext
                    Wend
                End If
                .Close
            End If
        End With
    End If
    
    brObj.ShowItemNum = brObj.ListCount
    
End Sub

Public Sub gsDepotCombo(ByVal brObj As XComboBox, Optional ByVal brNull As Boolean = False)
' 저장고 콤보설정
Dim sRow As Long

    brObj.Clear
    If brNull Then
        brObj.AddItem "", sRow
        sRow = sRow + 1
    End If

    gSql = "SELECT DEPOTCD,DEPOTNM FROM S2PIS092 ORDER BY DEPOTNM"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                While (Not .EOF)
                    brObj.AddItem .Fields("DEPOTNM").Value & Space(50 - HLen(.Fields("DEPOTNM").Value)) & gCboSplitStr & .Fields("DEPOTCD").Value, sRow
                    sRow = sRow + 1
                
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    
    brObj.ShowItemNum = brObj.ListCount
    
End Sub

Public Sub gsOperCombo(ByVal brObj As XComboBox, Optional ByVal brNull As Boolean = False)
' 저장고 콤보설정
Dim sRow As Long

    brObj.Clear
    If brNull Then
        brObj.AddItem "", sRow
        sRow = sRow + 1
    End If

    gSql = "SELECT OPERCD,OPERNM FROM S2PIS005 ORDER BY OPERNM"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                While (Not .EOF)
                    brObj.AddItem .Fields("OPERNM").Value & Space(50 - HLen(.Fields("OPERNM").Value)) & gCboSplitStr & .Fields("OPERCD").Value, sRow
                    sRow = sRow + 1
                
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    
    brObj.ShowItemNum = brObj.ListCount
    
End Sub

Public Sub gsReasonCombo(ByVal brObj As Object, ByVal brType As String, Optional ByVal brNull As Boolean = False)
' 저장고 콤보설정
Dim sRow As Long

    brObj.Clear
    If brNull Then
        brObj.AddItem "", sRow
        sRow = sRow + 1
    End If

    gSql = "SELECT REASONCD,REASONNM FROM S2PIS006 WHERE KINDFG IN ('0','" & brType & "') ORDER BY REASONNM"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                While (Not .EOF)
                    brObj.AddItem .Fields("REASONNM").Value & Space(50 - HLen(.Fields("REASONNM").Value)) & gCboSplitStr & .Fields("REASONCD").Value, sRow
                    sRow = sRow + 1
                
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    
    brObj.ShowItemNum = brObj.ListCount
    
End Sub

Public Function gfMagamCheck(ByVal brDt As String, Optional ByVal brEqual As Boolean = False) As Boolean
Dim sReturn As Boolean, sStr As String

    sReturn = True
    If brEqual Then
        sStr = ">="
        frmMain.stsBar.Panels(2).Text = ""
    Else
        sStr = ">"
    End If
    ' 이후날짜에 마감자료가 있을 경우 마감 못함
    gSql = "SELECT WORKDT FROM S2PIS311 WHERE WORKDT" & sStr & "'" & brDt & "' GROUP BY WORKDT" & vbNewLine & _
           "UNION ALL SELECT WORKDT FROM S2PIS312 WHERE WORKDT" & sStr & "'" & brDt & "' GROUP BY WORKDT" & vbNewLine & _
           "UNION ALL SELECT WORKDT FROM S2PIS313 WHERE WORKDT" & sStr & "'" & brDt & "' GROUP BY WORKDT"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = False
            End If
            .Close
        End If
    End With
    
    If sReturn = False And brEqual Then
        frmMain.stsBar.Panels(2).Text = "마감된 일자입니다. 작업하실 수 없습니다.!!"
    End If
    
    gfMagamCheck = sReturn

End Function

Public Function gfMagamMaxDate() As String
Dim sReturn As String

    gSql = "SELECT MAX(A.LASTDT) AS LASTDT FROM ( " & vbNewLine & _
           "    SELECT MAX(WORKDT) AS LASTDT FROM S2PIS311" & vbNewLine & _
           "    UNION ALL SELECT MAX(WORKDT) AS LASTDT FROM S2PIS312" & vbNewLine & _
           "    UNION ALL SELECT MAX(WORKDT) AS LASTDT FROM S2PIS313" & vbNewLine & _
           ") A "
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("LASTDT").Value
            End If
            .Close
        End If
    End With
    
    If Len(sReturn) = 0 Then
        sReturn = Format(gfSystemDate, "yyyy") & "0101"
    End If
    
    gfMagamMaxDate = sReturn
    
End Function
