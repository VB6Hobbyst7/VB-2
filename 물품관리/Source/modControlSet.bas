Attribute VB_Name = "modControlSet"
Option Explicit

Public Sub gsSetDutyCombo(ByVal brCbo As ComboBox, Optional ByVal brSort As Boolean = True)
Dim cDuty As clsMstDuty

    brCbo.Clear
    Set cDuty = New clsMstDuty
    With cDuty.cfList
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("dutynm").Value
                brCbo.ItemData(brCbo.NewIndex) = Val(.Fields("dutycd").Value)
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

Public Function gfPresentStkRmd(ByVal brStk As Integer, ByVal brDt As String, Optional ByVal brIOfg As Boolean = True) As Double
Dim sYm As String, sStDate As String, sReturn As Double
    
    ' 이전달까지의 수불량
    If Format(brDt, "MM") = "01" Then
        sYm = Format(brDt, "yyyy-MM")
        gSql = "select prevqty, 0 as inqty, 0 as outqty from stkRMD where stkcd = " & brStk & " and rmdym = '" & sYm & "'"
    Else
        sYm = Format(DateAdd("m", -1, brDt), "yyyy-MM")
        gSql = "select sum(prevqty) as prevqty, sum(inqty) as inqty, sum(outqty) as outqty from stkRMD where stkcd = " & brStk & " and rmdym = '" & sYm & "'" & _
               " group by stkcd"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("prevqty").Value) + Val("" & .Fields("inqty").Value) - Val("" & .Fields("outqty").Value)
            End If
            .Close
        End If
    End With
    
    sStDate = Format(brDt, "yyyy-MM") & "-01"
    
    ' 당월 입고량
    gSql = "select sum(ioqty) as inqty from buyL where stkcd = " & brStk & " and buydt between '" & sStDate & "' and '" & brDt & "'" & _
           " group by stkcd"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = sReturn + Val("" & .Fields("inqty").Value)
            End If
            .Close
        End If
    End With
    
    ' 당월 출고량
    gSql = "select sum(qty) as outqty from outL where stkcd = " & brStk & " and outdt between '" & sStDate & "' and '" & brDt & "'" & _
           " group by stkcd"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = sReturn - Val("" & .Fields("outqty").Value)
            End If
            .Close
        End If
    End With
    
    If brIOfg = False Then
        gSql = "select buyioqty from mstSTK where stkcd = " & brStk
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    If Val("" & .Fields("buyioqty").Value) > 0 Then
                        sReturn = sReturn / Val("" & .Fields("buyioqty").Value)
                    End If
                End If
                .Close
            End If
        End With
    End If
    
    gfPresentStkRmd = sReturn
    
End Function

Public Sub gsSetCustCombo(ByVal brCbo As ComboBox, Optional ByVal brSort As Boolean = True, Optional ByVal brDel As Byte = gAllData)

    brCbo.Clear
    brCbo.AddItem ""
    gSql = "select custcd, custnm from mstCUST"
    If brDel <> gAllData Then
        gSql = gSql & " where delfg = " & brDel
    End If
    gSql = gSql & " order by " & IIf(brSort, "custcd", "custnm")
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("custnm").Value
                brCbo.ItemData(brCbo.NewIndex) = .Fields("custcd").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With
    
End Sub

Public Sub gsSetCustComboPV(ByVal brCbo As PVComboBox, Optional ByVal brSort As Boolean = True, Optional ByVal brDel As Byte = gAllData)

    brCbo.ClearItems
    brCbo.AddItem ""
    gSql = "select custcd, custnm from mstCUST"
    If brDel <> gAllData Then
        gSql = gSql & " where delfg = " & brDel
    End If
    gSql = gSql & " order by " & IIf(brSort, "custcd", "custnm")
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("custcd").Value
                brCbo.SubItem(brCbo.NewIndex, 1) = .Fields("custnm").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With
    
End Sub

Public Sub gsSetKindCombo(ByVal brCbo As ComboBox, Optional ByVal brSort As Boolean = True, Optional ByVal brFg As Byte = gAllData)

    brCbo.Clear
    brCbo.AddItem ""
    gSql = "select kindcd, kindnm from mstSTKG where kindcd > 0"
    If brFg <> gAllData Then
        gSql = gSql & " and reagentfg = " & brFg
    End If
    gSql = gSql & " order by " & IIf(brSort, "kindcd", "kindnm")
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("kindnm").Value
                brCbo.ItemData(brCbo.NewIndex) = .Fields("kindcd").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

Public Sub gsSetKindComboPV(ByVal brCbo As PVComboBox, Optional ByVal brSort As Boolean = True, Optional ByVal brFg As Byte = gAllData)

    brCbo.ClearItems
    brCbo.AddItem ""
    gSql = "select kindcd, kindnm from mstSTKG where kindcd > 0"
    If brFg <> gAllData Then
        gSql = gSql & " and reagentfg = " & brFg
    End If
    gSql = gSql & " order by " & IIf(brSort, "kindcd", "kindnm")
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("kindcd").Value
                brCbo.SubItem(brCbo.NewIndex, 1) = .Fields("kindnm").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

Public Sub gsSetMachCombo(ByVal brCbo As ComboBox, Optional ByVal brSort As Boolean = True, Optional ByVal brFg As Byte = gAllData)

    brCbo.Clear
    brCbo.AddItem ""
    gSql = "select machcd, machnm from mstMACH where machcd > ''"
    If brFg <> gAllData Then
        gSql = gSql & " and delfg = " & brFg
    End If
    gSql = gSql & " order by " & IIf(brSort, "machcd", "machnm")
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("machnm").Value
                brCbo.ItemData(brCbo.NewIndex) = .Fields("machcd").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

Public Sub gsSetMachComboPV(ByVal brCbo As PVComboBox, Optional ByVal brSort As Boolean = True, Optional ByVal brFg As Byte = gAllData)

    brCbo.ClearItems
    brCbo.AddItem ""
    gSql = "select machcd, machnm from mstMACH where machcd > ''"
    If brFg <> gAllData Then
        gSql = gSql & " and delfg = " & brFg
    End If
    gSql = gSql & " order by " & IIf(brSort, "machcd", "machnm")
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("machcd").Value
                brCbo.SubItem(brCbo.NewIndex, 1) = .Fields("machnm").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

Public Sub gsSetOperCombo(ByVal brCbo As ComboBox, Optional ByVal brSort As Boolean = True)

    brCbo.Clear
    brCbo.AddItem ""
    gSql = "select opercd, opernm from mstOPER where opercd > 0"
    gSql = gSql & " order by " & IIf(brSort, "opercd", "opernm")
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                brCbo.AddItem .Fields("opernm").Value
                brCbo.ItemData(brCbo.NewIndex) = .Fields("opercd").Value
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

