Attribute VB_Name = "modMaster"
Option Explicit

Public Function gfStkName(ByVal brCd As Long) As String
Dim sReturn As String

    gSql = "select stknm from mstSTK where stkcd = " & brCd
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("stknm").Value
            End If
            .Close
        End If
    End With
    
    gfStkName = sReturn
    
End Function

Public Function gfCustName(ByVal brCd As Long) As String
Dim sReturn As String

    gSql = "select custnm from mstCUST where custcd = " & brCd
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("custnm").Value
            End If
            .Close
        End If
    End With
    
    gfCustName = sReturn
    
End Function

Public Function gfKindName(ByVal brCd As Integer) As String
Dim sReturn As String

    gSql = "select kindnm from mstSTKG where kindcd = " & brCd
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("kindnm").Value
            End If
            .Close
        End If
    End With
    
    gfKindName = sReturn
    
End Function

Public Function gfDutyName(ByVal brCd As Integer) As String
Dim sReturn As String

    gSql = "select dutynm from mstDUTY where dutycd = " & brCd
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("dutynm").Value
            End If
            .Close
        End If
    End With
    
    gfDutyName = sReturn
    
End Function

Public Function gfMachName(ByVal brCd As String) As String
Dim sReturn As String

    gSql = "select machnm from mstMACH where machcd = '" & brCd & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("machnm").Value
            End If
            .Close
        End If
    End With
    
    gfMachName = sReturn
    
End Function

Public Function gfOperName(ByVal brCd As Integer) As String
Dim sReturn As String

    gSql = "select opernm from mstOPER where opercd = " & brCd
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = "" & .Fields("opernm").Value
            End If
            .Close
        End If
    End With
    
    gfOperName = sReturn
    
End Function

Public Function gfTestName(ByVal brCd As String) As String
Dim sReturn As String

    If cDb.cfOraConnect Then
        gSql = "select itemcode, itemhnm from TWMED_ITEM where itemcode = '" & brCd & "' and visible = 1"
        With cDb.cfOraRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    sReturn = "" & .Fields("itemhnm").Value
                End If
                .Close
            End If
        End With
    End If
    
    gfTestName = sReturn
    
End Function
