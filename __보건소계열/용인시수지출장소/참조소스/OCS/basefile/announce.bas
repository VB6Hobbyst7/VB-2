Attribute VB_Name = "BasAnnounce"
Option Explicit

Global GstrSabun                As String
Global GstrSabunName            As String
Global GnMgrNo                  As Long

Global GnAnnounceGetCount       As Integer

Global GnaAnnounceMgrNOs()      As Long
Global GnaAnnouncePerson()      As Long
Global GsaAnnounceMemos()       As String
Global GsaAnnounceDateTime()    As String
Global GsaAnnounceGroup()       As String

Public Sub Read_Announce_Ment()
    
    GnAnnounceGetCount = 0
    
    strSql = "SELECT GbRetry FROM TWBAS_ANNOUNCESET " & _
             " WHERE AnnounceDate = TRUNC(SYSDATE)              " & _
             "   AND IDnumber     = " & Val(GstrPassIDnumber)
    
    If OpenRDO(strSql, 0) Then
        RdoSet(0).Close
    Else
        Call Read_Announce_Ment_Detail
        If GnAnnounceGetCount > 0 Then FrmAnnounce.Show 1
    End If
    
End Sub

Private Sub Read_Announce_Ment_Detail()
    Dim i, j, k     As Integer
    Dim strGROUP    As String
    Dim strAND1     As String
    Dim strAND2     As String
    Dim strAND3     As String
    
    Select Case UCase$(Mid$(GstrPassClass, 1, 3))
        Case "OCS": strGROUP = "OCS"
        Case "ADM": strGROUP = "ADM"
        Case "NRS": strGROUP = "NRS"
        Case "XRA": strGROUP = "XRAY"
        Case "EXA": strGROUP = "EXAM"
        Case Else:  strGROUP = "PMPA"
    End Select
    
    strAND1 = " (AnnounceGroup IN ( 'ALL ', '" & strGROUP & "' ))                            OR "
    strAND2 = " (AnnounceGroup = 'DEPT'  AND AnnounceDept = '" & GstrPassDept & "')          OR "
    strAND3 = " (AnnounceGroup = 'PERS'  AND AnnouncePerson = " & Val(GstrPassIDnumber) & ")    "
    
    strSql = ""
    strSql = strSql & "SELECT   TO_CHAR(EntDate, 'YYYY-MM-DD') EntDate,     "
    strSql = strSql & "         MgrNo,EntPerson,EntTime,Memos,AnnounceGroup "
    strSql = strSql & "  FROM   TWBAS_ANNOUNCEMENT              "
    strSql = strSql & " WHERE   ANNOUNCEDATE = TRUNC(SYSDATE)               "
    strSql = strSql & "   AND ( " & strAND1 & strAND2 & strAND3 & " )       "
    
    If OpenRDO(strSql, 0) Then
        GnAnnounceGetCount = RdoSet(0).RowCount
        ReDim GnaAnnounceMgrNOs(GnAnnounceGetCount)
        ReDim GnaAnnouncePerson(GnAnnounceGetCount)
        ReDim GsaAnnounceMemos(GnAnnounceGetCount)
        ReDim GsaAnnounceDateTime(GnAnnounceGetCount)
        ReDim GsaAnnounceGroup(GnAnnounceGetCount)
        
        For i = 1 To GnAnnounceGetCount
            GnaAnnounceMgrNOs(i) = RdoSet(0).rdoColumns("MgrNo")
            GnaAnnouncePerson(i) = RdoSet(0).rdoColumns("EntPerson")
            GsaAnnounceMemos(i) = RdoSet(0).rdoColumns("Memos")
            GsaAnnounceGroup(i) = RdoSet(0).rdoColumns("AnnounceGroup")
            GsaAnnounceDateTime(i) = RdoSet(0).rdoColumns("EntDate") & "  " & _
                                     RdoSet(0).rdoColumns("EntTime")
            RdoSet(0).MoveNext
        Next i
        
        RdoSet(0).Close
    End If
End Sub
