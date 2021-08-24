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
Dim rs                          As ADODB.Recordset

Public Sub Read_Announce_Ment()
    
    GnAnnounceGetCount = 0
    
    strSQL = "         SELECT  MgrNo                           " & vbLf
    strSQL = strSQL & "  FROM MIS_OCS_OPD9999.DBO.TWOCS_ANNOUNCEMENT    " & vbLf
    strSQL = strSQL & " WHERE AnnounceDate =  TRUNC(SYSDATE ) " & vbLf
'b    StrSql = StrSql & "   AND IDnumber     = " & Val(GstrPassIDnumber)
    Result = AdoOpenSet(rs, strSQL)
'b    If rowindicator > 0 Then
        Call Read_Announce_Ment_Detail
        FrmAnnounce.Show
'b    End If

'    If rowindicator > 0 Then
'        Call Read_Announce_Ment_Detail
'        If GnAnnounceGetCount > 0 Then FrmAnnounce.Show
'    Else
'        MfrmMain.PictureM.Visible = False
'        strPassOk = "OK"
'        Load FrmViewSlips       'SLIP View 시 Perform을 위해 미리 Load
'        FrmOrders.Show
'    End If
    
End Sub

Public Sub Read_Announce_Ment_Detail()
    Dim i, j, K     As Integer
    Dim strGROUP    As String
    Dim strAND1     As String
    Dim strAND2     As String
    Dim strAND3     As String
    
    Select Case UCase$(Mid$(GstrPassClass, 1, 3))
        Case "OCS": strGROUP = "OCS"
        Case "ADM": strGROUP = "ADM"
        Case "NRS": strGROUP = "NRS"
        Case "NUR": strGROUP = "NUR"
        Case "XRA": strGROUP = "XRAY"
        Case "EXA": strGROUP = "EXAM"
        Case Else:  strGROUP = "PMPA"
    End Select
    
    strAND1 = " (AnnounceGroup IN ( 'ALL ', '" & strGROUP & "' ))                            OR "
    strAND2 = " (AnnounceGroup = 'DEPT'  AND AnnounceDept = '" & GstrPassDept & "')          OR "
    strAND3 = " (AnnounceGroup = 'PERS'  AND AnnouncePerson = " & Val(GstrPassIDnumber) & ")    "
    
'B    StrSql = "FOR ALL "
    strSQL = "         SELECT   EntDate, MgrNo,EntPerson,EntTime,Memos,AnnounceGroup        "
    strSQL = strSQL & "  FROM   MIS_OCS_OPD9999.DBO.TWOCS_ANNOUNCEMENT   "
    strSQL = strSQL & " WHERE   ANNOUNCEDATE = TRUNC(SYSDATE )         "
    strSQL = strSQL & "   AND ( " & strAND1 & strAND2 & strAND3 & " )       "
    strSQL = strSQL & "   AND   MgrNo NOT IN ( SELECT MgrNo FROM MIS_OCS_OPD9999.DBO.TWOCS_ANNOUNCESET "
    strSQL = strSQL & "                         WHERE AnnounceDate = TRUNC(SYSDATE )"
    strSQL = strSQL & "                           AND IDnumber     = " & Val(GstrPassIDnumber)
    strSQL = strSQL & "                           AND GbReTry      = 'N' ) "
    Result = AdoOpenSet(rs, strSQL)
    
    If rowindicator > 0 Then
        GnAnnounceGetCount = rowindicator
        ReDim GnaAnnounceMgrNOs(GnAnnounceGetCount)
        ReDim GnaAnnouncePerson(GnAnnounceGetCount)
        ReDim GsaAnnounceMemos(GnAnnounceGetCount)
        ReDim GsaAnnounceDateTime(GnAnnounceGetCount)
        ReDim GsaAnnounceGroup(GnAnnounceGetCount)
        
        For i = 0 To (GnAnnounceGetCount - 1)
            GnaAnnounceMgrNOs(i + 1) = AdoGetNumber(rs, "MgrNo", i)
            GnaAnnouncePerson(i + 1) = AdoGetNumber(rs, "EntPerson", i)
            GsaAnnounceMemos(i + 1) = AdoGetString(rs, "Memos", i)
            GsaAnnounceGroup(i + 1) = AdoGetString(rs, "AnnounceGroup", i)
            GsaAnnounceDateTime(i + 1) = AdoGetString(rs, "EntDate", i) & "  " & _
                                         AdoGetString(rs, "EntTime", i)
        Next i
    End If

End Sub
