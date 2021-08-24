VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private AdoCn               As ADODB.Connection
Private AdoRs               As ADODB.Recordset

Public Sub SetAdoCn(ByVal new_AdoCn As ADODB.Connection)
    Set AdoCn = new_AdoCn
End Sub

'검사항목을 저장
Public Function LetTestInfo(ByVal Test_Property As Scripting.Dictionary) As Boolean
    
On Error GoTo ErrorTrap
    
    SQL = ""
    SQL = SQL & "SELECT EQUIPCD FROM EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
    SQL = SQL & "   AND SEQNO       = " & Test_Property.Item("SEQ") & vbCrLf
    SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
    SQL = SQL & "   AND TESTCODE    = " & STS(Test_Property.Item("TESTCD")) & vbCrLf
    
    If DBExists(AdoCn, SQL) Then
        SQL = ""
        SQL = SQL & "UPDATE EQPMASTER SET " & vbCrLf
        SQL = SQL & " SEQNO = " & Test_Property.Item("SEQ") & vbCrLf
        SQL = SQL & ",SENDCHANNEL   = " & STS(Test_Property.Item("OCH")) & vbCrLf
        SQL = SQL & ",RSLTCHANNEL   = " & STS(Test_Property.Item("RCH")) & vbCrLf
        SQL = SQL & ",TESTCODE      = " & STS(Test_Property.Item("TESTCD")) & vbCrLf
        SQL = SQL & ",TESTNAME      =  " & STS(Test_Property.Item("TESTNM")) & vbCrLf
        SQL = SQL & ",ABBRNAME      =  " & STS(Test_Property.Item("ABBRNM")) & vbCrLf
        SQL = SQL & ",RESPRECUSE    =  " & Test_Property.Item("RESUSE") & vbCrLf
        SQL = SQL & ",RESPREC       =  " & Test_Property.Item("RES") & vbCrLf
        SQL = SQL & ",REFMLOW       =  " & STS(Test_Property.Item("REFML")) & vbCrLf
        SQL = SQL & ",REFMHIGH      =  " & STS(Test_Property.Item("REFMH")) & vbCrLf
        SQL = SQL & ",REFFLOW       =  " & STS(Test_Property.Item("REFFL")) & vbCrLf
        SQL = SQL & ",REFFHIGH      =  " & STS(Test_Property.Item("REFFH")) & vbCrLf
        SQL = SQL & ",RESTYPE       =  " & STS(Test_Property.Item("USERESULT")) & vbCrLf
        '-- AMR
        SQL = SQL & ",AMRLIMIT1     =  " & STS(Test_Property.Item("AMRLIMIT1")) & vbCrLf
        SQL = SQL & ",AMRLIMIT2     =  " & STS(Test_Property.Item("AMRLIMIT2")) & vbCrLf
        SQL = SQL & ",AMRLIMIT3     =  " & STS(Test_Property.Item("AMRLIMIT3")) & vbCrLf
        SQL = SQL & ",AMRLIMIT4     =  " & STS(Test_Property.Item("AMRLIMIT4")) & vbCrLf
        SQL = SQL & ",AMRLIMIT5     =  " & STS(Test_Property.Item("AMRLIMIT5")) & vbCrLf
        SQL = SQL & ",AMRLIMIT6     =  " & STS(Test_Property.Item("AMRLIMIT6")) & vbCrLf
        SQL = SQL & ",AMRLIMIT7     =  " & STS(Test_Property.Item("AMRLIMIT7")) & vbCrLf
        SQL = SQL & ",AMRRESULT1    =  " & STS(Test_Property.Item("AMRRESULT1")) & vbCrLf
        SQL = SQL & ",AMRRESULT2    =  " & STS(Test_Property.Item("AMRRESULT2")) & vbCrLf
        SQL = SQL & ",AMRRESULT3    =  " & STS(Test_Property.Item("AMRRESULT3")) & vbCrLf
        SQL = SQL & ",AMRRESULT4    =  " & STS(Test_Property.Item("AMRRESULT4")) & vbCrLf
        SQL = SQL & ",AMRRESULT5    =  " & STS(Test_Property.Item("AMRRESULT5")) & vbCrLf
        SQL = SQL & ",AMRRESULT6    =  " & STS(Test_Property.Item("AMRRESULT6")) & vbCrLf
        SQL = SQL & ",AMRRESULT7    =  " & STS(Test_Property.Item("AMRRESULT7")) & vbCrLf
        SQL = SQL & ",AMRINRESULT   =  " & STS(Test_Property.Item("AMRINRESULT")) & vbCrLf
        
        SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
        SQL = SQL & "   AND SEQNO       = " & Test_Property.Item("SEQ") & vbCrLf
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
        SQL = SQL & "   AND TESTCODE    = " & STS(Test_Property.Item("TESTCD")) & vbCrLf
                          
    Else
        SQL = ""
        SQL = SQL & "INSERT INTO EQPMASTER " & vbCrLf
        SQL = SQL & "(EQUIPCD, SEQNO, SENDCHANNEL, RSLTCHANNEL"
        SQL = SQL & " , TESTCODE, TESTNAME, ABBRNAME, RESPRECUSE, RESPREC "
        SQL = SQL & " , REFMLOW, REFMHIGH, REFFLOW, REFFHIGH,RESTYPE " & vbCrLf
        '-- AMR
        SQL = SQL & " , AMRLimit1, AMRLimit2, AMRLimit3, AMRLimit4, AMRLimit5, AMRLimit6, AMRLimit7        " & vbCrLf
        SQL = SQL & " , AMRResult1, AMRResult2, AMRResult3, AMRResult4, AMRResult5, AMRResult6, AMRResult7   " & vbCrLf
        SQL = SQL & " , AMRINResult)                                              " & vbCrLf
        
        SQL = SQL & " VALUES (" & vbCrLf
        SQL = SQL & STS(Test_Property.Item("EQPCD"))
        SQL = SQL & "," & Test_Property.Item("SEQ")
        SQL = SQL & "," & STS(Test_Property.Item("OCH"))
        SQL = SQL & "," & STS(Test_Property.Item("RCH"))
        SQL = SQL & "," & STS(Test_Property.Item("TESTCD"))
        SQL = SQL & "," & STS(Test_Property.Item("TESTNM"))
        SQL = SQL & "," & STS(Test_Property.Item("ABBRNM"))
        SQL = SQL & "," & Test_Property.Item("RESUSE")
        SQL = SQL & "," & Test_Property.Item("RES")
        SQL = SQL & "," & STS(Test_Property.Item("REFML"))
        SQL = SQL & "," & STS(Test_Property.Item("REFMH"))
        SQL = SQL & "," & STS(Test_Property.Item("REFFL"))
        SQL = SQL & "," & STS(Test_Property.Item("REFFH"))
        SQL = SQL & "," & STS(Test_Property.Item("USERESULT"))
        '-- AMR
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT1"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT2"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT3"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT4"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT5"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT6"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT7"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT1"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT2"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT3"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT4"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT5"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT6"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT7"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRINRESULT"))
        SQL = SQL & ")" & vbCrLf
    End If
    Call DBExec(AdoCn_Local, SQL)

    LetTestInfo = True
Exit Function

ErrorTrap:
    LetTestInfo = False

End Function

'검사결과변환을 저장
Public Function LetAMRInfo(ByVal Test_Property As Scripting.Dictionary) As Boolean
    
On Error GoTo ErrorTrap
    
    SQL = ""
    SQL = SQL & "SELECT EQUIPCD FROM AMRMASTER " & vbCrLf
    SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
    SQL = SQL & "   AND SEQNO       = " & Test_Property.Item("SEQ") & vbCrLf
    SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
    SQL = SQL & "   AND TESTCODE    = " & STS(Test_Property.Item("TESTCD")) & vbCrLf
    
    If DBExists(AdoCn, SQL) Then
        SQL = ""
        SQL = SQL & "UPDATE AMRMASTER SET " & vbCrLf
        SQL = SQL & " AMRLIMIT8     =  " & STS(Test_Property.Item("AMRLIMIT8")) & vbCrLf
        SQL = SQL & ",AMRLIMIT9     =  " & STS(Test_Property.Item("AMRLIMIT9")) & vbCrLf
        SQL = SQL & ",AMRLIMIT10    =  " & STS(Test_Property.Item("AMRLIMIT10")) & vbCrLf
        SQL = SQL & ",AMRLIMIT11    =  " & STS(Test_Property.Item("AMRLIMIT11")) & vbCrLf
        SQL = SQL & ",AMRLIMIT12    =  " & STS(Test_Property.Item("AMRLIMIT12")) & vbCrLf
        SQL = SQL & ",AMRLIMIT13    =  " & STS(Test_Property.Item("AMRLIMIT13")) & vbCrLf
        SQL = SQL & ",AMRLIMIT14    =  " & STS(Test_Property.Item("AMRLIMIT14")) & vbCrLf
        SQL = SQL & ",AMRLIMIT15    =  " & STS(Test_Property.Item("AMRLIMIT15")) & vbCrLf
        SQL = SQL & ",AMRLIMIT16    =  " & STS(Test_Property.Item("AMRLIMIT16")) & vbCrLf
        SQL = SQL & ",AMRLIMIT17    =  " & STS(Test_Property.Item("AMRLIMIT17")) & vbCrLf
        SQL = SQL & ",AMRLIMIT18    =  " & STS(Test_Property.Item("AMRLIMIT18")) & vbCrLf
        SQL = SQL & ",AMRLIMIT19    =  " & STS(Test_Property.Item("AMRLIMIT19")) & vbCrLf
        SQL = SQL & ",AMRRESULT8    =  " & STS(Test_Property.Item("AMRRESULT8")) & vbCrLf
        SQL = SQL & ",AMRRESULT9    =  " & STS(Test_Property.Item("AMRRESULT9")) & vbCrLf
        SQL = SQL & ",AMRRESULT10   =  " & STS(Test_Property.Item("AMRRESULT10")) & vbCrLf
        SQL = SQL & ",AMRRESULT11   =  " & STS(Test_Property.Item("AMRRESULT11")) & vbCrLf
        SQL = SQL & ",AMRRESULT12   =  " & STS(Test_Property.Item("AMRRESULT12")) & vbCrLf
        SQL = SQL & ",AMRRESULT13   =  " & STS(Test_Property.Item("AMRRESULT13")) & vbCrLf
        SQL = SQL & ",AMRRESULT14   =  " & STS(Test_Property.Item("AMRRESULT14")) & vbCrLf
        SQL = SQL & ",AMRRESULT15   =  " & STS(Test_Property.Item("AMRRESULT15")) & vbCrLf
        SQL = SQL & ",AMRRESULT16   =  " & STS(Test_Property.Item("AMRRESULT16")) & vbCrLf
        SQL = SQL & ",AMRRESULT17   =  " & STS(Test_Property.Item("AMRRESULT17")) & vbCrLf
        SQL = SQL & ",AMRRESULT18   =  " & STS(Test_Property.Item("AMRRESULT18")) & vbCrLf
        SQL = SQL & ",AMRRESULT19   =  " & STS(Test_Property.Item("AMRRESULT19")) & vbCrLf
        SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
        SQL = SQL & "   AND SEQNO       = " & Test_Property.Item("SEQ") & vbCrLf
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
        SQL = SQL & "   AND TESTCODE    = " & STS(Test_Property.Item("TESTCD")) & vbCrLf
    Else
        SQL = ""
        SQL = SQL & "INSERT INTO AMRMASTER " & vbCrLf
        SQL = SQL & "(EQUIPCD, SEQNO, RSLTCHANNEL, TESTCODE "
        SQL = SQL & " , AMRLimit8, AMRLimit9, AMRLimit10, AMRLimit11, AMRLimit12, AMRLimit13, AMRLimit14            " & vbCrLf
        SQL = SQL & " , AMRLimit15, AMRLimit16, AMRLimit17, AMRLimit18, AMRLimit19                                  " & vbCrLf
        SQL = SQL & " , AMRResult8, AMRResult9, AMRResult10, AMRResult11, AMRResult12, AMRResult13, AMRResult14     " & vbCrLf
        SQL = SQL & " , AMRResult15, AMRResult16, AMRResult17, AMRResult18, AMRResult19)                            " & vbCrLf
        SQL = SQL & " VALUES (" & vbCrLf
        SQL = SQL & STS(Test_Property.Item("EQPCD"))
        SQL = SQL & "," & Test_Property.Item("SEQ")
        SQL = SQL & "," & STS(Test_Property.Item("RCH"))
        SQL = SQL & "," & STS(Test_Property.Item("TESTCD"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT8"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT9"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT10"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT11"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT12"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT13"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT14"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT15"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT16"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT17"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT18"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRLIMIT19"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT8"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT9"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT10"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT11"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT12"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT13"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT14"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT15"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT16"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT17"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT18"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRRESULT19"))
        SQL = SQL & ")" & vbCrLf
    End If
    Call DBExec(AdoCn_Local, SQL)

    LetAMRInfo = True
Exit Function

ErrorTrap:
    LetAMRInfo = False

End Function


'검사항목을 삭제
Public Function DelTestInfo(ByVal Test_Property As Scripting.Dictionary) As Boolean
    
On Error GoTo ErrorTrap
    
    SQL = ""
    SQL = SQL & "DELETE FROM EQPMASTER " & vbNewLine
    SQL = SQL & " WHERE EQUIPCD = " & STS(Test_Property.Item("EQPCD")) & vbNewLine
    
    If Test_Property.Item("OCH") <> "" Then
        SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH")) & vbNewLine
    End If
    
    If Test_Property.Item("RCH") <> "" Then
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbNewLine
    End If
    
    If Test_Property.Item("TESTCD") <> "" Then
        SQL = SQL & "   AND TESTCODE = " & STS(Test_Property.Item("TESTCD"))
    End If
    
    Call DBExec(AdoCn_Local, SQL)
    
    'AMR 지우기
    SQL = ""
    SQL = SQL & "DELETE FROM AMRMASTER " & vbNewLine
    SQL = SQL & " WHERE EQUIPCD = " & STS(Test_Property.Item("EQPCD")) & vbNewLine
    SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbNewLine
    SQL = SQL & "   AND TESTCODE = " & STS(Test_Property.Item("TESTCD"))
    
    Call DBExec(AdoCn_Local, SQL)
    
    DelTestInfo = True
Exit Function

ErrorTrap:
    DelTestInfo = False

End Function


