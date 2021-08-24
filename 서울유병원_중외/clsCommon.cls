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
Public Function LetEqpInfo(ByVal Test_Property As Scripting.Dictionary) As Boolean
    
On Error GoTo ErrorTrap
    
    SQL = ""
    SQL = SQL & "SELECT EQUIPCD FROM EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
    SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
    'SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH")) & vbCrLf
    'SQL = SQL & "   AND SEQNO       = " & Test_Property.Item("SEQ") & vbCrLf
    
    If DBExists(AdoCn, SQL) Then
        SQL = ""
        SQL = SQL & "UPDATE EQPMASTER SET " & vbCrLf
        SQL = SQL & "  SEQNO = " & Test_Property.Item("SEQ") & vbCrLf
        SQL = SQL & ", SENDCHANNEL   = " & STS(Test_Property.Item("OCH")) & vbCrLf
        SQL = SQL & ", RSLTCHANNEL   = " & STS(Test_Property.Item("RCH")) & vbCrLf
        SQL = SQL & ", TESTNAME      =  " & STS(Test_Property.Item("TESTNM")) & vbCrLf
        SQL = SQL & ", ABBRNAME      =  " & STS(Test_Property.Item("ABBRNM")) & vbCrLf
        SQL = SQL & ", RESPRECUSE    =  " & STS(Test_Property.Item("RESUSE")) & vbCrLf
        SQL = SQL & ", RESPREC       =  " & STS(Test_Property.Item("RES")) & vbCrLf
        SQL = SQL & ", REFMLOW       =  " & STS(Test_Property.Item("REFML")) & vbCrLf
        SQL = SQL & ", REFMHIGH      =  " & STS(Test_Property.Item("REFMH")) & vbCrLf
        SQL = SQL & ", REFFLOW       =  " & STS(Test_Property.Item("REFFL")) & vbCrLf
        SQL = SQL & ", REFFHIGH      =  " & STS(Test_Property.Item("REFFH")) & vbCrLf
        SQL = SQL & ", RESTYPE       =  " & STS(Test_Property.Item("USERESULT")) & vbCrLf
        SQL = SQL & ", CALYN         =  " & STS(Test_Property.Item("CALYN")) & vbCrLf
        SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
'        SQL = SQL & "   AND SEQNO       = " & Test_Property.Item("SEQ") & vbCrLf
'        SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH")) & vbCrLf
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
                          
    Else
        SQL = ""
        SQL = SQL & "INSERT INTO EQPMASTER " & vbCrLf
        SQL = SQL & "(EQUIPCD, RSLTCHANNEL, SENDCHANNEL, SEQNO"
        SQL = SQL & " , TESTNAME, ABBRNAME, RESPRECUSE, RESPREC "
        SQL = SQL & " , REFMLOW, REFMHIGH, REFFLOW, REFFHIGH,RESTYPE, CALYN)" & vbCrLf
        SQL = SQL & " VALUES (" & vbCrLf
        SQL = SQL & STS(Test_Property.Item("EQPCD"))
        SQL = SQL & "," & STS(Test_Property.Item("RCH"))
        SQL = SQL & "," & STS(Test_Property.Item("OCH"))
        SQL = SQL & "," & Test_Property.Item("SEQ")
        SQL = SQL & "," & STS(Test_Property.Item("TESTNM"))
        SQL = SQL & "," & STS(Test_Property.Item("ABBRNM"))
        SQL = SQL & "," & STS(Test_Property.Item("RESUSE"))
        SQL = SQL & "," & STS(Test_Property.Item("RES"))
        SQL = SQL & "," & STS(Test_Property.Item("REFML"))
        SQL = SQL & "," & STS(Test_Property.Item("REFMH"))
        SQL = SQL & "," & STS(Test_Property.Item("REFFL"))
        SQL = SQL & "," & STS(Test_Property.Item("REFFH"))
        SQL = SQL & "," & STS(Test_Property.Item("USERESULT"))
        SQL = SQL & "," & STS(Test_Property.Item("CALYN"))
        SQL = SQL & ")" & vbCrLf
    End If
    Call DBExec(AdoCn_Local, SQL)

    LetEqpInfo = True
Exit Function

ErrorTrap:
    LetEqpInfo = False

End Function

'검사항목을 저장
Public Function LetTestInfo(ByVal Test_Property As Scripting.Dictionary) As Boolean
    Dim varTestCode As Variant
    Dim i           As Integer
    
On Error GoTo ErrorTrap
    
    SQL = ""
    SQL = SQL & "SELECT RSLTCHANNEL FROM TESTMASTER " & vbCrLf
    SQL = SQL & " WHERE RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
    
    If DBExists(AdoCn, SQL) Then
        SQL = ""
        SQL = SQL & "DELETE FROM TESTMASTER " & vbCrLf
        SQL = SQL & " WHERE RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
    
        Call DBExec(AdoCn_Local, SQL)
    
    End If
    
    varTestCode = Split(Test_Property.Item("TESTCD"), "#")
        
    For i = 0 To UBound(varTestCode) - 1
        SQL = ""
        SQL = SQL & "INSERT INTO TESTMASTER         " & vbCrLf
        SQL = SQL & "(RSLTCHANNEL, SEQNO, TESTCODE, CALCULATE) " & vbCrLf
        SQL = SQL & " VALUES (" & vbCrLf
        SQL = SQL & STS(Test_Property.Item("RCH"))
        SQL = SQL & "," & i
        SQL = SQL & "," & STS(varTestCode(i))
        
        If Test_Property.Item("TESTCALCD") = varTestCode(i) Then
            SQL = SQL & "," & STS(Test_Property.Item("CALCULATE"))
        Else
            SQL = SQL & ",''"
        End If
        
        SQL = SQL & ")" & vbCrLf
        
        Call DBExec(AdoCn_Local, SQL)
    Next

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
    SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
    
    If DBExists(AdoCn, SQL) Then
        SQL = ""
        SQL = SQL & "UPDATE AMRMASTER SET " & vbCrLf
        SQL = SQL & "  AMRINRESULT      =  " & STS(Test_Property.Item("AMRINRESULT")) & vbCrLf
        SQL = SQL & ", AMRNUMLIMIT1     =  " & STS(Test_Property.Item("AMRNUMLIMIT1")) & vbCrLf
        SQL = SQL & ", AMRNUMLIMIT2     =  " & STS(Test_Property.Item("AMRNUMLIMIT2")) & vbCrLf
        SQL = SQL & ", AMRNUMLIMIT3     =  " & STS(Test_Property.Item("AMRNUMLIMIT3")) & vbCrLf
        SQL = SQL & ", AMRNUMLIMIT4     =  " & STS(Test_Property.Item("AMRNUMLIMIT4")) & vbCrLf
        SQL = SQL & ", AMRNUMLIMIT5     =  " & STS(Test_Property.Item("AMRNUMLIMIT5")) & vbCrLf
        SQL = SQL & ", AMRNUMRESULT1    =  " & STS(Test_Property.Item("AMRNUMRESULT1")) & vbCrLf
        SQL = SQL & ", AMRNUMRESULT2    =  " & STS(Test_Property.Item("AMRNUMRESULT2")) & vbCrLf
        SQL = SQL & ", AMRNUMRESULT3    =  " & STS(Test_Property.Item("AMRNUMRESULT3")) & vbCrLf
        SQL = SQL & ", AMRNUMRESULT4    =  " & STS(Test_Property.Item("AMRNUMRESULT4")) & vbCrLf
        SQL = SQL & ", AMRNUMRESULT5    =  " & STS(Test_Property.Item("AMRNUMRESULT5")) & vbCrLf
        SQL = SQL & ", AMRCHARLIMIT1    =  " & STS(Test_Property.Item("AMRCHARLIMIT1")) & vbCrLf
        SQL = SQL & ", AMRCHARLIMIT2    =  " & STS(Test_Property.Item("AMRCHARLIMIT2")) & vbCrLf
        SQL = SQL & ", AMRCHARLIMIT3    =  " & STS(Test_Property.Item("AMRCHARLIMIT3")) & vbCrLf
        SQL = SQL & ", AMRCHARLIMIT4    =  " & STS(Test_Property.Item("AMRCHARLIMIT4")) & vbCrLf
        SQL = SQL & ", AMRCHARLIMIT5    =  " & STS(Test_Property.Item("AMRCHARLIMIT5")) & vbCrLf
        SQL = SQL & ", AMRCHARRESULT1   =  " & STS(Test_Property.Item("AMRCHARRESULT1")) & vbCrLf
        SQL = SQL & ", AMRCHARRESULT2   =  " & STS(Test_Property.Item("AMRCHARRESULT2")) & vbCrLf
        SQL = SQL & ", AMRCHARRESULT3   =  " & STS(Test_Property.Item("AMRCHARRESULT3")) & vbCrLf
        SQL = SQL & ", AMRCHARRESULT4   =  " & STS(Test_Property.Item("AMRCHARRESULT4")) & vbCrLf
        SQL = SQL & ", AMRCHARRESULT5   =  " & STS(Test_Property.Item("AMRCHARRESULT5")) & vbCrLf
        SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
    Else
        SQL = ""
        SQL = SQL & "INSERT INTO AMRMASTER " & vbCrLf
        SQL = SQL & "(  EQUIPCD,        RSLTCHANNEL,    AMRINResult                                     " & vbCrLf
        SQL = SQL & " , AMRNUMLIMIT1,   AMRNUMLIMIT2,   AMRNUMLIMIT3,   AMRNUMLIMIT4,   AMRNUMLIMIT5    " & vbCrLf
        SQL = SQL & " , AMRNUMRESULT1,  AMRNUMRESULT2,  AMRNUMRESULT3,  AMRNUMRESULT4,  AMRNUMRESULT5   " & vbCrLf
        SQL = SQL & " , AMRCHARLIMIT1,  AMRCHARLIMIT2,  AMRCHARLIMIT3,  AMRCHARLIMIT4,  AMRCHARLIMIT5   " & vbCrLf
        SQL = SQL & " , AMRCHARRESULT1, AMRCHARRESULT2, AMRCHARRESULT3, AMRCHARRESULT4, AMRCHARRESULT5) " & vbCrLf
        SQL = SQL & " VALUES (" & vbCrLf
        SQL = SQL & STS(Test_Property.Item("EQPCD"))
        SQL = SQL & "," & STS(Test_Property.Item("RCH"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRINRESULT"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMLIMIT1"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMLIMIT2"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMLIMIT3"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMLIMIT4"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMLIMIT5"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMRESULT1"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMRESULT2"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMRESULT3"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMRESULT4"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRNUMRESULT5"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARLIMIT1"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARLIMIT2"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARLIMIT3"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARLIMIT4"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARLIMIT5"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARRESULT1"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARRESULT2"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARRESULT3"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARRESULT4"))
        SQL = SQL & "," & STS(Test_Property.Item("AMRCHARRESULT5"))
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
    
    If Test_Property.Item("SEQ") <> "" Then
        SQL = SQL & "   AND SEQNO = " & Test_Property.Item("SEQ") & vbNewLine
    End If
    
    If Test_Property.Item("OCH") <> "" Then
        SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH")) & vbNewLine
    End If
    
    If Test_Property.Item("RCH") <> "" Then
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbNewLine
    End If
    
'    If Test_Property.Item("TESTCD") <> "" Then
'        SQL = SQL & "   AND TESTCODE = " & STS(Test_Property.Item("TESTCD"))
'    End If
    
    If Test_Property.Item("TESTNM") <> "" Then
        SQL = SQL & "   AND TESTNAME = " & STS(Test_Property.Item("TESTNM"))
    End If
    
    If Test_Property.Item("ABBRNM") <> "" Then
        SQL = SQL & "   AND ABBRNAME = " & STS(Test_Property.Item("ABBRNM"))
    End If
    
    Call DBExec(AdoCn_Local, SQL)
    
    'AMR MASTER 지우기
    SQL = ""
    SQL = SQL & "DELETE FROM AMRMASTER " & vbNewLine
    SQL = SQL & " WHERE EQUIPCD = " & STS(Test_Property.Item("EQPCD")) & vbNewLine
    SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbNewLine
    'SQL = SQL & "   AND TESTCODE = " & STS(Test_Property.Item("TESTCD"))
    
    Call DBExec(AdoCn_Local, SQL)
    
    'TEST MASTER 지우기
    SQL = ""
    SQL = SQL & "DELETE FROM TESTMASTER " & vbNewLine
    SQL = SQL & " WHERE RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbNewLine
    
    Call DBExec(AdoCn_Local, SQL)
    
    DelTestInfo = True
Exit Function

ErrorTrap:
    DelTestInfo = False

End Function


