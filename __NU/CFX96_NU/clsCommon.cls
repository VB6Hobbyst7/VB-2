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

'�˻��׸��� ����
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
        SQL = SQL & ",REFMLOW        =  " & STS(Test_Property.Item("REFML")) & vbCrLf
        SQL = SQL & ",REFMHIGH      =  " & STS(Test_Property.Item("REFMH")) & vbCrLf
        SQL = SQL & ",REFFLOW      =  " & STS(Test_Property.Item("REFFL")) & vbCrLf
        SQL = SQL & ",REFFHIGH      =  " & STS(Test_Property.Item("REFFH")) & vbCrLf
        SQL = SQL & " WHERE EQUIPCD     = " & STS(Test_Property.Item("EQPCD")) & vbCrLf
        SQL = SQL & "   AND SEQNO       = " & Test_Property.Item("SEQ") & vbCrLf
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbCrLf
        SQL = SQL & "   AND TESTCODE    = " & STS(Test_Property.Item("TESTCD")) & vbCrLf
                          
    Else
        SQL = ""
        SQL = SQL & "INSERT INTO EQPMASTER " & vbCrLf
        SQL = SQL & "(EQUIPCD, SEQNO, SENDCHANNEL, RSLTCHANNEL"
        SQL = SQL & " ,TESTCODE, TESTNAME, ABBRNAME, RESPRECUSE, RESPREC"
        SQL = SQL & " , REFMLOW, REFMHIGH, REFFLOW, REFFHIGH)" & vbCrLf
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
        SQL = SQL & ")" & vbCrLf
    End If
    Call DBExec(AdoCn_Local, SQL)

    LetTestInfo = True
Exit Function

ErrorTrap:
    LetTestInfo = False

End Function



'�˻��׸��� ����
Public Function DelTestInfo(ByVal Test_Property As Scripting.Dictionary) As Boolean
    
On Error GoTo ErrorTrap
    
    SQL = ""
    SQL = SQL & "DELETE FROM EQPMASTER " & vbNewLine
    SQL = SQL & " WHERE EQUIPCD = " & STS(Test_Property.Item("EQPCD")) & vbNewLine
    If Test_Property.Item("OCH") = "" Then
        SQL = SQL & "   AND (SENDCHANNEL IS NULL OR SENDCHANNEL = '')" & vbNewLine
    Else
        SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH")) & vbNewLine
    End If
    
    If Test_Property.Item("RCH") = "" Then
        SQL = SQL & "   AND (RSLTCHANNEL IS NULL OR RSLTCHANNEL = '')" & vbNewLine
    Else
        SQL = SQL & "   AND RSLTCHANNEL = " & STS(Test_Property.Item("RCH")) & vbNewLine
    End If
    
    If Test_Property.Item("TESTCD") = "" Then
        SQL = SQL & "   AND (TESTCODE IS NULL OR TESTCODE = '')"
    Else
        SQL = SQL & "   AND TESTCODE = " & STS(Test_Property.Item("TESTCD"))
    End If
    Call DBExec(AdoCn_Local, SQL)
    
    DelTestInfo = True
Exit Function

ErrorTrap:
    DelTestInfo = False

End Function

