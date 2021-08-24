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
    SQL = SQL & "SELECT EQUIPCD FROM EQPMASTER " & vbNewLine
    SQL = SQL & " WHERE EQUIPCD = " & STS(Test_Property.Item("EQPCD")) & vbNewLine
    SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH"))
    SQL = SQL & "   AND TESTCODE = " & STS(Test_Property.Item("TESTCD"))
    
    If DBExists(AdoCn, SQL) Then
        SQL = ""
        SQL = SQL & "UPDATE EQPMASTER SET "
        SQL = SQL & " SEQNO = " & Test_Property.Item("SEQ")
        SQL = SQL & ",SENDCHANNEL = " & STS(Test_Property.Item("OCH"))
        SQL = SQL & ",RSLTCHANNEL = " & STS(Test_Property.Item("RCH"))
        SQL = SQL & ",TESTCODE = " & STS(Test_Property.Item("TESTCD"))
        SQL = SQL & ",TESTNAME =  " & STS(Test_Property.Item("TESTNM"))
        SQL = SQL & ",ABBRNAME=  " & STS(Test_Property.Item("ABBRNM"))
        SQL = SQL & ",RESPREC =  " & Test_Property.Item("RES")
        SQL = SQL & ",REFLOW =  " & STS(Test_Property.Item("REFL"))
        SQL = SQL & ",REFHIGH =  " & STS(Test_Property.Item("REFH"))
        SQL = SQL & ",RESULTTYPE = " & STS(Test_Property.Item("RSTTYPE"))
        SQL = SQL & ",CUTOFFUSE = " & STS(Test_Property.Item("CUTUSE"))
        SQL = SQL & ",COLIN = " & STS(Test_Property.Item("COLIN"))
        SQL = SQL & ",COLCOMP = " & STS(Test_Property.Item("COLCP"))
        SQL = SQL & ",COLOUT = " & STS(Test_Property.Item("COLOUT"))
        SQL = SQL & ",COHIN = " & STS(Test_Property.Item("COHIN"))
        SQL = SQL & ",COHCOMP = " & STS(Test_Property.Item("COHCP"))
        SQL = SQL & ",COHOUT = " & STS(Test_Property.Item("COHOUT"))
        SQL = SQL & ",COMOUT = " & STS(Test_Property.Item("COMOUT"))
        '-- QC
        SQL = SQL & ",QCLab = " & STS(Test_Property.Item("LAB"))
        SQL = SQL & ",QCLot = " & STS(Test_Property.Item("LOT"))
        SQL = SQL & ",QCAnalyte = " & STS(Test_Property.Item("ANALYTE"))
        SQL = SQL & ",QCMethod = " & STS(Test_Property.Item("METHOD"))
        SQL = SQL & ",QCInstrument = " & STS(Test_Property.Item("INSTRUMENT"))
        SQL = SQL & ",QCReagent = " & STS(Test_Property.Item("REAGENT"))
        SQL = SQL & ",QCUnit = " & STS(Test_Property.Item("UNIT"))
        SQL = SQL & ",QCTemp = " & STS(Test_Property.Item("TEMP"))
        SQL = SQL & " WHERE EQUIPCD = " & STS(Test_Property.Item("EQPCD"))
        SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH"))
        SQL = SQL & "   AND TESTCODE = " & STS(Test_Property.Item("TESTCD"))
                          
    Else
        SQL = ""
        SQL = SQL & "INSERT INTO EQPMASTER (EQUIPCD,SEQNO,SENDCHANNEL,RSLTCHANNEL,TESTCODE,TESTNAME,ABBRNAME,RESPREC,"
        SQL = SQL & " REFLOW,REFHIGH,RESULTTYPE,CUTOFFUSE,COLIN,COLCOMP,COLOUT,COHIN,COHCOMP,COHOUT,COMOUT,"
        '-- QC
        SQL = SQL & " QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument, QCReagent, QCUnit, QCTemp)"
        SQL = SQL & " VALUES ("
        SQL = SQL & STS(Test_Property.Item("EQPCD"))
        SQL = SQL & "," & Test_Property.Item("SEQ")
        SQL = SQL & "," & STS(Test_Property.Item("OCH"))
        SQL = SQL & "," & STS(Test_Property.Item("RCH"))
        SQL = SQL & "," & STS(Test_Property.Item("TESTCD"))
        SQL = SQL & "," & STS(Test_Property.Item("TESTNM"))
        SQL = SQL & "," & STS(Test_Property.Item("ABBRNM"))
        SQL = SQL & "," & Test_Property.Item("RES")
        SQL = SQL & "," & STS(Test_Property.Item("REFL"))
        SQL = SQL & "," & STS(Test_Property.Item("REFH"))
        SQL = SQL & "," & STS(Test_Property.Item("RSTTYPE"))
        SQL = SQL & "," & STS(Test_Property.Item("CUTUSE"))
        SQL = SQL & "," & STS(Test_Property.Item("COLIN"))
        SQL = SQL & "," & STS(Test_Property.Item("COLCP"))
        SQL = SQL & "," & STS(Test_Property.Item("COLOUT"))
        SQL = SQL & "," & STS(Test_Property.Item("COHIN"))
        SQL = SQL & "," & STS(Test_Property.Item("COHCP"))
        SQL = SQL & "," & STS(Test_Property.Item("COHOUT"))
        SQL = SQL & "," & STS(Test_Property.Item("COMOUT"))
        '-- QC
        SQL = SQL & "," & STS(Test_Property.Item("LAB"))
        SQL = SQL & "," & STS(Test_Property.Item("LOT"))
        SQL = SQL & "," & STS(Test_Property.Item("ANALYTE"))
        SQL = SQL & "," & STS(Test_Property.Item("METHOD"))
        SQL = SQL & "," & STS(Test_Property.Item("INSTRUMENT"))
        SQL = SQL & "," & STS(Test_Property.Item("REAGENT"))
        SQL = SQL & "," & STS(Test_Property.Item("UNIT"))
        SQL = SQL & "," & STS(Test_Property.Item("TEMP"))
        SQL = SQL & ")"
    End If
    Call DBExec(AdoCn_Local, SQL)

    LetTestInfo = True
Exit Function

ErrorTrap:
    LetTestInfo = False

End Function

'검사항목을 삭제
Public Function DelTestInfo(ByVal Test_Property As Scripting.Dictionary) As Boolean
    
On Error GoTo ErrorTrap
    
    SQL = ""
    SQL = SQL & "DELETE FROM EQPMASTER " & vbNewLine
    SQL = SQL & " WHERE EQUIPCD = " & STS(Test_Property.Item("EQPCD")) & vbNewLine
    SQL = SQL & "   AND SENDCHANNEL = " & STS(Test_Property.Item("OCH")) & vbNewLine
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


