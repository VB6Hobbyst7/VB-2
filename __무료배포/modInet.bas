Attribute VB_Name = "modInet"
Option Explicit

Private Const CHUNK_SIZE& = 4096&
Private Const CP_UTF8 As Long = 65001
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long


Type XML_Select
'<worklist>
    '<bcno><![CDATA[3010700030]]></bcno>
    '<patnm><![CDATA[박성일]]></patnm>
    '<prgstno><![CDATA[400321-1******]]></prgstno>
    '<pid><![CDATA[000132623]]></pid>
    '<sex><![CDATA[M]]></sex>
    '<age><![CDATA[78]]></age>
    '<spcnm><![CDATA[Throat swab]]></spcnm>
    '<spccd><![CDATA[023]]></spccd>
    '<tclscd><![CDATA[VB6012A]]></tclscd>
    '<spcstat><![CDATA[4]]></spcstat>
    '<rsltstat><![CDATA[-]]></rsltstat>
    '<workno><![CDATA[20181217I20002]]></workno>
    '<testcd><![CDATA[VB6012A]]></testcd>
    '<execprcpuniqno><![CDATA[2002638354]]></execprcpuniqno>
    '<spcacptdt><![CDATA[20181217094414]]></spcacptdt>
    '<prcpdd><![CDATA[20181217]]></prcpdd>
    '<retestyn><![CDATA[N]]></retestyn>
    '<testlrgcd><![CDATA[I]]></testlrgcd>
    '<orddeptcd><![CDATA[NU]]></orddeptcd>
'</worklist>

    BCNO                 As String
    PATNM                As String
    PRGSTNO              As String
    PID                  As String
    SEX                  As String
    AGE                  As String
    SPCNM                As String
    SPCCD                As String
    TCLSCD               As String
    SPCSTAT              As String
    RSLTSTAT             As String
    WORKNO               As String
    TESTCD               As String
    EXECprcpuniqno       As String
    SPCACPTDT            As String
    PRCPDD               As String
    RETESTYN             As String
    TESTLRGCD            As String
    ORDDEPTCD            As String
End Type

Public XmlSelect  As XML_Select


Type XML_SelectS
'<worklist>
    '<bcno><![CDATA[3010700030]]></bcno>
    '<patnm><![CDATA[박성일]]></patnm>
    '<prgstno><![CDATA[400321-1******]]></prgstno>
    '<pid><![CDATA[000132623]]></pid>
    '<sex><![CDATA[M]]></sex>
    '<age><![CDATA[78]]></age>
    '<spcnm><![CDATA[Throat swab]]></spcnm>
    '<spccd><![CDATA[023]]></spccd>
    '<tclscd><![CDATA[VB6012A]]></tclscd>
    '<spcstat><![CDATA[4]]></spcstat>
    '<rsltstat><![CDATA[-]]></rsltstat>
    '<workno><![CDATA[20181217I20002]]></workno>
    '<testcd><![CDATA[VB6012A]]></testcd>
    '<execprcpuniqno><![CDATA[2002638354]]></execprcpuniqno>
    '<spcacptdt><![CDATA[20181217094414]]></spcacptdt>
    '<prcpdd><![CDATA[20181217]]></prcpdd>
    '<retestyn><![CDATA[N]]></retestyn>
    '<testlrgcd><![CDATA[I]]></testlrgcd>
    '<orddeptcd><![CDATA[NU]]></orddeptcd>
'</worklist>

    BCNO()               As String
    PATNM()              As String
    PRGSTNO()            As String
    PID()                As String
    SEX()                As String
    AGE()                As String
    SPCNM()              As String
    SPCCD()              As String
    TCLSCD()             As String
    SPCSTAT()            As String
    RSLTSTAT()           As String
    WORKNO()             As String
    TESTCD()             As String
    EXECprcpuniqno()     As String
    SPCACPTDT()          As String
    PRCPDD()             As String
    RETESTYN()           As String
    TESTLRGCD()          As String
    ORDDEPTCD()          As String
End Type

Public XmlSelectS  As XML_SelectS

Public Function OpenURLWithIE2(ByVal URL As String, ByRef Inet As Inet) As String
     Dim TotBuf() As Byte, ChunkedBuf() As Byte, Converted() As Byte, ni As Long
    
     With Inet
          .Cancel
          .URL = URL
          .Execute , "GET", inputhdrs:="User-agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)" & vbCrLf
          
          Do While .StillExecuting
               DoEvents
          Loop
          
          ChunkedBuf() = .GetChunk(CHUNK_SIZE, icByteArray)
          
          Do While UBound(ChunkedBuf) >= 0
               ni = ni + UBound(ChunkedBuf) + 1
               ReDim Preserve TotBuf(ni - 1)
               RtlMoveMemory TotBuf(ni - UBound(ChunkedBuf) - 1), ChunkedBuf(0), UBound(ChunkedBuf) + 1&
               ChunkedBuf() = .GetChunk(CHUNK_SIZE, icByteArray)
          Loop
     End With
    
     Dim lSize As Long
     lSize = MultiByteToWideChar(CP_UTF8, 0&, TotBuf(0), UBound(TotBuf) + 1&, ByVal 0&, 0&)
    
     ReDim Converted(lSize * 2 - 1)
     MultiByteToWideChar CP_UTF8, 0&, TotBuf(0), UBound(TotBuf) + 1&, Converted(0), lSize
     
     OpenURLWithIE2 = Converted
     
End Function


Public Sub DisplayNode_InfoS(asPath As String, asCnt As Integer)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    Dim i, j            As Integer
    Dim intNodeLen      As Integer
    
On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        MsgBox ("You have error " & myErr.reason)
    Else
        ReDim Preserve XmlSelectS.AGE(asCnt)
        ReDim Preserve XmlSelectS.BCNO(asCnt)
        ReDim Preserve XmlSelectS.EXECprcpuniqno(asCnt)
        ReDim Preserve XmlSelectS.ORDDEPTCD(asCnt)
        ReDim Preserve XmlSelectS.PATNM(asCnt)
        ReDim Preserve XmlSelectS.PID(asCnt)
        ReDim Preserve XmlSelectS.PRCPDD(asCnt)
        ReDim Preserve XmlSelectS.PRGSTNO(asCnt)
        ReDim Preserve XmlSelectS.RETESTYN(asCnt)
        ReDim Preserve XmlSelectS.RSLTSTAT(asCnt)
        ReDim Preserve XmlSelectS.SEX(asCnt)
        ReDim Preserve XmlSelectS.SPCACPTDT(asCnt)
        ReDim Preserve XmlSelectS.SPCCD(asCnt)
        ReDim Preserve XmlSelectS.SPCNM(asCnt)
        ReDim Preserve XmlSelectS.SPCSTAT(asCnt)
        ReDim Preserve XmlSelectS.TCLSCD(asCnt)
        ReDim Preserve XmlSelectS.TESTCD(asCnt)
        ReDim Preserve XmlSelectS.TESTLRGCD(asCnt)
        ReDim Preserve XmlSelectS.WORKNO(asCnt)
            
        '<bcno><![CDATA[3010700030]]></bcno>
        '<patnm><![CDATA[박성일]]></patnm>
        '<prgstno><![CDATA[400321-1******]]></prgstno>
        '<pid><![CDATA[000132623]]></pid>
        '<sex><![CDATA[M]]></sex>
        '<age><![CDATA[78]]></age>
        '<spcnm><![CDATA[Throat swab]]></spcnm>
        '<spccd><![CDATA[023]]></spccd>
        '<tclscd><![CDATA[VB6012A]]></tclscd>
        '<spcstat><![CDATA[4]]></spcstat>
        '<rsltstat><![CDATA[-]]></rsltstat>
        '<workno><![CDATA[20181217I20002]]></workno>
        '<testcd><![CDATA[VB6012A]]></testcd>
        '<execprcpuniqno><![CDATA[2002638354]]></execprcpuniqno>
        '<spcacptdt><![CDATA[20181217094414]]></spcacptdt>
        '<prcpdd><![CDATA[20181217]]></prcpdd>
        '<retestyn><![CDATA[N]]></retestyn>
        '<testlrgcd><![CDATA[I]]></testlrgcd>
        '<orddeptcd><![CDATA[NU]]></orddeptcd>
        
        
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                For intNodeLen = 0 To xNode.childNodes.Length - 1
                    For i = 0 To xNode.childNodes.Item(intNodeLen).childNodes.Length - 1
                        'Debug.Print xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName & ":" & xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue
                        Select Case UCase(xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName)
                            Case "AGE":             XmlSelectS.AGE(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '나이           [78]
                            Case "BCNO":            XmlSelectS.BCNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                '검체번호       [3010700030]
                            Case "EXECprcpuniqno":  XmlSelectS.EXECprcpuniqno(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue      '챠트번호?      [2002638354]
                            Case "ORDDEPTCD":       XmlSelectS.ORDDEPTCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '처방부서코드?  [NU]
                            
                            Case "PATNM":           XmlSelectS.PATNM(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '환자명         [박성일]
                            Case "PID":             XmlSelectS.PID(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '환자번호       [000132623]
                            Case "PRCPDD":          XmlSelectS.PRCPDD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '처방일?        [20181217]
                            Case "PRGSTNO":         XmlSelectS.PRGSTNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '주민번호       [400321-1******]
                            
                            Case "RETESTYN":        XmlSelectS.RETESTYN(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '재검여부       [N]
                            Case "RSLTSTAT":        XmlSelectS.RSLTSTAT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '결과상태       [-]
                            
                            Case "SEX":             XmlSelectS.SEX(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '성별           [M]
                            
                            Case "SPCACPTDT":       XmlSelectS.SPCACPTDT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '검체채취시간?  [20181217094414]
                            Case "SPCCD":           XmlSelectS.SPCCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '검체코드       [023]
                            Case "SPCNM":           XmlSelectS.SPCNM(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '검체명         [Throat swab]
                            Case "SPCSTAT":         XmlSelectS.SPCSTAT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '검체상태       [4]
                            
                            Case "TCLSCD":          XmlSelectS.TCLSCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '처방코드       [VB6012A]
                            Case "TESTCD":          XmlSelectS.TESTCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '검사코드       [VB6012A]
                            Case "TESTLRGCD":       XmlSelectS.TESTLRGCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '결과그룹코드?  [I]
                            Case "WORKNO":          XmlSelectS.WORKNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '워크번호       [20181217I20002]
                        End Select
                    Next
                    j = j + 1
                Next
            End If
        Next
       
        Set Child_Node = Nothing
        
    End If

    Exit Sub
    
ErrXML:
    Exit Sub
    
End Sub
