Attribute VB_Name = "Library_XML"
Option Explicit

Type WorkList
    
    pid         As String
    kornm       As String
    bcno        As String
    barcode     As String
    deptname    As String
    OK          As Integer
    HJList      As String
    ExamName    As String
End Type

Public gWorkList() As WorkList
Public giIndex As Integer
Public gStrXML As String

Public Const INTERNET_FLAG_ASYNC = &H10000000
Public Const INTERNET_SUCCESS  As Long = 1
Public Const INTERNET_OPEN_TYPE_PRECONFIG  As Long = 0
Public Const INTERNET_OPEN_TYPE_DIRECT  As Long = 1
Public Const INTERNET_OPEN_TYPE_PROXY  As Long = 3
Public Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY  As Long = 4
Public Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet" (ByVal hEnumHandle As Long) As Long

Private Function DB_WebSend(RunstrUrl As String) As String

    Dim strBuffer As String * 1000
    Dim lngHandle1 As Long
    Dim lngHandle2 As Long
    Dim intRetVal As Long
    Dim intRetSize As Long
    Dim strOut As String

    DB_WebSend = ""

'On Error GoTo dbWebSendError

    If RunstrUrl = "" Then Exit Function
    DoEvents

    lngHandle1 = InternetOpen("Web Setting DbUpdate", INTERNET_OPEN_TYPE_DIRECT, "", "", 0)
    lngHandle2 = InternetOpenUrl(lngHandle1, RunstrUrl, "", 0, INTERNET_FLAG_ASYNC, 0)

    Do While True
        DoEvents
        strBuffer = Space(1000)

        intRetVal = InternetReadFile(lngHandle2, strBuffer, 1000, intRetSize)
        If intRetSize = 0 Then
            Exit Do
        End If

        strOut = strOut & strBuffer
    Loop
    InternetCloseHandle (lngHandle2)
    InternetCloseHandle (lngHandle1)
    If Not strOut = "" Then
        DB_WebSend = strOut
    End If
dbWebSendError:

End Function

Public Sub URLstart(ByVal URL As String)
Dim rtn As String
gStrXML = ""
rtn = DB_WebSend(URL)
gStrXML = Trim(rtn)

End Sub




Public Sub Clear_Worklist()
    giIndex = -1
    ReDim gWorkList(0)
End Sub

Public Function Online_Sch(ByVal asStr As String) As String

    Dim sRetStr As String
    Dim sFilename As String
    Dim sParam As String
    
    Online_Sch = ""

    sRetStr = asStr
    

    Dim xDoc As MSXML2.DOMDocument
    Set xDoc = New MSXML2.DOMDocument
    If xDoc.Load(asStr) Then
        ' Data Load, Start Parsing
        display_online_parsing_Barcode xDoc.childNodes, 0

    Else
        ' 문서를 로드하지 못했습니다.
        Dim strErrText As String
        Dim xPE As MSXML2.IXMLDOMParseError
       ' ParseError 개체를 가져옵니다
        Set xPE = xDoc.parseError
        With xPE
    
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL

        End With

'''        SaveXML_Data strErrText
    End If
    

    
    Set xPE = Nothing
    Set xDoc = Nothing

End Function

Public Sub display_online_parsing_Barcode(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
    ByVal Indent As Integer)
'''pid, kornm, bcno, barcode, deptname
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2
    
    For Each xNode In Nodes
        If Trim(xNode.parentNode.nodeName) = "ifordcd" Then
            Save_Raw_Data xNode.childNodes.nextNode.nodeValue
            
        End If
        
        If xNode.hasChildNodes Then
            If Trim(xNode.parentNode.nodeName) <> "ifordcd" Then
                display_online_parsing_Barcode xNode.childNodes, Indent
            End If
        End If
        
    Next
   
End Sub

Public Function Online_Param(ByVal asStr As String) As String

    Dim sRetStr As String
    Dim sFilename As String
    Dim sParam As String
    
    Online_Param = ""
    sFilename = "List"
    
    sRetStr = asStr
    
    'SaveXMLFile sRetStr
    Xml_Log sRetStr, sFilename
    
'    Xml_Log sRetStr, asProc
    Clear_Worklist
    
    
    Dim xDoc As MSXML2.DOMDocument
    Set xDoc = New MSXML2.DOMDocument
    If xDoc.Load(App.Path & "\XML\" & sFilename & ".xml") Then
        ' Data Load, Start Parsing




        display_online_parsing_Rece xDoc.childNodes, 0

    Else
        ' 문서를 로드하지 못했습니다.
        Dim strErrText As String
        Dim xPE As MSXML2.IXMLDOMParseError
       ' ParseError 개체를 가져옵니다
        Set xPE = xDoc.parseError
        With xPE
    
            strErrText = "Your XML Document failed to load" & _
                         "due the following error." & vbCrLf & _
                         "Error #: " & .errorCode & ": " & xPE.reason & _
                         "Line #: " & .Line & vbCrLf & _
                         "Line Position: " & .linepos & vbCrLf & _
                         "Position In File: " & .filepos & vbCrLf & _
                         "Source Text: " & .srcText & vbCrLf & _
                         "Document URL: " & .URL

        End With

'        SaveXML_Data strErrText
    End If
    
'''    Xml_Log Online_Param, "TLA"
    
    
    Set xPE = Nothing
    Set xDoc = Nothing
    
'''    If InStr(1, gOnline_Ret, vbTab) > 0 Then
'''        Online_Param = Left(gOnline_Ret, InStr(1, gOnline_Ret, vbTab) - 1)
'''    End If
    
End Function

Public Sub display_online_parsing_Rece(ByRef Nodes As MSXML2.IXMLDOMNodeList, _
    ByVal Indent As Integer)
'''pid, kornm, bcno, barcode, deptname
    Dim xNode As MSXML2.IXMLDOMNode
    Indent = Indent + 2

    For Each xNode In Nodes
            
'''        If xNode.nodeType = 4 Then
            Select Case Trim(xNode.parentNode.nodeName)
''            Case "pid"
''                giIndex = giIndex + 1
''                ReDim Preserve gWorkList(giIndex)
''                gWorkList(giIndex).pid = xNode.nodeValue
''            Case "kornm"
''                ReDim Preserve gWorkList(giIndex)
''                gWorkList(giIndex).kornm = xNode.nodeValue
''            Case "bcno"
''                ReDim Preserve gWorkList(giIndex)
''                gWorkList(giIndex).bcno = xNode.nodeValue
''            Case "barcode"
''                ReDim Preserve gWorkList(giIndex)
''                gWorkList(giIndex).barcode = xNode.nodeValue
''            Case "deptname"
''                ReDim Preserve gWorkList(giIndex)
''                gWorkList(giIndex).deptname = xNode.nodeValue
            Case "HJList"
                giIndex = giIndex + 1
                ReDim Preserve gWorkList(giIndex)
                gWorkList(giIndex).HJList = xNode.nodeValue
            End Select

'''        End If
        If xNode.hasChildNodes Then
            display_online_parsing_Rece xNode.childNodes, Indent
        End If
    Next xNode
End Sub

Public Sub Xml_Log(argSQL As String, argFileName As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFilename As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\" & "XML", vbDirectory) <> "XML" Then
        MkDir (App.Path & "\" & "XML")
    End If
    
    sFilename = argFileName
    If Dir(App.Path & "\" & "XML" & "\" & sFilename & ".xml") <> "" Then
        Kill App.Path & "\" & "XML" & "\" & sFilename & ".xml"
    End If
    
    Open App.Path & "\" & "XML" & "\" & sFilename & ".xml" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub


