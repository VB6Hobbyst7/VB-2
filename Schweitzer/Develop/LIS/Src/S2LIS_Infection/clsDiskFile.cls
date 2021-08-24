VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiskFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_FileName As String
Private m_FileDesc As String

Public Property Let FileDesc(s As String)
    m_FileDesc = s
End Property

Public Property Get FileDesc() As String
    FileDesc = m_FileDesc
End Property

'  I.E.  c:\windows\cloulds.jpg is clouds.jpg
Public Property Get File() As String
    File = ParseFile()
End Property

Public Property Let FileName(s As String)
    m_FileName = s
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Get FileSize() As Long
    FileSize = FileLen(m_FileName)
End Property

Public Property Get FileExt() As String
    FileExt = ParseExt()
End Property

Public Property Get Directory() As String
    Directory = ParseDir()
End Property

Public Property Get FileDate() As Date
    FileDate = FileDateTime(m_FileName)
End Property

Private Function ParseFile() As String
Dim N As Integer
    ParseFile = ""
    For N = Len(m_FileName) To 1 Step -1
       If Mid(m_FileName, N, 1) = "\" Then
          ParseFile = Right(m_FileName, Len(m_FileName) - N)
          N = -1
      End If
    Next N
    '
End Function

Private Function ParseDir() As String
Dim N As Integer
    '
    ParseDir = ""
    For N = Len(m_FileName) To 1 Step -1
       If Mid(m_FileName, N, 1) = "\" Then
           ParseDir = Left(m_FileName, N)
          N = -1
       End If
    Next N
    '
End Function

Private Function ParseExt() As String
'FILE PATH중 확장자를 가져온다.
Dim N As Integer
   '
   ParseExt = "(N/A)"
   For N = Len(m_FileName) To 1 Step -1
       If Mid(m_FileName, N, 1) = "." Then
          ParseExt = Right(m_FileName, Len(m_FileName) - N)
          N = -1
       End If
   Next N
   '
End Function

Public Sub CopyFile(NewLocation)
'파일을 카피한다.
    FileCopy m_FileName, NewLocation + ParseFile
End Sub

Public Sub DeleteFile()
'FILE을 삭제한다.
    Kill m_FileName
End Sub

Private Function ParsePath(strPath As String, lngPart As opgParsePath) As String
'FILE PATH중 파일,패스,드라이브부분을 읽어낸다.
Dim lngPos              As Long
Dim strPart             As String
Dim blnIncludesFile     As Boolean
    lngPos = InStrRev(strPath, "\")
    blnIncludesFile = InStrRev(strPath, ".") > lngPos
    '
    If lngPos > 0 Then
        Select Case lngPart
            Case opgParsePath.FILE_ONLY
                If blnIncludesFile Then
                    strPart = Right$(strPath, Len(strPath) - lngPos)
                Else
                    strPart = ""
                End If
            Case opgParsePath.PATH_ONLY
                If blnIncludesFile Then
                    strPart = Left$(strPath, lngPos)
                Else
                    strPart = strPath
                End If
            Case opgParsePath.DRIVE_ONLY
                strPart = Left$(strPath, 3)
            Case opgParsePath.FILEEXT_ONLY
                If blnIncludesFile Then
                    strPart = Mid(strPath, InStrRev(strPath, ".") + 1, 3)
                Else
                    strPart = ""
                End If
            Case Else
                strPart = ""
        End Select
    End If
    ParsePath = strPart
ParsePath_End:
    Exit Function
End Function

Public Function TrimSpace(strInput As String) As String
'문장중 불필요한 SPACE제거
Dim astrInput()     As String
Dim astrText()      As String
Dim strElement      As String
Dim lngCount        As Long
Dim lngIncr         As Long
    '
    astrInput = Split(strInput)
    ReDim astrText(UBound(astrInput))
    lngIncr = LBound(astrInput)
    For lngCount = LBound(astrInput) To UBound(astrInput)
        strElement = astrInput(lngCount)
        If Len(strElement) > 0 Then
            astrText(lngIncr) = strElement
            lngIncr = lngIncr + 1
        End If
    Next
    '
    ReDim Preserve astrText(LBound(astrText) To lngIncr - 1)
    '
    TrimSpace = Join(astrText)
    '
End Function

Public Function ReplaceWord(strText As String, strFind As String, _
                     strReplace As String) As String
'문자열중 해당하는 패턴의(strFind)의 문자열을 교체한다.
Dim astrText()  As String
Dim lngCount    As Long
    astrText = Split(strText)
    '
    For lngCount = LBound(astrText) To UBound(astrText)
        If astrText(lngCount) Like strFind Then
            astrText(lngCount) = strReplace
        End If
    Next
    ReplaceWord = Join(astrText)
    '
End Function

Private Sub Class_Initialize()
   '
   m_FileName = "Uninitialized"
   m_FileDesc = "N/A"
   '
End Sub

Public Function GetFiles(strPath As String, dctDict As Dictionary, _
   Optional blnRecursive As Boolean) As Boolean
'해당패스와 하위패스의 모든 파일을 리턴한다.
Dim fsoSysObj       As FileSystemObject
Dim fdrFolder       As Folder
Dim fdrSubFolder    As Folder
Dim filFile         As File
Dim objFiles        As Files

   '
   Set fsoSysObj = New FileSystemObject
   '
   On Error Resume Next
   Set fdrFolder = fsoSysObj.GetFolder(strPath)
   Set objFiles = fsoSysObj.GetFile("c\scandisk.log")
   If Err <> 0 Then
       GetFiles = False
       GoTo GetFiles_End
   End If
   On Error GoTo 0
   '
   For Each filFile In objFiles
       dctDict.Add filFile.Path, filFile.Path
   Next filFile
   '
   'For Each filFile In fdrFolder.Files
   '    dctDict.Add filFile.Path, filFile.Path
   'Next filFile

   If blnRecursive Then
       For Each fdrSubFolder In fdrFolder.SubFolders
           GetFiles fdrSubFolder.Path, dctDict, True
       Next fdrSubFolder
   End If
   '
   GetFiles = True
   '
GetFiles_End:
    Exit Function
End Function

Public Sub LogError(strErrFile As String, errX As ErrObject, Optional strProcName As String)
'파일에 에러로그를 기록한다.
Dim fsoSysObj   As FileSystemObject
Dim filFile     As File
Dim txsStream   As TextStream
Dim lngErrNum   As Long
Dim strPath     As String
Dim strErrText  As String
    '
    Set fsoSysObj = New FileSystemObject
    lngErrNum = errX.Number
    strErrText = errX.Description
    errX.Clear
    '
    'strPath = GetTempDir
    If Len(strPath) = 0 Then
        GoTo LogError_End
    End If
    '
    On Error Resume Next
    Set filFile = fsoSysObj.GetFile(strPath & strErrFile)
    If Err <> 0 Then
        Set filFile = fsoSysObj.CreateTextFile(strPath & strErrFile)
    End If
    On Error GoTo 0
    '
    Set txsStream = filFile.OpenAsTextStream(ForAppending)
    With txsStream
        .WriteLine lngErrNum
        .WriteLine strErrText
        If Len(strProcName) > 0 Then .WriteLine strProcName
        .WriteLine Now
        .WriteBlankLines 1
        .Close
    End With
LogError_End:
    Exit Sub
End Sub

Public Function FilterExactMatch(astrItems() As String, _
                          strSearch As String) As String()
'배열중 문자열 배열검색
'astrResult = FilterExactMatch(astrTest, strSearch)
Dim astrFilter()    As String
Dim astrTemp()      As String
Dim lngUpper        As Long
Dim lngLower        As Long
Dim lngIndex        As Long
Dim lngCount        As Long
    '
    astrFilter = Filter(astrItems, strSearch)
    '
    lngUpper = UBound(astrFilter)
    lngLower = LBound(astrFilter)
    '
    ReDim astrTemp(lngLower To lngUpper)
    '
    For lngIndex = lngLower To lngUpper
        If astrFilter(lngIndex) = strSearch Then
            astrTemp(lngCount) = strSearch
            lngCount = lngCount + 1
        End If
    Next lngIndex
    '
    ReDim Preserve astrTemp(lngLower To lngCount - 1)
    '
    FilterExactMatch = astrTemp
End Function

