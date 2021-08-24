VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSlideImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private RS              As Recordset
Private RS1             As Recordset
Private objDiskFile     As clsDiskFile

Private mvarPtID        As String
Private mvarBusiDir     As String
Private mvarAccSeq      As String
Private mvarServerPath  As String
Private mvarImageDiv    As String
Private mvarLastIndex   As Integer
Private mvarRecordCount As Long

Public Property Get ptid() As String
    ptid = mvarPtID
End Property

Public Property Get BusiDir() As String
    BusiDir = mvarBusiDir
End Property

Public Property Get AccSeq() As String
    AccSeq = mvarAccSeq
End Property

Public Property Get ServerPath() As String
    ServerPath = mvarServerPath
End Property

Public Property Get FileName() As String
    If (RS.RecordCount > 0) Then
       FileName = ""
    End If
    FileName = "" & RS!FileName
End Property

Public Property Get SlideNo() As String
    If (RS.RecordCount > 0) Then
       SlideNo = ""
    End If
    SlideNo = "" & RS!SlideNo
End Property

Public Property Get ImageDiv() As String
    If (RS.RecordCount > 0) Then
       ImageDiv = ""
    End If
    ImageDiv = "" & RS!ImageDiv
    '
End Property

Public Property Get LastIndex() As String
    LastIndex = mvarLastIndex
End Property

Public Property Get EOF() As Boolean
    If (RS.RecordCount > 0) Then
       EOF = False
    End If
    EOF = RS.EOF
End Property

Public Property Get BOF() As Boolean
    If (RS.RecordCount > 0) Then
       BOF = False
    End If
    BOF = RS.BOF
End Property

Public Property Get RecordCount() As Long
   
    If RS Is Nothing Then
       mvarRecordCount = 0
       RecordCount = mvarRecordCount
       Exit Sub
    End If
    '
    If (RS.RecordCount > 0) Then
       mvarRecordCount = RS.RecordCount
    Else
       mvarRecordCount = 0
    End If
    RecordCount = mvarRecordCount
   '
End Property

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    ' Assign the Recordset to the Data object
    Set Data = RS
End Sub

Private Sub Class_Initialize()
    
    ' Create an instance of the Recordset
    Set RS = New ADODB.Recordset
    Set objDiskFile = New clsDiskFile
    mvarLastIndex = 0
    '
    ' Set the properties of the Recordset
    With RS
        .Fields.Append "Index", adVarChar, 10
        .Fields.Append "SlideNo", adBSTR, 40
        .Fields.Append "ImageDiv", adVarChar, 10
        .Fields.Append "FileSize", adVarChar, 20
        .Fields.Append "FileDate", adVarChar, 40
        .Fields.Append "FileDesc", adVarChar, 255
        .Fields.Append "FileName", adVarChar, 255
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    P_SLIDE_SERVER_PATH = "C:\Schweitzer\Develop\" 'GetConstVariable("P_SLIDE_SERVER_PATH")         ' "C:\Schweitzer\Server\"
End Sub

'Private Function GetConstVariable(ByVal qKeyString As String) As Variant
'    Dim strTmp  As String
'
'    If Dir(INIPath) = "" Then Exit Function
'    strTmp = medGetINI("LIS_CONST", qKeyString, INIPath)
'
'    Select Case medGetP(strTmp, 2, LINE_DIV)
'        Case "0": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
'        Case "1": GetConstVariable = CLng(medGetP(strTmp, 1, LINE_DIV))
'        Case "2": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
'    End Select
'
'End Function

Public Sub LoadSlide(ByVal pPath As String, ByVal pPtid As String, _
   ByVal pBusiDir As String, ByVal pAccseq As String)
   Dim strPath As String, strName As String, strFile As String
   Dim strFileHead As String
   Dim i As Integer
   Dim strTmp As String
   Dim strDesc As String
   
    ' Image File = '1000001-S1999-1902-*.*"
    mvarPtID = pPtid                        '1000001
    mvarBusiDir = pBusiDir                  'S1999
    mvarAccSeq = pAccseq                    '1902
    mvarServerPath = pPath                  '\\SEA\skyD\Documents\Project\Kcl\Slide\
    'strName                                '1000001-S1999-1902-G-1.jpg
    '
    strFileHead = pPtid & "-" & pBusiDir & "-"
    If pAccseq = "" Then
      pAccseq = strFileHead & "*.*"
    Else
      pAccseq = strFileHead & pAccseq & "-" & "*.*"
    End If
    
    If P_SLIDE_SERVER_PATH = "" Then
        strPath = pPath
    Else
        strPath = pPath & pBusiDir & "\"
    End If
'
    
    'strName = Dir(strPath, vbFileName)
    strName = Dir(strPath & pAccseq)
    i = 0
    Do While strName <> ""
        If strName <> "." And strName <> ".." And strName <> "pagefile.sys" Then
            If (GetAttr(strPath & strName) And vbArchive) = vbArchive Then
            'If (GetAttr(strPath & strName) And vbFileName) = vbFileName Then
                i = i + 1
                With RS
                    .AddNew
                    .Fields.Item("Index") = Val(ParseFileNo(strName))
                    If Val(ParseFileNo(strName)) > mvarLastIndex Then
                       mvarLastIndex = Val(ParseFileNo(strName))
                    End If
                    .Fields.Item("SlideNo") = Mid(medGetP(strName, 1, "."), 8 + 2, Len(strName))
                    
                    If P_SLIDE_SERVER_PATH = "" Then
                        .Fields.Item("FileName") = pPath & strName
                        objDiskFile.FileName = pPath & strName
                    Else
                        .Fields.Item("FileName") = pPath & pBusiDir & "\" & strName
                        objDiskFile.FileName = pPath & pBusiDir & "\" & strName
                    End If
'
                    
                    .Fields.Item("FileSize") = objDiskFile.FileSize & " byte"
                    .Fields.Item("FileDate") = objDiskFile.FileDate
                    
                    strDesc = GetDesc(pPath & pBusiDir & "\", strName)
                    
                    .Fields.Item("FileDesc") = medGetP(strDesc, 1, vbTab)
                    
                    strTmp = medGetP(medGetP(strName, 1, "."), 5, "-")
'                    If strTmp = "G" Then
'                       strTmp = SLIDE_GROSS_IMAGE
'                    Else
'                       strTmp = SLIDE_DIAGNOSIS_IMAGE
'                    End If
                    strTmp = "이미지"
                    .Fields.Item("ImageDiv") = strTmp & COL_DIV & medGetP(strDesc, 2, vbTab)
                    .Update
                End With
            End If
        End If
        strName = Dir
    Loop
    
    ' Return to the first record
    If RS.RecordCount > 0 Then
      RS.MoveFirst
    End If
End Sub

Private Function GetDesc(ByVal pPath As String, ByVal pstrName As String) As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim lngSeq      As Long
    Dim RS          As New Recordset
       
    strWorkArea = medGetP(Trim(mvarBusiDir), 1, "-")
    strAccDt = medGetP(Trim(mvarBusiDir), 2, "-")
    strAccSeq = Trim(mvarAccSeq)
    lngSeq = Val(medGetP(medGetP(pstrName, 1, "."), 5, "-"))
    
    If Mid$(strAccDt, 1, 1) = "9" Then
       strAccDt = "19" & strAccDt
    Else
       strAccDt = "20" & strAccDt
    End If
    
    RS.Open " select rmk,prtfg from " & TB_LAB310 & _
               "  where workarea = '" & strWorkArea & "' " & _
               "    and accdt    = '" & strAccDt & "' " & _
               "    and accseq   = '" & strAccSeq & "' " & _
               "    and seq      = " & lngSeq & " " & _
               "    and ptid     = '" & mvarPtID & "' ", DBConn
    
    
    If RS.RecordCount > 0 Then GetDesc = RS.Fields("rmk").Value & "" & vbTab & RS.Fields("prtfg").Value & ""
    
    Set RS = Nothing
End Function

Public Sub MoveLast()
    If RS.RecordCount = 0 Then Exit Sub
    RS.MoveLast
End Sub

Public Sub MoveFirst()
    If RS.RecordCount = 0 Then Exit Sub
    RS.MoveFirst
End Sub

Public Sub MoveNext()
    ' Cycle through the Recordset
    If RS.RecordCount = 0 Then Exit Sub
    RS.MoveNext
End Sub

Public Sub MovePrevious()
    If RS.RecordCount = 0 Then Exit Sub
    RS.MovePrevious
End Sub

Public Function LvwString(Optional ByVal pLastRow As Boolean = False) As String
    Dim aryTmp() As String
    
    If RS.RecordCount = 0 Then
       LvwString = ""
    Else
       RS.MoveFirst
       LvwString = ConvGetString(RS.GetString(adClipString, -1, vbTab, vbNewLine), vbNewLine)
       If pLastRow = True Then
          aryTmp = Split(LvwString, vbNewLine)
          LvwString = aryTmp(UBound(aryTmp))
       End If
    End If
End Function

Private Sub Class_Terminate()
   If Not (RS Is Nothing) Then
      Set RS = Nothing
   End If
   Set objDiskFile = Nothing
   
End Sub

Private Function ParseFileNo(ByVal pFileName) As String
    Dim ii As Integer
   '
'   ParseFileNo = medGetP(pFileName, 4, "-")
    ParseFileNo = medGetP(pFileName, 1, ".")
    ParseFileNo = medGetP(ParseFileNo, 5, "-")
   '
End Function
Public Function ParseFile(ByVal pFileName) As String

    Dim N As Integer
    
    ParseFile = ""
    For N = Len(pFileName) To 1 Step -1
    
        If Mid(pFileName, N, 1) = "\" Then
            ParseFile = Right(pFileName, Len(pFileName) - N)
            N = -1
        End If
        
    Next N
    
End Function

Public Function ParseDir(ByVal pFileName) As String

    Dim N As Integer

    ParseDir = ""
    For N = Len(pFileName) To 1 Step -1
    
        If Mid(pFileName, N, 1) = "\" Then
            ParseDir = Left(pFileName, N)
            N = -1
        End If
        
    Next N
    
End Function

Public Function ParseExt(ByVal pFileName) As String

    Dim N As Integer

    ParseExt = "(N/A)"
    For N = Len(pFileName) To 1 Step -1
    
        If Mid(pFileName, N, 1) = "." Then
            ParseExt = Right(pFileName, Len(pFileName) - N)
            N = -1
        End If
        
    Next N

End Function

Public Function GetNextFileName(ByVal pExt As String, ByVal pImageDiv As String) As String
    Dim strNewFileNo As String
    Dim strName As String
    Dim strPath As String
    
    If RS.RecordCount > 0 Then
       RS.MoveLast
       strNewFileNo = mvarLastIndex
       If strNewFileNo = "" Then
          strNewFileNo = "1"
       Else
          strNewFileNo = CStr(Val(strNewFileNo) + 1)
       End If
    Else
       strNewFileNo = "1"
    End If
    '
    If Val(strNewFileNo) < 10 Then
       strNewFileNo = "0" & Val(strNewFileNo)
    End If
     With Me
         If P_SLIDE_SERVER_PATH = "" Then
             strPath = Me.ServerPath
         Else
             strPath = Me.ServerPath & Me.BusiDir & "\"
         End If
         
         strName = .ptid & "-" & .BusiDir & "-" & .AccSeq & "-" & strNewFileNo & _
                     "." & pExt
    '
    '            strName = .ptid & "-" & .BusiDir & "-" & .AccSeq & "-" & strNewFileNo & _
    '                "-" & pImageDiv & "." & pExt
    '
              
     '
         GetNextFileName = strPath & strName
     End With
     mvarLastIndex = mvarLastIndex + 1
    '
End Function

Public Sub ListAdd(ByVal pFileNo As String, ByVal pExt As String, ByVal pImgDiv As String)
    Dim strName As String
    Dim strTmp As String
   '
    With Me
        If Val(pFileNo) < 10 Then
           pFileNo = "0" & pFileNo
        End If
        
      
        strName = .ptid & "-" & .BusiDir & "-" & .AccSeq & "-" & pFileNo & _
             pImgDiv & "." & pExt
    End With
   '
   With RS
        .AddNew
        .Fields.Item("Index") = mvarLastIndex
        .Fields.Item("SlideNo") = Mid(medGetP(strName, 1, "."), 8 + 2, Len(strName))
      
        If P_SLIDE_SERVER_PATH = "" Then
            .Fields.Item("FileName") = Me.ServerPath & strName
            objDiskFile.FileName = Me.ServerPath & strName
        Else
            .Fields.Item("FileName") = Me.ServerPath & Me.BusiDir & "\" & strName
            objDiskFile.FileName = Me.ServerPath & Me.BusiDir & "\" & strName
        End If
'

        .Fields.Item("FileSize") = objDiskFile.FileSize & " byte"
        .Fields.Item("FileDate") = objDiskFile.FileDate
        .Fields.Item("FileDesc") = ""
        strTmp = medGetP(medGetP(strName, 1, "."), 5, "-")

        strTmp = "이미지"

        .Fields.Item("ImageDiv") = strTmp
        .Update
   End With
   '
End Sub

'문자열 조건식
'rs.Find "S[SlideNo]] = 'S1999-12-1'"

Public Function GetFiles(strPath As String, dctDict As Dictionary, _
   Optional blnRecursive As Boolean) As Boolean
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

