Attribute VB_Name = "modFolder"
Option Explicit
Option Compare Text

Public Enum opgParsePath
    FILE_ONLY
    PATH_ONLY
    DRIVE_ONLY
    FILEEXT_ONLY
End Enum
'Global Const SLIDE_CLIENT_PATH = "E:\Schweizer\"

Public Function ConvGetString(ByVal pData As String, pDelimeter As String)
Dim aryTmp() As String
   aryTmp = Split(pData, pDelimeter)
   ReDim Preserve aryTmp(UBound(aryTmp) - 1)
   ConvGetString = Join(aryTmp, pDelimeter)
End Function

Public Function ParsePath(strPath As String, lngPart As opgParsePath) As String
    Dim lngPos          As Long
    Dim strPart         As String
    Dim blnIncludesFile As Boolean
    On Error GoTo ParsePath_End
    lngPos = InStrRev(strPath, "\")
    strPath = CStr(strPath)
    blnIncludesFile = InStrRev(strPath, ".") > lngPos
    If lngPos > 0 Then
        Select Case lngPart
            Case opgParsePath.FILE_ONLY
                If blnIncludesFile Then
                    '                    strPart = Right$(strPath, Len(strPath) - lngPos)
                    strPart = Mid(strPath, lngPos + 1, Len(strPath))
                Else
                    strPart = ""
                End If
            Case opgParsePath.PATH_ONLY
                If blnIncludesFile Then
                    '
                    '                     strPart = Left$(strPath, lngPos)
                    strPart = Mid(strPath, 1, lngPos)
                    
                Else
                    strPart = strPath
                End If
            Case opgParsePath.DRIVE_ONLY
                'strPart = Left$(strPath, 3)
                strPart = Mid(strPath, 1, InStr(strPath, "\"))
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

Public Function FillDictionary(ByVal strPath As String) As Scripting.Dictionary
    Dim fsoSysObj As Scripting.FileSystemObject
    Dim fdrFolder As Scripting.Folder
    Dim filFile As Scripting.File
    Dim dctImages As Scripting.Dictionary
    Set fsoSysObj = New FileSystemObject
    'Set fdrFolder = fsoSysObj.GetFolder("C:\ANA\Anatomic\SlideImage\")
    Set fdrFolder = fsoSysObj.GetFolder(strPath)
    Set dctImages = New Scripting.Dictionary
    
    For Each filFile In fdrFolder.Files
        Select Case ParsePath(filFile.Path, FILEEXT_ONLY)

            Case "bmp", "gif", "jpg"
                dctImages.Add filFile.Path, filFile.Name
            Case Else
                
        End Select
    Next
    Set FillDictionary = dctImages
End Function

