VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls°ø¿ë_SFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CloseConnection()
    Call sftp.Disconnect
End Sub

Public Function CreateFTPDirectory(sDirectory As String) As Boolean
    Dim success As Long
    '  Create a new directory:
    success = sftp.CreateDir(sDirectory)
    If (success <> 1) Then
        CreateFTPDirectory = False
'''        MsgBox sftp.LastErrorText
'''        Exit Sub
    Else
        CreateFTPDirectory = True
    End If
End Function

Public Function FTPDownloadFile(sLocal As String, sRemote As String) As Boolean
    Dim handle  As String
    Dim success As Long
    
    FTPDownloadFile = False
    
    '  Open a file on the server:
    handle = sftp.OpenFile(sRemote, "readOnly", "openExisting")
    If (handle = vbNullString) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    '  Download the file:
    success = sftp.DownloadFile(handle, sLocal)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    '  Close the file.
    success = sftp.CloseHandle(handle)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    FTPDownloadFile = True
End Function

Public Sub FtpScanDirectory(strIMGFILEPATH As String, Optional ByVal sParams$ = "*.txt")
    Dim success As Long
    Dim handle As String
    Dim intFileCnt      As Integer
    Dim intFileCnt_IMG  As Integer
       
    intFileCnt = 0
    ReDim Preserve FtpScanFileName(intFileCnt)
    ReDim Preserve FtpScanFileDate(intFileCnt)
    
    intFileCnt_IMG = 0
    ReDim Preserve FtpScanFileName_IMG(intFileCnt_IMG)
    ReDim Preserve FtpScanFileName_IMG(intFileCnt_IMG)
    
    handle = sftp.OpenDir(strIMGFILEPATH)
    
    If (handle = vbNullString) Then
        MsgBox sftp.LastErrorText
        Exit Sub
    End If
    
    '  Download the directory listing:
    Dim dirListing As ChilkatSFtpDir
    
    Set dirListing = sftp.ReadDir(handle)
    
    If (dirListing Is Nothing) Then
        MsgBox sftp.LastErrorText
        Exit Sub
    End If
    
    '  Iterate over the files.
    Dim i As Long
    Dim n As Long
    n = dirListing.NumFilesAndDirs
    If (n = 0) Then
'''        Text1.Text = Text1.Text & "No entries found in this directory." & vbCrLf
    Else
        For i = 0 To n - 1
            
            Dim fileObj As ChilkatSFtpFile
            
            Set fileObj = dirListing.GetFileObject(i)
            
            If fileObj.Filename <> "." And fileObj.Filename <> ".." Then
                intFileCnt = intFileCnt + 1
                ReDim Preserve FtpScanFileName(intFileCnt)
                ReDim Preserve FtpScanFileDate(intFileCnt)
                
                FtpScanFileName(intFileCnt) = fileObj.Filename
                FtpScanFileDate(intFileCnt) = fileObj.LastModifiedTime
                
                If InStr(UCase(fileObj.Filename), ".TIF") > 0 Or InStr(UCase(fileObj.Filename), ".TIFF") > 0 Or InStr(UCase(fileObj.Filename), ".PDF") > 0 Or InStr(UCase(fileObj.Filename), ".JPG") > 0 Or InStr(UCase(fileObj.Filename), ".JPEG") > 0 Or InStr(UCase(fileObj.Filename), ".IMG") > 0 Then
                    intFileCnt_IMG = intFileCnt_IMG + 1
                    
                    ReDim Preserve FtpScanFileName_IMG(intFileCnt_IMG)
                    ReDim Preserve FtpScanFileDate_IMG(intFileCnt_IMG)
                    
                    FtpScanFileName_IMG(intFileCnt_IMG) = fileObj.Filename
                    FtpScanFileDate_IMG(intFileCnt_IMG) = fileObj.LastModifiedTime
                End If
            End If
        Next
    End If
    
    '  Close the directory
    success = sftp.CloseHandle(handle)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Sub
    End If
End Sub

Public Function FTPUploadFile(sLocal As String, sRemote As String) As Boolean
    Dim success As Long

    FTPUploadFile = False
    
    '  Open a file for writing on the SSH server.
    '  If the file already exists, it is overwritten.
    '  (Specify "createNew" instead of "createTruncate" to
    '  prevent overwriting existing files.)
    Dim handle As String
    handle = sftp.OpenFile(sRemote, "writeOnly", "createTruncate")
    If (handle = vbNullString) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    '  Upload from the local file to the SSH server.
    success = sftp.UploadFile(handle, sLocal)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    '  Close the file.
    success = sftp.CloseHandle(handle)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If

    FTPUploadFile = True
End Function

Public Function OpenConnection(sServer As String, sPort As String, sUser As String, sPassword As String) As Boolean
    Dim success As Long
    
    OpenConnection = False
    
    '''success = sftp.UnlockComponent("Anything for 30-day trial")
    success = sftp.UnlockComponent("SSH12345678_3B73CC038O1G")
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    '  Set some timeouts, in milliseconds:
    sftp.ConnectTimeoutMs = 5000
    sftp.IdleTimeoutMs = 10000
    
    '  Connect to the SSH server.
    '  The standard SSH port = 22
    '  The hostname may be a hostname or IP address.
    success = sftp.Connect(gstrFTP_RH, gstrFTP_RP)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If

    '  Authenticate with the SSH server.  Chilkat SFTP supports
    '  both password-based authenication as well as public-key
    '  authentication.  This example uses password authenication.
    success = sftp.AuthenticatePw(gstrFTP_UN, gstrFTP_PW)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    '  After authenticating, the SFTP subsystem must be initialized:
    success = sftp.InitializeSftp()
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
    
    OpenConnection = True
End Function

Public Function RemoveFTPDirectory(sDirectory As String) As Boolean
    Dim success As Long
    '  Delete a directory on the server.
    '  The directory should be empty prior to deleting.
    success = sftp.RemoveDir(sDirectory)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
End Function

Public Function RenameFTPFile(sExisting As String, sNewName As String) As Boolean
    Dim success As Long
    
    '  Rename the file or directory:
    success = sftp.RenameFileOrDir(sExisting, sNewName)
    If (success <> 1) Then
        MsgBox sftp.LastErrorText
        Exit Function
    End If
End Function

Public Function SetFTPDirectory(sDir As String)
    Dim handle As String
    
    handle = sftp.OpenDir(sDir)
    If (handle = vbNullString) Then
        SetFTPDirectory = False
        '''MsgBox sftp.LastErrorText
        '''Exit Sub
    Else
        SetFTPDirectory = True
    End If
End Function

