Attribute VB_Name = "ModuleFile"
Option Explicit
Dim F() As String

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function ifFileExists(strfilename As String) As Integer


    Dim i As Integer
    On Error Resume Next
    
    i = Len(Dir$(strfilename))
    
    If Err Or i = 0 Then
        ifFileExists = False
    Else
        ifFileExists = True
    End If
    
End Function
Public Sub Delete_File(Pos_File As String)
   Kill Pos_File   ' file삭제
End Sub
Public Sub Rfile_mdb(deldir As String)
   Const ATTR_DIRECTORY = 16  ' Declare form constant.

   Dim Count, i, DirName ' Declare variables.
   DirName = Dir(deldir & "*.?db", ATTR_DIRECTORY) ' Get first directory name.
   'Iterate through PATH, caching all subdirectories in D()
   
   Do While DirName <> ""
      If DirName <> "." And DirName <> ".." Then
         If GetAttr(deldir + DirName) And ATTR_DIRECTORY = ATTR_DIRECTORY Then
            If (Count Mod 10) = 0 Then
               ReDim Preserve F(Count + 10)  ' Resize the array.
            End If
            Count = Count + 1 ' Increment counter.
            F(Count) = DirName

         End If
      End If
      DirName = Dir  ' Get another directory name.
   Loop

   Counter = Count
End Sub
Public Function del_mdb_file(deldir As String, del_date As Integer) As Integer
   Dim orgdate, FileDate
   Dim i          As Integer
   Dim Datestamp

' False
   del_mdb_file = False

' Date of Client
   orgdate = CVDate(Date)

' Read filename of *.mdb
   Call Rfile_mdb(deldir)

' Read Date of File
   For i = 1 To Counter
      Datestamp = FileDateTime(deldir & F(i))
      FileDate = Format(Datestamp, "YY-MM-DD")
      FileDate = CVDate(FileDate)
     
' Compare of Date
      'If (orgdate - filedate) > 7 Then   del_date
      If (orgdate - FileDate) > del_date Then
' Delete Old File
      Delete_File (deldir & F(i))
      End If
   Next i

' True
   del_mdb_file = True
End Function


Public Sub delcheck(dbpath As String, host As String)

    Dim RetVal%, tmpstr$, tmpstr2$, Tmp$, slipdate$
    
    slipdate = String(255, 0)
    ddate = 0
    
    tmpstr = host & "delete_date"
    ddate = GetPrivateProfileInt(ByVal "Slip setting", ByVal tmpstr, 0, "SLIP.INI")
    tmpstr2 = host & "slip_delete"
    RetVal = GetPrivateProfileString(ByVal "Slip setting", ByVal tmpstr2, "Not Used", slipdate, Len(slipdate), "SLIP.INI")
    
    If ddate = 0 Then
        RetVal = WritePrivateProfileString("Slip setting", ByVal tmpstr, ByVal "7", "SLIP.INI")
        ddate = 7
    End If
    
    If RetVal And Left(slipdate, 8) = "Not Used" Then
        slipdate = Mid$(Date$, 1, 4) & Mid$(Date$, 6, 2) & Mid$(Date$, 9, 2)
        RetVal = WritePrivateProfileString("Slip setting", ByVal tmpstr2, ByVal Left$(slipdate, 8), "SLIP.INI")
    Else
        Tmp = Mid$(Date$, 1, 4) & Mid$(Date$, 6, 2) & Mid$(Date$, 9, 2)
        
        If Left$(slipdate, 8) <> Tmp Then
            Dim rt
            rt = MsgBox("저장된지 " & Str$(ddate) & "일이 지난 모든 자료는 삭제됩니다. 삭제하시겠습니까?", 4, "AxSYM")
            If rt = 6 Then
                rt = del_mdb_file(dbpath, ddate)
                RetVal = WritePrivateProfileString("slip setting", ByVal tmpstr2, ByVal Tmp, "SLIP.INI")
            End If
        End If
    End If
    
End Sub
