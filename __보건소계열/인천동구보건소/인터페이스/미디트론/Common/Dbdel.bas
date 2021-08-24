Attribute VB_Name = "Del_MDB"
Option Explicit
    
    Dim Counter As Integer
    Dim F() As String
    
    Global slipdate  As String
    Global delhost   As String

Function Del_MDB_File(deldir As String, del_date As Integer) As Integer
    
    Dim orgdate, FileDate
    Dim i   As Integer
    Dim Datestamp
    
    ' False
    Del_MDB_File = False
    
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
        If (orgdate - FileDate) > del_date Then
            ' Delete Old File
            Delete_File (deldir & F(i))
        End If
    Next i
    
    ' True
    Del_MDB_File = True

End Function

Sub DelCheck2(cpath As String)
    
    Dim RetVal  As Integer
    Dim tmp     As String
    Dim retval1 As Integer
    Dim rt  As Integer
    Dim Filename    As String
    
    If ddate = "0" Or ddate = "" Then
        ddate = "7"
    End If
    If RetVal And Left(slipdate, 8) = "Not Used" Then
        slipdate = Mid$(Date$, 9, 2) & Mid$(Date$, 1, 2) & Mid$(Date$, 4, 2)
        'RetVal = WritePrivateProfileString("Slip Setting", ByVal "slip_delete", ByVal Left$(slipdate, 6), "SLIP.INI")
    Else
        tmp = Mid$(Date$, 9, 2) & Mid$(Date$, 1, 2) & Mid$(Date$, 4, 2)
    
        If Left$(slipdate, 6) <> tmp Then
            rt = MsgBox("저장된지 " & Str$(ddate) & "일이 지난 모든 자료는 삭제됩니다. 삭제하시겠습니까?", MB_YESNO)
            If rt = IDYES Then
                rt = Del_MDB_File(cpath, Val(ddate))
                'RetVal = WritePrivateProfileString("Slip Setting", ByVal "slip_delete", ByVal tmp, "SLIP.INI")
            End If
        End If
    End If

End Sub



Sub Delete_File(Pos_File As String)

    Kill Pos_File   ' file삭제
      
End Sub

Sub Rfile_mdb(deldir As String)

   Const ATTR_DIRECTORY = 16  ' Declare form constant.

   Dim count, i, DirName ' Declare variables.
   
   DirName = Dir(deldir & "*.?db", ATTR_DIRECTORY) ' Get first directory name.
   'Iterate through PATH, caching all subdirectories in D()
   
   Do While DirName <> ""
      If DirName <> "." And DirName <> ".." Then
         If GetAttr(deldir + DirName) And ATTR_DIRECTORY = ATTR_DIRECTORY Then
            If (count Mod 10) = 0 Then
               ReDim Preserve F(count + 10)  ' Resize the array.
            End If
            count = count + 1 ' Increment counter.
            F(count) = DirName

         End If
      End If
      DirName = Dir  ' Get another directory name.
   Loop

   Counter = count
   
End Sub

