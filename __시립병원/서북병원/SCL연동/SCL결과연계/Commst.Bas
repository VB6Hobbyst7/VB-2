Attribute VB_Name = "Commst"
Option Explicit

Public Function mDelete(sPrmName As String, sPrmKey As String) As Integer

   ' If bGblAppSecPowerUpdate Then
    
        If sPrmKey = "" Then mDelete = False
    
        mvbFrm.Mvb1.P0 = sPrmName                       ' file name
        mvbFrm.Mvb1.P1 = sPrmKey                        ' key value
    
        mvbFrm.Mvb1.Code = "d ^mDelete(.P0,P1)"
        mvbFrm.Mvb1.ExecFlag = 1
    
        If mvbFrm.Mvb1.P0 = "" Then
            mDelete = False
        Else
            mDelete = True
        End If
        
   ' End If
    
End Function

Public Function mKill(sPrmName As String) As Integer
    
   ' If bGblAppSecPowerUpdate Then
    
        If sPrmName = "" Then mKill = False
    
        mvbFrm.Mvb1.P0 = sPrmName                       ' file name
    
        mvbFrm.Mvb1.Code = "d ^mKill(P0)"
        mvbFrm.Mvb1.ExecFlag = 1
    
        mKill = True
    
    'End If
        
End Function

Public Function mPrint(sRcpNum As String, iMaxNum As Integer)    ', sStateKey As String)
    
    mvbFrm.Mvb1.P0 = sRcpNum
    mvbFrm.Mvb1.P1 = iMaxNum
    
    mvbFrm.Mvb1.Code = "d ^mRcpPrint(P0,P1)"
    mvbFrm.Mvb1.ExecFlag = 1

    'sStateKey = MvbFrm.MVB1.P1

End Function

Public Function mreading(sPrmName As String, sPrmCurrentKey As String, sPrmCompareKey As String, sPrmReturnValue As String, sPrmOrder As String) As String
    
    'If bGblAppSecPowerUpdate Or bGblAppSecPowerRead Then
    
        mvbFrm.Mvb1.P0 = sPrmName                       ' file name
        mvbFrm.Mvb1.P1 = sPrmCurrentKey                 ' 이전에 읽은 Key Value
        mvbFrm.Mvb1.P2 = sPrmCompareKey                 ' 비교 대상 Key Value
        mvbFrm.Mvb1.P3 = sPrmOrder                      ' Order
        mvbFrm.Mvb1.P4 = ""                             ' 결과값 저장
        
        mvbFrm.Mvb1.Code = "d ^mReading(P0,.P1,P2,P3,.P4)"
        mvbFrm.Mvb1.ExecFlag = 1
    
        sPrmReturnValue = mvbFrm.Mvb1.P4                ' return value
        
        mreading = mvbFrm.Mvb1.P1                       ' key value save
    
    'Else
    '    mreading = ""
    '    Exit Function
   
    'End If
    
End Function

Public Function mReadNext(sPrmName As String, sPrmCurrentKey As String, sPrmCompareKey As String, sPrmReturnValue As String) As String
    mReadNext = mreading(sPrmName, sPrmCurrentKey, sPrmCompareKey, sPrmReturnValue, "1")
    
    
End Function

Public Function mReadPrev(sPrmName As String, sPrmCurrentKey As String, sPrmCompareKey As String, sPrmReturnValue As String) As String
    mReadPrev = mreading(sPrmName, sPrmCurrentKey, sPrmCompareKey, sPrmReturnValue, "-1")
End Function

Public Function mSetNext(sPrmName As String, sPrmCurrentKey As String) As String
    
    mvbFrm.Mvb1.P0 = sPrmName                       ' file name
    mvbFrm.Mvb1.P1 = sPrmCurrentKey                 ' Key Value
    
    mvbFrm.Mvb1.Code = "d ^mSetNext(P0,.P1)"
    mvbFrm.Mvb1.ExecFlag = 1
    
    mSetNext = mvbFrm.Mvb1.P1                       ' key value save
    
End Function

Public Function mSetPrev(sPrmName As String, sPrmCurrentKey As String) As String
    
    mvbFrm.Mvb1.P0 = sPrmName                       ' file name
    mvbFrm.Mvb1.P1 = sPrmCurrentKey                 ' Key Value
    
    mvbFrm.Mvb1.Code = "d ^mSetPrev(P0,.P1)"
    mvbFrm.Mvb1.ExecFlag = 1

    mSetPrev = mvbFrm.Mvb1.P1                       ' key value save

End Function

Public Function mSetReadEqual(sPrmName As String, sPrmCurrentKey As String, sPrmReturnValue As String) As String

    'If bGblAppSecPowerUpdate Or bGblAppSecPowerRead Then
    
        mvbFrm.Mvb1.P0 = sPrmName                       ' file name
        mvbFrm.Mvb1.P1 = sPrmCurrentKey                 ' 이전에 읽은 Key Value
        mvbFrm.Mvb1.P2 = ""                             ' 결과값 저장
        
        mvbFrm.Mvb1.Code = "d ^mReadEqual(P0,.P1,.P2)"
        mvbFrm.Mvb1.ExecFlag = 1
        
        sPrmReturnValue = mvbFrm.Mvb1.P2
        'Debug.Print sPrmReturnValue ' return value
        mSetReadEqual = mvbFrm.Mvb1.P1                  ' key value save
        'Debug.Print mSetReadEqual
    'Else
    '    mSetReadEqual = ""
    'End If
    
End Function

Public Function mSetReadNext(sPrmName As String, sPrmCurrentKey As String, sPrmCompareKey As String, sPrmReturnValue As String) As String
    sPrmCurrentKey = mSetNext(sPrmName, sPrmCurrentKey)
    mSetReadNext = mreading(sPrmName, sPrmCurrentKey, sPrmCompareKey, sPrmReturnValue, "1")
End Function

Public Function mSetReadPrev(sPrmName As String, sPrmCurrentKey As String, sPrmCompareKey As String, sPrmReturnValue As String) As String
    sPrmCurrentKey = mSetPrev(sPrmName, sPrmCurrentKey)
    mSetReadPrev = mreading(sPrmName, sPrmCurrentKey, sPrmCompareKey, sPrmReturnValue, "-1")
End Function

Public Function mUpdate(sPrmName As String, sPrmKey As String, sPrmValue As String) As Boolean

    'If bGblAppSecPowerUpdate Then
    
        If sPrmKey = "" Then mUpdate = False
    
        mvbFrm.Mvb1.P0 = sPrmName                       ' file name
        mvbFrm.Mvb1.P1 = sPrmKey                        ' key value
        mvbFrm.Mvb1.P2 = sPrmValue                      ' data value
    
        mvbFrm.Mvb1.Code = "d ^mUpdate(.P0,P1,P2)"
        mvbFrm.Mvb1.ExecFlag = 1
    
        If mvbFrm.Mvb1.P0 = "" Then
            mUpdate = False
        Else
            mUpdate = True
        End If
    'Else
    '    mUpdate = False
    'End If
    
End Function

Public Function mWrite(sPrmName As String, sPrmKey As String, sPrmValue As String) As Integer

    'If bGblAppSecPowerUpdate Then
    
        If sPrmKey = "" Then mWrite = False
    
        mvbFrm.Mvb1.P0 = sPrmName                       ' file name
        mvbFrm.Mvb1.P1 = sPrmKey                        ' key value
        mvbFrm.Mvb1.P2 = sPrmValue                      ' data value
    
        mvbFrm.Mvb1.Code = "d ^mWrite(.P0,P1,P2)"
        mvbFrm.Mvb1.ExecFlag = 1
    
        If mvbFrm.Mvb1.P0 = "" Then
            mWrite = False
        Else
            mWrite = True
        End If
    'Else
    '    mWrite = False
    'End If
    
End Function

Public Function mGetIndexYon(sPrmName As String) As Boolean

    If sPrmName = "" Then
        mGetIndexYon = False
        Exit Function
    End If
        

    mvbFrm.Mvb1.P0 = sPrmName                       ' file name

    mvbFrm.Mvb1.Code = "d ^mGetIndexYon(.P0)"
    mvbFrm.Mvb1.ExecFlag = 1

    If mvbFrm.Mvb1.P0 = "" Then
        mGetIndexYon = False
    Else
        mGetIndexYon = True
    End If

End Function

Public Function mGetGblLst() As String

    mvbFrm.Mvb1.P0 = ""                 ' file name
    mvbFrm.Mvb1.P1 = ""                 ' Key Value
    
    mvbFrm.Mvb1.Code = "d ^mGetGblLst(P0,.P1)"
    mvbFrm.Mvb1.ExecFlag = 1
    
    mGetGblLst = mvbFrm.Mvb1.P1                       ' key value save

End Function


Sub mGD()

    mvbFrm.Mvb1.Code = "d ^mGD"
    mvbFrm.Mvb1.ExecFlag = 1

End Sub


Public Function mIndexRead(sPrmName As String, sPrmReturnValue As String) As String
    
    mvbFrm.Mvb1.P0 = sPrmName                       ' file name
    mvbFrm.Mvb1.P1 = ""                             ' 결과값 저장
    
    mvbFrm.Mvb1.Code = "d ^IdxLstRead(P0,.P1)"
    mvbFrm.Mvb1.ExecFlag = 1
    
    sPrmReturnValue = mvbFrm.Mvb1.P1                ' return value
    mIndexRead = mvbFrm.Mvb1.P1                     ' key value save

End Function

Public Function mGblNamFromIndex(sPrmIndexNam As String, sPrmReturnVal As String) As String

    mvbFrm.Mvb1.P0 = sPrmIndexNam                   ' file name
    mvbFrm.Mvb1.P1 = ""                             ' 결과값 저장
    
    mvbFrm.Mvb1.Code = "d ^mGetGblNameFromIndex(P0,.P1)"
    mvbFrm.Mvb1.ExecFlag = 1
    
    sPrmReturnVal = mvbFrm.Mvb1.P1                ' return value
    mGblNamFromIndex = mvbFrm.Mvb1.P1                     ' key value
    
End Function
