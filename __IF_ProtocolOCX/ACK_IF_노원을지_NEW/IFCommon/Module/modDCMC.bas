Attribute VB_Name = "modDCMC"
Option Explicit

Public giIFFlagCnt  As Integer
Type IFFLAGINFO
    SEQ     As String
    FLAGCD  As String
    FLAGINFO    As String
    DISPCD  As String
    USEYN   As String
    REMARK  As String
End Type
Public gIFFlagInfo()    As IFFLAGINFO

Public Sub MakeIFFlagStruct(ByVal sIFFlag As String, ByVal iCnt As Integer)
    Dim ii%
    Dim aRow()  As String
    Dim aData() As String
    
    ReDim gIFFlagInfo(iCnt)
    
    aRow() = Split(sIFFlag, Chr(3))
    
    '0        1       2         3       4      5
    'flagseq, flagcd, flaginfo, dispcd, useyn, remark
    
    For ii = 0 To UBound(aRow())
        If Trim(aRow(ii)) = "" Then Exit For
    
        Erase aData()
        aData() = Split(aRow(ii), Chr(124))
        
        With gIFFlagInfo(ii + 1)
            .SEQ = Trim(aData(0))
            .FLAGCD = Trim(aData(1))
            .FLAGINFO = Trim(aData(2))
            .DISPCD = Trim(aData(3))
            .USEYN = Trim(aData(4))
            .REMARK = Trim(aData(5))
        End With
    Next ii
    
End Sub

Public Function GetIFFlagInfo(ByVal sFlag As String) As String
    
    Dim ii%
    
    GetIFFlagInfo = sFlag
    
    For ii = 1 To giIFFlagCnt
        With gIFFlagInfo(ii)
            If .FLAGCD = sFlag And Trim(.FLAGCD) <> "" And Trim(.FLAGINFO) <> "" Then
                GetIFFlagInfo = Trim(.FLAGINFO)
                Exit For
            End If
        End With
    Next ii
    
End Function

Public Sub GetIFFlagInfoDB()
    On Error GoTo ErrHandler

    Dim objDB As Object
    Dim sRetVal3$
    Dim iItemCnt%

    Set objDB = CreateObject("AIFLD" & Left(fCurVerObject("LocalDB", gsMachineCd), 2) & ".DCIFLD" & fCurVerObject("LocalDB", gsMachineCd))

    'flag info
    sRetVal3 = objDB.Get_IFflaginfo(gsMachineCd)

    If sRetVal3 <> "NONE" Then
        iItemCnt = GetByOneUserSymbol(sRetVal3, sRetVal3, Chr$(3))
        giIFFlagCnt = iItemCnt
        Call MakeIFFlagStruct(sRetVal3, iItemCnt)
    End If

ErrHandler:
    If Err <> 0 Then
        Set objDB = Nothing
        ViewMsg "GetIFFlagInfo - " & Err.Description
    End If
End Sub
