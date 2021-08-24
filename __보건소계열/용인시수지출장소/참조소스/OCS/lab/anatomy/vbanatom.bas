Attribute VB_Name = "VbAnatomy"
Option Explicit

Global GstrJinName    As String

Global GsLoginOk      As String
Global GsDiagNo       As String
Global GsSpecial      As String
Global GsGoFlag       As String
Global GsHistology       As String
Global GsCytology     As String
Global GsGross        As String
Global GsFirst        As String
Global GsComplete     As String
Global GsJSHistology     As String
Global GsJSCytology   As String
''Global LsDeptNO       As String

Public gSFrDate       As String
Public gSToDate       As String

Global GReceptSeq     As Integer
Global GAnato_Jeobsu_View As Boolean

Global GobjectSS      As Object


Public Function Specode_Get(ByVal sCodeky As String, ByVal sCodeGu As String) As String
    Dim Vrs                 As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & "   FROM TWEXAM_SPECODE "
    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
    strSQL = strSQL & "  WHERE CodeGu = '" & sCodeGu & "' "
    strSQL = strSQL & "    AND Codeky = '" & Trim(sCodeky) & "' "
    
    Result = AdoOpenSet(Vrs, strSQL)
    
    If Result Then
        Specode_Get = Vrs.Fields("itemnm").Value & ""
        Exit Function
    End If
    
    Specode_Get = " "
    
    Vrs.Close
    If Not Vrs Is Nothing Then
        Set Vrs = Nothing
    End If
        
    Exit Function

End Function


Public Function DiagCodeSearch(ByVal sCodeky As String) As String
    Dim Vrs                 As ADODB.Recordset

    If sCodeky = "" Then DiagCodeSearch = "": Exit Function
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_Dict "
    strSQL = strSQL & "  WHERE CODE = '" & sCodeky & "' "

    Result = AdoOpenSet(Vrs, strSQL)
    
    If Result Then
        DiagCodeSearch = Vrs.Fields("dxdict").Value & ""
    Else
        DiagCodeSearch = ""
    End If
    
    Vrs.Close
    If Not Vrs Is Nothing Then
        Set Vrs = Nothing
    End If

End Function

