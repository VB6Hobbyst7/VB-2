VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Base 1
Option Explicit

Private objBarInfo      As clsDictionary
Private objBarHan       As clsDictionary
Private objMySql        As New clsBarSqlStmt
Private mudtStatInfo    As tpStatInfo
Private mudtBarcode     As tpBarcode
Private mvarFontKor     As enFontKor
Private mvarRotation    As enRotation

Private mudtBarData(1 To 14) As tpBarData

Private mvarPortNo      As String
Private mvarWidth       As String
Private mvarLength      As String
Private mvarTotLength   As String
Private mvarAccCheck    As String
Private mvarBarKind     As String

Private mvarMainKind    As String
Private mvarMainPort    As String


Private strSqlStmt()    As String

Friend Property Let StatInfo(ByRef vData As tpStatInfo)
'-- 응급여부 출력양식
    mudtStatInfo = vData
End Property
Friend Property Get StatInfo() As tpStatInfo
    StatInfo = mudtStatInfo
End Property

Friend Property Let Barcode(ByRef vData As tpBarcode)
'-- 바코드 출력양식
    mudtBarcode = vData
End Property
Friend Property Get Barcode() As tpBarcode
    Barcode = mudtBarcode
End Property

Friend Property Let BarData(ByVal iIndex As enBarData, ByRef vData As tpBarData)
'-- 각종 정보 출력양식
    mudtBarData(iIndex) = vData
End Property
Friend Property Get BarData(ByVal iIndex As enBarData) As tpBarData
    BarData = mudtBarData(iIndex)
End Property

Public Property Let Width(ByVal vData As String)
'-- 라벨 너비
    mvarWidth = vData
End Property
Public Property Get Width() As String
    Width = mvarWidth
End Property

Public Property Let Length(ByVal vData As String)
'-- 라벨 길이
    mvarLength = vData
End Property
Public Property Get Length() As String
    Length = mvarLength
End Property

Public Property Let TotLength(ByVal vData As String)
'-- 라벨 전체길이
    mvarTotLength = vData
End Property
Public Property Get TotLength() As String
    TotLength = mvarTotLength
End Property

Public Property Let PortNo(ByVal vData As Long)
'-- 시리얼 포트 번호
    mvarPortNo = vData
End Property
Public Property Get PortNo() As Long
    PortNo = Val(mvarPortNo)
End Property

Public Property Let AccCheck(ByVal vData As String)
'-- 접수번호 체크
    mvarAccCheck = vData
End Property
Public Property Get AccCheck() As String
    AccCheck = mvarAccCheck
End Property

Friend Property Let FontKor(ByVal vData As enFontKor)
'-- 한글 폰트
    mvarFontKor = vData
End Property
Friend Property Get FontKor() As enFontKor
    FontKor = mvarFontKor
End Property

Friend Property Let Rotation(ByVal vData As enRotation)
'-- Rotation
    mvarRotation = vData
End Property
Friend Property Get Rotation() As enRotation
    Rotation = mvarRotation
End Property


Public Property Let BarKind(ByVal vData As String)
'-- BarCode Kind
    mvarBarKind = vData
End Property

Public Property Get BarKind() As String
    BarKind = mvarBarKind
End Property

Public Property Let MainBarKind(ByVal vData As String)
'-- BarCode Kind
    mvarMainKind = vData
End Property

Public Property Get MainBarKind() As String
    MainBarKind = mvarMainKind
End Property

Public Sub GetPortNO(ByRef objBarcode As clsBarcode)
    Dim Rs As Recordset
    
    Set Rs = New Recordset
    
    On Error GoTo ErrTrap
    
    Rs.Open objMySql.Get_Port, DBConn
    
    With objBarcode
        .PortNo = Val(Rs.Fields("field1").Value & "")
    End With
    
    Set Rs = Nothing
    
ErrTrap:
    Set Rs = Nothing
    MsgBox Err.Description, vbCritical
End Sub

Public Sub ReadConfig(ByVal strPrjNm As String)
    Dim objRs    As Recordset
    Dim strKey   As String
    Dim ii       As Integer
    
    ' 기존 테이블 때문에 어쩔수 없다..
    ' 기존 병원의 경우는 바코드기종을 한대밖에 설정할수 없다.

    mvarMainKind = medGetP(strPrjNm, 2, COL_DIV)
    

    Set objRs = New Recordset
    objRs.Open objMySql.SQL_ReadConfig(strPrjNm), DBConn
    
    
'    If objRs.DBerror Then
'        dbconn.DisplayErrors
'        Set objRs = Nothing
'        Exit Sub
'    End If

    While Not objRs.EOF
        strKey = objRs.Fields("cdval2").Value & ""
        Select Case strKey
            Case BAR_PORT:    mvarPortNo = objRs.Fields("field1").Value & ""
                              mvarMainPort = mvarMainPort
            Case BAR_WIDTH:   mvarWidth = objRs.Fields("field1").Value & ""
            Case BAR_LENGTH:  mvarLength = objRs.Fields("field1").Value & ""
            Case BAR_TOTLEN:  mvarTotLength = objRs.Fields("field1").Value & ""
            Case BAR_BARCODE:
                 With mudtBarcode
                    .PrtFg = objRs.Fields("field1").Value & ""
                    .PosX = medGetP(objRs.Fields("field2").Value & "", 1, COL_DIV)
                    .PosY = medGetP(objRs.Fields("field2").Value & "", 2, COL_DIV)
                    .Length = objRs.Fields("field3").Value & ""
                    .Height = objRs.Fields("field4").Value & ""
                    .Style = objRs.Fields("field5").Value & ""
                    
                 End With
            Case BAR_ACCFG:  mvarAccCheck = objRs.Fields("field1").Value & ""
'---------------------------------
'바코드 기종이 Zebra일때
'한글 폰트를 Dictionary에 담아논다.
'---------------------------------
            Case BAR_KIND:   mvarBarKind = objRs.Fields("field1").Value & ""
                    If mvarBarKind = "2" Then
                        objBarHan.Clear
                        objBarHan.FieldInialize "han", "hex,val"
                    End If
                    mvarMainKind = mvarBarKind
            Case BAR_STAT:
                 With mudtStatInfo
                    .PrtLineFg = medGetP(objRs.Fields("field1").Value & "", 1, COL_DIV)
                    .PrtReverseFg = medGetP(objRs.Fields("field1").Value & "", 2, COL_DIV)
                    .PosX = medGetP(objRs.Fields("field2").Value & "", 1, COL_DIV)
                    .PosY = medGetP(objRs.Fields("field2").Value & "", 2, COL_DIV)
                    .Width = medGetP(objRs.Fields("field3").Value & "", 1, COL_DIV)
                    .Length = medGetP(objRs.Fields("field3").Value & "", 2, COL_DIV)
                    .ErDeptCd = objRs.Fields("field4").Value & ""
                    .ReverseFld = objRs.Fields("field5").Value & ""
                 End With
            Case DAT_LOCATION, DAT_WORKAREA, DAT_ACCDT, DAT_ACCSEQ, DAT_SPCNO, _
                 DAT_DEPT, DAT_ORDDT, DAT_COLTM, DAT_PTNM, DAT_PTID, DAT_SPCNM, _
                 DAT_ORDNM1, DAT_ORDNM2, DAT_FROZEN, DAT_BLDPTNM, DAT_BLDPTID, _
                 DAT_WARDID, DAT_COMPNM, DAT_VOLUME, DAT_BLDTYPE, DAT_BLDNO, DAT_TESTDT:
                 With mudtBarData(strKey)
                    .Key = Val(strKey)
                    .PrtFg = objRs.Fields("field1").Value & ""
                    .PosX = medGetP(objRs.Fields("field2").Value & "", 1, COL_DIV)
                    .PosY = medGetP(objRs.Fields("field2").Value & "", 2, COL_DIV)
                    .Length = objRs.Fields("field3").Value & ""
                    .FontX = medGetP(objRs.Fields("field4").Value & "", 1, COL_DIV)
                    .FontY = medGetP(objRs.Fields("field4").Value & "", 2, COL_DIV)
                    .BoldFg = medGetP(objRs.Fields("field5").Value & "", 1, COL_DIV)
                    .ReverseFg = medGetP(objRs.Fields("field5").Value & "", 2, COL_DIV)
                 End With
        End Select
        objRs.MoveNext
    Wend

'    objRs.RsClose
    Set objRs = Nothing

    If Dir(INIPath) = "" Then
        Call medSetINI(medGetP(strPrjNm, 1, COL_DIV), "PORT", mvarPortNo, INIPath)
        Call medSetINI(medGetP(strPrjNm, 1, COL_DIV), "KIND", mvarBarKind, INIPath)
    Else
        If medGetINI(medGetP(strPrjNm, 1, COL_DIV), "PORT", mvarPortNo, INIPath) = "" Then
            Call medSetINI(medGetP(strPrjNm, 1, COL_DIV), "PORT", mvarPortNo, INIPath)
            Call medSetINI(medGetP(strPrjNm, 1, COL_DIV), "KIND", mvarBarKind, INIPath)
        End If
    End If
    mvarPortNo = medGetINI(medGetP(strPrjNm, 1, COL_DIV), "PORT", INIPath)
    mvarBarKind = medGetINI(medGetP(strPrjNm, 1, COL_DIV), "KIND", INIPath)
 

End Sub

Public Function SaveConfig(ByVal strPrjNm As String) As Boolean

    Dim i As Long
    Dim strFields(1 To 5) As String

    ' 기존 테이블 때문에 어쩔수 없다..

    Erase strSqlStmt
    
    ReDim strSqlStmt(1)

    Call MakeSqlStmt(strPrjNm, BAR_PORT, mvarPortNo, "", "", "", "")
    Call MakeSqlStmt(strPrjNm, BAR_WIDTH, mvarWidth, "", "", "", "")
    Call MakeSqlStmt(strPrjNm, BAR_LENGTH, mvarLength, "", "", "", "")
    Call MakeSqlStmt(strPrjNm, BAR_TOTLEN, mvarTotLength, "", "", "", "")
    '바코드 기종
    Call MakeSqlStmt(strPrjNm, BAR_KIND, mvarBarKind, "", "", "", "")
    
    Call MakeSqlStmt(strPrjNm, BAR_BARCODE, mudtBarcode.PrtFg, _
                     Format(mudtBarcode.PosX, "0###") & COL_DIV & Format(mudtBarcode.PosY, "0###"), _
                     Format(mudtBarcode.Length, "0#"), Format(mudtBarcode.Height, "0###"), _
                     mudtBarcode.Style)
    Call MakeSqlStmt(strPrjNm, BAR_ACCFG, mvarAccCheck, "", "", "", "")
    Call MakeSqlStmt(strPrjNm, BAR_STAT, mudtStatInfo.PrtReverseFg & COL_DIV & mudtStatInfo.PrtLineFg, _
                     Format(mudtStatInfo.PosX, "0###") & COL_DIV & Format(mudtStatInfo.PosY, "0###"), _
                     Format(mudtStatInfo.Width, "0###") & COL_DIV & Format(mudtStatInfo.Length, "0###"), _
                     mudtStatInfo.ErDeptCd, mudtStatInfo.ReverseFld)

    For i = 1 To 14
        With mudtBarData(i)

            Erase strFields

            strFields(1) = .PrtFg                                   '출력여부
            strFields(2) = Format(.PosX, "0###") & COL_DIV          'PosX
            strFields(2) = strFields(2) & Format(.PosY, "0###")     'PosY
            strFields(3) = Format(.Length, "0#")                    'Length
            strFields(4) = Format(.FontX, "0") & COL_DIV            'FontX
            strFields(4) = strFields(4) & Format(.FontY, "0")       'FontY
            strFields(5) = Format(.BoldFg, "0") & COL_DIV           'BoldFg
            strFields(5) = strFields(5) & Format(.ReverseFg, "0")   'ReverseFg

            Call MakeSqlStmt(strPrjNm, CStr(i), strFields(1), strFields(2), _
                             strFields(3), strFields(4), strFields(5))
        End With
    Next

On Error GoTo Err_Trap

    DBConn.BeginTrans
    For i = 1 To UBound(strSqlStmt)
        If strSqlStmt(i) <> "" Then
            DBConn.Execute strSqlStmt(i)
            
        End If
    Next
    DBConn.CommitTrans
    SaveConfig = True
    
    Exit Function
    
Err_Trap:
    DBConn.RollbackTrans
    SaveConfig = False
    MsgBox Err.Description, vbExclamation

End Function


Private Sub MakeSqlStmt(ByVal strPrjNm As String, ByVal strCdVal2 As String, ParamArray strFields() As Variant)

    '기존 데이터 Delete
    ReDim Preserve strSqlStmt(UBound(strSqlStmt) + 1)
    strSqlStmt(UBound(strSqlStmt)) = _
        "delete from " & objTables.T_COM002 & " where cdindex = " & DBS(COM2_BAR_CONFIG) & _
        "and cdval1 = " & DBS(strPrjNm) & " and cdval2 = " & DBS(strCdVal2)

    '변경된 데이터 Insert
    ReDim Preserve strSqlStmt(UBound(strSqlStmt) + 1)
    strSqlStmt(UBound(strSqlStmt)) = _
        "insert into " & objTables.T_COM002 & _
        "(cdindex, cdval1, cdval2, field1, field2, field3, field4, field5) " & _
        "values (" & DBS(COM2_BAR_CONFIG, 1) & DBS(strPrjNm, 1) & DBS(strCdVal2, 1) & _
        DBS(strFields(0), 1) & DBS(strFields(1), 1) & DBS(strFields(2), 1) & _
        DBS(strFields(3), 1) & DBS(strFields(4)) & ")"

End Sub

Friend Sub GetConfig(ByVal strPrjNm As String, ByRef objBarcode As clsBarcode)
    Dim strSql  As String
    Dim i       As Long
    
    If Trim(strPrjNm) = "" Then Exit Sub

    Call ReadConfig(strPrjNm)
    
    With objBarcode
        .PortNo = Val(mvarPortNo)
        .LabelWidth = mvarWidth
        .LabelLength = mvarLength
        .LabelTotLength = mvarTotLength
        .FontKor = FONT_GULIM
        .Rotation = ROT_ZERO
        .AccCheck = mvarAccCheck
        .StatInfo = mudtStatInfo

        .Barcode = mudtBarcode
        .BarKind = mvarBarKind
        '한글 이미지......
        If .BarKind = "2" Then
             Set .ObjBarH = objBarHan
        End If
        
        For i = 1 To 14
            .BarData(i) = mudtBarData(i)
        Next
    End With


End Sub

Private Sub Class_Initialize()
    Set objBarInfo = New clsDictionary
    Set objBarHan = New clsDictionary
End Sub

Private Sub Class_Terminate()
    Set objBarInfo = Nothing
    Set objBarHan = Nothing
End Sub
