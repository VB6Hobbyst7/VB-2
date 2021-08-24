VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCREGRST0218"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RegServer(ByVal iPersonCnt As Integer, ByVal sWDate As String, ByVal sTWSeq As String, _
                        ByVal sTJDate As String, ByVal sTJGbn As String, ByVal sTJNo As String, _
                        ByVal sTRack As String, ByVal sTPos As String, _
                        ByVal sTRegNo As String, ByVal sTName As String, ByVal sTSex As String, _
                        ByVal sTEmer As String, ByVal sTReRun As String, ByVal sTOther As String, _
                        ByVal sTItemCnt As String, ByVal sTIFSeq As String, ByVal sTServerCd As String, _
                        ByVal sTRst1 As String, ByVal sTRst2 As String, ByVal sTFlag As String, _
                        Optional ByVal iMode As Integer) As String
    On Error GoTo ErrHandler
    
    Dim i%, j%, iITcnt%, k%, m%, iTRegCnt%
    Dim sItemCnt$, sCServerCd$, sCRst1$, sCRst2$, sCFlag$, sTmpServerCd$, sTmpRst1$, sTmpRst2$, sTmpFlag$, sTmpSex$
    Dim sTRegSeq$, sCWSeq$, sCJNo$, sCJDate$, sCJGbn$, sCRegNo$, sCOther$
    Dim sBarCd$, sTmpOther$, sTmpRef$, sCRef$
            
    Call GetTestItem
    Call GetOrdRstCfg
    Call GetTestCdSeq
    
'    Call GetIFFlagInfoDB    'flag 조회
    
    Dim ADOCN   As New ADODB.Connection
    Dim SqlStr$
    Dim sRtnVal$
    Dim lngTransOK&
    Dim lngTransOK1&
    Dim lngTransOK2&
    
    Dim bTrans  As Boolean
    
    Dim sCIFSeqCd$, sTmpIFSeqCd$, sTmpIFRstCd$      '2006/5/29 yk
    
    Dim sEqCd$
    Dim aEqCd() As String
    
    Dim sAPTNO$, sSPRDT$, sYear$
    
    aEqCd() = Split(gsMachineNm, "-")
    sEqCd = Trim(aEqCd(0))

    RegServer = "OK"
        
    ADOCN.Open gRstcfg.sPath, "hanaromf", "hanaro"
    
    iTRegCnt = 0
    sTRegSeq = ""
    
    For i = 1 To iPersonCnt
        sCWSeq = GetByOne(sTWSeq, sTWSeq)
        sCJNo = GetByOne(sTJNo, sTJNo)
        sCRegNo = GetByOne(sTRegNo, sTRegNo)
        
        sYear = Left(Now, 2)
        sAPTNO = Left(sCJNo, 8)
'        sSPRDT = sYear & Right(sCJNo, 6)
        sSPRDT = sYear & Right(sCJNo, 5)
        
        sTmpServerCd = GetByOneUserSymbol(sTServerCd, sTServerCd, Chr$(3))
        sTmpIFSeqCd = GetByOneUserSymbol(sTIFSeq, sTIFSeq, Chr$(3))
        sTmpRst1 = GetByOneUserSymbol(sTRst1, sTRst1, Chr$(3))
        sTmpRst2 = GetByOneUserSymbol(sTRst2, sTRst2, Chr$(3))
        sTmpFlag = GetByOneUserSymbol(sTFlag, sTFlag, Chr$(3))
                
        sItemCnt = GetByOne(sTItemCnt, sTItemCnt)
        sTmpSex = GetByOne(sTSex, sTSex)
        
        sTmpRef = Get_RefVal_HANARO(ADOCN, sTmpServerCd, sTmpRst1, sTmpSex)
        
        ADOCN.BeginTrans: bTrans = True
        lngTransOK1 = 0
        
        For j = 1 To Val(sItemCnt)
            sCServerCd = GetByOne(sTmpServerCd, sTmpServerCd)
            sCIFSeqCd = GetByOne(sTmpIFSeqCd, sTmpIFSeqCd)
            sCRst1 = GetByOne(sTmpRst1, sTmpRst1)
            sCRst2 = GetByOne(sTmpRst2, sTmpRst2)

            If Trim(sTmpRef) <> "" Then
                sCRef = GetByOne(sTmpRef, sTmpRef)
            Else
                sCRef = ""
            End If
       
            '결과 있는 경우만 등록
            If Trim(sCRst1) = "" And Trim(sCRst2) = "" Then
                lngTransOK = 0
            Else
                If Left(sCRst1, 1) = ">" Or Left(sCRst1, 1) = "<" Then
                    sCRst1 = Mid(sCRst1, 2)
                End If
                
                If Not IsNumeric(sCRst1) Then
                    sCRst2 = sCRst1
                    sCRst1 = "0"
'                ElseIf sCRst1 = "0" Then
'                    sCRst1 = "0"
'                    sCRst2 = "0"
                End If
                
                'UPDATE
                SqlStr = "  UPDATE SPL_SZDATA01T  " _
                       & "  SET NRSLT = " & sCRst1 & ", " _
                       & "      CRSLT = '" & sCRst2 & "', " _
                       & "      WORKDT = '" & Trim(sWDate) & "', " _
                       & "      STATUS = '3', " _
                       & "      HIGHLOW = '" & Trim(sCRef) & "' " _
                       & "WHERE WORKCD = '" & Trim(sEqCd) & "'" _
                       & "  AND ACPTNO = '" & Trim(sAPTNO) & "'" _
                       & "  AND BCDDATE = '" & Trim(sSPRDT) & "' " _
                       & "  AND CODE = '" & sCServerCd & "'" _
                       & "  AND STATUS IN ('1','2') "
                       
                ADOCN.Execute SqlStr, lngTransOK
            
                lngTransOK1 = lngTransOK1 + lngTransOK
            End If
        Next j
        
        If lngTransOK1 = 0 Then
            ADOCN.RollbackTrans: bTrans = False
        Else
            '한 행이라도 UPDATE 된 경우
            ADOCN.CommitTrans: bTrans = False
            sTRegSeq = sTRegSeq & sCWSeq & "|"
            iTRegCnt = iTRegCnt + 1
        End If
    Next i
    
    ADOCN.Close: Set ADOCN = Nothing
        
    Call EditRegState(iTRegCnt, sWDate, sTRegSeq)
        
ErrHandler:
    If Err <> 0 Then
        If bTrans = True Then
            ADOCN.RollbackTrans
        End If
        If ADOCN.State = 1 Then
            ADOCN.Close: Set ADOCN = Nothing
        End If
        ViewMsg "RegServer (Err - " & Err.Description & ")"
    End If
End Function

Private Function Get_RefVal_HANARO(ByVal ADOConn As ADODB.Connection, _
                                ByVal sTotServerCd As String, ByVal sTotRst As String, ByVal sSex As String) As String
    On Error GoTo ErrRtn
    
    Dim ADORS   As New ADODB.Recordset
    Dim SqlStr  As String
    Dim ii      As Integer
    Dim tmpOrdCd()  As String
    Dim tmpRst()    As String
    Dim sOrdCd  As String
    Dim sSubCd  As String
    Dim sRst    As String
    Dim sgLow   As Single
    Dim sgHigh  As Single
    Dim tmpRef  As String
    Dim sTotRef As String
    
    sTotRef = ""
    
    tmpOrdCd() = Split(sTotServerCd, Chr(124))
    tmpRst() = Split(sTotRst, Chr(124))
    
    For ii = 0 To UBound(tmpOrdCd()) - 1
        sOrdCd = Trim(tmpOrdCd(ii))
        sRst = Trim(tmpRst(ii))
        
        tmpRef = ""
        
        '참고치 테이블 조회
        SqlStr = " SELECT JRANGEML, JRANGEMH, JRANGEWL, JRANGEWH " _
                & "  FROM BAG_MEDITEMSUGA " _
                & " WHERE MEDITEM = '" & sOrdCd & "'"
    
        ADORS.Open SqlStr, ADOConn, adOpenStatic
    
        If ADORS.RecordCount <> 0 Then
            ADORS.MoveFirst
    
            If Trim(sSex) = "F" Then
                '여
                sgLow = Val(ADORS.Fields("JRANGEML") & "")
                sgHigh = Val(ADORS.Fields("JRANGEMH") & "")
            Else
                '남 or 성별 없는 경우
                sgLow = Val(ADORS.Fields("JRANGEWL") & "")
                sgHigh = Val(ADORS.Fields("JRANGEWH") & "")
            End If
            
            If sgLow = 0 And sgHigh = 0 Then
            Else
                '참고치 비교
                If Val(sRst) > sgHigh Then
                    tmpRef = "H"
                ElseIf Val(sRst) < sgLow Then
                    tmpRef = "L"
                Else
                    tmpRef = ""
                End If
            End If
        End If
        ADORS.Close: Set ADORS = Nothing
                
        '참고치 누적
        sTotRef = sTotRef & tmpRef & Chr(124)
    Next ii
    
    Get_RefVal_HANARO = sTotRef
    
ErrRtn:
    If Err <> 0 Then
        If ADORS.State <> 0 Then
            ADORS.Close: Set ADORS = Nothing
        End If
        Get_RefVal_HANARO = ""
        ViewMsg "참고치 판정 오류 - " & Err.Description
    End If
End Function

Public Function GetDeviceCD() As String
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "if.var1")
    
    If Trim(sBuf) = "" Then
        GetDeviceCD = "ZC1"     'HITACHI7600
    Else
        GetDeviceCD = Trim(sBuf)
    End If
    
    Exit Function
    
ErrHandler:
    ViewMsg "GetDeviceCD - Err(" & Err.Description & ")"
End Function


Private Function ConvLabnoToExpand$(ByVal sComp5$)
    ConvLabnoToExpand = Format(DateAdd("d", Val(sComp5), "1999-01-01"), "YYYYMMDD")
End Function

Public Sub SetMachineInfo(ByVal sMCd$, ByVal sMNm$)
    gsMachineCd = sMCd
    gsMachineNm = sMNm
End Sub

