VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlCodeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------- 공통코드 마스터 --------

Public Function GetComCdIndex(ByVal cdindex As String, Optional ByVal CdVal1 As String = "", _
                              Optional ByVal CdVal2 As String = "") As String
'Coding By Legends
'FROM 공통코드 인덱스(COM001) SELECT
    
    GetComCdIndex = "SELECT * FROM " & T_COM001 & _
                   " WHERE  " & DBW("cdindex", cdindex, 2)
                 
    If CdVal1 <> "" Then
        GetComCdIndex = GetComCdIndex & " AND " & DBW("cdval1", CdVal1, 2)
        
        If CdVal2 <> "" Then
            GetComCdIndex = GetComCdIndex & " AND " & DBW("cdval2", CdVal2, 2)
        End If
    End If
End Function

Public Function SetComCdIndex(ByVal pUpdateFg As Boolean, ByVal pCdindex As String, _
                              Optional ByVal pCdVal1 As String, Optional ByVal pCdVal2 As String, _
                              Optional ByVal pField1 As String, Optional ByVal pField2 As String, _
                              Optional ByVal pField3 As String, Optional ByVal pField4 As String, _
                              Optional ByVal pField5 As String, Optional ByVal pText1 As String, _
                              Optional ByVal pText2 As String)
'Coding By Legends
'To 공통코드 인덱스(COM001)
'pUpdateFg True : Update, False : Insert
    If pUpdateFg Then
        SetComCdIndex = " update " & T_COM001 & _
                        " set " & _
                                    DBW("field1", pField1, 3) & _
                                    DBW("field2", pField2, 3) & _
                                    DBW("field3", pField3, 3) & _
                                    DBW("field4", pField4, 3) & _
                                    DBW("field5", pField5, 3) & _
                                    DBW("text1", pText1, 3) & _
                                    DBW("text2", pText2, 2) & _
                        " WHERE " & DBW("cdindex", pCdindex, 2) & _
                        "   AND " & DBW("cdval1", pCdVal1, 2) & _
                        "   AND " & DBW("cdval2", pCdVal2, 2)
    Else
        SetComCdIndex = "insert into " & T_COM001 & _
                        " (cdindex,cdval1,cdval2,field1,field2,field3,field4,field5,text1,text2) values ( " & _
                        DBV("cdindex", pCdindex, 1) & DBV("cdval1", pCdVal1, 1) & DBV("cdval2", pCdVal2, 1) & _
                        DBV("field1", pField1, 1) & DBV("field2", pField2, 1) & DBV("field3", pField3, 1) & _
                        DBV("field4", pField4, 1) & DBV("field5", pField5, 1) & DBV("text1", pText1, 1) & _
                        DBV("text2", pText2) & ")"
    End If
End Function

Public Function DelComCdIndex(ByVal cdindex As String, ByVal CdVal1 As String, _
                              ByVal CdVal2 As String) As String
'Coding By Legends
'To 공통코드 마스터1(COM001) Delete
    DelComCdIndex = "delete FROM " & T_COM001 & _
                   " WHERE " & DBW("cdindex", cdindex, 2) & _
                   "   AND " & DBW("cdval1", CdVal1, 2) & _
                   "   AND " & DBW("cdval2", CdVal2, 2)
End Function
                            
Public Function DelALLComCdIndex(ByVal cdindex As String, ByVal CdVal1 As String) As String
'Coding By Legends
'To 공통코드 마스터1(COM001) Delete
    DelALLComCdIndex = "delete FROM " & T_COM001 & _
                       " WHERE " & DBW("cdindex", cdindex, 2) & _
                       "   AND " & DBW("cdval1", CdVal1, 2)
End Function

Public Function GetComCdMST1(ByVal cdindex As String, Optional ByVal CdVal1 As String = "", _
                             Optional ByVal CdVal2 As String = "") As String
'Coding By Legends
'FROM 공통코드 마스터1(COM002) SELECT
    
    Select Case cdindex
        Case LC2_MultiSpc
             GetComCdMST1 = "SELECT a.cdindex,a.cdval1,a.cdval2,a.field1,a.field2,a.field3,a.field4,a.field5,a.text1,a.text2," & _
                           "       b.field4 as spcnm FROM " & T_LAB031 & " a," & T_LAB032 & " b" & _
                           " WHERE " & DBW("a.cdindex", cdindex, 2) & _
                           " AND " & DBW("b.cdindex=", LC3_Specimen)
                           
            If CdVal1 <> "" Then
                GetComCdMST1 = GetComCdMST1 & " AND " & DBW("a.cdval1", CdVal1, 2)
                
                If CdVal2 <> "" Then
                    GetComCdMST1 = GetComCdMST1 & " AND " & DBW("a.cdval2", CdVal2, 2)
                End If
                GetComCdMST1 = GetComCdMST1 & " AND a.field1=b.cdval1"
            Else
                GetComCdMST1 = GetComCdMST1 & " AND a.cdval1=b.cdval1"
            End If
       
        Case LC2_Detail, LC2_Panel, LC2_ItemResult, LC2_Calculation, LC2_RelTest, LC2_TempletTest, LC2_TempletText1, LC2_TempletText2, _
            LC2_TempletText3, LC2_DoctTest
        
            GetComCdMST1 = "SELECT a.cdindex,a.cdval1,a.cdval2,a.field1,a.field2,a.field3,a.field4,a.field5,a.text1,a.text2," & _
                           "       b.testnm FROM " & T_LAB031 & " a," & T_LAB001 & " b" & _
                           " WHERE " & DBW("a.cdindex", cdindex, 2)
                         
            If CdVal1 <> "" Then
                GetComCdMST1 = GetComCdMST1 & " AND " & DBW("a.cdval1", CdVal1, 2)
                
                If CdVal2 <> "" Then
                    GetComCdMST1 = GetComCdMST1 & " AND " & DBW("a.cdval2", CdVal2, 2)
                End If
                
                Select Case cdindex
                    Case LC2_Detail, LC2_RelTest:    GetComCdMST1 = GetComCdMST1 & " AND a.cdval2=b.testcd"
                    Case LC2_Panel:     GetComCdMST1 = GetComCdMST1 & " AND a.field1=b.testcd"
                    Case LC2_ItemResult, LC2_Calculation, LC2_TempletTest, LC2_TempletText1, LC2_TempletText2, LC2_TempletText3:
                        GetComCdMST1 = GetComCdMST1 & " AND a.cdval1=b.testcd"
                    Case LC2_DoctTest:
                        GetComCdMST1 = GetComCdMST1 & " AND a.cdval2=b.testcd"
                End Select
                
            Else
                If cdindex = LC2_DoctTest Then
                    GetComCdMST1 = GetComCdMST1 & " AND a.cdval2=b.testcd"
                Else
                    GetComCdMST1 = GetComCdMST1 & " AND a.cdval1=b.testcd"
                End If
            End If
        
        Case LC2_Warning
            GetComCdMST1 = "SELECT a.cdindex,a.cdval1,a.cdval2,a.field1,a.field2,a.field3,a.field4,a.field5,a.text1,a.text2," & _
                           "       b.testnm, c.field1 as rstnm " & _
                           "  FROM " & T_LAB031 & " a," & T_LAB001 & " b, " & T_LAB031 & " c " & _
                           " WHERE " & DBW("a.cdindex", cdindex, 2) & _
                           "   AND a.cdval1 = b.testcd " & _
                           "   AND b.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                           "                       WHERE testcd = b.testcd) " & _
                           "   AND " & DBW("c.cdindex", LC2_ItemResult, 2) & _
                           "   AND a.cdval1 = c.cdval1 AND a.cdval2 = c.cdval2 "
            If CdVal1 <> "" Then
                GetComCdMST1 = GetComCdMST1 & " AND " & DBW("a.cdval1", CdVal1, 2)
            End If
            If CdVal2 <> "" Then
                GetComCdMST1 = GetComCdMST1 & " AND " & DBW("a.cdval2", CdVal2, 2)
            End If

        Case Else
            GetComCdMST1 = "SELECT * FROM " & T_LAB031 & _
                           " WHERE " & DBW("cdindex", cdindex, 2)
                         
            If CdVal1 <> "" Then
                GetComCdMST1 = GetComCdMST1 & " AND " & DBW("cdval1", CdVal1, 2)
                
                If CdVal2 <> "" Then
                    GetComCdMST1 = GetComCdMST1 & " AND " & DBW("cdval2", CdVal2, 2)
                End If
            End If
    End Select
End Function


Public Function SetComCdMST1(ByVal pUpdateFg As Boolean, ByVal cdindex As String, _
                             Optional ByVal CdVal1 As String = "", Optional ByVal CdVal2 As String = "", _
                             Optional ByVal Field1 As String = "", Optional ByVal Field2 As String = "", _
                             Optional ByVal Field3 As String = "", Optional ByVal Field4 As String = "", _
                             Optional ByVal Field5 As String = "", Optional ByVal Text1 As String = "", _
                             Optional ByVal Text2 As String = "") As String
'Coding By Legends
'To 공통코드 마스터1(COM002) pUpdateFg = True : Update, False : Insert
    If pUpdateFg Then
        SetComCdMST1 = " update " & T_LAB031 & _
                       " set " & _
                                   DBW("field1", Field1, 3) & _
                                   DBW("field2", Field2, 3) & _
                                   DBW("field3", Field3, 3) & _
                                   DBW("field4", Field4, 3) & _
                                   DBW("field5", Field5, 3) & _
                                   DBW("text1", Text1, 3) & _
                                   DBW("text2", Text2, 2) & _
                       " WHERE " & DBW("cdindex", cdindex, 2) & _
                       "   AND " & DBW("cdval1", CdVal1, 2) & _
                       "   AND " & DBW("cdval2", CdVal2, 2)
    Else
        SetComCdMST1 = "insert into " & T_LAB031 & _
                     " (cdindex,cdval1,cdval2,field1,field2,field3,field4,field5,text1,text2) values ( " & _
                     DBV("cdindex", cdindex, 1) & DBV("cdval1", CdVal1, 1) & DBV("cdval2", CdVal2, 1) & _
                     DBV("field1", Field1, 1) & DBV("field2", Field2, 1) & DBV("field3", Field3, 1) & _
                     DBV("field4", Field4, 1) & DBV("field5", Field5, 1) & DBV("text1", Text1, 1) & _
                     DBV("text2", Text2) & ")"
    End If
End Function

Public Function DelComCdMST1(ByVal cdindex As String, ByVal CdVal1 As String, _
                             ByVal CdVal2 As String) As String
'Coding By Legends
'To 공통코드 마스터1(COM002) Delete
    DelComCdMST1 = "delete FROM " & T_LAB031 & _
                   " WHERE " & DBW("cdindex", cdindex, 2) & _
                   "   AND " & DBW("cdval1", CdVal1, 2) & _
                   "   AND " & DBW("cdval2", CdVal2, 2)
End Function
                            
Public Function DelALLComCdMST1(ByVal cdindex As String, ByVal CdVal1 As String) As String
'Coding By Legends
'To 공통코드 마스터1(COM002) Delete
    DelALLComCdMST1 = "delete FROM " & T_LAB031 & _
                   " WHERE " & DBW("cdindex", cdindex, 2) & _
                   "   AND " & DBW("cdval1", CdVal1, 2)
End Function

Public Function GetComCdMST2(ByVal cdindex As String, Optional ByVal CdVal1 As String = "", _
                             Optional ByVal Field1 As String = "") As String
    '## 5.0.18: 이상대(2005-01-04)
    '   - LC3_WBCCode, LC3_NRBCCode 추가
    Select Case cdindex
        Case LC3_HighItem, LC3_ByPass, LC3_ICUTestCd, LC3_POCTestCd, LC3_ReportTesctCd, _
             LC3_WBCDiffCode, LC3_WBCCode, LC3_NRBCCode, "01", "02 "
            GetComCdMST2 = " SELECT a.cdindex,a.cdval1,a.field1,a.field2,a.field3,a.field4,a.field5,a.text1,a.text2," & _
                           "       b.testnm,b.abbrnm10 " & _
                           " FROM " & T_LAB032 & " a," & T_LAB001 & " b" & _
                           " WHERE " & DBW("a.cdindex", cdindex, 2) & _
                           " AND a.cdval1=b.testcd"
            If CdVal1 <> "" Then
                GetComCdMST2 = GetComCdMST2 & " AND " & DBW("a.cdval1", CdVal1, 2)
            End If
            
            If Field1 <> "" Then
                GetComCdMST2 = GetComCdMST2 & " AND " & DBW("a.field1", Field1, 2)
            End If
        Case "03", _
             "04"
            GetComCdMST2 = " SELECT a.cdindex,a.cdval1,a.field1,a.field2,a.field3,a.field4,a.field5,a.text1,a.text2," & _
                           "       b.testnm,b.abbrnm10 " & _
                           " FROM " & T_LAB032 & " a," & T_LAB001 & " b" & _
                           " WHERE " & DBW("a.cdindex", cdindex, 2) & _
                           " AND a.cdval1=b.testcd"
            If CdVal1 <> "" Then
                GetComCdMST2 = GetComCdMST2 & " AND " & DBW("a.cdval1", CdVal1, 2)
            End If
            
            If Field1 <> "" Then
                GetComCdMST2 = GetComCdMST2 & " AND " & DBW("a.field1", Field1, 2)
            End If
        
        Case Else
            GetComCdMST2 = "SELECT * FROM " & T_LAB032 & _
                           " WHERE " & DBW("cdindex", cdindex, 2)
            If CdVal1 <> "" Then
                GetComCdMST2 = GetComCdMST2 & " AND " & DBW("cdval1", CdVal1, 2)
            End If
            
            If Field1 <> "" Then
                GetComCdMST2 = GetComCdMST2 & " AND " & DBW("field1", Field1, 2)
            End If
    End Select
    
    GetComCdMST2 = GetComCdMST2 & " ORDER BY text1, cdval1"
End Function

Public Function SetComCdMST2(ByVal pUpdateFg As Boolean, ByVal cdindex As String, _
                             Optional ByVal CdVal1 As String = "", _
                             Optional ByVal Field1 As String = "", Optional ByVal Field2 As String = "", _
                             Optional ByVal Field3 As String = "", Optional ByVal Field4 As String = "", _
                             Optional ByVal Field5 As String = "", Optional ByVal Text1 As String = "", _
                             Optional ByVal Text2 As String = "") As String
'Coding By Legends
'To 공통코드 마스터2(COM003) pUpdateFg = True : Update, False : Insert
    If pUpdateFg Then
        SetComCdMST2 = " update " & T_LAB032 & _
                       " set " & _
                                   DBW("field1", Field1, 3) & _
                                   DBW("field2", Field2, 3) & _
                                   DBW("field3", Field3, 3) & _
                                   DBW("field4", Field4, 3) & _
                                   DBW("field5", Field5, 3) & _
                                   DBW("text1", Text1, 3) & _
                                   DBW("text2", Text2, 2) & _
                       " WHERE " & DBW("cdindex", cdindex, 2) & _
                       "   AND " & DBW("cdval1", CdVal1, 2)
    Else
        SetComCdMST2 = "insert into " & T_LAB032 & _
                       " (cdindex,cdval1,field1,field2,field3,field4,field5,text1,text2) values ( " & _
                       DBV("cdindex", cdindex, 1) & DBV("cdval1", CdVal1, 1) & DBV("field1", Field1, 1) & _
                       DBV("field2", Field2, 1) & DBV("field3", Field3, 1) & DBV("field4", Field4, 1) & _
                       DBV("field5", Field5, 1) & DBV("text1", Text1, 1) & DBV("text2", Text2) & ")"
    End If
End Function

Public Function DelComCdMST2(ByVal cdindex As String, ByVal CdVal1 As String) As String
'Coding By Legends
'To 공통코드 마스터2(COM003) Delete
    DelComCdMST2 = "delete FROM " & T_LAB032 & _
                   " WHERE " & DBW("cdindex", cdindex, 2) & _
                   "   AND " & DBW("cdval1", CdVal1, 2)
End Function

Public Function GetComCdTemp(ByVal cdindex As String, Optional ByVal CdVal1 As String = "", _
                           Optional ByVal Field1 As String = "") As String
'Coding By Legends
'FROM 템플릿 마스터(lab034)  SELECT
    
    GetComCdTemp = "SELECT * FROM " & T_LAB034 & _
                   " WHERE " & DBW("cdindex", cdindex, 2)
    If CdVal1 <> "" Then
        GetComCdTemp = GetComCdTemp & " AND " & DBW("cdval1", CdVal1, 2)
    End If
    
    If Field1 <> "" Then
        GetComCdTemp = GetComCdTemp & " AND " & DBW("field1", Field1, 2)
    End If
End Function

Public Function SetComCdTemp(ByVal pUpdateFg As Boolean, ByVal pCdindex As String, _
                             Optional pCdVal1 As String, Optional pField1 As String, _
                             Optional pField2 As String, Optional pText1 As String, _
                             Optional pText2 As String)
'Coding By Legends
'To 템플릿 마스터(COM004) pUpdatFg False : Insert, True : Update
    If pUpdateFg Then
        SetComCdTemp = " update " & T_LAB034 & _
                       " set " & _
                                   DBW("field1", pField1, 3) & _
                                   DBW("field2", pField2, 3) & _
                                   DBW("text1", pText1, 3) & _
                                   DBW("text2", pText2, 2) & _
                       " WHERE " & DBW("cdindex", pCdindex, 2) & _
                       "   AND " & DBW("cdval1", pCdVal1, 2)
    Else
        SetComCdTemp = "insert into " & T_LAB034 & _
                       " (cdindex,cdval1,field1,field2,text1,text2) values ( " & _
                       DBV("cdindex", pCdindex, 1) & DBV("cdval1", pCdVal1, 1) & DBV("field1", pField1, 1) & DBV("field2", pField2, 1) & _
                       DBV("text1", pText1, 1) & DBV("text2", pText2) & ")"
    End If
End Function

Public Function DelComCdTemp(ByVal pCdindex As String, ByVal pCdVal1 As String) As String
'Coding By Legends
'To 템플릿 마스터(COM004) Delete
    DelComCdTemp = "delete FROM " & T_LAB034 & _
                   " WHERE " & DBW("cdindex", pCdindex, 2) & _
                   "   AND " & DBW("cdval1", pCdVal1, 2)
End Function

Public Function GetComCdIdxComCd123(ByVal cdindex As String, ByVal CdVal1 As String, _
                                     Optional CdVal2 As String = "0") As String
'Coding By Legends
'FROM 공통코드인덱스(COM001), 공통코드마스터1(COM002), 2(COM003), 진단템플릿 마스터(COM004) SELECT
'For 공통코드마스터 출력

    Dim strComTbl As String
    
    Select Case cdindex
        Case "AC2"
            strComTbl = T_LAB031
        Case "AC3"
            strComTbl = T_LAB032
        Case "AC4"
            strComTbl = T_LAB034
    End Select

    GetComCdIdxComCd123 = "SELECT a.field1 as title, a.text1 as header, b.* " & _
                          " FROM " & T_COM001 & " a, " & strComTbl & " b " & _
                          " WHERE " & DBW("a.cdindex", cdindex, 2) & _
                          "   AND " & DBW("a.cdval1", CdVal1, 2) & _
                          "   AND " & DBW("a.cdval2", CdVal2, 2) & _
                          "   AND a.cdval1 = b.cdindex "
End Function



'+--------------------------------------------------------------------------------+"
'+ 프로젝트 설정에서 사용하는 Sql 문장의 집합                                     |
'+--------------------------------------------------------------------------------+"

Public Function GetSQLShowIndex2(Optional ByVal CdVal1 As String = "") As String
    If CdVal1 = "" Then
        GetSQLShowIndex2 = " SELECT * FROM " & T_LAB034 & _
                           " WHERE " & DBW("cdindex", LC5_Index, 2)
    Else
        GetSQLShowIndex2 = " SELECT * FROM " & T_LAB034 & _
                           " WHERE " & DBW("cdindex", LC5_Index, 2) & _
                           " AND " & DBW("cdval1", CdVal1, 2)
    End If
End Function

Public Function DB_SQLShowIndex2(ByVal CdVal1 As String, Optional ByVal FIEDL1 As String = "", _
                                 Optional ByVal Text1 As String = "", Optional ByVal Text2 As String = "", _
                                 Optional ByVal CHKFIELDS As String = "") As String
    If CHKFIELDS = "insert" Then
        DB_SQLShowIndex2 = " INSERT INTO " & T_LAB034 & " (cdindex,cdval1,field1,text1,text2) " & _
                           " VALUES (" & _
                             DBV("cdindex", LC5_Index, 1) & DBV("cdval1", CdVal1, 1) & DBV("field1", FIEDL1, 1) & _
                             DBV("text1", Text1, 1) & DBV("text2", Text2) & _
                                   ")"
    ElseIf CHKFIELDS = "delete" Then
        DB_SQLShowIndex2 = " delete " & T_LAB034 & _
                           " WHERE " & _
                                     DBW("cdindex", LC5_Index, 2) & _
                           " AND " & DBW("cdval1", CdVal1, 2)
    End If
End Function

