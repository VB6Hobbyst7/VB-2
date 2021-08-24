Attribute VB_Name = "modQuery"
Option Explicit

Public SQL  As String
Public RS   As ADODB.Recordset


'-- 사용자ID로 사용자명을 찾아온다.
Public Function Get_UserName(ByVal strUserID As String, Optional ByVal strUserPW As String) As String
    Dim pAdoRS      As ADODB.Recordset

On Error GoTo ErrorRoutine

    Get_UserName = ""
    
    SQL = ""
    SQL = SQL & "SELECT USER_NAME,USER_PW,USER_DEPART,USER_COMP " & vbCrLf
    SQL = SQL & "  FROM LBL_M_USER                              " & vbCrLf
    SQL = SQL & " WHERE USER_CD  = '" & strUserID & "'          " & vbCrLf
    SQL = SQL & "   AND USED_YN  = 'Y'                          " & vbCrLf
    If strUserPW <> "" Then
        SQL = SQL & "   AND USER_PW = '" & strUserPW & "'       " & vbCrLf
    End If
    
    Set pAdoRS = New ADODB.Recordset
    Call GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_UserName(ByVal strUserID As String, Optional ByVal strUserPW As String) As String")
    If Not pAdoRS Is Nothing Then
        If pAdoRS.EOF Then
            Get_UserName = ""
        Else
            Get_UserName = Trim(pAdoRS("USER_NAME") & "")
                            
            gKUKDO.USERID = strUserID
            gKUKDO.USERNM = Trim(pAdoRS("USER_NAME") & "")
            gKUKDO.USERGRD = Trim(pAdoRS("USER_COMP") & "")

        End If
        
        pAdoRS.Close
        Set pAdoRS = Nothing
    Else
        GoTo ErrorRoutine
    End If
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_UserName(ByVal strUserID As String, Optional ByVal strUserPW As String) As String")


End Function

'-- 사용자리스트 찾아온다.
Public Function Get_UserList(Optional ByVal pUserID As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset

On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT USER_CD,USER_NAME,USER_PW,USER_DEPART,USER_COMP "
    SQL = SQL & "     , USED_YN,REGIST_ID,REGIST_DT,MODIFY_ID,MODIFY_DT " & vbCrLf
    SQL = SQL & "  FROM LBL_M_USER                                      " & vbCrLf
    If pUserID <> "" Then
        SQL = SQL & " WHERE USER_CD =   '" & pUserID & "'               " & vbCrLf
    End If
    SQL = SQL & " ORDER BY USER_CD, USER_NAME                           " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_UserList(Optional ByVal pUserID As String) As ADODB.Recordset") Then
        Set Get_UserList = pAdoRS
    Else
        Set Get_UserList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_UserList(Optional ByVal pUserID As String) As ADODB.Recordset")

End Function

'-- 고객사리스트 찾아온다.
Public Function Get_CompList(Optional ByVal pCompCd As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    SQL = ""
    SQL = SQL & "SELECT COMP_CD,COMP_NAME,COMP_LINE,COMP_VIEW,COMP_DIS_NO "
    SQL = SQL & "     , USED_YN,REGIST_ID,REGIST_DT,MODIFY_ID,MODIFY_DT " & vbCrLf
    SQL = SQL & "  FROM LBL_M_COMP                                      " & vbCrLf
    If pCompCd <> "" Then
        SQL = SQL & " WHERE COMP_CD =   '" & pCompCd & "'               " & vbCrLf
    End If
    SQL = SQL & " ORDER BY COMP_DIS_NO,COMP_CD                          " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_CompList(Optional ByVal pCompCD As String) As ADODB.Recordset") Then
        Set Get_CompList = pAdoRS
    Else
        Set Get_CompList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_CompList(Optional ByVal pCompCD As String) As ADODB.Recordset")

End Function

'-- 고객사명만 찾아온다.
Public Function Get_CompList_Name(Optional pCompCd As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT COMP_NAME  " & vbCrLf
    SQL = SQL & "  FROM LBL_M_COMP                  " & vbCrLf
    SQL = SQL & " WHERE USED_YN = 'Y'               " & vbCrLf
    If pCompCd <> "" Then
        SQL = SQL & "   AND COMP_CD = '" & pCompCd & "' " & vbCrLf
    End If
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_CompList_Name(Optional pCompCd As String) As ADODB.Recordset") Then
        Set Get_CompList_Name = pAdoRS
    Else
        Set Get_CompList_Name = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_CompList_Name(Optional pCompCd As String) As ADODB.Recordset")

End Function

'-- TEMP 테이블
Public Function Get_TempMaster(ByVal pGubunCd As String, Optional pCode1 As String, Optional pCode2 As String, Optional pCode3 As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT CODE1,CODE2,CODE3,NAME1,NAME2,NAME3  " & vbCrLf
    SQL = SQL & "  FROM TEMP_MASTER                  " & vbCrLf
    SQL = SQL & " WHERE GUBUN_CD = '" & pGubunCd & "'               " & vbCrLf
    If pCode1 <> "" Then
        SQL = SQL & "   AND CODE1 = '" & pCode1 & "' " & vbCrLf
    End If
    If pCode2 <> "" Then
        SQL = SQL & "   AND CODE1 = '" & pCode2 & "' " & vbCrLf
    End If
    If pCode3 <> "" Then
        SQL = SQL & "   AND CODE1 = '" & pCode3 & "' " & vbCrLf
    End If
    SQL = SQL & " ORDER BY CODE1,CODE2,CODE3 "
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_TempMaster(ByVal pGubunCd As String, Optional pCode1 As String, Optional pCode2 As String, Optional pCode3 As String) As ADODB.Recordset") Then
        Set Get_TempMaster = pAdoRS
    Else
        Set Get_TempMaster = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_CompList_Name(Optional pCompCd As String) As ADODB.Recordset")

End Function


'-- 전체 고객사 코드/명만 찾아온다.
Public Function Get_CompList_CodeName() As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT COMP_CD,COMP_LINE,COMP_NAME    " & vbCrLf
    SQL = SQL & "  FROM LBL_M_COMP                              " & vbCrLf
    SQL = SQL & " WHERE USED_YN = 'Y'                           " & vbCrLf
    SQL = SQL & " ORDER BY COMP_NAME,COMP_LINE                  " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_CompList_CodeName() As ADODB.Recordset") Then
        Set Get_CompList_CodeName = pAdoRS
    Else
        Set Get_CompList_CodeName = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_CompList_CodeName() As ADODB.Recordset")

End Function

'-- 선택한 제품의 고객사 코드/명만 찾아온다.
Public Function Get_Comp_CodeName(ByVal pProdCd As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT C.COMP_CD,C.COMP_NAME, P.PROD_LENGTH   " & vbCrLf
    SQL = SQL & "  FROM LBL_M_COMP C, LBL_M_PROD P                      " & vbCrLf
    SQL = SQL & " WHERE C.COMP_CD = P.COMP_CD                           " & vbCrLf
    SQL = SQL & "   AND P.PROD_CD = '" & pProdCd & "'                   " & vbCrLf
    SQL = SQL & "   AND C.USED_YN = 'Y'                                 " & vbCrLf
    SQL = SQL & "   AND P.USED_YN = 'Y'                                 " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_Comp_CodeName(ByVal pProdCd As String) As ADODB.Recordset") Then
        Set Get_Comp_CodeName = pAdoRS
    Else
        Set Get_Comp_CodeName = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_Comp_CodeName(ByVal pProdCd As String) As ADODB.Recordset")

End Function


'-- 제품리스트 찾아온다.
Public Function Get_PackList(Optional ByVal pPackCD As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    SQL = ""
    SQL = SQL & "SELECT PACK_CD,PACK_NAME,PACK_CORE,PACK_DIA,PACK_DIS_NO "
    SQL = SQL & "     , PACK_CAT_WIDTH,PACK_PRO_WIDTH,PACK_PRO_LENGTH,PACK_CAT_GU" & vbCrLf
    SQL = SQL & "     , USED_YN,REGIST_ID,REGIST_DT,MODIFY_ID,MODIFY_DT " & vbCrLf
    SQL = SQL & "  FROM LBL_M_PACK                                      " & vbCrLf
    If pPackCD <> "" Then
        SQL = SQL & " WHERE PACK_CD =   '" & pPackCD & "'               " & vbCrLf
    End If
    SQL = SQL & " ORDER BY PACK_DIS_NO,PACK_CD                          " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_PackList(Optional ByVal pPackCD As String) As ADODB.Recordset") Then
        Set Get_PackList = pAdoRS
    Else
        Set Get_PackList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_PackList(Optional ByVal pPackCD As String) As ADODB.Recordset")

End Function

'-- 제품코드 리스트 찾아온다.
Public Function Get_ProdList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT PROD_CD , PROD_NAME, COMP_CD"
    SQL = SQL & ", PROD_LENGTH,PROD_MATERIAL_CD,EXPIR_MONTH,PROD_STOR_TEMP,PROD_SIZE,PROD_CHIMEI_PN"
    SQL = SQL & ", VENDER_CD,PROD_LINE_FA,PROD_SLIT_FA,PROD_CONTROL_YN,PROD_PCN_NO,USED_YN,ITEM_BARCODE"
    SQL = SQL & ", REGIST_ID,REGIST_DT,MODIFY_ID,MODIFY_DT"
    SQL = SQL & "  FROM LBL_M_PROD                                      " & vbCrLf
    SQL = SQL & " WHERE 1=1"
    If pProdCd <> "" Then
        SQL = SQL & "   AND PROD_CD =   '" & pProdCd & "'               " & vbCrLf
    End If
    If pCompCd <> "" And pCompCd <> "전체" Then
        SQL = SQL & "   AND COMP_CD =   '" & pCompCd & "'               " & vbCrLf
    End If
    SQL = SQL & " ORDER BY PROD_CD,COMP_CD                          " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Get_ProdList(Optional ByVal pProdCD As String, Optional ByVal pCompCD As String) As ADODB.Recordset") Then
        Set Get_ProdList = pAdoRS
    Else
        Set Get_ProdList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Get_ProdList(Optional ByVal pProdCD As String, Optional ByVal pCompCD As String) As ADODB.Recordset")

End Function

'-- 제품코드 리스트 찾아온다.
Public Function Get_MaxProdCode() As String
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT MAX(PROD_CD) AS PROD_CD    " & vbCrLf
    SQL = SQL & "  FROM LBL_M_PROD      " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    Call GetRecordset(AdoCn, SQL, pAdoRS, "")
    If Not pAdoRS Is Nothing Then
        If pAdoRS.EOF Then
            Get_MaxProdCode = "P0001"
        Else
            Get_MaxProdCode = Trim(pAdoRS("PROD_CD") & "")
            Get_MaxProdCode = Mid(Get_MaxProdCode, 2)
            Get_MaxProdCode = Get_MaxProdCode + 1
            Get_MaxProdCode = Format(Get_MaxProdCode, "0000")
            Get_MaxProdCode = "P" & Get_MaxProdCode
        End If
        
        pAdoRS.Close
        Set pAdoRS = Nothing
    Else
        GoTo ErrorRoutine
    End If
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "")

End Function


'-- 라벨정보 리스트 찾아온다.
Public Function Get_LabelList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    Dim pCallForm   As String
    
    pCallForm = "Public Function Get_LabelList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset"
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT LM.PROD_LABEL_CD            "
    SQL = SQL & "     , LM.PROD_CD                  "
    SQL = SQL & "     , LM.COMP_CD                  "
    SQL = SQL & "     , P.PROD_NAME                 "
    SQL = SQL & "     , P.PROD_LENGTH               " & vbCrLf
    SQL = SQL & "     , C.COMP_NAME                 "
    SQL = SQL & "     , LM.LABEL_PRT_NO             "
    SQL = SQL & "     , LM.LABEL_PRT_SIDE           " & vbCrLf
    SQL = SQL & "     , LM.PROD_LABEL_TYPE          " & vbCrLf
    SQL = SQL & "     , LM.LABEL_BAR_SIDE01_TYPE    "
    SQL = SQL & "     , LM.LABEL_BAR_SIDE02_TYPE    "
    SQL = SQL & "     , LM.PROD_MAX_TOT             " & vbCrLf
    SQL = SQL & "     , LM.USED_YN                  "
    SQL = SQL & "     , LM.REGIST_ID                "
    SQL = SQL & "     , LM.REGIST_DT                "
    SQL = SQL & "     , LM.MODIFY_ID                "
    SQL = SQL & "     , LM.MODIFY_DT                " & vbCrLf
    SQL = SQL & "  FROM LBL_LABEL_MASTER LM         "
    SQL = SQL & "     , LBL_M_PROD P                "
    SQL = SQL & "     , LBL_M_COMP C                " & vbCrLf
    SQL = SQL & " WHERE LM.PROD_CD   =   P.PROD_CD                                           " & vbCrLf
    SQL = SQL & "   AND LM.COMP_CD   =   P.COMP_CD                                           " & vbCrLf
    SQL = SQL & "   AND LM.COMP_CD   =   C.COMP_CD                                           " & vbCrLf
    If pProdCd <> "" And pProdCd <> "전체" Then
        SQL = SQL & "   AND LM.PROD_CD          =   '" & pProdCd & "'                               " & vbCrLf
    End If
    If pCompCd <> "" And pCompCd <> "전체" Then
        SQL = SQL & "   AND LM.COMP_CD          =   '" & pCompCd & "'                               " & vbCrLf
    End If
    If pLblType <> "" Then
        SQL = SQL & "   AND LM.PROD_LABEL_TYPE  =   '" & pLblType & "'                      " & vbCrLf
    End If
    SQL = SQL & " ORDER BY LM.PROD_CD, LM.COMP_CD, LM.PROD_LABEL_TYPE                          " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, pCallForm) Then
        Set Get_LabelList = pAdoRS
    Else
        Set Get_LabelList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, pCallForm)

End Function

'-- 바코드정보 리스트 찾아온다.
Public Function Get_BarList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    Dim pCallForm   As String
    
    pCallForm = "Public Function Get_BarList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset"
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT P.PROD_CD,P.PROD_NAME,P.PROD_LENGTH,C.COMP_CD,C.COMP_NAME " & vbCrLf
    SQL = SQL & ", B.BAR_CD,B.BAR_TYPE,B.BAR_GU   " & vbCrLf
    SQL = SQL & ", B.USED_YN, B.REGIST_ID, B.REGIST_DT, B.MODIFY_ID, B.MODIFY_DT            " & vbCrLf
    SQL = SQL & "  FROM LBL_BAR_INFO B, LBL_M_PROD P, LBL_M_COMP C                         " & vbCrLf
    SQL = SQL & " WHERE B.PROD_CD   =   P.PROD_CD                                           " & vbCrLf
    SQL = SQL & "   AND B.COMP_CD   =   P.COMP_CD                                           " & vbCrLf
    SQL = SQL & "   AND B.COMP_CD   =   C.COMP_CD                                           " & vbCrLf
    If pProdCd <> "" Then
        SQL = SQL & "   AND B.PROD_CD   =   '" & pProdCd & "'                               " & vbCrLf
    End If
    If pCompCd <> "" And pCompCd <> "전체" Then
        SQL = SQL & "   AND B.COMP_CD   =   '" & pCompCd & "'                               " & vbCrLf
    End If
    If pLblType <> "" Then
        SQL = SQL & "   AND B.PROD_LABEL_TYPE   =   '" & pLblType & "'                      " & vbCrLf
    End If
    SQL = SQL & " ORDER BY B.PROD_CD, C.COMP_CD, B.BAR_TYPE                          " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, pCallForm) Then
        Set Get_BarList = pAdoRS
    Else
        Set Get_BarList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, pCallForm)

End Function

'-- 라벨정보 마스터 찾아온다.
'Public Function Get_LabelMaster(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset
'    Dim pAdoRS      As ADODB.Recordset
'    Dim pCallForm   As String
'
'    pCallForm = "Public Function Get_LabelMaster(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset"
'
'On Error GoTo ErrorRoutine
'
'    SQL = ""
'    SQL = SQL & "SELECT "
'    SQL = SQL & "  M.PROD_LABEL_CD"
'    SQL = SQL & ", M.LABEL_ITEM_NO"
'    SQL = SQL & ", M.LABEL_ITEM_SEQ                           " & vbCrLf
'    SQL = SQL & ", M.LABEL_ITEM_NAME"
'    SQL = SQL & ", M.LABEL_ITEM_NAME_PRT      " & vbCrLf
'    SQL = SQL & ", M.BAR_CD                         " & vbCrLf
'    SQL = SQL & ", M.LABEL_ITEM_GU"
'    SQL = SQL & ", M.LABEL_ITEM_X_COORD"
'    SQL = SQL & ", M.LABEL_ITEM_Y_COORD                       " & vbCrLf
'    SQL = SQL & ", M.LABEL_ITEM_FONT"
'    SQL = SQL & ", M.LABEL_ITEM_BOLD"
'    SQL = SQL & ", M.LABEL_ITEM_ROT                              " & vbCrLf
'    SQL = SQL & ", M.USED_YN"
'    SQL = SQL & ", M.REGIST_ID"
'    SQL = SQL & ", M.REGIST_DT"
'    SQL = SQL & ", M.MODIFY_ID"
'    SQL = SQL & ", M.MODIFY_DT    " & vbCrLf
'    SQL = SQL & "  FROM LBL_LABEL_MASTER I"
'    SQL = SQL & "     , LBL_LABEL_DETAIL M                          " & vbCrLf
'    SQL = SQL & " WHERE I.PROD_CD           =   M.PROD_CD                           " & vbCrLf
'    SQL = SQL & "   AND I.COMP_CD           =   M.COMP_CD                           " & vbCrLf
'    SQL = SQL & "   AND I.PROD_LABEL_TYPE   =   M.PROD_LABEL_TYPE                   " & vbCrLf
'    If pProdCd <> "" Then
'        SQL = SQL & "   AND I.PROD_CD   =   '" & pProdCd & "'                       " & vbCrLf
'    End If
'    If pCompCd <> "" And pCompCd <> "전체" Then
'        SQL = SQL & "   AND I.COMP_CD   =   '" & pCompCd & "'                       " & vbCrLf
'    End If
'    If pLblType <> "" Then
'        SQL = SQL & "   AND I.PROD_LABEL_TYPE   =   '" & pLblType & "'              " & vbCrLf
'    End If
'    SQL = SQL & " ORDER BY M.LABEL_ITEM_NO, M.LABEL_ITEM_SEQ                        " & vbCrLf
'
'    Set pAdoRS = New ADODB.Recordset
'
'    If GetRecordset(AdoCn, SQL, pAdoRS, pCallForm) Then
'        Set Get_LabelMaster = pAdoRS
'    Else
'        Set Get_LabelMaster = Nothing
'    End If
'
'    Set pAdoRS = Nothing
'
'Exit Function
'
'ErrorRoutine:
'    Set pAdoRS = Nothing
'    Call DBErrorSet(AdoCn, SQL, pCallForm)
'
'End Function

'-- 라벨정보 마스터 찾아온다.
Public Function Get_LabelMasterList(ByVal pProdLabelCd As String, Optional ByVal pItemNo As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    Dim pCallForm   As String
    
    pCallForm = "Public Function Get_LabelMasterList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset"
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  M.PROD_LABEL_CD"
    SQL = SQL & ", D.LABEL_ITEM_NO"
    SQL = SQL & ", D.LABEL_ITEM_SEQ                           " & vbCrLf
    SQL = SQL & ", D.LABEL_ITEM_NAME"
    SQL = SQL & ", D.LABEL_ITEM_NAME_PRT      " & vbCrLf
    SQL = SQL & ", D.BAR_CD                         " & vbCrLf
    SQL = SQL & ", D.LABEL_ITEM_GU"
    SQL = SQL & ", D.LABEL_ITEM_X_COORD"
    SQL = SQL & ", D.LABEL_ITEM_Y_COORD                       " & vbCrLf
    SQL = SQL & ", D.LABEL_ITEM_FONT"
    SQL = SQL & ", D.LABEL_ITEM_BOLD"
    SQL = SQL & ", D.LABEL_ITEM_ROT                              " & vbCrLf
    SQL = SQL & ", D.USED_YN"
    SQL = SQL & ", D.REGIST_ID"
    SQL = SQL & ", D.REGIST_DT"
    SQL = SQL & ", D.MODIFY_ID"
    SQL = SQL & ", D.MODIFY_DT    " & vbCrLf
    SQL = SQL & "  FROM LBL_LABEL_MASTER M"
    SQL = SQL & "     , LBL_LABEL_DETAIL D                          " & vbCrLf
    SQL = SQL & " WHERE M.PROD_LABEL_CD     =  D.PROD_LABEL_CD           " & vbCrLf
    SQL = SQL & "   AND M.PROD_LABEL_CD     =  '" & pProdLabelCd & "'                       " & vbCrLf
    If pItemNo <> "" Then
        SQL = SQL & "   AND D.LABEL_ITEM_NO   =   '" & pItemNo & "'              " & vbCrLf
    End If
    SQL = SQL & " ORDER BY M.PROD_LABEL_CD, D.LABEL_ITEM_SEQ                        " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, pCallForm) Then
        Set Get_LabelMasterList = pAdoRS
    Else
        Set Get_LabelMasterList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, pCallForm)

End Function

'-- 라벨정보 마스터 찾아온다.
Public Function Get_LabelDetail(ByVal pProdLabelCd As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    Dim pCallForm   As String
    
    pCallForm = "Public Function Get_LabelDetail(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset"
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  M.LABEL_ITEM_NO"
    SQL = SQL & ", M.LABEL_ITEM_SEQ                           " & vbCrLf
    SQL = SQL & ", M.LABEL_ITEM_NAME"
'    SQL = SQL & ", M.LABEL_ITEM_MEMO"
    SQL = SQL & ", M.LABEL_ITEM_NAME_PRT      " & vbCrLf
    SQL = SQL & ", M.BAR_CD                         " & vbCrLf
    SQL = SQL & ", M.LABEL_ITEM_GU"
    SQL = SQL & ", M.LABEL_ITEM_X_COORD"
    SQL = SQL & ", M.LABEL_ITEM_Y_COORD                       " & vbCrLf
    SQL = SQL & ", M.LABEL_ITEM_FONT"
    'SQL = SQL & ", M.LABEL_ITEM_FONTSIZE                     " & vbCrLf
    'SQL = SQL & ", M.LABEL_ITEM_BOLD"
    SQL = SQL & ", M.LABEL_ITEM_ROT                              " & vbCrLf
    SQL = SQL & ", M.USED_YN"
    SQL = SQL & ", M.REGIST_ID"
    SQL = SQL & ", M.REGIST_DT"
    SQL = SQL & ", M.MODIFY_ID"
    SQL = SQL & ", M.MODIFY_DT    " & vbCrLf
    SQL = SQL & "  FROM LBL_LABEL_MASTER I"
    SQL = SQL & "     , LBL_LABEL_DETAIL M                          " & vbCrLf
    SQL = SQL & " WHERE I.PROD_LABEL_CD     =   M.PROD_LABEL_CD       " & vbCrLf
'    SQL = SQL & "   AND I.COMP_CD           =   M.COMP_CD                           " & vbCrLf
'    SQL = SQL & "   AND I.PROD_LABEL_TYPE   =   M.PROD_LABEL_TYPE                   " & vbCrLf
    If pProdLabelCd <> "" Then
        SQL = SQL & "   AND I.PROD_LABEL_CD   =   '" & pProdLabelCd & "'                       " & vbCrLf
    End If
'    If pCompCd <> "" And pCompCd <> "전체" Then
'        SQL = SQL & "   AND I.COMP_CD   =   '" & pCompCd & "'                       " & vbCrLf
'    End If
'    If pLblType <> "" Then
'        SQL = SQL & "   AND I.PROD_LABEL_TYPE   =   '" & pLblType & "'              " & vbCrLf
'    End If
    SQL = SQL & " ORDER BY M.LABEL_ITEM_NO, M.LABEL_ITEM_SEQ                        " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, pCallForm) Then
        Set Get_LabelDetail = pAdoRS
    Else
        Set Get_LabelDetail = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, pCallForm)

End Function

'-- 라벨마스터 리스트 찾아온다.
Public Function Get_LabelMaster(ByVal pProdLabelCd As String) As ADODB.Recordset

'Public Function Get_LabelMasterList(Optional ByVal pProdCd As String, _
                                    Optional ByVal pCompCd As String, _
                                    Optional ByVal pLblType As String, _
                                    Optional ByVal pItemNo As String) As ADODB.Recordset
    
    Dim pAdoRS      As ADODB.Recordset
    Dim pCallForm   As String
    
    pCallForm = "Get_LabelMaster(ByVal pProdLabelCd As String, Optional ByVal pItemNo As String)"
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  PROD_LABEL_CD"
    SQL = SQL & ", LABEL_ITEM_NO " & vbCrLf
    SQL = SQL & ", LABEL_ITEM_SEQ"
    SQL = SQL & ", LABEL_ITEM_NAME "
    SQL = SQL & ", LABEL_ITEM_NAME_PRT"
    SQL = SQL & ", BAR_CD"
    SQL = SQL & ", LABEL_ITEM_GU"
    SQL = SQL & ", LABEL_ITEM_X_COORD"
    SQL = SQL & ", LABEL_ITEM_Y_COORD"
    SQL = SQL & ", LABEL_ITEM_FONT"
    SQL = SQL & ", LABEL_ITEM_ROT"
    SQL = SQL & ", USED_YN"
    SQL = SQL & ", REGIST_ID"
    SQL = SQL & ", REGIST_DT"
    SQL = SQL & ", MODIFY_ID"
    SQL = SQL & ", MODIFY_DT"
    SQL = SQL & "  FROM LBL_LABEL_DETAIL                                                 " & vbCrLf
    SQL = SQL & " WHERE 1 = 1                                                           " & vbCrLf
    If pProdLabelCd <> "" Then
        SQL = SQL & "   AND PROD_LABEL_CD = '" & pProdLabelCd & "'                       " & vbCrLf
    End If
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, pCallForm) Then
        Set Get_LabelMaster = pAdoRS
    Else
        Set Get_LabelMaster = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, pCallForm)

End Function

'-- 바코드마스터 리스트 찾아온다.
Public Function Get_BarMasterList(Optional ByVal pProdCd As String, _
                                    Optional ByVal pCompCd As String, _
                                    Optional ByVal pLblType As String, _
                                    Optional ByVal pItemNo As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    Dim pCallForm   As String
    
    pCallForm = "Public Function Get_BarMasterList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLblType As String) As ADODB.Recordset"
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT PROD_CD, COMP_CD, PROD_LABEL_TYPE, LABEL_ITEM_NO " & vbCrLf
    SQL = SQL & ", LABEL_ITEM_SEQ, LABEL_ITEM_NAME , LABEL_ITEM_NAME_PRT"
    SQL = SQL & ", LABEL_ITEM_BAR_GU, LABEL_ITEM_BAR_CD, LABEL_ITEM_X_COORD, LABEL_ITEM_Y_COORD"
    SQL = SQL & ", LABEL_ITEM_FONTNAME,LABEL_ITEM_FONTSIZE,LABEL_ITEM_BOLD,LABEL_ITEM_ROT"
    SQL = SQL & ", USED_YN, REGIST_ID, REGIST_DT, MODIFY_ID, MODIFY_DT"
    SQL = SQL & "  FROM LBL_LABEL_DETAIL                                                 " & vbCrLf
    SQL = SQL & " WHERE 1 = 1                                                           " & vbCrLf
    If pProdCd <> "" Then
        SQL = SQL & "   AND PROD_CD         =   '" & pProdCd & "'                       " & vbCrLf
    End If
    If pCompCd <> "" And pCompCd <> "전체" Then
        SQL = SQL & "   AND COMP_CD         =   '" & pCompCd & "'                       " & vbCrLf
    End If
    If pLblType <> "" Then
        SQL = SQL & "   AND PROD_LABEL_TYPE =   '" & pLblType & "'                      " & vbCrLf
    End If
    If pItemNo <> "" Then
        SQL = SQL & "   AND LABEL_ITEM_NO   =   '" & pItemNo & "'                       " & vbCrLf
    End If
    SQL = SQL & " ORDER BY PROD_CD,COMP_CD,PROD_LABEL_TYPE                              " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, pCallForm) Then
        Set Get_BarMasterList = pAdoRS
    Else
        Set Get_BarMasterList = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, pCallForm)

End Function

'-- 제품코드 리스트 찾아온다.
Public Function Get_ProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String) As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  PROD_CD "
    SQL = SQL & ", PROD_NAME"
    SQL = SQL & ", COMP_CD"
    SQL = SQL & ", PROD_LENGTH"
    SQL = SQL & "  FROM LBL_M_PROD                                      " & vbCrLf
    SQL = SQL & " WHERE 1=1"
    If pProdCd <> "" Then
        SQL = SQL & "   AND PROD_CD =   '" & pProdCd & "'               " & vbCrLf
    End If
    If pCompCd <> "" And pCompCd <> "전체" Then
        SQL = SQL & "   AND COMP_CD =   '" & pCompCd & "'               " & vbCrLf
    End If
    SQL = SQL & "   AND USED_YN = 'Y'                               " & vbCrLf
    SQL = SQL & " ORDER BY PROD_CD,COMP_CD                          " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Get_ProdList_CodeName(Optional ByVal pProdCD As String, Optional ByVal pCompCD As String) As ADODB.Recordset") Then
        Set Get_ProdList_CodeName = pAdoRS
    Else
        Set Get_ProdList_CodeName = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Get_ProdList_CodeName(Optional ByVal pProdCD As String, Optional ByVal pCompCD As String) As ADODB.Recordset")

End Function


'-- 자재리스트를  찾아온다.
Public Function Get_Material() As ADODB.Recordset
    Dim pAdoRS      As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    SQL = ""
    SQL = SQL & "SELECT MAT_CD,MAT_NAME,MAT_DIS_NO " & vbCrLf
    SQL = SQL & "  FROM LBL_M_MATERIAL             " & vbCrLf
    SQL = SQL & " ORDER BY MAT_DIS_NO ,MAT_CD      " & vbCrLf
    
    Set pAdoRS = New ADODB.Recordset
    
    If GetRecordset(AdoCn, SQL, pAdoRS, "Public Function Get_Material() As ADODB.Recordset") Then
        Set Get_Material = pAdoRS
    Else
        Set Get_Material = Nothing
    End If
    
    Set pAdoRS = Nothing
    
Exit Function

ErrorRoutine:
    Set pAdoRS = Nothing
    Call DBErrorSet(AdoCn, SQL, "Public Function Get_Material() As ADODB.Recordset")

End Function


'-- 사용자 저장
Public Function Set_User(ByVal pState As String) As Boolean
    
On Error GoTo ErrorRoutine
    
    Set_User = False
    
    If pState = "IN" Then
        SQL = ""
        SQL = SQL & "INSERT INTO LBL_M_USER "
        SQL = SQL & "(USER_CD,USER_NAME,USER_PW,USER_DEPART,USER_COMP"
        SQL = SQL & ",USED_YN,REGIST_ID,REGIST_DT)  " & vbCrLf
        SQL = SQL & "  VALUES                       " & vbCrLf
        SQL = SQL & "('" & gUSER.ID & "'"
        SQL = SQL & ",'" & gUSER.NAME & "'"
        SQL = SQL & ",'" & gUSER.PW & "'"
        SQL = SQL & ",'" & gUSER.DEPT & "'"
        SQL = SQL & ",'" & gUSER.COMP & "'"
        SQL = SQL & ",'" & gUSER.YN & "'"
        SQL = SQL & ",'" & gKUKDO.USERID & "'"
        SQL = SQL & ",'" & gsDBDateTime & "')"
    ElseIf pState = "UP" Then
        SQL = ""
        SQL = SQL & "UPDATE LBL_M_USER SET" & vbCrLf
        SQL = SQL & "  USER_NAME    = '" & gUSER.NAME & "'" & vbCrLf
        SQL = SQL & ", USER_PW      = '" & gUSER.PW & "'" & vbCrLf
        SQL = SQL & ", USER_DEPART  = '" & gUSER.DEPT & "'" & vbCrLf
        SQL = SQL & ", USER_COMP    = '" & gUSER.COMP & "'" & vbCrLf
        SQL = SQL & ", USED_YN      = '" & gUSER.YN & "'" & vbCrLf
        SQL = SQL & ", MODIFY_ID    = '" & gKUKDO.USERID & "'" & vbCrLf
        SQL = SQL & ", MODIFY_DT    = '" & gsDBDateTime & "'" & vbCrLf
        SQL = SQL & " WHERE USER_CD = '" & gUSER.ID & "'" & vbCrLf
    ElseIf pState = "DEL" Then
        SQL = ""
        SQL = SQL & "DELETE FROM LBL_M_USER " & vbCrLf
        SQL = SQL & " WHERE USER_CD = '" & gUSER.ID & "'" & vbCrLf
    End If
    
    Call DBExec(AdoCn, SQL)
    Set_User = True
Exit Function

ErrorRoutine:
    Set_User = False
    Call DBErrorSet(AdoCn, SQL, "Public Function Set_User(ByVal pState As String) As Boolean")

End Function

'-- 자재코드 저장
Public Function Set_Mat(ByVal pState As String) As Boolean
    
On Error GoTo ErrorRoutine
    
    Set_Mat = False
    
    If pState = "IN" Then
        SQL = ""
        SQL = SQL & "INSERT INTO LBL_M_MATERIAL "
        SQL = SQL & "(MAT_CD,MAT_NAME,MAT_DIS_NO"
        SQL = SQL & ",USED_YN,REGIST_ID,REGIST_DT)  " & vbCrLf
        SQL = SQL & "  VALUES                       " & vbCrLf
        SQL = SQL & "('" & gMAT.CD & "'"
        SQL = SQL & ",'" & gMAT.NAME & "'"
        SQL = SQL & ",'" & gMAT.DISNO & "'"
        SQL = SQL & ",'" & gMAT.YN & "'"
        SQL = SQL & ",'" & gKUKDO.USERID & "'"
        SQL = SQL & ",'" & gsDBDateTime & "')"
    ElseIf pState = "UP" Then
        SQL = ""
        SQL = SQL & "UPDATE LBL_M_MATERIAL SET" & vbCrLf
        SQL = SQL & "  MAT_NAME     = '" & gMAT.NAME & "'" & vbCrLf
        SQL = SQL & ", MAT_DIS_NO   = '" & gMAT.DISNO & "'" & vbCrLf
        SQL = SQL & ", USED_YN      = '" & gMAT.YN & "'" & vbCrLf
        SQL = SQL & ", MODIFY_ID    = '" & gKUKDO.USERID & "'" & vbCrLf
        SQL = SQL & ", MODIFY_DT    = '" & gsDBDateTime & "'" & vbCrLf
        SQL = SQL & " WHERE MAT_CD = '" & gMAT.CD & "'" & vbCrLf
    ElseIf pState = "DEL" Then
        SQL = ""
        SQL = SQL & "DELETE FROM LBL_M_MATERIAL " & vbCrLf
        SQL = SQL & " WHERE MAT_CD = '" & gMAT.CD & "'" & vbCrLf
    End If
    
    Call DBExec(AdoCn, SQL)
    Set_Mat = True
Exit Function

ErrorRoutine:
    Set_Mat = False
    Call DBErrorSet(AdoCn, SQL, "Public Function Set_Mat(ByVal pState As String) As Boolean")

End Function


'-- 고객사 저장
Public Function Set_Comp(ByVal pState As String) As Boolean
    
On Error GoTo ErrorRoutine
    
    Set_Comp = False
    
    If pState = "IN" Then
        SQL = ""
        SQL = SQL & "INSERT INTO LBL_M_COMP "
        SQL = SQL & "(COMP_CD,COMP_NAME,COMP_LINE,COMP_VIEW,COMP_DIS_NO"
        SQL = SQL & ",USED_YN,REGIST_ID,REGIST_DT)  " & vbCrLf
        SQL = SQL & "  VALUES                       " & vbCrLf
        SQL = SQL & "('" & gComp.CD & "'"
        SQL = SQL & ",'" & gComp.NAME & "'"
        SQL = SQL & ",'" & gComp.LINE & "'"
        SQL = SQL & ",'" & gComp.VIEW & "'"
        SQL = SQL & ",'" & gComp.DISNO & "'"
        SQL = SQL & ",'" & gComp.YN & "'"
        SQL = SQL & ",'" & gKUKDO.USERID & "'"
        SQL = SQL & ",'" & gsDBDateTime & "')"
    ElseIf pState = "UP" Then
        SQL = ""
        SQL = SQL & "UPDATE LBL_M_COMP SET" & vbCrLf
        SQL = SQL & "  COMP_NAME    = '" & gComp.NAME & "'  " & vbCrLf
        SQL = SQL & ", COMP_LINE    = '" & gComp.LINE & "'  " & vbCrLf
        SQL = SQL & ", COMP_VIEW    = '" & gComp.VIEW & "'  " & vbCrLf
        SQL = SQL & ", COMP_DIS_NO  = '" & gComp.DISNO & "' " & vbCrLf
        SQL = SQL & ", USED_YN      = '" & gComp.YN & "'    " & vbCrLf
        SQL = SQL & ", MODIFY_ID    = '" & gKUKDO.USERID & "'" & vbCrLf
        SQL = SQL & ", MODIFY_DT    = '" & gsDBDateTime & "'" & vbCrLf
        SQL = SQL & " WHERE COMP_CD = '" & gComp.CD & "'" & vbCrLf
    ElseIf pState = "DEL" Then
        SQL = ""
        SQL = SQL & "DELETE FROM LBL_M_COMP " & vbCrLf
        SQL = SQL & " WHERE COMP_CD = '" & gComp.CD & "'" & vbCrLf
    End If
    
    Call DBExec(AdoCn, SQL)
    Set_Comp = True
Exit Function

ErrorRoutine:
    Set_Comp = False
    Call DBErrorSet(AdoCn, SQL, "Public Function Set_Comp(ByVal pState As String) As Boolean")

End Function

'-- 포장코드 저장
Public Function Set_Pack(ByVal pState As String) As Boolean
    
On Error GoTo ErrorRoutine
    
    Set_Pack = False
    
    If pState = "IN" Then
        SQL = ""
        SQL = SQL & "INSERT INTO LBL_M_PACK "
        SQL = SQL & "(PACK_CD,PACK_NAME,PACK_CORE,PACK_DIA"
        SQL = SQL & ",PACK_CAT_WIDTH,PACK_PRO_WIDTH,PACK_PRO_LENGTH,PACK_CAT_GU"
        SQL = SQL & ",PACK_DIS_NO,USED_YN,REGIST_ID,REGIST_DT)  " & vbCrLf
        SQL = SQL & "  VALUES                       " & vbCrLf
        SQL = SQL & "('" & gPack.CD & "'"
        SQL = SQL & ",'" & gPack.NAME & "'"
        SQL = SQL & ",'" & gPack.CORE & "'"
        SQL = SQL & ",'" & gPack.DIA & "'"
        SQL = SQL & ",'" & gPack.CWID & "'"
        SQL = SQL & ",'" & gPack.PWID & "'"
        SQL = SQL & ",'" & gPack.pLen & "'"
        SQL = SQL & ",'" & gPack.CGBN & "'"
        SQL = SQL & ",'" & gPack.DISNO & "'"
        SQL = SQL & ",'" & gPack.YN & "'"
        SQL = SQL & ",'" & gKUKDO.USERID & "'"
        SQL = SQL & ",'" & gsDBDateTime & "')"
    ElseIf pState = "UP" Then
        SQL = ""
        SQL = SQL & "UPDATE LBL_M_PACK SET" & vbCrLf
        SQL = SQL & "  PACK_NAME    = '" & gPack.NAME & "'  " & vbCrLf
        SQL = SQL & ", PACK_CORE    = '" & gPack.CORE & "'  " & vbCrLf
        SQL = SQL & ", PACK_DIA     = '" & gPack.DIA & "'   " & vbCrLf
        SQL = SQL & ", PACK_CAT_WIDTH  = '" & gPack.CWID & "'  " & vbCrLf
        SQL = SQL & ", PACK_PRO_WIDTH  = '" & gPack.PWID & "'   " & vbCrLf
        SQL = SQL & ", PACK_PRO_LENGTH = '" & gPack.pLen & "'  " & vbCrLf
        SQL = SQL & ", PACK_CAT_GU     = '" & gPack.CGBN & "'   " & vbCrLf
        SQL = SQL & ", PACK_DIS_NO  = '" & gPack.DISNO & "' " & vbCrLf
        SQL = SQL & ", USED_YN      = '" & gPack.YN & "'    " & vbCrLf
        SQL = SQL & ", MODIFY_ID    = '" & gKUKDO.USERID & "'" & vbCrLf
        SQL = SQL & ", MODIFY_DT    = '" & gsDBDateTime & "'" & vbCrLf
        SQL = SQL & " WHERE PACK_CD = '" & gPack.CD & "'" & vbCrLf
    ElseIf pState = "DEL" Then
        SQL = ""
        SQL = SQL & "DELETE FROM LBL_M_PACK " & vbCrLf
        SQL = SQL & " WHERE PACK_CD = '" & gPack.CD & "'" & vbCrLf
    End If
    
    Call DBExec(AdoCn, SQL)
    Set_Pack = True
Exit Function

ErrorRoutine:
    Set_Pack = False
    Call DBErrorSet(AdoCn, SQL, "Public Function Set_Pack(ByVal pState As String) As Boolean")

End Function

'-- 제품마스터 저장
Public Function Set_Prod(ByVal pState As String) As Boolean
    
On Error GoTo ErrorRoutine
    
    Set_Prod = False
        
    If pState = "IN" Then
        SQL = ""
        SQL = SQL & "INSERT INTO LBL_M_PROD "
        SQL = SQL & "(PROD_CD, PROD_NAME, COMP_CD"
        SQL = SQL & ",PROD_LENGTH, PROD_MATERIAL_CD, EXPIR_MONTH, PROD_STOR_TEMP, PROD_SIZE, PROD_CHIMEI_PN"
        SQL = SQL & ",VENDER_CD, PROD_LINE_FA, PROD_SLIT_FA, PROD_CONTROL_YN, PROD_PCN_NO"
        SQL = SQL & ",USED_YN, REGIST_ID,REGIST_DT)  " & vbCrLf
        SQL = SQL & "  VALUES                       " & vbCrLf
        SQL = SQL & "('" & gProd.CD & "'"
        SQL = SQL & ",'" & gProd.NAME & "'"
        SQL = SQL & ",'" & gProd.COMPCD & "'"
        SQL = SQL & ",'" & gProd.LEN & "'"
        SQL = SQL & ",'" & gProd.METCD & "'"
        SQL = SQL & ",'" & gProd.MONTH & "'"
        SQL = SQL & ",'" & gProd.TEMP & "'"
        SQL = SQL & ",'" & gProd.SIZE & "'"
        SQL = SQL & ",'" & gProd.CHPN & "'"
        SQL = SQL & ",'" & gProd.VDCD & "'"
        SQL = SQL & ",'" & gProd.LINEFA & "'"
        SQL = SQL & ",'" & gProd.SLITFA & "'"
        SQL = SQL & ",'" & gProd.CTYN & "'"
        SQL = SQL & ",'" & gProd.PCNNO & "'"
        SQL = SQL & ",'" & gProd.YN & "'"
'        SQL = SQL & ",'" & gProd.BAR & "'"
        SQL = SQL & ",'" & gKUKDO.USERID & "'"
        SQL = SQL & ",'" & gsDBDateTime & "')"
    ElseIf pState = "UP" Then
        If gProd.CD <> "" And gProd.COMPCD <> "" Then
            SQL = ""
            SQL = SQL & "UPDATE LBL_M_PROD SET" & vbCrLf
            SQL = SQL & "  PROD_NAME        = '" & gProd.NAME & "'  " & vbCrLf
            SQL = SQL & ", PROD_LENGTH      = '" & gProd.LEN & "'  " & vbCrLf
            SQL = SQL & ", PROD_MATERIAL_CD = '" & gProd.METCD & "'   " & vbCrLf
            SQL = SQL & ", EXPIR_MONTH      = '" & gProd.MONTH & "'  " & vbCrLf
            SQL = SQL & ", PROD_STOR_TEMP   = '" & gProd.TEMP & "'   " & vbCrLf
            SQL = SQL & ", PROD_SIZE        = '" & gProd.SIZE & "'  " & vbCrLf
            SQL = SQL & ", PROD_CHIMEI_PN   = '" & gProd.CHPN & "'   " & vbCrLf
            SQL = SQL & ", VENDER_CD        = '" & gProd.VDCD & "' " & vbCrLf
            SQL = SQL & ", PROD_LINE_FA     = '" & gProd.LINEFA & "' " & vbCrLf
            SQL = SQL & ", PROD_SLIT_FA     = '" & gProd.SLITFA & "' " & vbCrLf
            SQL = SQL & ", PROD_CONTROL_YN  = '" & gProd.CTYN & "' " & vbCrLf
            SQL = SQL & ", PROD_PCN_NO      = '" & gProd.PCNNO & "' " & vbCrLf
            SQL = SQL & ", USED_YN          = '" & gProd.YN & "'    " & vbCrLf
'            SQL = SQL & ", ITEM_BARCODE     = '" & gProd.BAR & "' " & vbCrLf
            SQL = SQL & ", MODIFY_ID        = '" & gKUKDO.USERID & "'" & vbCrLf
            SQL = SQL & ", MODIFY_DT        = '" & gsDBDateTime & "'" & vbCrLf
            SQL = SQL & " WHERE PROD_CD     = '" & gProd.CD & "'" & vbCrLf
            SQL = SQL & "   AND COMP_CD     = '" & gProd.COMPCD & "'" & vbCrLf
        End If
    ElseIf pState = "DEL" Then
        If gProd.CD <> "" And gProd.COMPCD <> "" Then
            SQL = ""
            SQL = SQL & "DELETE FROM LBL_M_PROD " & vbCrLf
            SQL = SQL & " WHERE PROD_CD = '" & gProd.CD & "'" & vbCrLf
            SQL = SQL & "   AND COMP_CD = '" & gProd.COMPCD & "'" & vbCrLf
        End If
    End If
    
    Call DBExec(AdoCn, SQL)
    
    Set_Prod = True

Exit Function

ErrorRoutine:
    Set_Prod = False
    Call DBErrorSet(AdoCn, SQL, "Public Function Set_Prod(ByVal pState As String) As Boolean")

End Function

'-- 라벨정보 저장
Public Function Set_Label_Master(ByVal pState As String) As Boolean
    
On Error GoTo ErrorRoutine
    
    Set_Label_Master = False
        
    If pState = "IN" Then
        SQL = ""
        SQL = SQL & "INSERT INTO LBL_LABEL_MASTER "
        SQL = SQL & "(PROD_LABEL_CD"
        SQL = SQL & ",PROD_CD"
        SQL = SQL & ",COMP_CD"
        SQL = SQL & ",PROD_LABEL_TYPE"
        SQL = SQL & ",LABEL_PRT_NO"
        SQL = SQL & ",LABEL_PRT_SIDE"
        SQL = SQL & ",LABEL_BAR_SIDE01_TYPE"
        SQL = SQL & ",LABEL_BAR_SIDE02_TYPE"
        SQL = SQL & ",PROD_MAX_TOT"
        SQL = SQL & ",USED_YN"
        SQL = SQL & ",REGIST_ID"
        SQL = SQL & ",REGIST_DT)  " & vbCrLf
        SQL = SQL & "  VALUES                       " & vbCrLf
        SQL = SQL & "('" & gLblMaster.LABELCD & "'"
        SQL = SQL & ",'" & gLblMaster.PRODCD & "'"
        SQL = SQL & ",'" & gLblMaster.COMPCD & "'"
        SQL = SQL & ",'" & gLblMaster.LBLTYPE & "'"
        SQL = SQL & "," & IIf(gLblMaster.LBLPRTNO = "", "0", gLblMaster.LBLPRTNO)
        SQL = SQL & ",'" & gLblMaster.LBLPRTSIDE & "'"
        SQL = SQL & ",'" & gLblMaster.LBLBARSIDE1 & "'"
        SQL = SQL & ",'" & gLblMaster.LBLBARSIDE2 & "'"
        SQL = SQL & "," & IIf(gLblMaster.PRODMAXTOT = "", "0", gLblMaster.PRODMAXTOT)
        SQL = SQL & ",'" & gLblMaster.YN & "'"
        SQL = SQL & ",'" & gKUKDO.USERID & "'"
        SQL = SQL & ",'" & gsDBDateTime & "')"
    ElseIf pState = "UP" Then
        'If gLblMaster.PRODCD <> "" And gLblMaster.COMPCD <> "" And gLblMaster.LBLTYPE <> "" Then
            SQL = ""
            SQL = SQL & "UPDATE LBL_LABEL_MASTER SET" & vbCrLf
            SQL = SQL & "  PROD_CD                  = '" & gLblMaster.PRODCD & "'" & vbCrLf
            SQL = SQL & ", COMP_CD                  = '" & gLblMaster.COMPCD & "'" & vbCrLf
            SQL = SQL & ", LABEL_PRT_NO             = " & gLblMaster.LBLPRTNO & vbCrLf
            SQL = SQL & ", LABEL_PRT_SIDE           = '" & gLblMaster.LBLPRTSIDE & "'   " & vbCrLf
            SQL = SQL & ", LABEL_BAR_SIDE01_TYPE    = '" & gLblMaster.LBLBARSIDE1 & "'  " & vbCrLf
            SQL = SQL & ", LABEL_BAR_SIDE02_TYPE    = '" & gLblMaster.LBLBARSIDE2 & "'   " & vbCrLf
            SQL = SQL & ", PROD_MAX_TOT             = " & gLblMaster.PRODMAXTOT & vbCrLf
            SQL = SQL & ", USED_YN                  = '" & gLblMaster.YN & "'    " & vbCrLf
            SQL = SQL & ", MODIFY_ID                = '" & gKUKDO.USERID & "'" & vbCrLf
            SQL = SQL & ", MODIFY_DT                = '" & gsDBDateTime & "'" & vbCrLf
            SQL = SQL & " WHERE PROD_LABEL_CD       = '" & gLblMaster.LABELCD & "'" & vbCrLf
        'End If
    ElseIf pState = "DEL" Then
        'If gLblMaster.PRODCD <> "" And gLblMaster.COMPCD <> "" And gLblMaster.LBLTYPE <> "" Then
            SQL = ""
            SQL = SQL & "DELETE FROM LBL_LABEL_MASTER " & vbCrLf
            SQL = SQL & " WHERE PROD_LABEL_CD       = '" & gLblMaster.LABELCD & "'" & vbCrLf
        'End If
    End If
    
    Call DBExec(AdoCn, SQL)
    
    Set_Label_Master = True

Exit Function

ErrorRoutine:
    Set_Label_Master = False
    Call DBErrorSet(AdoCn, SQL, "Public Function Set_Label(ByVal pState As String) As Boolean")

End Function

'-- 라벨정보 저장
'Public Function Set_Bar(ByVal pState As String) As Boolean
'
'On Error GoTo ErrorRoutine
'
'    Set_Bar = False
'
'    If pState = "IN" Then
'        SQL = ""
'        SQL = SQL & "INSERT INTO LBL_LABEL_MASTER "
'        SQL = SQL & "(PROD_CD,COMP_CD,PROD_LABEL_TYPE"
'        SQL = SQL & ",LABEL_PRT_NO,LABEL_PRT_DEFAULT_NO,LABEL_PRT_SIDE"
'        SQL = SQL & ",LABEL_BAR_SIDE01_TYPE,LABEL_BAR_SIDE02_TYPE,LABEL_BAR_SIDE03_TYPE,LABEL_BAR_SIDE04_TYPE"
'        SQL = SQL & ",PROD_MAX_TOT,USED_YN"
'        SQL = SQL & ",REGIST_ID,REGIST_DT)  " & vbCrLf
'        SQL = SQL & "  VALUES                       " & vbCrLf
'        SQL = SQL & "('" & gLblMaster.PRODCD & "'"
'        SQL = SQL & ",'" & gLblMaster.COMPCD & "'"
'        SQL = SQL & ",'" & gLblMaster.LBLTYPE & "'"
'        SQL = SQL & "," & gLblMaster.LBLPRTNO
'        SQL = SQL & "," & gLblMaster.LBLPRTDEFAULTNO
'        SQL = SQL & ",'" & gLblMaster.LBLPRTSIDE & "'"
'        SQL = SQL & ",'" & gLblMaster.LBLBARSIDE1 & "'"
'        SQL = SQL & ",'" & gLblMaster.LBLBARSIDE2 & "'"
'        SQL = SQL & ",'" & gLblMaster.LBLBARSIDE3 & "'"
'        SQL = SQL & ",'" & gLblMaster.LBLBARSIDE4 & "'"
'        If gLblMaster.PRODMAXTOT <> "" And IsNumeric(gLblMaster.PRODMAXTOT) Then
'            SQL = SQL & "," & gLblMaster.PRODMAXTOT
'        Else
'            SQL = SQL & ",0"
'        End If
'        SQL = SQL & ",'" & gLblMaster.YN & "'"
'        SQL = SQL & ",'" & gKUKDO.USERID & "'"
'        SQL = SQL & ",'" & gsDBDateTime & "')"
'    ElseIf pState = "UP" Then
'        If gLblMaster.PRODCD <> "" And gLblMaster.COMPCD <> "" And gLblMaster.LBLTYPE <> "" Then
'            SQL = ""
'            SQL = SQL & "UPDATE LBL_LABEL_MASTER SET" & vbCrLf
'            SQL = SQL & "  LABEL_PRT_NO             = " & gLblMaster.LBLPRTNO & vbCrLf
'            SQL = SQL & ", LABEL_PRT_DEFAULT_NO     = " & gLblMaster.LBLPRTDEFAULTNO & vbCrLf
'            SQL = SQL & ", LABEL_PRT_SIDE           = '" & gLblMaster.LBLPRTSIDE & "'   " & vbCrLf
'            SQL = SQL & ", LABEL_BAR_SIDE01_TYPE    = '" & gLblMaster.LBLBARSIDE1 & "'  " & vbCrLf
'            SQL = SQL & ", LABEL_BAR_SIDE02_TYPE    = '" & gLblMaster.LBLBARSIDE2 & "'   " & vbCrLf
'            SQL = SQL & ", LABEL_BAR_SIDE03_TYPE    = '" & gLblMaster.LBLBARSIDE3 & "'  " & vbCrLf
'            SQL = SQL & ", LABEL_BAR_SIDE04_TYPE    = '" & gLblMaster.LBLBARSIDE4 & "'   " & vbCrLf
'            SQL = SQL & ", PROD_MAX_TOT             = " & gLblMaster.PRODMAXTOT & vbCrLf
'            SQL = SQL & ", USED_YN                  = '" & gLblMaster.YN & "'    " & vbCrLf
'            SQL = SQL & ", MODIFY_ID                = '" & gKUKDO.USERID & "'" & vbCrLf
'            SQL = SQL & ", MODIFY_DT                = '" & gsDBDateTime & "'" & vbCrLf
'            SQL = SQL & " WHERE PROD_CD             = '" & gLblMaster.PRODCD & "'" & vbCrLf
'            SQL = SQL & "   AND COMP_CD             = '" & gLblMaster.COMPCD & "'" & vbCrLf
'            SQL = SQL & "   AND PROD_LABEL_TYPE     = '" & gLblMaster.LBLTYPE & "'" & vbCrLf
'        End If
'    ElseIf pState = "DEL" Then
'        If gLblMaster.PRODCD <> "" And gLblMaster.COMPCD <> "" And gLblMaster.LBLTYPE <> "" Then
'            SQL = ""
'            SQL = SQL & "DELETE FROM LBL_LABEL_MASTER " & vbCrLf
'            SQL = SQL & " WHERE PROD_CD         = '" & gLblMaster.PRODCD & "'" & vbCrLf
'            SQL = SQL & "   AND COMP_CD         = '" & gLblMaster.COMPCD & "'" & vbCrLf
'            SQL = SQL & "   AND PROD_LABEL_TYPE = '" & gLblMaster.LBLTYPE & "'" & vbCrLf
'        End If
'    End If
'
'    Call DBExec(AdoCn, SQL)
'
'    Set_Bar = True
'
'Exit Function
'
'ErrorRoutine:
'    Set_Bar = False
'    Call DBErrorSet(AdoCn, SQL, "Public Function Set_Label(ByVal pState As String) As Boolean")
'
'End Function

'-- 라벨정보 저장
Public Function Set_Label_Detail(ByVal pState As String, Optional ByVal pIdx As Integer) As Boolean
    
    Dim i       As Integer
    
    
On Error GoTo ErrorRoutine
    
    Set_Label_Detail = False
        
    If pState = "IN" Then
        'For i = 1 To UBound(gLblDetail.LBLITEM_NO)
            SQL = ""
            SQL = SQL & "INSERT INTO LBL_LABEL_DETAIL "
            SQL = SQL & "(PROD_LABEL_CD"
            SQL = SQL & ",LABEL_ITEM_NO"
            SQL = SQL & ",LABEL_ITEM_SEQ"
            SQL = SQL & ",LABEL_ITEM_NAME"
            SQL = SQL & ",LABEL_ITEM_NAME_PRT"
            SQL = SQL & ",BAR_CD"
            SQL = SQL & ",LABEL_ITEM_GU"
            SQL = SQL & ",LABEL_ITEM_X_COORD"
            SQL = SQL & ",LABEL_ITEM_Y_COORD"
            SQL = SQL & ",LABEL_ITEM_FONT"
            SQL = SQL & ",LABEL_ITEM_ROT"
            SQL = SQL & ",USED_YN"
            SQL = SQL & ",REGIST_ID"
            SQL = SQL & ",REGIST_DT)  " & vbCrLf
            SQL = SQL & "  VALUES                       " & vbCrLf
            SQL = SQL & "('" & gLblMaster.LABELCD & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_NO(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_SEQ(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_NAME(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_NMPRT(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_BARCD(pIdx) & "'"    'code128
            SQL = SQL & ",'" & gLblDetail.LBLITEM_BARGU(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_X(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_Y(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_FONT(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.LBLITEM_ROT(pIdx) & "'"
            SQL = SQL & ",'" & gLblDetail.YN(pIdx) & "'"
            SQL = SQL & ",'" & gKUKDO.USERID & "'"
            SQL = SQL & ",'" & gsDBDateTime & "')"
            
            Call DBExec(AdoCn, SQL)
            
        'Next
        
    ElseIf pState = "UP" Then
        'For i = 1 To UBound(gLblDetail.LBLITEM_NO)
            'If gLblDetail.PRODCD <> "" And gLblDetail.COMPCD <> "" And gLblDetail.LBLTYPE <> "" Then
                SQL = ""
                SQL = SQL & "UPDATE LBL_LABEL_DETAIL SET" & vbCrLf
                SQL = SQL & "  LABEL_ITEM_SEQ       = '" & gLblDetail.LBLITEM_SEQ(pIdx) & "'   " & vbCrLf
                SQL = SQL & ", LABEL_ITEM_NAME      = '" & gLblDetail.LBLITEM_NAME(pIdx) & "'  " & vbCrLf
                SQL = SQL & ", LABEL_ITEM_NAME_PRT  = '" & gLblDetail.LBLITEM_NMPRT(pIdx) & "' " & vbCrLf
                SQL = SQL & ", BAR_CD               = '" & gLblDetail.LBLITEM_BARCD(pIdx) & "' " & vbCrLf
                SQL = SQL & ", LABEL_ITEM_GU        = '" & gLblDetail.LBLITEM_BARGU(pIdx) & "' " & vbCrLf
                SQL = SQL & ", LABEL_ITEM_X_COORD   = '" & gLblDetail.LBLITEM_X(pIdx) & "'     " & vbCrLf
                SQL = SQL & ", LABEL_ITEM_Y_COORD   = '" & gLblDetail.LBLITEM_Y(pIdx) & "'     " & vbCrLf
                SQL = SQL & ", LABEL_ITEM_FONT      = '" & gLblDetail.LBLITEM_FONT(pIdx) & "'  " & vbCrLf
                SQL = SQL & ", LABEL_ITEM_ROT       = " & gLblDetail.LBLITEM_ROT(pIdx) & vbCrLf
                SQL = SQL & ", USED_YN              = '" & gLblDetail.YN(pIdx) & "'            " & vbCrLf
                SQL = SQL & ", MODIFY_ID            = '" & gKUKDO.USERID & "'               " & vbCrLf
                SQL = SQL & ", MODIFY_DT            = '" & gsDBDateTime & "'                " & vbCrLf
                SQL = SQL & " WHERE PROD_LABEL_CD   = '" & gLblDetail.LABELCD & "'              " & vbCrLf
                SQL = SQL & "   AND LABEL_ITEM_NO   = '" & gLblDetail.LBLITEM_NO(pIdx) & "'    " & vbCrLf
            
                Call DBExec(AdoCn, SQL)
            
            'End If
        'Next
        
    ElseIf pState = "DEL" Then
        
        'If gLblDetail.PRODCD <> "" And gLblDetail.COMPCD <> "" And gLblDetail.LBLTYPE <> "" Then
            SQL = ""
            SQL = SQL & "DELETE FROM LBL_LABEL_DETAIL                        " & vbCrLf
            SQL = SQL & " WHERE PROD_LABEL_CD    = '" & gLblDetail.LABELCD & "'" & vbCrLf
            'SQL = SQL & "   AND LABEL_ITEM_NO   = '" & pIdx & "'" & vbCrLf
                    
            Call DBExec(AdoCn, SQL)

        'End If
        
    End If
    
'    Call DBExec(AdoCn, SQL)
    
    Set_Label_Detail = True

Exit Function

ErrorRoutine:
    Set_Label_Detail = False
    Call DBErrorSet(AdoCn, SQL, "Public Function Set_Label(ByVal pState As String) As Boolean")

End Function


'Data Base의 현재일자시간
Public Function gsDBDateTime() As Date
    
    Dim sRs As ADODB.Recordset
    
On Error GoTo ErrorRoutine
    
    If gDBCONN = "1" Then
        Set sRs = New ADODB.Recordset
        SQL = "select format(date(), 'YYYY-MM-DD') + ' ' + format(time(), 'HH:mm:ss') as SYSDATE FROM LBL_M_USER"
        sRs.Open SQL, AdoCn, adOpenStatic, adLockReadOnly
        If Not sRs.EOF Then
            gsDBDateTime = sRs("SYSDATE")
        Else
            gsDBDateTime = Now
        End If
        sRs.Close
        Set sRs = Nothing
        
    Else
        Set sRs = New ADODB.Recordset
        SQL = "select sysdate from dual"
        sRs.Open SQL, AdoCn, adOpenStatic, adLockReadOnly
        If Not sRs.EOF Then
            gsDBDateTime = sRs("SYSDATE")
        Else
            gsDBDateTime = Now
        End If
        sRs.Close
        Set sRs = Nothing
    End If
    
Exit Function

ErrorRoutine:
    gsDBDateTime = Now
    Call DBErrorSet(AdoCn, SQL, "gsDBDateTime")
    
End Function



