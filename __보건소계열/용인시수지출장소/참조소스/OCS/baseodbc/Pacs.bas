Attribute VB_Name = "Pacs"

Sub Pacs_Write()
    
     strSQL = "   Create or Replace view mwl      " & vbLf
     strSQL = strSQL & " (                        " & vbLf
     strSQL = strSQL & " character_set,           " & vbLf
     strSQL = strSQL & " scheduled_aetitle,       " & vbLf
     strSQL = strSQL & " scheduled_dttm,          " & vbLf
     strSQL = strSQL & " scheduled_modality,      " & vbLf
     strSQL = strSQL & " scheduled_station,       " & vbLf
     strSQL = strSQL & " scheduled_location,      " & vbLf
     strSQL = strSQL & " scheduled_proc_id,       " & vbLf
     strSQL = strSQL & " scheduled_proc_desc,     " & vbLf
     strSQL = strSQL & " scheduled_proc_status,   " & vbLf
     strSQL = strSQL & " premedication,           " & vbLf
     strSQL = strSQL & " contrast_agent,          " & vbLf
     strSQL = strSQL & " requested_proc_id,       " & vbLf
     strSQL = strSQL & " requested_proc_desc,     " & vbLf
     strSQL = strSQL & " requested_proc_priority, " & vbLf
     strSQL = strSQL & " study_instance_uid,      " & vbLf
     strSQL = strSQL & " accession_no,            " & vbLf
     strSQL = strSQL & " perform_doctor,          " & vbLf
     strSQL = strSQL & " request_doctor,          " & vbLf
     strSQL = strSQL & " refer_doctor,            " & vbLf
     strSQL = strSQL & " admission_id,            " & vbLf
     strSQL = strSQL & " patient_transport,       " & vbLf
     strSQL = strSQL & " patient_location,        " & vbLf
     strSQL = strSQL & " patient_name,            " & vbLf
     strSQL = strSQL & " patient_id,              " & vbLf
     strSQL = strSQL & " patient_birth_dttm,      " & vbLf
     strSQL = strSQL & " patient_sex,             " & vbLf
     strSQL = strSQL & " patient_weight,          " & vbLf
     strSQL = strSQL & " patient_state,           " & vbLf
     strSQL = strSQL & " confidentiality,         " & vbLf
     strSQL = strSQL & " pregnancy_status,        " & vbLf
     strSQL = strSQL & " medical_alerts,          " & vbLf
     strSQL = strSQL & " contrast_allergies,      " & vbLf
     strSQL = strSQL & " special_needs      ) as  " & vbLf
     
     strSQL = strSQL & " Select  'ANY'," & vbLf
     strSQL = strSQL & "         'ANY'," & vbLf
     strSQL = strSQL & "         TO_DATE(TO_CHAR(SEEKDATE,'YYYYMMDD') " & vbLf
     strSQL = strSQL & "         || SUBSTR(SEEKTIME,1,2)              " & vbLf
     strSQL = strSQL & "         || SUBSTR(SEEKTIME,4,2),'YYYYMMDDHH24MI')," & vbLf                    '촬영일자,시간
     strSQL = strSQL & "         DECODE(X.XJONG,'1','CR','2','CT','3','MR','5','CR','6','US')," & vbLf      'CR,CT,MR,US,OT,ES '07 확인
     strSQL = strSQL & "         'ANY'," & vbLf
     strSQL = strSQL & "         'ANY'," & vbLf
     strSQL = strSQL & "         'ANY'," & vbLf         '검사명 'ANY'
     strSQL = strSQL & "         NULL, " & vbLf
     strSQL = strSQL & "         NULL, " & vbLf
     strSQL = strSQL & "         NULL, " & vbLf
     strSQL = strSQL & "         NULL, " & vbLf
     strSQL = strSQL & "         'ANY', " & vbLf          '검사번호 'ANY'
     strSQL = strSQL & "         NULL, " & vbLf
     strSQL = strSQL & "         NULL, " & vbLf
     strSQL = strSQL & "         '1.2.410.200001.20.' || X.PTNO || X.DRCODE || TO_CHAR(X.SEEKDATE,'YYYYMMDD') || X.ORDERNO ," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf      '접수 의사코드
     strSQL = strSQL & "         X.DRCODE," & vbLf   '의뢰 의사코드
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         P.SNAME," & vbLf
     strSQL = strSQL & "         X.PTNO, " & vbLf
     strSQL = strSQL & "         P.BIRTHDATE," & vbLf       '생년월일
     strSQL = strSQL & "         DECODE(SUBSTR(P.JUMIN2,1,1),'1','M','2','F','3','M','4','F','O')," & vbLf  '남여구분
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         NULL," & vbLf
     strSQL = strSQL & "         X.REMARK" & vbLf
     strSQL = strSQL & " FROM TW_MIS_OCS.TWXRAY_DETAIL X,   " & vbLf
     strSQL = strSQL & "      TW_MIS_PMPA.TWBAS_PATIENT P   " & vbLf
     strSQL = strSQL & " WHERE X.PTNO       = P.PTNO        " & vbLf
     strSQL = strSQL & "   AND X.XJONG not in ('4','7')     " & vbLf '골밀도,핵의학영상 제외
     strSQL = strSQL & "   AND X.SEEKDATE   = TRUNC(SYSDATE)" & vbLf
     strSQL = strSQL & "   AND X.GbPrint    = 'A'           " & vbLf
     strSQL = strSQL & "   AND X.GbReserved = '7'           " & vbLf
     
     strSQL = strSQL & " UNION ALL       " & vbLf
       
     strSQL = strSQL & "   Select  'ANY'," & vbLf
     strSQL = strSQL & "           'ANY'," & vbLf
     strSQL = strSQL & "           TO_DATE(TO_CHAR(RDATE,'YYYYMMDD')  ,'YYYYMMDD'),"
    ' strSQL = strSQL & "           || SUBSTR(SEEKTIME,1,2)"                      '촬영일자,시간
    ' strSQL = strSQL & "           || SUBSTR(SEEKTIME,4,2),'YYYYMMDDHH24MI'),
     strSQL = strSQL & "           'ES', " & vbLf
     strSQL = strSQL & "           'ANY'," & vbLf
     strSQL = strSQL & "           'ANY'," & vbLf
     strSQL = strSQL & "           'ANY'," & vbLf
     strSQL = strSQL & "           NULL, " & vbLf
     strSQL = strSQL & "           NULL, " & vbLf
     strSQL = strSQL & "           NULL, " & vbLf
     strSQL = strSQL & "           NULL, " & vbLf
     strSQL = strSQL & "           'ANY', " & vbLf
     strSQL = strSQL & "           NULL, " & vbLf
     strSQL = strSQL & "           NULL, " & vbLf
     strSQL = strSQL & "           '1.2.410.200001.20.' || J.PTNO || J.DRCODE || TO_CHAR(J.JDATE,'YYYYMMDD') || J.ORDERNO ," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           J.DRCODE," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           P.SNAME," & vbLf
     strSQL = strSQL & "           J.PTNO," & vbLf
     strSQL = strSQL & "           P.BIRTHDATE," & vbLf
     strSQL = strSQL & "           DECODE(SUBSTR(P.JUMIN2,1,1),'1','M','2','F','3','M','4','F','O')," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           NULL," & vbLf
     strSQL = strSQL & "           j.REMARK" & vbLf
     strSQL = strSQL & "   FROM   TW_MIS_OCS.TWENDO_JUPMST J, " & vbLf
     strSQL = strSQL & "          TW_MIS_PMPA.TWBAS_PATIENT P " & vbLf
     strSQL = strSQL & "   Where  j.Ptno  = P.Ptno            " & vbLf
     strSQL = strSQL & "     and  j.jdate = trunc(sysdate)    " & vbLf
     Result = adoSQL(strSQL)
        
    If Result = -1 Then
        MsgBox "PACS VIEW TABLE을 생성하지 못했습니다.전산실 연락요망!!", , "Table 생성 Error"
        Call DbAdoDisConnect
        End
    End If
    
End Sub
