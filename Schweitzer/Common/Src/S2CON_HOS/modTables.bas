Attribute VB_Name = "modTables"
Option Explicit

Public IsSetTable As Boolean
'-- Common
Public mT_COM001 As String
Public mT_COM002 As String
Public mT_COM003 As String
Public mT_COM004 As String
Public mT_COM005 As String
Public mT_COM006 As String
Public mT_COM007 As String
Public mT_COM008 As String
Public mT_COM009 As String
Public mT_COM010 As String
Public mT_COM011 As String
Public mT_COM012 As String
Public mT_COM013 As String
Public mT_COM014 As String
Public mT_COM015 As String
Public mT_COM016 As String
Public mT_COM017 As String
Public mT_COM099 As String
Public mT_COM101 As String
'-- HIS
' 테이블명 상수 ( LIS Tables )
Public mT_HIS001 As String       '환자마스터
Public mT_HIS002 As String       '입원환자마스터
Public mT_HIS003 As String       '부서마스터
Public mT_HIS004 As String       '병상마스터
Public mT_HIS005 As String       '의사마스터
Public mT_HIS006 As String       '상병마스터
Public mT_HIS007 As String       '수술코드마스터
Public mT_HIS008 As String       '수가마스터
Public mT_HIS009 As String
Public mT_HIS010 As String
'-- LIS
Public mT_LAB001 As String
Public mT_LAB002 As String
Public mT_LAB003 As String
Public mT_LAB004 As String
Public mT_LAB005 As String
Public mT_LAB006 As String
Public mT_LAB007 As String
Public mT_LAB008 As String
Public mT_LAB009 As String
Public mT_LAB010 As String
Public mT_LAB011 As String
Public mT_LAB012 As String
Public mT_LAB013 As String
Public mT_LAB014 As String
Public mT_LAB015 As String

Public mT_LAB021 As String
Public mT_LAB022 As String
Public mT_LAB023 As String
Public mT_LAB024 As String
Public mT_LAB025 As String
Public mT_LAB026 As String
Public mT_LAB027 As String
Public mT_LAB028 As String
Public mT_LAB029 As String
Public mT_LAB030 As String

Public mT_LAB031 As String
Public mT_LAB032 As String
Public mT_LAB033 As String
Public mT_LAB034 As String
Public mT_LAB035 As String
Public mT_LAB036 As String
Public mT_LAB041 As String
Public mT_LAB099 As String

Public mT_LAB101 As String
Public mT_LAB102 As String
Public mT_LAB103 As String
Public mT_LAB104 As String
Public mT_LAB105 As String
Public mT_LAB106 As String

Public mT_LAB195 As String
Public mT_LAB201 As String
Public mT_LAB202 As String
Public mT_LAB203 As String
Public mT_LAB204 As String
Public mT_LAB205 As String
Public mT_LAB206 As String

Public mT_LAB301 As String
Public mT_LAB302 As String
Public mT_LAB303 As String
Public mT_LAB304 As String
Public mT_LAB305 As String
Public mT_LAB306 As String
Public mT_LAB307 As String
Public mT_LAB308 As String
Public mT_LAB309 As String
Public mT_LAB310 As String

Public mT_LAB350 As String
Public mT_LAB351 As String
Public mT_LAB352 As String
Public mT_LAB353 As String
Public mT_LAB354 As String
Public mT_LAB360 As String
Public mT_LAB361 As String

Public mT_LAB401 As String
Public mT_LAB402 As String
Public mT_LAB403 As String
Public mT_LAB404 As String
Public mT_LAB405 As String
Public mT_LAB406 As String
Public mT_LAB407 As String

Public mT_LAB501 As String
Public mT_LAB502 As String
Public mT_LAB503 As String
Public mT_LAB504 As String
Public mT_LAB505 As String
Public mT_LAB506 As String

Public mT_LAB601 As String
Public mT_LAB602 As String
Public mT_LAB603 As String
Public mT_LAB604 As String

Public mT_LAB701 As String
Public mT_LAB702 As String
Public mT_LAB703 As String
Public mT_LAB704 As String

Public mT_LAB801 As String

Public mT_LAB901 As String
Public mT_LAB902 As String
Public mT_LAB903 As String
Public mT_LAB904 As String

Public mT_LAB999 As String

Public mT_INT001 As String       'OCS INTERFACE 처방내역
Public mT_INT002 As String       'OCS INTERFACE 수가계산내역
Public mT_INT003 As String       'OCS INTERFACE 혈액불출내역


'-- APS
'Public mT_APS001 As String       '검사항목 마스터
'Public mT_APS002 As String       '검체 마스터
'Public mT_APS003 As String       '지정검체 마스터
'Public mT_APS004 As String       'Snomedcode(진단) 마스터
'Public mT_APS005 As String       'Snomedcode(조직) 마스터
'Public mT_APS006 As String       '장기 채취 방법 마스터
'Public mT_APS007 As String
'Public mT_APS008 As String       '위크쉬트 마스터
'Public mT_APS009 As String       '워크쉬트 검사정보
'Public mT_APS011 As String       '의뢰처 기관마스터
'
'Public mT_APS101 As String       '처방Header
'Public mT_APS102 As String       '처방Body
'Public mT_APS103 As String       '처방별 검체 변환내역
'Public mT_APS104 As String
'Public mT_APS105 As String       '수가계산내역
'Public mT_APS106 As String       '처방 변경내역
'
'
'Public mT_APS201 As String       '채혈접수내역
'Public mT_APS202 As String       '추가검체내역
'Public mT_APS203 As String       '결과보고대기내역
'Public mT_APS204 As String       '일괄채혈내역
'Public mT_APS205 As String       '슬라이드 대출내역
'Public mT_APS206 As String       '외부의뢰 내역(Send out)
'Public mT_APS207 As String       '수탁검사 내역(Referral)
'
'Public mT_APS301 As String       '워크쉬트 생성내역
'Public mT_APS302 As String       '조직병리 결과내역
'Public mT_APS303 As String       '세포병리 결과내역
'Public mT_APS304 As String       'Snomed coding(조직) 결과내역
'Public mT_APS305 As String       'Snomed coding(진단) 결과내역
'Public mT_APS306 As String       'ICD9(10)CM 코딩내역
'Public mT_APS307 As String       '시드니시스템 결과내역
'Public mT_APS308 As String       '면역현광 결과내역
'Public mT_APS309 As String       '효소조직 유방암 결과내역
'Public mT_APS310 As String       '슬라이드 결과이미지 내역
'Public mT_APS311 As String       '동일 검체 병리번호내역
'Public mT_APS312 As String       '블록,슬라이드 작업내역
'Public mT_APS313 As String       '컨트롤 슬라이드 QC내역
'Public mT_APS317 As String       '부검 결과내역
'Public mT_APS318 As String       '전자현미경 결과 참고문헌 내역
'Public mT_APS319 As String       '진단협의의사 내역
'Public mT_APS320 As String       '진단병리일반결과수정내역
'Public mT_APS321 As String       '진단병리 접수취소내역
'
'Public mT_APS401 As String       '부검접수내역
'Public mT_APS402 As String       '부검육안소견내역
'Public mT_APS403 As String       '부검진단결과내역
'
'Public mT_APS901 As String       '과거결과조회내역

'-- BBS
Public mT_BBS001 As String       '검사항목 마스터(수혈처방)
Public mT_BBS002 As String       'Donor Screen 검사 적격치 마스터
Public mT_BBS003 As String       '검체 보관 장소 마스터
Public mT_BBS004 As String       '
Public mT_BBS005 As String       '
Public mT_BBS006 As String       'KIT마스터
Public mT_BBS007 As String       '
Public mT_BBS008 As String       '
Public mT_BBS009 As String       '
Public mT_BBS010 As String       '
                            
Public mT_BBS101 As String       '수혈사유내역
Public mT_BBS102 As String       '채혈 & 처방 중간다리
Public mT_BBS103 As String       'XM결과등록 환자별 리마크

Public mT_BBS201 As String       '채혈 접수내역
Public mT_BBS202 As String       '처방 접수내역
Public mT_BBS203 As String       '처방 STATUS
Public mT_BBS204 As String       '수혈요청 전송내역
Public mT_BBS206 As String       '보관 검체마스터
Public mT_BBS207 As String       '검체 추가요청내역

Public mT_BBS302 As String       'Cross Matching 결과내역
Public mT_BBS303 As String       'ABO 검사 결과내역
Public mT_BBS304 As String       'FILTER 출고내역

Public mT_BBS401 As String       '혈액 입고내역
Public mT_BBS402 As String       '혈액 출고내역
Public mT_BBS403 As String       '혈액 반환내역
Public mT_BBS404 As String       '혈액 폐기내역
Public mT_BBS405 As String       '혈액 Bag 회수내역
Public mT_BBS409 As String       '
Public mT_BBS411 As String       '본원 헌혈증 수령내역
Public mT_BBS412 As String       '본원 센터별 헌혈증 수령내역
Public mT_BBS413 As String       '자병원 헌혈증 분배내역
Public mT_BBS414 As String       '현혈증 반납내역

Public mT_BBS501 As String       'Record of Transfusion
Public mT_BBS502 As String       'Reaction(수혈전 XM)
Public mT_BBS503 As String       'Reaction(수혈전 기타검사)
Public mT_BBS504 As String       'Reaction(수혈후 XM)
Public mT_BBS505 As String       'Reaction(수혈후 기타검사)
Public mT_BBS506 As String       '수혈부작용등록
Public mT_BBS601 As String       '헌혈자 마스터
Public mT_BBS602 As String       '헌혈자 접수내역
Public mT_BBS603 As String       '적격여부 판정내역
Public mT_BBS604 As String       '헌혈자 문진내역
Public mT_BBS605 As String       '헌혈자 검사의뢰내역
Public mT_BBS606 As String       '헌혈자 추가재료 사용내역
Public mT_BBS607 As String       '헌혈자 판정 사유테이블
Public mT_BBS901 As String       '혈액마감내역
Public mT_BBS902 As String       '혈액형 입력테이블
Public mT_BBS903 As String       '혈액형 입력테이블
'-- ICS
Public mT_ICS001 As String       '법정감염관리 테이블
Public mT_ICS002 As String       '원내감염관리 테이블
Public mT_ICS101 As String       '중간결과테이블(법정감염)
Public mT_ICS102 As String       '중간결과테이블(원내감염)
Public mT_ICS103 As String       '수혈부작용등록테이블
Public mT_ICS201 As String       '최종결과테이블(법정감염)
Public mT_ICS202 As String       '중간결과테이블(원내감염)
Public mT_ICS301 As String       'History(법정감염)
Public mT_ICS302 As String

Public Sub SetTable()
    mT_COM001 = ReadINI("TABLE", "T_COM001", "")
    mT_COM002 = ReadINI("TABLE", "T_COM002", "")
    mT_COM003 = ReadINI("TABLE", "T_COM003", "")
    mT_COM004 = ReadINI("TABLE", "T_COM004", "")
    mT_COM005 = ReadINI("TABLE", "T_COM005", "")
    mT_COM006 = ReadINI("TABLE", "T_COM006", "")
    mT_COM007 = ReadINI("TABLE", "T_COM007", "")
    mT_COM008 = ReadINI("TABLE", "T_COM008", "")
    mT_COM009 = ReadINI("TABLE", "T_COM009", "")
    mT_COM010 = ReadINI("TABLE", "T_COM010", "")
    mT_COM011 = ReadINI("TABLE", "T_COM011", "")
    mT_COM012 = ReadINI("TABLE", "T_COM012", "")
    mT_COM013 = ReadINI("TABLE", "T_COM013", "")
    mT_COM014 = ReadINI("TABLE", "T_COM014", "")
    mT_COM015 = ReadINI("TABLE", "T_COM015", "")
    mT_COM016 = ReadINI("TABLE", "T_COM016", "")
    mT_COM017 = ReadINI("TABLE", "T_COM017", "")
    mT_COM099 = ReadINI("TABLE", "T_COM099", "")
    mT_COM101 = ReadINI("TABLE", "T_COM101", "")
    
    mT_HIS001 = ReadINI("TABLE", "T_HIS001", "")
    mT_HIS002 = ReadINI("TABLE", "T_HIS002", "")
    mT_HIS003 = ReadINI("TABLE", "T_HIS003", "")
    mT_HIS004 = ReadINI("TABLE", "T_HIS004", "")
    mT_HIS005 = ReadINI("TABLE", "T_HIS005", "")
    mT_HIS006 = ReadINI("TABLE", "T_HIS006", "")
    mT_HIS007 = ReadINI("TABLE", "T_HIS007", "")
    mT_HIS008 = ReadINI("TABLE", "T_HIS008", "")
    mT_HIS009 = ReadINI("TABLE", "T_HIS009", "")
    mT_HIS010 = ReadINI("TABLE", "T_HIS010", "")

    mT_LAB001 = ReadINI("TABLE", "T_LAB001", "")
    mT_LAB002 = ReadINI("TABLE", "T_LAB002", "")
    mT_LAB003 = ReadINI("TABLE", "T_LAB003", "")
    mT_LAB004 = ReadINI("TABLE", "T_LAB004", "")
    mT_LAB005 = ReadINI("TABLE", "T_LAB005", "")
    mT_LAB006 = ReadINI("TABLE", "T_LAB006", "")
    mT_LAB007 = ReadINI("TABLE", "T_LAB007", "")
    mT_LAB008 = ReadINI("TABLE", "T_LAB008", "")
    mT_LAB009 = ReadINI("TABLE", "T_LAB009", "")
    mT_LAB010 = ReadINI("TABLE", "T_LAB010", "")
    mT_LAB011 = ReadINI("TABLE", "T_LAB011", "")
    mT_LAB012 = ReadINI("TABLE", "T_LAB012", "")
    mT_LAB013 = ReadINI("TABLE", "T_LAB013", "")
    mT_LAB014 = ReadINI("TABLE", "T_LAB014", "")
    mT_LAB015 = ReadINI("TABLE", "T_LAB015", "")

    mT_LAB021 = ReadINI("TABLE", "T_LAB021", "")
    mT_LAB022 = ReadINI("TABLE", "T_LAB022", "")
    mT_LAB023 = ReadINI("TABLE", "T_LAB023", "")
    mT_LAB024 = ReadINI("TABLE", "T_LAB024", "")
    mT_LAB025 = ReadINI("TABLE", "T_LAB025", "")
    mT_LAB026 = ReadINI("TABLE", "T_LAB026", "")
    mT_LAB027 = ReadINI("TABLE", "T_LAB027", "")
    mT_LAB028 = ReadINI("TABLE", "T_LAB028", "")
    mT_LAB029 = ReadINI("TABLE", "T_LAB029", "")
    mT_LAB030 = ReadINI("TABLE", "T_LAB030", "")

    mT_LAB031 = ReadINI("TABLE", "T_LAB031", "")
    mT_LAB032 = ReadINI("TABLE", "T_LAB032", "")
    mT_LAB033 = ReadINI("TABLE", "T_LAB033", "")
    mT_LAB034 = ReadINI("TABLE", "T_LAB034", "")
    mT_LAB035 = ReadINI("TABLE", "T_LAB035", "")
    mT_LAB036 = ReadINI("TABLE", "T_LAB036", "")
    mT_LAB041 = ReadINI("TABLE", "T_LAB041", "")
    mT_LAB099 = ReadINI("TABLE", "T_LAB099", "")

    mT_LAB101 = ReadINI("TABLE", "T_LAB101", "")
    mT_LAB102 = ReadINI("TABLE", "T_LAB102", "")
    mT_LAB103 = ReadINI("TABLE", "T_LAB103", "")
    mT_LAB104 = ReadINI("TABLE", "T_LAB104", "")
    mT_LAB105 = ReadINI("TABLE", "T_LAB105", "")
    mT_LAB106 = ReadINI("TABLE", "T_LAB106", "")
    
    mT_LAB195 = ReadINI("TABLE", "T_LAB195", "")
    
    mT_LAB201 = ReadINI("TABLE", "T_LAB201", "")
    mT_LAB202 = ReadINI("TABLE", "T_LAB202", "")
    mT_LAB203 = ReadINI("TABLE", "T_LAB203", "")
    mT_LAB204 = ReadINI("TABLE", "T_LAB204", "")
    mT_LAB205 = ReadINI("TABLE", "T_LAB205", "")
    mT_LAB206 = ReadINI("TABLE", "T_LAB206", "")

    mT_LAB301 = ReadINI("TABLE", "T_LAB301", "")
    mT_LAB302 = ReadINI("TABLE", "T_LAB302", "")
    mT_LAB303 = ReadINI("TABLE", "T_LAB303", "")
    mT_LAB304 = ReadINI("TABLE", "T_LAB304", "")
    mT_LAB305 = ReadINI("TABLE", "T_LAB305", "")
    mT_LAB306 = ReadINI("TABLE", "T_LAB306", "")
    mT_LAB307 = ReadINI("TABLE", "T_LAB307", "")
    mT_LAB308 = ReadINI("TABLE", "T_LAB308", "")
    mT_LAB309 = ReadINI("TABLE", "T_LAB309", "")
    mT_LAB310 = ReadINI("TABLE", "T_LAB310", "")

    mT_LAB350 = ReadINI("TABLE", "T_LAB350", "")
    mT_LAB351 = ReadINI("TABLE", "T_LAB351", "")
    mT_LAB352 = ReadINI("TABLE", "T_LAB352", "")
    mT_LAB353 = ReadINI("TABLE", "T_LAB353", "")
    mT_LAB354 = ReadINI("TABLE", "T_LAB354", "")
    
    mT_LAB360 = ReadINI("TABLE", "T_LAB360", "")
    mT_LAB361 = ReadINI("TABLE", "T_LAB361", "")

    mT_LAB401 = ReadINI("TABLE", "T_LAB401", "")
    mT_LAB402 = ReadINI("TABLE", "T_LAB402", "")
    mT_LAB403 = ReadINI("TABLE", "T_LAB403", "")
    mT_LAB404 = ReadINI("TABLE", "T_LAB404", "")
    mT_LAB405 = ReadINI("TABLE", "T_LAB405", "")
    mT_LAB406 = ReadINI("TABLE", "T_LAB406", "")
    mT_LAB407 = ReadINI("TABLE", "T_LAB407", "")

    mT_LAB601 = ReadINI("TABLE", "T_LAB601", "")
    mT_LAB602 = ReadINI("TABLE", "T_LAB602", "")
    mT_LAB603 = ReadINI("TABLE", "T_LAB603", "")
    mT_LAB604 = ReadINI("TABLE", "T_LAB604", "")

    mT_LAB501 = ReadINI("TABLE", "T_LAB501", "")
    mT_LAB502 = ReadINI("TABLE", "T_LAB502", "")
    mT_LAB503 = ReadINI("TABLE", "T_LAB503", "")
    mT_LAB504 = ReadINI("TABLE", "T_LAB504", "")
    mT_LAB505 = ReadINI("TABLE", "T_LAB505", "")
    mT_LAB506 = ReadINI("TABLE", "T_LAB506", "")

    mT_LAB701 = ReadINI("TABLE", "T_LAB701", "")
    mT_LAB702 = ReadINI("TABLE", "T_LAB702", "")
    mT_LAB703 = ReadINI("TABLE", "T_LAB703", "")
    mT_LAB704 = ReadINI("TABLE", "T_LAB704", "")

    mT_LAB801 = ReadINI("TABLE", "T_LAB801", "")
    
    mT_LAB901 = ReadINI("TABLE", "T_LAB901", "")
    mT_LAB902 = ReadINI("TABLE", "T_LAB902", "")
    mT_LAB903 = ReadINI("TABLE", "T_LAB903", "")
    mT_LAB904 = ReadINI("TABLE", "T_LAB904", "")
    
    mT_LAB999 = ReadINI("TABLE", "T_LAB999", "")

    mT_INT001 = ReadINI("TABLE", "T_INT001", "")
    mT_INT002 = ReadINI("TABLE", "T_INT002", "")
    mT_INT003 = ReadINI("TABLE", "T_INT003", "")

'-- APS
'    mT_APS001 = ReadINI("TABLE", "T_APS001", "")
'    mT_APS002 = ReadINI("TABLE", "T_APS002", "")
'    mT_APS003 = ReadINI("TABLE", "T_APS003", "")
'    mT_APS004 = ReadINI("TABLE", "T_APS004", "")
'    mT_APS005 = ReadINI("TABLE", "T_APS005", "")
'    mT_APS006 = ReadINI("TABLE", "T_APS006", "")
'    mT_APS007 = ReadINI("TABLE", "T_APS007", "")
'    mT_APS008 = ReadINI("TABLE", "T_APS008", "")
'    mT_APS009 = ReadINI("TABLE", "T_APS009", "")
'    mT_APS011 = ReadINI("TABLE", "T_APS011", "")
'
'    mT_APS101 = ReadINI("TABLE", "T_APS101", "")
'    mT_APS102 = ReadINI("TABLE", "T_APS102", "")
'    mT_APS103 = ReadINI("TABLE", "T_APS103", "")
'    mT_APS106 = ReadINI("TABLE", "T_APS106", "")
'
'    mT_APS201 = ReadINI("TABLE", "T_APS201", "")
'    mT_APS202 = ReadINI("TABLE", "T_APS202", "")
'    mT_APS203 = ReadINI("TABLE", "T_APS203", "")
'    mT_APS204 = ReadINI("TABLE", "T_APS204", "")
'    mT_APS205 = ReadINI("TABLE", "T_APS205", "")
'    mT_APS206 = ReadINI("TABLE", "T_APS206", "")
'    mT_APS207 = ReadINI("TABLE", "T_APS207", "")
'
'    mT_APS301 = ReadINI("TABLE", "T_APS301", "")
'    mT_APS302 = ReadINI("TABLE", "T_APS302", "")
'    mT_APS303 = ReadINI("TABLE", "T_APS303", "")
'    mT_APS304 = ReadINI("TABLE", "T_APS304", "")
'    mT_APS305 = ReadINI("TABLE", "T_APS305", "")
'    mT_APS306 = ReadINI("TABLE", "T_APS306", "")
'    mT_APS307 = ReadINI("TABLE", "T_APS307", "")
'    mT_APS308 = ReadINI("TABLE", "T_APS308", "")
'    mT_APS309 = ReadINI("TABLE", "T_APS309", "")
'    mT_APS310 = ReadINI("TABLE", "T_APS310", "")
'    mT_APS311 = ReadINI("TABLE", "T_APS311", "")
'    mT_APS312 = ReadINI("TABLE", "T_APS312", "")
'    mT_APS313 = ReadINI("TABLE", "T_APS313", "")
'    mT_APS317 = ReadINI("TABLE", "T_APS317", "")
'    mT_APS318 = ReadINI("TABLE", "T_APS318", "")
'    mT_APS319 = ReadINI("TABLE", "T_APS319", "")
'    mT_APS320 = ReadINI("TABLE", "T_APS320", "")
'    mT_APS321 = ReadINI("TABLE", "T_APS321", "")
'
'    mT_APS401 = ReadINI("TABLE", "T_APS401", "")
'    mT_APS402 = ReadINI("TABLE", "T_APS402", "")
'    mT_APS403 = ReadINI("TABLE", "T_APS403", "")
'
'    mT_APS901 = ReadINI("TABLE", "T_APS901", "")

'-- BBS
    mT_BBS001 = ReadINI("TABLE", "T_BBS001", "")
    mT_BBS002 = ReadINI("TABLE", "T_BBS002", "")
    mT_BBS003 = ReadINI("TABLE", "T_BBS003", "")
    mT_BBS004 = ReadINI("TABLE", "T_BBS004", "")
    mT_BBS005 = ReadINI("TABLE", "T_BBS005", "")
    mT_BBS006 = ReadINI("TABLE", "T_BBS006", "")
    mT_BBS007 = ReadINI("TABLE", "T_BBS007", "")
    mT_BBS008 = ReadINI("TABLE", "T_BBS008", "")
    mT_BBS009 = ReadINI("TABLE", "T_BBS009", "")
    mT_BBS010 = ReadINI("TABLE", "T_BBS010", "")

    mT_BBS101 = ReadINI("TABLE", "T_BBS101", "")
    mT_BBS102 = ReadINI("TABLE", "T_BBS102", "")
    mT_BBS103 = ReadINI("TABLE", "T_BBS103", "")
    
    mT_BBS201 = ReadINI("TABLE", "T_BBS201", "")
    mT_BBS202 = ReadINI("TABLE", "T_BBS202", "")
    mT_BBS203 = ReadINI("TABLE", "T_BBS203", "")
    mT_BBS204 = ReadINI("TABLE", "T_BBS204", "")
    mT_BBS206 = ReadINI("TABLE", "T_BBS206", "")
    mT_BBS207 = ReadINI("TABLE", "T_BBS207", "")

    mT_BBS302 = ReadINI("TABLE", "T_BBS302", "")
    mT_BBS303 = ReadINI("TABLE", "T_BBS303", "")
    mT_BBS304 = ReadINI("TABLE", "T_BBS304", "")

    mT_BBS401 = ReadINI("TABLE", "T_BBS401", "")
    mT_BBS402 = ReadINI("TABLE", "T_BBS402", "")
    mT_BBS403 = ReadINI("TABLE", "T_BBS403", "")
    mT_BBS404 = ReadINI("TABLE", "T_BBS404", "")
    mT_BBS405 = ReadINI("TABLE", "T_BBS405", "")
    mT_BBS409 = ReadINI("TABLE", "T_BBS409", "")
    mT_BBS411 = ReadINI("TABLE", "T_BBS411", "")
    mT_BBS412 = ReadINI("TABLE", "T_BBS412", "")
    mT_BBS413 = ReadINI("TABLE", "T_BBS413", "")
    mT_BBS414 = ReadINI("TABLE", "T_BBS414", "")

    mT_BBS501 = ReadINI("TABLE", "T_BBS501", "")
    mT_BBS502 = ReadINI("TABLE", "T_BBS502", "")
    mT_BBS503 = ReadINI("TABLE", "T_BBS503", "")
    mT_BBS504 = ReadINI("TABLE", "T_BBS504", "")
    mT_BBS505 = ReadINI("TABLE", "T_BBS505", "")
    mT_BBS506 = ReadINI("TABLE", "T_BBS506", "")
    
    mT_BBS601 = ReadINI("TABLE", "T_BBS601", "")
    mT_BBS602 = ReadINI("TABLE", "T_BBS602", "")
    mT_BBS603 = ReadINI("TABLE", "T_BBS603", "")
    mT_BBS604 = ReadINI("TABLE", "T_BBS604", "")
    mT_BBS605 = ReadINI("TABLE", "T_BBS605", "")
    mT_BBS606 = ReadINI("TABLE", "T_BBS606", "")
    mT_BBS607 = ReadINI("TABLE", "T_BBS607", "")
    mT_BBS901 = ReadINI("TABLE", "T_BBS901", "")
    mT_BBS902 = ReadINI("TABLE", "T_BBS902", "")
    mT_BBS903 = ReadINI("TABLE", "T_BBS903", "")

'-- ICS
    mT_ICS001 = ReadINI("TABLE", "T_ICS001", "")
    mT_ICS002 = ReadINI("TABLE", "T_ICS002", "")
    mT_ICS101 = ReadINI("TABLE", "T_ICS101", "")
    mT_ICS102 = ReadINI("TABLE", "T_ICS102", "")
    mT_ICS103 = ReadINI("TABLE", "T_ICS103", "")
    mT_ICS201 = ReadINI("TABLE", "T_ICS201", "")
    mT_ICS202 = ReadINI("TABLE", "T_ICS202", "")
    mT_ICS301 = ReadINI("TABLE", "T_ICS301", "")
    mT_ICS302 = ReadINI("TABLE", "T_ICS302", "")
End Sub
