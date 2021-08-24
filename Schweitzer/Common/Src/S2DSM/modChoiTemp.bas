Attribute VB_Name = "modChoiTemp"
Option Explicit

'임시 Table 변수 선언 ===================
Global Const T_COM006 = "com006"
Global Const T_COM007 = "com007"
Global Const T_COM008 = "com008"
Global Const T_COM009 = "com009"
Global Const T_COM010 = "com010"
'========================================

'인사 마스터 Table ====================================
Global Const Temp006_Fields0 = "empid"       '직원ID
Global Const Temp006_Fields1 = "emplngnm"    '직원 긴이름
Global Const Temp006_Fields7 = "deptcd"      '부서코드
'======================================================

'폼 마스터 Table ======================================
Global Const Temp007_Fields1 = "formid"      '폼ID
Global Const Temp007_Fields2 = "formnm"      '폼이름
Global Const Temp007_Fields3 = "formdesc"    '폼설명
'======================================================

'그룹 등록 마스터 Header Table ========================
Global Const Temp008_Fields0 = "groupid"     '그룹ID
Global Const Temp008_Fields1 = "groupnm"     '그룹이름
Global Const Temp008_Fields2 = "groupdesc"   '그룹설명
Global Const Temp008_Fields3 = "userfg"      '사용자 구분 'M':Manager, 'D':Developer, 'S':Supervisor
Global Const Temp008_Fields4 = "apsfg"       '진단병리 여:'1', 부:'0'
Global Const Temp008_Fields5 = "bbsfg"       '혈액은행 여:'1', 부:'0'
Global Const Temp008_Fields6 = "lisfg"       'LIS 여:'1', 부:'0'
'======================================================

'그룹 등록 마스터 Body Table ==========================
Global Const Temp009_Fields0 = "groupid"     '그룹ID
Global Const Temp009_Fields1 = "deptfg"      '부서구분
Global Const Temp009_Fields2 = "formid"      '폼ID
Global Const Temp009_Fields3 = "readfg"      '읽기권한 '0':없음, '1':있음
Global Const Temp009_Fields4 = "writefg"     '쓰기권한 '0':없음, '1':있음
Global Const Temp009_Fields5 = "printfg"     '출력권한 '0':없음, '1':있음
'======================================================

'사용자 관리 마스터 Table =============================
Global Const Temp010_Fields0 = "loginid"    '로그인ID
Global Const Temp010_Fields1 = "loginnm"    '로그인이름
Global Const Temp010_Fields2 = "empid"      '직원ID
Global Const Temp010_Fields3 = "logindesc"  '로그인 설명
Global Const Temp010_Fields4 = "groupid"    '그룹ID
'======================================================
