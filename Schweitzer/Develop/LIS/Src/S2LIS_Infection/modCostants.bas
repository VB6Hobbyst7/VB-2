Attribute VB_Name = "modCostants"
Option Explicit

Global Const COL_DIV = ";"
Global Const END_DIV = "◀"


Global Const TB_COM001 = "s2com001"
Global Const TB_COM006 = "s2com006"
Global Const TB_COM007 = "s2com007"     '폼관리마스터
Global Const TB_COM008 = "s2com008"
Global Const TB_COM009 = "s2com009"
Global Const TB_COM010 = "s2com010"
Global Const TB_LAB001 = "s2lab001"     '검사항목마스터
Global Const TB_LAB002 = "s2lab002"     '검사동의어마스터
Global Const TB_LAB003 = "s2lab003"     '검사별장비마스터
Global Const TB_LAB004 = "s2lab004"     '지정검체마스터
Global Const TB_LAB005 = "s2lab005"     '기준치마스터
Global Const TB_LAB006 = "s2lab006"     '장비마스터
Global Const TB_LAB007 = "s2lab007"     '외부검사마스터
Global Const TB_LAB008 = "s2lab008"     'Worksheet마스터
Global Const TB_LAB009 = "s2lab009"     '공지사항내역
Global Const TB_LAB011 = "s2lab011"     'QC마스터
Global Const TB_LAB012 = "s2lab012"     'QC검사정보마스터
Global Const TB_LAB013 = "s2lab013"     '미생물QC마스터
Global Const TB_LAB014 = "s2lab014"     'QC컨트롤마스터
Global Const TB_LAB015 = "s2lab015"     '직원마스터
Global Const TB_LAB031 = "s2lab031"     '공통마스터1
Global Const TB_LAB032 = "s2lab032"     '공통마스터2
Global Const TB_LAB034 = "s2lab034"     '공통Template마스터
Global Const TB_LAB310 = "s2lab310"     '
Global Const TB_LAB315 = "s2lab315"     '감염관리 Header
Global Const TB_LAB316 = "s2lab316"     '감염관리 Body

Global Const TB_HIS005 = "orac1.ccusermt" '의사마스터(jikjong 의사:'HAA', 간호사: 'HAB')


'** OCS Master =======================================
Global Const TB_Dept = "orac1.ccdeptct" '부서마스터
'=====================================================

Global Const LC3_INFECTION = "C259"      ' 감염관리 의뢰검체
Global Const LC3_INFECTIONTEST = "C260"  ' 감염관리 균종관리
Global Const LC4_Infection = "C428"      ' 감염관리 검사방법
Global Const LC3_ElectronicSign = "C245" ' 전자서명
Global Const LC4_FootWard = "R100"       ' 감염관리 레포트 바닥글
    
Global Const StsCd_LIS_MidRst = 4       '중간보고
Global Const StsCd_LIS_FinRst = 5       '결과확인(최종결과)
Global Const StsCd_LIS_ModRst = 6       '결과수정
    
' 공통코드2 (LAB032) Index 상수
Global Const CD2_OrdSpc = "C259"        '의뢰검체
Global Const CD2_Micro = "C260"         '균종


Public Const DCM_Black = vbBlack          '검정색
Public Const DCM_White = vbWhite         '하얀색
Public Const DCM_Yellow = vbYellow       '노란색
Public Const DCM_Red = vbRed             '빨간색
Public Const DCM_Green = vbGreen         '녹색
Public Const DCM_Blue = vbBlue           '파란색
Public Const DCM_Magenta = vbMagenta     '자홍색
Public Const DCM_Cyan = vbCyan           '청록색

Public Const DCM_Grey = &H808080         '회색 --> 이건 쓰지말것. 나중에 지울것임.
Public Const DCM_Gray = &H808080         '회색
Public Const DCM_MidGray = &HC0C0C0
Public Const DCM_LightGrey = &HE0E0E0    '옅은회색
Public Const DCM_LightGray = &HE0E0E0    '옅은회색 --> 이것두..스펠이 틀려서리...
Public Const DCM_LightPink = &HF7F3F8    '옅은 분홍색
Public Const DCM_LightRed = &H7477EF     '옅은 빨간색
Public Const DCM_LightBlue = &HDF6A3E    '옅은 파란색
Public Const DCM_MidBlue = &HB9602F      '옅은 파란색

Public Const DCM_Brown = &H4A4189        '갈색

'Title Color
Public Const DCM_Title_Green = &HCDD19E  '연두색비스무리.. ^^;
Public Const DCM_Title_Pink = &HD8A9D6   '분홍색비스무리.. ^^;
Public Const DCM_Title_Blue = &HF9A071   '파란색비스무리.. ^^;

Global Const INIT_USER_SEC = "USER"
Global Const INIT_UID_KEY = "UID"
Global Const INIT_UNM_KEY = "UNM"
Global Const INIT_PWD_KEY = "PWD"
Global Const gintMAX_SIZE = 255

Public P_SLIDE_SERVER_PATH As String

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long

