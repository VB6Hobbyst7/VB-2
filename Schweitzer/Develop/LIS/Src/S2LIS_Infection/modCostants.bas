Attribute VB_Name = "modCostants"
Option Explicit

Global Const COL_DIV = ";"
Global Const END_DIV = "��"


Global Const TB_COM001 = "s2com001"
Global Const TB_COM006 = "s2com006"
Global Const TB_COM007 = "s2com007"     '������������
Global Const TB_COM008 = "s2com008"
Global Const TB_COM009 = "s2com009"
Global Const TB_COM010 = "s2com010"
Global Const TB_LAB001 = "s2lab001"     '�˻��׸񸶽���
Global Const TB_LAB002 = "s2lab002"     '�˻絿�Ǿ����
Global Const TB_LAB003 = "s2lab003"     '�˻纰��񸶽���
Global Const TB_LAB004 = "s2lab004"     '������ü������
Global Const TB_LAB005 = "s2lab005"     '����ġ������
Global Const TB_LAB006 = "s2lab006"     '��񸶽���
Global Const TB_LAB007 = "s2lab007"     '�ܺΰ˻縶����
Global Const TB_LAB008 = "s2lab008"     'Worksheet������
Global Const TB_LAB009 = "s2lab009"     '�������׳���
Global Const TB_LAB011 = "s2lab011"     'QC������
Global Const TB_LAB012 = "s2lab012"     'QC�˻�����������
Global Const TB_LAB013 = "s2lab013"     '�̻���QC������
Global Const TB_LAB014 = "s2lab014"     'QC��Ʈ�Ѹ�����
Global Const TB_LAB015 = "s2lab015"     '����������
Global Const TB_LAB031 = "s2lab031"     '���븶����1
Global Const TB_LAB032 = "s2lab032"     '���븶����2
Global Const TB_LAB034 = "s2lab034"     '����Template������
Global Const TB_LAB310 = "s2lab310"     '
Global Const TB_LAB315 = "s2lab315"     '�������� Header
Global Const TB_LAB316 = "s2lab316"     '�������� Body

Global Const TB_HIS005 = "orac1.ccusermt" '�ǻ縶����(jikjong �ǻ�:'HAA', ��ȣ��: 'HAB')


'** OCS Master =======================================
Global Const TB_Dept = "orac1.ccdeptct" '�μ�������
'=====================================================

Global Const LC3_INFECTION = "C259"      ' �������� �Ƿڰ�ü
Global Const LC3_INFECTIONTEST = "C260"  ' �������� ��������
Global Const LC4_Infection = "C428"      ' �������� �˻���
Global Const LC3_ElectronicSign = "C245" ' ���ڼ���
Global Const LC4_FootWard = "R100"       ' �������� ����Ʈ �ٴڱ�
    
Global Const StsCd_LIS_MidRst = 4       '�߰�����
Global Const StsCd_LIS_FinRst = 5       '���Ȯ��(�������)
Global Const StsCd_LIS_ModRst = 6       '�������
    
' �����ڵ�2 (LAB032) Index ���
Global Const CD2_OrdSpc = "C259"        '�Ƿڰ�ü
Global Const CD2_Micro = "C260"         '����


Public Const DCM_Black = vbBlack          '������
Public Const DCM_White = vbWhite         '�Ͼ��
Public Const DCM_Yellow = vbYellow       '�����
Public Const DCM_Red = vbRed             '������
Public Const DCM_Green = vbGreen         '���
Public Const DCM_Blue = vbBlue           '�Ķ���
Public Const DCM_Magenta = vbMagenta     '��ȫ��
Public Const DCM_Cyan = vbCyan           'û�ϻ�

Public Const DCM_Grey = &H808080         'ȸ�� --> �̰� ��������. ���߿� �������.
Public Const DCM_Gray = &H808080         'ȸ��
Public Const DCM_MidGray = &HC0C0C0
Public Const DCM_LightGrey = &HE0E0E0    '����ȸ��
Public Const DCM_LightGray = &HE0E0E0    '����ȸ�� --> �̰͵�..������ Ʋ������...
Public Const DCM_LightPink = &HF7F3F8    '���� ��ȫ��
Public Const DCM_LightRed = &H7477EF     '���� ������
Public Const DCM_LightBlue = &HDF6A3E    '���� �Ķ���
Public Const DCM_MidBlue = &HB9602F      '���� �Ķ���

Public Const DCM_Brown = &H4A4189        '����

'Title Color
Public Const DCM_Title_Green = &HCDD19E  '���λ��񽺹���.. ^^;
Public Const DCM_Title_Pink = &HD8A9D6   '��ȫ���񽺹���.. ^^;
Public Const DCM_Title_Blue = &HF9A071   '�Ķ����񽺹���.. ^^;

Global Const INIT_USER_SEC = "USER"
Global Const INIT_UID_KEY = "UID"
Global Const INIT_UNM_KEY = "UNM"
Global Const INIT_PWD_KEY = "PWD"
Global Const gintMAX_SIZE = 255

Public P_SLIDE_SERVER_PATH As String

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long

