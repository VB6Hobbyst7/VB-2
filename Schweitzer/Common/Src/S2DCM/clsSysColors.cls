VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysColors"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'** VB���� �����ϴ� ��������...
Public DCM_Black As Long
Public DCM_White As Long
Public DCM_Yellow As Long
Public DCM_Red As Long
Public DCM_Green As Long
Public DCM_Blue As Long
Public DCM_Magenta As Long
Public DCM_Cyan As Long
Public DCM_MidGray As Long

'** ���⼭���ʹ� User-Defined Color
'Public DCM_Grey As Long
Public DCM_Gray As Long
'Public DCM_LightGrey As Long
Public DCM_LightGray As Long
Public DCM_LightPink As Long
Public DCM_LightRed As Long
Public DCM_LightBlue As Long
Public DCM_MidBlue As Long
Public DCM_Brown As Long
Public DCN_Green As Long

Public DCM_Title_Green As Long
Public DCM_Title_Pink As Long
Public DCM_Title_Blue As Long

'�ϴû� : &H00FEDECD&

Private Sub Class_Initialize()

    DCM_Black = vbBlack         '������
    DCM_White = vbWhite         '�Ͼ��
    DCM_Yellow = vbYellow       '�����
    DCM_Red = vbRed             '������
    DCM_Green = vbGreen         '���
    DCM_Blue = vbBlue           '�Ķ���
    DCM_Magenta = vbMagenta     '��ȫ��
    DCM_Cyan = vbCyan           'û�ϻ�
    
'    DCM_Grey = &H808080         'ȸ�� --> �̰� ��������. ���߿� �������.
    DCM_Gray = &H808080         'ȸ��
    DCM_MidGray = &HC0C0C0
'    DCM_LightGrey = &HE0E0E0    '����ȸ��
    DCM_LightGray = &HE0E0E0    '����ȸ�� --> �̰͵�..������ Ʋ������...
    DCM_LightPink = &HF7F3F8    '���� ��ȫ��
    DCM_LightRed = &H7477EF     '���� ������
    DCM_LightBlue = &HDF6A3E    '���� �Ķ���
    DCM_MidBlue = &HB9602F      '���� �Ķ���
    
    DCM_Brown = &H4A4189        '����
    DCM_Green = &H8000&         '���λ�
    
    'Title Color
    DCM_Title_Green = &HCDD19E  '���λ��񽺹���.. ^^;
    DCM_Title_Pink = &HD8A9D6   '��ȫ���񽺹���.. ^^;
    DCM_Title_Blue = &HF9A071   '�Ķ����񽺹���.. ^^;
End Sub