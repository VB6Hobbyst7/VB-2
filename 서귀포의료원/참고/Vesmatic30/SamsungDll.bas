Attribute VB_Name = "SamsungDll"
Option Explicit

'�������� ��ȸ
Declare Function ExaminfoList2 Lib "P_SLDLL.dll" (ByVal In_Spcid As String, out_order As Variant) As Integer

'��� �Է�
Declare Function ResultList2 Lib "P_SLDLL.dll" (ByVal out_eqipcode As String, ByVal out_cnt As Integer, ByVal out_Result As String) As Integer
