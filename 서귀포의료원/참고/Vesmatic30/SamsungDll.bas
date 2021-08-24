Attribute VB_Name = "SamsungDll"
Option Explicit

'접수정보 조회
Declare Function ExaminfoList2 Lib "P_SLDLL.dll" (ByVal In_Spcid As String, out_order As Variant) As Integer

'결과 입력
Declare Function ResultList2 Lib "P_SLDLL.dll" (ByVal out_eqipcode As String, ByVal out_cnt As Integer, ByVal out_Result As String) As Integer
