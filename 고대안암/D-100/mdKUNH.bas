Attribute VB_Name = "mdKUNH"
Option Explicit

'결과 입력
Declare Function ResultList Lib "P_SLDLL2.dll" (ByVal out_jobsect As String, ByVal out_Userid As String, ByVal out_Result As Variant, _
                                                ByVal out_eqipcd As String, ByVal out_autoryn As String) As Integer

'접수정보 조회
Declare Function ExaminfoList Lib "P_SLDLL2.dll" (ByVal in_Flag As String, ByVal In_Spcid As String, ByVal in_execdate As String, out_order As Variant) As Integer


'연결
Declare Function TuxedoInit Lib "P_SLDLL2.dll" (ByVal in_usrname As String, ByVal in_cltname As String, ByVal in_svrid As String) As Integer

'종료
Declare Function TuxedoTerm Lib "P_SLDLL2.dll" () As Integer

'로그인
Declare Function UserChk Lib "P_SLDLL2.dll" (ByVal in_userid As String, ByVal in_pass As String, ByVal in_locate As String, out_usernm As Variant) As Integer

'''Declare Function UserChk1 Lib "P_SLDLL2.dll" (ByVal in_userid As String, ByVal in_pass As String, ByVal in_locate As String, ByVal out_usernm As Variant) As Integer


'function UserChk(in_userid, in_pass, in_locate: AnsiString; var out_usernm:variant): Integer; StdCall;

'(in_usrname,in_cltname,in_svrid:

Public sEMRUser As String * 8
Public sEMRID As String * 6
Public sEMRPW As String * 2

'''Dim sEMRUser As String * 8
'''Dim sEMRID As String * 6
'''Dim sEMRPW As String * 2

