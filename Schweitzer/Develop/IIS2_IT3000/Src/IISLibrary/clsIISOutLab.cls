VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISOutLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WorkArea As String     '/* Work Area(������ȣ��) */
Public AccDt As String        '/* ������(������ȣ��)    */
Public AccSeq As Integer      '/* ��������(������ȣ��)  */
Public TestCd As String       '/* �˻��׸�        */
Public PtId As String         '/* ȯ��ID  */
Public SpcCd As String        '/* ��ü�ڵ�  */
Public OutLabCd As String     '/* ����ڵ�        */
Public StsCd As String        '/* STATUS          */
Public RcvDt As String        '/* ��������        */
Public SendDt As String       '/* SEND Date       */
Public ChargeDt As String     '/* û������        */
Public SendId As String       '/* ������ ID   */

