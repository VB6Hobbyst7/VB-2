VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollectioin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Class �� : clsCollection(DLL��)
'|  2. ��    �� : ä��/��������(LAB201)�� �����Ѵ�.
'|  3. �� �� �� : ��̰�
'|
'|  CopyRight(C) 1999 ��ÿ�Ƽ����
'|
'+--------------------------------------------------------------------------------------+


'%  Ŭ���� clsCollection�� Data Attributes
Public SpcYY As String      '/* ��ü��ȣ-�⵵   */
Public SpcNo As Long        '/* ��ü��ȣ        */
Public PtId As String       '/* ȯ�� ID         */
Public PtNm As String       '/* ȯ�ڸ�         */
Public AgeDay As Long       '/* ȯ�� �Ϸ�       */
Public Sex As String        '/* ȯ�� ����   */
Public BedinDt As String    '/* �Կ���          */
Public DeptCd As String     '/* �����          */
Public OrdDoct As Integer   '/* ó����          */
Public MajDoct As Integer   '/* ��ġ��          */
Public WorkArea As String   '/* Work Area(������ȣ��) */
Public AccDt As String      '/* ������(������ȣ��)    */
Public AccSeq As Integer    '/* ��������(������ȣ��)  */
Public StsCd As String      '/* STATUS - '0 ':ó��,1:ä��,2:����,3:In-Process,4:���,5:Ȯ��,6:���� */
Public ReqTotCnt As Integer '/* Required Item Total Count */
Public ReqInputCnt As Integer  '/* Required Item Input Count */
Public VfyDt As String      '/* ���Ȯ������    */
Public VfyTm As String      '/* ���Ȯ�νð�    */
Public VfyId As Integer     '/* ���Ȯ����      */
Public ColDt As String      '/* ä������        */
Public ColTm As String      '/* ä���ð�        */
Public ColId As Integer     '/* ä����          */
Public RcvDt As String      '/* ��������        */
Public RcvTm As String      '/* �����ð�        */
Public RcvId As Integer     '/* ������          */
Public EntDt As String      '/* �Է�����        */
Public EntTm As String      '/* �Է½ð�        */
Public EntId As Integer     '/* �Է���          */
Public SpcCd As String      '/* ��ü�ڵ�        */
Public SpcNm As String      '/* ��ü��        */
Public MultiFg As String    '/* ������ü����  */
Public OrgAccNo As String   '/* �� ������ȣ     */
Public WardId As String     '/* �����ڵ�        */
Public RoomId As String     '/* �����ڵ�     */
Public BedId As String      '/* ħ���ڵ�    */
Public HosilId As String    '/* ȣ���ڵ�    */
Public FootNoteFg As String '/* FOOTNOTE����('0':��,'1':��) */
Public StoreCd As String    '/* ��ü���� �ڵ�   */
Public RptFg As String      '/* REPORT��¿���('0':��,'1':��) */
Public TestDiv As String    '/* �˻籸�� - 1:��Ÿ�˻�,2:�̻����˻� */
Public QcFg As String       '/* QC����('0':��,'1':��) */
Public RmkCd As String      '/* ��ü Remark �ڵ� */
Public StatFg As String     '/* ���޿���    */
Public BuildCd As String    '/* �ǹ��ڵ�    */
Public BuildNm As String    '/* �ǹ��ڵ��    */
Public OrgBuildCd As String '/* ä���� �ǹ��ڵ�    */
Public TestNames As String  '/* Barcode�� ��µ� �˻���� */

Public OrdDt As String
Public OrdNo As Integer
Public OrdSeq As Integer
Public OrdCd As String

Public SpcGrp As String     '/* ��ü��  */
Public SpcSeq As String     '/* ������ü�� Seq */

'= Lab ä������ ���� Data Type ='
Private Type TpLabGroup
   BuildCd As String
   TestDiv As String
   WorkArea As String
   SpcCd As String
   StoreCd As String
   StatFg As String
   ColDate As String
   MultiFg As String
   SpcGrp As String
   OrdDt As String
   OrdNo As Integer
   OrdSeq As Integer
   OrdCd As String
   DeptCd As String     '/* �����          */
   OrdDoct As Integer   '/* ó����          */
   MajDoct As Integer   '/* ��ġ��          */
   LabelCnt As Integer
   TestNm As String     '/* �˻� ����   */
End Type

Private Type LabNumbers
   WorkArea As String
   AccDt As String
   AccSeq As Integer
   BuildCd As String
End Type
   
Private Type BarcodeLabel
   WorkArea As String
   AccDt As String
   AccSeq As Integer
   ColDt As String
   ColTm As String
   OrdDt As String
   StatFg As String
   SpcNo As String
   Location As String
   PtNm As String
   PtId As String
   SpcNm As String
   StoreCd As String
   WardId As String
   TestNames As String
   LabelCnt As String
End Type
   
   
Public ErrMsg As String
Public ColCount As Integer

'���뺯�� - �����ϴ� ������Ʈ�κ��� ������
'Public MyOraSE As Object     'OraSession
Public MyDb As Object     'OraDatabase

Public CollectDone As Boolean

Private blnBatchCol As Boolean
Private ColDone As Boolean  '������ü ä���Ϸ� ����
Private ReadyToCollect As Boolean
Private KeyCount As Integer
Private MultiSpcFlag As String
Private BarCount As Integer
Private SortList As ListBox
Private LabGroup() As TpLabGroup
Private CreateLabNo() As LabNumbers
Private BarcodeBuffer() As BarcodeLabel
Private SqlBuffer() As String
Private MySql As New clsLISSqlStatement
'Private MyBar As New clsBarcode

Private sWorkDt As String
Private sWorkTm As String
Private iWorkSeq As Integer



'% Method 1 : CollectQuery
'%                 Parameter�� ���� Sql�� �����ϰ�, �� �ʵ��� ����
'%                 Ŭ���� clsCollection�� Data Attribute�� �����Ѵ�.

Public Function CollectQuery(ByVal pWorkArea As String, _
                                         ByVal pAccDt As String, ByVal pAccSeq As Integer) As Boolean

   Dim SqlStmt As String
   Dim MyOraDS As Object
   
   SqlStmt = "Select a.* " & _
                 "From " & T_LAB201 & " a " & _
                 "Where a.workarea = '" & pWorkArea & "' " & _
                 "and     a.accdt = '" & pAccDt & "' " & _
                 "and     a.accseq = " & pAccSeq
   
   Set MyOraDS = OpenRecordSet(SqlStmt)  'Sql ����
   If MyOraDS.EOF Then
      CollectQuery = False
      Exit Function
   End If
   
   With MyOraDS
      
      .MoveFirst
         
      SpcYY = Trim("" & .Fields("SpcYy").Value)
      SpcNo = Val("" & .Fields("SpcNo").Value)
      PtId = Trim("" & .Fields("PtId").Value)
      Sex = Trim("" & .Fields("Sex").Value)
      AgeDay = Val("" & .Fields("AgeDay").Value)
      BedinDt = Trim("" & .Fields("BedInDt").Value)
      DeptCd = Trim("" & .Fields("DeptCd").Value)
      OrdDoct = Val("" & .Fields("OrdDoct").Value)
      MajDoct = Val("" & .Fields("MajDoct").Value)
      WorkArea = Trim("" & .Fields("WorkArea").Value)
      AccDt = Trim("" & .Fields("AccDt").Value)
      AccSeq = Val("" & .Fields("AccSeq").Value)
      StsCd = Trim("" & .Fields("StsCd").Value)
      ReqTotCnt = Val("" & .Fields("ReqTotCnt").Value)
      ReqInputCnt = Val("" & .Fields("ReqInputCnt").Value)
      VfyDt = Trim("" & .Fields("VfyDt").Value)
      VfyTm = Trim("" & .Fields("VfyTm").Value)
      VfyId = Val("" & .Fields("VfyId").Value)
      ColDt = Trim("" & .Fields("ColDt").Value)
      ColTm = Trim("" & .Fields("ColTm").Value)
      ColId = Val("" & .Fields("ColId").Value)
      RcvDt = Trim("" & .Fields("RcvDt").Value)
      RcvTm = Trim("" & .Fields("RcvTm").Value)
      RcvId = Val("" & .Fields("RcvId").Value)
      EntDt = Trim("" & .Fields("EntDt").Value)
      EntTm = Trim("" & .Fields("EntTm").Value)
      EntId = Val("" & .Fields("EntId").Value)
      SpcCd = Trim("" & .Fields("SpcCd").Value)
      MultiFg = Trim("" & .Fields("MultiFg").Value)
      OrgAccNo = Trim("" & .Fields("OrgAccNo").Value)
      WardId = Trim("" & .Fields("WardId").Value)
      RoomId = Trim("" & .Fields("RoomId").Value)
      BedId = Trim("" & .Fields("BedId").Value)
      HosilId = Trim("" & .Fields("HosilId").Value)
      FootNoteFg = Trim("" & .Fields("FootNoteFg").Value)
      StoreCd = Trim("" & .Fields("StoreCd").Value)
      RptFg = Trim("" & .Fields("RptFg").Value)
      TestDiv = Trim("" & .Fields("TestDiv").Value)
      QcFg = Trim("" & .Fields("QcFg").Value)
      RmkCd = Trim("" & .Fields("RmkCd").Value)
      StatFg = Trim("" & .Fields("StatFg").Value)
      BuildCd = Trim("" & .Fields("BuildCd").Value)
      
   End With
   
End Function


Private Sub Class_Initialize()
   blnBatchCol = False
   Set MySql = New clsLISSqlStatement
   Set SortList = frmControls.lstList
   Call InitRtn
End Sub

Private Sub Class_Terminate()
   Set MySql = Nothing
   Set SortList = Nothing
End Sub

Public Sub InitRtn()
   SortList.Clear
   KeyCount = 0
   ColCount = 0
   ReadyToCollect = False
   CollectDone = False
   ColDone = False
   Erase SqlBuffer
   ErrMsg = ""
End Sub









