VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISEqpChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISEqpChoice.cls
'   작성자  : 오세원
'   내  용  : 현재 PC에서 사용할 검사장비 선택 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

'## EqpType Enum
Public Enum EqpCommTypeEnum
    ccBidirection = 0   '양방향
    ccUnidirection = 1  '단방향
End Enum

'## Barcode Enum
Public Enum BarcodeUseEnum
    ccNoBarcode = 0     '바코드 사용안함
    ccUseBarcode = 1    '바코드 사용
End Enum

'## AutoVfy Enum
Public Enum AutoVfyEnum
    ccNo = 0            '사용안함
    ccYes = 1           '사용
End Enum

'## HLCheck Enum
Public Enum HLCheckEnum
    ccNoCheck = 0       'H/L에 상관없이 자동결과등록
    ccCheck = 1         'H/L에 해당되면 자동결과등록 불가능
End Enum

'## BarPos Enum
Public Enum BarPosEnum
    ccPC = 0            'PC
    ccEqp = 1           '장비
End Enum

'## AutoRerun Enum
Public Enum AutoRerunEnum
    ccNoAutoRerun = 0
    ccAutoRerun = 1
End Enum

Private mSQL            As clsIISMasterSql

Private mEqpCd1         As String           '장비코드1
Private mEqpNm1         As String           '장비명1
Private mEqpKey1        As String           '장비 아이콘 Key1
Private mAutoVfy1       As AutoVfyEnum      '자동결과등록 사용여부1
Private mBarcode1       As BarcodeUseEnum   '장비1의 바코드 사용유무
Private mHLCheck1       As HLCheckEnum      '장비1의 자동결과등록시 H/L을 체크유무
Private mBarPos1        As BarPosEnum       '장비1의 바코드 리더기 위치
Private mEqpCd2         As String           '장비코드2
Private mEqpNm2         As String           '장비명2
Private mEqpKey2        As String           '장비 아이콘 Key2
Private mAutoVfy2       As AutoVfyEnum      '자동결과등록 사용여부2
Private mBarcode2       As BarcodeUseEnum   '장비2의 바코드 사용유무
Private mHLCheck2       As HLCheckEnum      '장비2의 자동결과등록시 H/L을 체크유무
Private mBarPos2        As BarPosEnum       '장비2의 바코드 리더기 위치

'   - 자동재검 유뮤 변수 추가
Private mAutoRerun1     As AutoRerunEnum    '장비1의 자동재검 유무
Private mAutoRerun2     As AutoRerunEnum    '장비2의 자동재검 유무

Public Property Get EqpCd1() As String
    EqpCd1 = mEqpCd1
End Property

Public Property Let EqpCd1(ByVal vData As String)
    mEqpCd1 = vData
End Property

Public Property Get EqpNm1() As String
    EqpNm1 = mEqpNm1
End Property

Public Property Let EqpNm1(ByVal vData As String)
    mEqpNm1 = vData
End Property

Public Property Get EqpKey1() As String
    EqpKey1 = mEqpKey1
End Property

Public Property Let EqpKey1(ByVal vData As String)
    mEqpKey1 = vData
End Property

Public Property Get AutoVfy1() As String
    AutoVfy1 = mAutoVfy1
End Property

Public Property Let AutoVfy1(ByVal vData As String)
    mAutoVfy1 = vData
End Property

Public Property Get Barcode1() As BarcodeUseEnum
    Barcode1 = mBarcode1
End Property

Public Property Let Barcode1(ByVal vData As BarcodeUseEnum)
    mBarcode1 = vData
End Property

Public Property Get HLCheck1() As HLCheckEnum
    HLCheck1 = mHLCheck1
End Property

Public Property Let HLCheck1(ByVal vData As HLCheckEnum)
    mHLCheck1 = vData
End Property

Public Property Get BarPos1() As BarPosEnum
    BarPos1 = mBarPos1
End Property

Public Property Let BarPos1(ByVal vData As BarPosEnum)
    mBarPos1 = vData
End Property

Public Property Get EqpCd2() As String
    EqpCd2 = mEqpCd2
End Property

Public Property Let EqpCd2(ByVal vData As String)
    mEqpCd2 = vData
End Property

Public Property Get EqpNm2() As String
    EqpNm2 = mEqpNm2
End Property

Public Property Let EqpNm2(ByVal vData As String)
    mEqpNm2 = vData
End Property

Public Property Get EqpKey2() As String
    EqpKey2 = mEqpKey2
End Property

Public Property Let EqpKey2(ByVal vData As String)
    mEqpKey2 = vData
End Property

Public Property Get AutoVfy2() As String
    AutoVfy2 = mAutoVfy2
End Property

Public Property Let AutoVfy2(ByVal vData As String)
    mAutoVfy2 = vData
End Property

Public Property Get Barcode2() As BarcodeUseEnum
    Barcode2 = mBarcode2
End Property

Public Property Let Barcode2(ByVal vData As BarcodeUseEnum)
    mBarcode2 = vData
End Property

Public Property Get HLCheck2() As HLCheckEnum
    HLCheck2 = mHLCheck2
End Property

Public Property Let HLCheck2(ByVal vData As HLCheckEnum)
    mHLCheck2 = vData
End Property

Public Property Get BarPos2() As BarPosEnum
    BarPos2 = mBarPos2
End Property

Public Property Let BarPos2(ByVal vData As BarPosEnum)
    mBarPos2 = vData
End Property

'   - 자동재검 유무 프로퍼티 추가
Public Property Get AutoRerun1() As AutoRerunEnum
    AutoRerun1 = mAutoRerun1
End Property

Public Property Let AutoRerun1(ByVal vData As AutoRerunEnum)
    mAutoRerun1 = vData
End Property

Public Property Get AutoRerun2() As AutoRerunEnum
    AutoRerun2 = mAutoRerun2
End Property

Public Property Let AutoRerun2(ByVal vData As AutoRerunEnum)
    mAutoRerun2 = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 현재 PC에서 사용할 장비 설정
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function SetEqp() As Boolean
    Dim strFileNm   As String   'ini파일 경로+이름
    Dim strAppName  As String   '레지의 AppName
    
    '## 1.ini파일에서 해당 장비코드의 아이콘 Key를 구함
    '## 2.레지에 저장
On Error GoTo Errors
    strFileNm = IniPath & "\" & "IIS.ini"
    strAppName = AppName
    
    '   - 장비 1, 2의 자동재검 옵션추가
    '## 검사장비1
    If mEqpCd1 <> "" Then
        EqpKey1 = mReadINI(strFileNm, "EQPKEY", mEqpCd1)
        SaveSetting strAppName, "Machine1", "EqpCd", mEqpCd1
        SaveSetting strAppName, "Machine1", "EqpKey", mEqpKey1
        SaveSetting strAppName, "Machine1", "AutoVfy", mAutoVfy1
        SaveSetting strAppName, "Machine1", "Barcode", mBarcode1
        SaveSetting strAppName, "Machine1", "HLCheck", mHLCheck1
        SaveSetting strAppName, "Machine1", "BarPos", mBarPos1
        SaveSetting strAppName, "Machine1", "AutoRerun", mAutoRerun1
    Else
        SaveSetting strAppName, "Machine1", "EqpCd", ""
        SaveSetting strAppName, "Machine1", "EqpKey", ""
        SaveSetting strAppName, "Machine1", "AutoVfy", ""
        SaveSetting strAppName, "Machine1", "Barcode", ""
        SaveSetting strAppName, "Machine1", "HLCheck", ""
        SaveSetting strAppName, "Machine1", "BarPos", ""
        SaveSetting strAppName, "Machine1", "AutoRerun", ""
    End If
    
    '## 검사장비2
    If mEqpCd2 <> "" Then
        EqpKey2 = mReadINI(strFileNm, "EQPKEY", mEqpCd2)
        SaveSetting strAppName, "Machine2", "EqpCd", mEqpCd2
        SaveSetting strAppName, "Machine2", "EqpKey", mEqpKey2
        SaveSetting strAppName, "Machine2", "AutoVfy", mAutoVfy2
        SaveSetting strAppName, "Machine2", "Barcode", mBarcode2
        SaveSetting strAppName, "Machine2", "HLCheck", mHLCheck2
        SaveSetting strAppName, "Machine2", "BarPos", mBarPos2
        SaveSetting strAppName, "Machine2", "AutoRerun", mAutoRerun2
    Else
        SaveSetting strAppName, "Machine2", "EqpCd", ""
        SaveSetting strAppName, "Machine2", "EqpKey", ""
        SaveSetting strAppName, "Machine2", "AutoVfy", ""
        SaveSetting strAppName, "Machine2", "Barcode", ""
        SaveSetting strAppName, "Machine2", "HLCheck", ""
        SaveSetting strAppName, "Machine2", "BarPos", ""
        SaveSetting strAppName, "Machine2", "AutoRerun", ""
    End If
    SetEqp = True
    Exit Function
    
Errors:
    Error.SetLog App.EXEName, "clsEqpChoice", "SetEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 현재 PC에서 사용하는 장비 조회
'   반환 : True(성공), False(실패)
'-----------------------------------------------------------------------------'
Public Function GetEqp() As Boolean
    Dim strAppName  As String   '레지의 AppName
    
On Error GoTo Errors
    strAppName = AppName
    
    '## NOTE: 프로그램적으로는 AutoVfy, Barcode, HLCheck에 '0','1'이외의 값을
    '   입력할수 없지만 사용자가 강제로 레지스트리값을 변경할경우에는 런타임에러가 발생하기
    '   때문에 프로그램을 변경해야 할것 같음.
    '   - 장비 1, 2의 자동재검 옵션추가
    '## 검사장비1
    mEqpCd1 = GetSetting(strAppName, "Machine1", "EqpCd", "")
    If mEqpCd1 <> "" Then
        mEqpNm1 = GetEqpNm(mEqpCd1)
        mEqpKey1 = GetSetting(strAppName, "Machine1", "EqpKey", "")
        mAutoVfy1 = CLng(GetSetting(strAppName, "Machine1", "AutoVfy", "0"))
        mBarcode1 = CLng(GetSetting(strAppName, "Machine1", "Barcode", "0"))
        mHLCheck1 = CLng(GetSetting(strAppName, "Machine1", "HLCheck", "0"))
        mBarPos1 = CLng(GetSetting(strAppName, "Machine1", "BarPos", "0"))
        mAutoRerun1 = CLng(GetSetting(strAppName, "Machine1", "AutoRerun", "0"))
    End If
    
    '## 검사장비2
    mEqpCd2 = GetSetting(strAppName, "Machine2", "EqpCd", "")
    If mEqpCd2 <> "" Then
        mEqpNm2 = GetEqpNm(mEqpCd2)
        mEqpKey2 = GetSetting(strAppName, "Machine2", "EqpKey", "")
        mAutoVfy2 = CLng(GetSetting(strAppName, "Machine2", "AutoVfy", "0"))
        mBarcode2 = CLng(GetSetting(strAppName, "Machine2", "Barcode", "0"))
        mHLCheck2 = CLng(GetSetting(strAppName, "Machine2", "HLCheck", "0"))
        mBarPos2 = CLng(GetSetting(strAppName, "Machine2", "BarPos", "0"))
        mAutoRerun2 = CLng(GetSetting(strAppName, "Machine2", "AutoRerun", "0"))
    End If
    GetEqp = True
    Exit Function
    
Errors:
    Error.SetLog App.EXEName, "clsEqpChoice", "GetEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 사용중인 장비코드, 장비명 조회
'   반환 : Recordset
'-----------------------------------------------------------------------------'
Public Function GetUsingEqp() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectUsingEqp, , adCmdText)
    Set GetUsingEqp = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsEqpChoice", "GetUsingEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드의 장비명 조회
'   인수 :
'       1.pEqpCd : 장비코드
'   반환 : 존재하는 장비이면 장비명 반환, 존재하지 않으면 Null반환
'-----------------------------------------------------------------------------'
Public Function GetEqpNm(ByVal pEqpCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectEqpCd(pEqpCd), , adCmdText)
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetEqpNm = Rs.Fields("EQPNM").Value
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsEqpChoice", "GetEqpNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub
