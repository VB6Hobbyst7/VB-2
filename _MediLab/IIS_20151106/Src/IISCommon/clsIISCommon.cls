VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCommon"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISCommon.cls
'   작성자  : 오세원
'   내  용  : 프로젝트에서 공통으로 사용하는 함수 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'
Option Explicit

'## AlwaysOnEnum
Public Enum AlwaysOnEnum
    ccOff = 0
    ccOn = 1
End Enum

'-----------------------------------------------------------------------------'
'   기능 : 컬럼 구분자
'-----------------------------------------------------------------------------'
Public Property Get DIV() As String
    DIV = Chr(19)
End Property

'-----------------------------------------------------------------------------'
'   기능 : 폼을 최상위로 원도우로 표시
'   인수 :
'       1.frmFrom   : 해당폼
'       2.OnOff     : 0(해제), 1(설정)
'-----------------------------------------------------------------------------'
Public Sub mAlwaysOn(ByVal pForm As Object, ByVal pFlag As AlwaysOnEnum)
    Dim hWndMode As Integer

    hWndMode = Choose(pFlag + 1, HWND_NOTOPMOST, HWND_TOPMOST)
    SetWindowPos pForm.hwnd, hWndMode, 0, 0, 10, 10, SWP_NOMOVE Or SWP_NOSIZE
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컴퓨터를 이름을 얻는 함수
'   반환 : 컴퓨터 이름
'-----------------------------------------------------------------------------'
Public Function mGetComNm() As String
   Dim sBuffer$
   Dim nSize    As Long
   Dim rtn      As Long
   
   sBuffer = String(256, Chr(0))
   rtn = GetComputerName(sBuffer$, Len(sBuffer))
   mGetComNm = sBuffer
End Function

'-----------------------------------------------------------------------------'
'   기능 : INI파일에서 정보읽기
'   인수 :
'       1.pFileNm   : INI파일명(경로포함)
'       2.pSection  : INI파일의 Section
'       3.pKey      : INI파일의 Key
'       4.pDefault  : 해당 키값이 없을때 Default값
'-----------------------------------------------------------------------------'
Public Function mReadINI(ByVal pFileNm As String, ByVal pSection As String, _
                         ByVal pKey As String, Optional ByVal pDefault As String = "") As String
    Dim p As String
    
    p = Space$(256)
    GetPrivateProfileString pSection, pKey, pDefault, p, 256, pFileNm
    mReadINI = Mid(Trim(p), 1, Len(Trim(p)) - 1)
End Function

'-----------------------------------------------------------------------------'
'   기능 : INI파일에서 정보쓰기 - 1.2.1: (2005-09-13)
'   인수 :
'       1.pFileNm   : INI파일명(경로포함)
'       2.pSection  : INI파일의 Section
'       3.pKey      : INI파일의 Key
'       4.pValue    : INI파일의 Key Value
'-----------------------------------------------------------------------------'
Public Sub mWriteINI(ByVal pFileNm, ByVal pSection As String, ByVal pKey As String, _
                     ByVal pValue As String)

    WritePrivateProfileString pSection, pKey, pValue, pFileNm
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

'-----------------------------------------------------------------------------'
'   기능 : 콤보박스에서 해당문자열을 찾기1
'   인수 :
'       1.pCombo    : ComboBox 컨트롤
'       2.pSerach   : 찾을문자열
'   반환 : -1(문자열을 찾지 못한경우), 이외에는 찾는 문자열의 ListIndex
'-----------------------------------------------------------------------------'
Public Function mFindCombo(ByRef pCombo As Object, ByVal pSearch As String) As Long
    Dim i As Long
    
    If pSearch = "" Then
        mFindCombo = -1
        Exit Function
    End If
    
    With pCombo
        For i = 0 To .ListCount - 1
            If .List(i) Like (pSearch & "*") Then
                mFindCombo = i
                Exit Function
            End If
        Next i
    End With
    mFindCombo = -1
End Function

'-----------------------------------------------------------------------------'
'   기능 : 콤보박스에서 해당문자열을 찾기2
'   인수 :
'       1.pCombo    : ComboBox 컨트롤
'       2.pSerach   : 찾을문자열
'   반환 : -1(문자열을 찾지 못한경우), 이외에는 찾는 문자열의 ListIndex
'-----------------------------------------------------------------------------'
Public Function mFindComboX(ByRef pCombo As Object, ByVal pSearch As String) As Long
    Dim i As Long
    
    If pSearch = "" Then
        mFindComboX = -1
        Exit Function
    End If
    
    With pCombo
        For i = 0 To .ListCount - 1
            If Trim(mGetP(.List(i), 2, DIV)) = pSearch Then
                mFindComboX = i
                Exit Function
            End If
        Next i
    End With
    mFindComboX = -1
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 Spread Sheet의 Data를 Clear
'   인수 :
'       - pSpread : 대상 Spread Sheet
'       - pCol    : Column Header를 Clear할 것인지
'       - pRow    : Row Header를 Clear할 것인지
'-----------------------------------------------------------------------------'
Public Sub mTblClear(ByVal pSpread As Object, Optional ByVal pCol As Boolean, _
                     Optional ByVal pRow As Boolean)
Dim ii As Integer
    
    If pSpread.MaxRows > 0 Then
        pSpread.Col = 1
        pSpread.Col2 = pSpread.MaxCols
        pSpread.Row = 1
        pSpread.Row2 = pSpread.MaxRows
        pSpread.BlockMode = True
        pSpread.Action = 12    '## ActionClearText
        pSpread.BlockMode = False
    End If
    
    If IsMissing(pCol) Then pCol = False
    If IsMissing(pRow) Then pCol = False
    
    '## Column Header Clear
    If pCol = True Then
        For ii = 1 To pSpread.MaxCols
            pSpread.Row = 0
            pSpread.Col = ii
            pSpread.Value = " "
        Next ii
    End If
    '## Row Header Clear
    If pRow = True Then
        For ii = 1 To pSpread.MaxRows
            pSpread.Col = 0
            pSpread.Row = ii
            pSpread.Value = " "
        Next ii
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : SetParent API의 래퍼함수
'   인수 :
'       - hWndChild  : 자식원도우의 핸들
'       - hWndParent : 부모원도우의 핸들
'-----------------------------------------------------------------------------'
Public Function mSetParent(ByVal hWndChild As Long, ByVal hWndParent As Long)
    Call SetParent(hWndChild, hWndParent)
End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당시간동안 시간을 지연
'   인수 :
'       - pInterval : 지연시간(1000분의 1초)
'-----------------------------------------------------------------------------'
Public Sub mSleep(ByVal pInterval As Long)
    Call Sleep(pInterval)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Workarea, AccDt, AccSeq를 이용해 접수번호를 구함(Workarea-AccDt-AccSeq)
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'   반환 : Workarea-AccDt-AccSeq 형식의 접수번호
'-----------------------------------------------------------------------------'
Public Function mGetAccNo(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As Long) As String
    mGetAccNo = pWorkarea & "-" & pAccDt & "-" & CStr(pAccSeq)
End Function

'-----------------------------------------------------------------------------'
'   기능 : 참고치의 From, To 값을 이용해 From-To 형식으로 반환
'   인수 :
'       - pRefFrVal : Reference From Value
'       - pRefToVal : Reference From Value
'   반환 : "From-To"형식의 참고치
'-----------------------------------------------------------------------------'
Public Function mGetRef(ByVal pRefFrVal As Single, ByVal pRefToVal As Single) As String
    If pRefFrVal = -1 And pRefToVal = -1 Then
        mGetRef = ""
    Else
        mGetRef = CStr(pRefFrVal) & " - " & CStr(pRefToVal)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 주민번호 앞자리를 이용해 나이계산 (By 김동렬)
'   인수 :
'       - pBirth : 주민번호 앞자리
'   반환 : 나이
'-----------------------------------------------------------------------------'
Public Function mGetAge(ByVal pBirth As String) As String
    Dim strYear         As String
    Dim strMonth        As String
    Dim strDay          As String
    Dim strPtYear       As String   '환자 생년월일의 년도
    Dim strPtMonth      As String   '환자 생년월일의 월
    Dim strPtDay        As String   '환자 생년월일의 일자
    Dim lngYear         As Long
    Dim lngMonth        As Long
    Dim lngThisYear     As Long
    Dim lngThisMonth    As Long
    Dim lngThisDay      As Long
    
    If Trim$(pBirth) = "" Then
        mGetAge = ""
        Exit Function
    End If
    
    strPtYear = Mid$(pBirth, 1, 2)
    strPtMonth = Mid$(pBirth, 3, 2)
    strPtDay = Mid$(pBirth, 5, 2)
    
    lngThisYear = CLng(Format$(Now, "YY"))
    If (CLng(strPtYear) > lngThisYear) Then lngThisYear = lngThisYear + 100
    
    lngThisMonth = CLng(Format(Now, "MM"))
    If (CLng(strPtMonth) > lngThisMonth) Then lngThisMonth = lngThisMonth + 12
    
    strYear = CStr(lngThisYear - CLng(strPtYear))
    strMonth = CStr(lngThisMonth - CLng(strPtMonth))
    
    '## 1.0.2: (2004-12-14)
    '   - 생년월일에 "700000"등으로 전달할때 DateDiff함수에서 런타임에러 발생을 방지하기 위해
    '     강제로 월,일을 변경!
    '## 1.1.5: (2005-03-25)
    '   - 생년월일중 월,일에 해당값이 각각 13,32 이상인 경우 강제로 월,일을 "01"로 변경
    If strPtMonth = "00" Or CLng(strPtMonth) > 12 Then strPtMonth = "01"
    If strPtDay = "00" Or CLng(strPtDay) > 31 Then strPtDay = "01"
    strDay = DateDiff("d", strPtYear & "-" & strPtMonth & "-" & strPtDay, Now)
    
    If Trim$(strYear) <> "0" Then
        mGetAge = strYear & "Y"
    ElseIf Trim$(strMonth) > "1" Then
        mGetAge = strMonth & "M"
    Else
        mGetAge = strDay & "D"
    End If
End Function
