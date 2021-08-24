VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISEqpTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISEqpTrans.cls (자사 LIS와 조인할때)

'   내  용  : 장비전송내역 클래스
'   작성일  : 2004-04-16
'   버  전  :
'-----------------------------------------------------------------------------'

'## 속성
'   - 장비코드
'   - 전송일자
'   - 전송시간
'   - 검체연도
'   - 검체번호
'   - 장비에서 검체위치

'## 메소드
'   - 전송내역 입력
'   - 전송내역 삭제 (장비코드, 검체연도+번호가 일치하는것만)

Option Explicit

Private mSQL        As clsIISLibrarySql

Private mEqpCd      As String   '장비코드
Private mTransDt    As String   '전송일자
Private mTransTm    As String   '전송시간
Private mSpcYy      As String   '검체연도
Private mSpcNo      As String   '검체번호
Private mTransNo    As String   '장비에서의 검체위치

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get TransDt() As String
    TransDt = mTransDt
End Property

Public Property Let TransDt(ByVal vData As String)
    mTransDt = vData
End Property

Public Property Get TransTm() As String
    TransTm = mTransTm
End Property

Public Property Let TransTm(ByVal vData As String)
    mTransTm = vData
End Property

Public Property Get SpcYy() As String
    SpcYy = mSpcYy
End Property

Public Property Let SpcYy(ByVal vData As String)
    mSpcYy = vData
End Property

Public Property Get SpcNo() As String
    SpcNo = mSpcNo
End Property

Public Property Let SpcNo(ByVal vData As String)
    mSpcNo = vData
End Property

Public Property Get TransNo() As String
    TransNo = mTransNo
End Property

Public Property Let TransNo(ByVal vData As String)
    mTransNo = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 장비 전송내역 입력
'   인수 :
'       - pEqpCd    : 장비코드
'       - pTransDt  : 장비에서 결과 전송일자
'       - pTransTm  : 장비에서 결과 전송시간
'       - pSpcYy    : 검체연도
'       - pSpcNo    : 검체번호
'       - pTransNo  : 장비에서 검체위치
'   반환 : 장비 전송내역 입력쿼리
'-----------------------------------------------------------------------------'
Public Function GetAddEqpTransSql(ByVal pEqpCd As String, ByVal pTransDt As String, _
                    ByVal pTransTm As String, ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pTransNo As String) As String
    Dim SQL As String
    
    SQL = mSQL.InsertEqpTrans(pEqpCd, pTransDt, pTransTm, pSpcYy, pSpcNo, pTransNo)
    
    GetAddEqpTransSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비 전송내역 삭제
'   인수 :
'       - pEqpCd    : 장비코드
'       - pSpcYy    : 검체연도
'       - pSpcNo    : 검체번호
'   반환 : 장비 전송내역 삭제쿼리
'-----------------------------------------------------------------------------'
Public Function GetDelEqpTransSql(ByVal pEqpCd As String, ByVal pSpcYy As String, _
                    ByVal pSpcNo As String) As String
    Dim SQL As String
    
    SQL = mSQL.DeleteEqpTrans(pEqpCd, pSpcYy, pSpcNo)
                
    GetDelEqpTransSql = SQL
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISLibrarySql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub
