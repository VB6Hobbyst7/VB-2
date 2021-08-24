VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarstkcd As String '지역 복사
Private mvarrmdfg As String '지역 복사
Private mvarrmdyear As String '지역 복사
Private mvarprevqty As Double '지역 복사
Public Property Let prevqty(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.prevqty = 5
    mvarprevqty = vData
End Property


Public Property Get prevqty() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.prevqty
    prevqty = mvarprevqty
End Property



Public Property Let rmdyear(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rmdyear = 5
    mvarrmdyear = vData
End Property


Public Property Get rmdyear() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rmdyear
    rmdyear = mvarrmdyear
End Property



Public Property Let rmdfg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.rmdfg = 5
    mvarrmdfg = vData
End Property


Public Property Get rmdfg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.rmdfg
    rmdfg = mvarrmdfg
End Property



Public Property Let stkcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.stkcd = 5
    mvarstkcd = vData
End Property


Public Property Get stkcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.stkcd
    stkcd = mvarstkcd
End Property

Public Function cfStkPrevQty(ByVal brStk As String, ByVal brDt As String) As Double
Dim sPrevQty As Double, sYear As String

    sYear = Mid(brDt, 1, 4)
    ' 년도이월재고
    gSql = "SELECT PREVQTY FROM S2PIS409                                            " & vbNewLine & _
           " WHERE STKCD='" & brStk & "' AND RMDYEAR='" & sYear & "' AND RMDFG='0'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sPrevQty = Val("" & .Fields("PREVQTY").Value)
            End If
            .Close
        End If
    End With
    
    ' 이전 입고량
    gSql = "SELECT SUM(IQTY_SO) AS PREVQTY FROM S2PIS201 WHERE STKCD='" & brStk & "'" & vbNewLine & _
           "  AND SUBSTR(ENTDT,1,4)='" & sYear & "' AND ENTDT<'" & brDt & "' GROUP BY STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sPrevQty = sPrevQty + Val("" & .Fields("PREVQTY").Value)
            End If
            .Close
        End If
    End With

    ' 이전 출고량
    gSql = "SELECT SUM(NVL(TESTQTY,0)+NVL(FREEQTY,0)+NVL(QCQTY,0)+NVL(RETESTQTY,0)+NVL(MANUQTY,0)+NVL(MACHQTY,0)+NVL(HANDQTY,0)) AS OUTQTY" & vbNewLine & _
           "  FROM S2PIS313                                                                                 " & vbNewLine & _
           " WHERE STKCD='" & brStk & "' AND SUBSTR(WORKDT,1,4)='" & sYear & "' AND WORKDT<'" & brDt & "'   " & vbNewLine & _
           " GROUP BY STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sPrevQty = sPrevQty - Val("" & .Fields("OUTQTY").Value)
            End If
            .Close
        End If
    End With
    
    cfStkPrevQty = sPrevQty

End Function
