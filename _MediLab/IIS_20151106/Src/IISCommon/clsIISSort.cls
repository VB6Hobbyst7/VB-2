VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISSort.cls
'   작성자  : 오세원
'   내  용  : 컬렉션 클래스를 정렬시키는 기능을 갖는 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'   #NOTE   :
'       1.현재 정렬 알고리즘은 키가 있는 컬렉션 클래스는 정렬이 불가능하여 일단은
'         키가 없는 상태로 정렬하고(AddEx), 다른 컬렉션 클래스를 생성하여 정렬된
'         항목을 키를 이용해 입력한다.(SortAdd)
'       2.알고리즘을 변형하여 키가 있는 컬렉션도 정렬할수 있도록 해야한다.
'       3.각 업무에 맞는 알고리즘을 만들어 필요에 따라 사용할수 있도록 해야한다.
'-----------------------------------------------------------------------------'

Option Explicit

'-----------------------------------------------------------------------------'
'   기능 : 전달된 컬렉션 클래스를 키를 이용해 오름차순 정렬(Shell Sort Algorithm)
'   인수 :
'       1.pCol:컬렉션 클래스
'   필요사항 :
'       1.컬렉션안의 클래스는 Compare 메소드를 구현해야 한다.
'       2.컬렉션 클래스는 AddEx, SortAdd 메소드를 구현해야 한다.
'-----------------------------------------------------------------------------'
Public Sub Sort(ByRef pCol As Object)
    Dim objTemp1    As Object
    Dim objTemp2    As Object
    Dim bResult     As Boolean
    Dim nGap        As Long
    Dim i           As Long
    Dim j           As Long

    nGap = pCol.Count / 2
    Do While nGap > 0
        For i = nGap To pCol.Count - 1
            j = i
            Set objTemp1 = pCol(i + 1)
            bResult = (objTemp1.Compare(pCol(j - nGap + 1)) = -1)

            Do While j >= nGap And bResult
                Set objTemp2 = pCol(j - nGap + 1)
                pCol.Remove j + 1
                If j + 1 > pCol.Count Then
                    Call pCol.AddEx(objTemp2)
                Else
                    Call pCol.AddEx(objTemp2, j + 1)
                End If

                j = j - nGap
                If j >= nGap Then
                    bResult = (objTemp1.Compare(pCol(j - nGap + 1)) = -1)
                End If
            Loop

            pCol.Remove j + 1
            If j + 1 > pCol.Count Then
                Call pCol.AddEx(objTemp1)
            Else
                Call pCol.AddEx(objTemp1, j + 1)
            End If
        Next i

        nGap = nGap / 2
    Loop

    Call pCol.SortAdd(pCol)

    Set objTemp1 = Nothing
    Set objTemp2 = Nothing
End Sub
