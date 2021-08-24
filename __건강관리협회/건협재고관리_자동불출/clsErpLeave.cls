VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErpLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarCD_COMPANY As String '지역 복사
Private mvarCD_PLANT As String '지역 복사
Private mvarCD_BIZAREA As String '지역 복사
Private mvarNO_IO As String '지역 복사
Private mvarNO_IOLINE As Long '지역 복사
Private mvarCD_SL As String '지역 복사
Private mvarDT_IO As String '지역 복사
Private mvarFG_IO As String '지역 복사
Private mvarCD_QTIOTP As String '지역 복사
Private mvarCD_ITEM As String '지역 복사
Private mvarQT_IO As Double '지역 복사
Private mvarNO_LOT As String '지역 복사
Private mvarDT_LIMIT As String '지역 복사
Private mvarNO_EMP As String '지역 복사
Private mvarYN_ERP As String '지역 복사
Private mvarID_ERP As String '지역 복사
Private mvarDTS_ERP As String '지역 복사
Private mvarNO_IOLINE2 As Long '지역 복사
Private mvarQT_IO_SUM As Double '지역 복사
Private mvarNO_IO_MGMT As String '지역 복사
Private mvarNO_IOLINE_MGMT As Long '지역 복사
Private mvarNO_IOLINE_MGMT2 As Long '지역 복사
Public Property Let NO_IOLINE_MGMT2(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IOLINE_MGMT2 = 5
    mvarNO_IOLINE_MGMT2 = vData
End Property


Public Property Get NO_IOLINE_MGMT2() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IOLINE_MGMT2
    NO_IOLINE_MGMT2 = mvarNO_IOLINE_MGMT2
End Property



Public Property Let NO_IOLINE_MGMT(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IOLINE_MGMT = 5
    mvarNO_IOLINE_MGMT = vData
End Property


Public Property Get NO_IOLINE_MGMT() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IOLINE_MGMT
    NO_IOLINE_MGMT = mvarNO_IOLINE_MGMT
End Property



Public Property Let NO_IO_MGMT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IO_MGMT = 5
    mvarNO_IO_MGMT = vData
End Property


Public Property Get NO_IO_MGMT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IO_MGMT
    NO_IO_MGMT = mvarNO_IO_MGMT
End Property



Public Property Let QT_IO_SUM(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.QT_IO_SUM = 5
    mvarQT_IO_SUM = vData
End Property


Public Property Get QT_IO_SUM() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.QT_IO_SUM
    QT_IO_SUM = mvarQT_IO_SUM
End Property



Public Property Let NO_IOLINE2(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IOLINE2 = 5
    mvarNO_IOLINE2 = vData
End Property


Public Property Get NO_IOLINE2() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IOLINE2
    NO_IOLINE2 = mvarNO_IOLINE2
End Property




Public Property Let DTS_ERP(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.DTS_ERP = 5
    mvarDTS_ERP = vData
End Property


Public Property Get DTS_ERP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.DTS_ERP
    DTS_ERP = mvarDTS_ERP
End Property



Public Property Let ID_ERP(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.ID_ERP = 5
    mvarID_ERP = vData
End Property


Public Property Get ID_ERP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.ID_ERP
    ID_ERP = mvarID_ERP
End Property



Public Property Let YN_ERP(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.YN_ERP = 5
    mvarYN_ERP = vData
End Property


Public Property Get YN_ERP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.YN_ERP
    YN_ERP = mvarYN_ERP
End Property



Public Property Let NO_EMP(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_EMP = 5
    mvarNO_EMP = vData
End Property


Public Property Get NO_EMP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_EMP
    NO_EMP = mvarNO_EMP
End Property



Public Property Let DT_LIMIT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.DT_LIMIT = 5
    mvarDT_LIMIT = vData
End Property


Public Property Get DT_LIMIT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.DT_LIMIT
    DT_LIMIT = mvarDT_LIMIT
End Property



Public Property Let NO_LOT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_LOT = 5
    mvarNO_LOT = vData
End Property


Public Property Get NO_LOT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_LOT
    NO_LOT = mvarNO_LOT
End Property



Public Property Let QT_IO(ByVal vData As Double)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.QT_IO = 5
    mvarQT_IO = vData
End Property


Public Property Get QT_IO() As Double
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.QT_IO
    QT_IO = mvarQT_IO
End Property



Public Property Let CD_ITEM(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_ITEM = 5
    mvarCD_ITEM = vData
End Property


Public Property Get CD_ITEM() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_ITEM
    CD_ITEM = mvarCD_ITEM
End Property



Public Property Let CD_QTIOTP(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_QTIOTP = 5
    mvarCD_QTIOTP = vData
End Property


Public Property Get CD_QTIOTP() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_QTIOTP
    CD_QTIOTP = mvarCD_QTIOTP
End Property



Public Property Let FG_IO(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.FG_IO = 5
    mvarFG_IO = vData
End Property


Public Property Get FG_IO() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.FG_IO
    FG_IO = mvarFG_IO
End Property



Public Property Let DT_IO(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.DT_IO = 5
    mvarDT_IO = vData
End Property


Public Property Get DT_IO() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.DT_IO
    DT_IO = mvarDT_IO
End Property



Public Property Let CD_SL(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_SL = 5
    mvarCD_SL = vData
End Property


Public Property Get CD_SL() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_SL
    CD_SL = mvarCD_SL
End Property



Public Property Let NO_IOLINE(ByVal vData As Long)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IOLINE = 5
    mvarNO_IOLINE = vData
End Property


Public Property Get NO_IOLINE() As Long
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IOLINE
    NO_IOLINE = mvarNO_IOLINE
End Property



Public Property Let NO_IO(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.NO_IO = 5
    mvarNO_IO = vData
End Property


Public Property Get NO_IO() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.NO_IO
    NO_IO = mvarNO_IO
End Property



Public Property Let CD_BIZAREA(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_BIZAREA = 5
    mvarCD_BIZAREA = vData
End Property


Public Property Get CD_BIZAREA() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_BIZAREA
    CD_BIZAREA = mvarCD_BIZAREA
End Property



Public Property Let CD_PLANT(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_PLANT = 5
    mvarCD_PLANT = vData
End Property


Public Property Get CD_PLANT() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_PLANT
    CD_PLANT = mvarCD_PLANT
End Property



Public Property Let CD_COMPANY(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.CD_COMPANY = 5
    mvarCD_COMPANY = vData
End Property


Public Property Get CD_COMPANY() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.CD_COMPANY
    CD_COMPANY = mvarCD_COMPANY
End Property

Private Function cfLastNo() As Long
Dim sReturn As Long

    sReturn = 1
    gSql = "SELECT MAX(NO_IOLINE) AS LASTSEQ FROM " & gTBLleave & " WHERE CD_COMPANY='" & CD_COMPANY & "'" & vbNewLine & _
           "   AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "'" & vbNewLine & _
           "   AND NO_IO='" & NO_IO & "' GROUP BY CD_COMPANY,CD_PLANT,CD_BIZAREA,NO_IO"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("LASTSEQ").Value) + 1
            End If
            .Close
        End If
    End With
    
    cfLastNo = sReturn

End Function

Private Function cfLastNo2() As Long
Dim sReturn As Long

    sReturn = 1
    gSql = "SELECT MAX(NO_IOLINE2) AS LASTSEQ FROM " & gTBLleave & " WHERE CD_COMPANY='" & CD_COMPANY & "'" & vbNewLine & _
           "   AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "'" & vbNewLine & _
           "   AND NO_IO='" & NO_IO & "' AND NO_IOLINE = " & NO_IOLINE & " GROUP BY CD_COMPANY,CD_PLANT,CD_BIZAREA,NO_IO,NO_IOLINE"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("LASTSEQ").Value) + 1
            End If
            .Close
        End If
    End With
    
    cfLastNo2 = sReturn

End Function

Public Function cfSave() As Boolean

    On Error GoTo errSave
    If NO_IOLINE = 0 Then NO_IOLINE = cfLastNo
    
    If NO_IOLINE2 = 0 Then NO_IOLINE2 = cfLastNo2
    
    gSql = "SELECT * FROM " & gTBLleave & " WHERE CD_COMPANY='" & CD_COMPANY & "'" & vbNewLine & _
           "   AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "'" & vbNewLine & _
           "   AND NO_IO='" & NO_IO & "' AND NO_IOLINE=" & NO_IOLINE & " AND NO_IOLINE2=" & NO_IOLINE2
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If .EOF Then
                gSql = "INSERT INTO " & gTBLleave & " ("
                gSql = gSql & "   CD_COMPANY "
                gSql = gSql & " , CD_PLANT "
                gSql = gSql & " , CD_BIZAREA "
                gSql = gSql & " , NO_IO "
                gSql = gSql & " , NO_IOLINE "
                gSql = gSql & " , NO_IOLINE2 "
                gSql = gSql & " , CD_SL "
                gSql = gSql & " , DT_IO "
                gSql = gSql & " , FG_IO "
                gSql = gSql & " , CD_QTIOTP "
                gSql = gSql & " , CD_ITEM "
                gSql = gSql & " , NO_LOT "
                gSql = gSql & " , DT_LIMIT "
                gSql = gSql & " , NO_EMP "
                gSql = gSql & " , YN_ERP "
                gSql = gSql & " , QT_IO "
                gSql = gSql & " , QT_IO_SUM "
                gSql = gSql & " , DTS_INSERT "
                gSql = gSql & " , NO_IO_MGMT "
                gSql = gSql & " , NO_IOLINE_MGMT "
                gSql = gSql & " , NO_IOLINE_MGMT2 " & vbNewLine
                gSql = gSql & ") VALUES ( " & vbNewLine
                gSql = gSql & "   '" & CD_COMPANY & "'"
                gSql = gSql & " , '" & CD_PLANT & "'"
                gSql = gSql & " , '" & CD_BIZAREA & "'"
                gSql = gSql & " , '" & NO_IO & "'"
                gSql = gSql & " ,  " & NO_IOLINE & " "
                gSql = gSql & " ,  0"                       ' 사용하지 않는 필드로 무조건 0으로 처리
                'gSql = gSql & " ,  " & NO_IOLINE2 & " "
                gSql = gSql & " , '" & CD_SL & "'"
                gSql = gSql & " , '" & DT_IO & "'"
                gSql = gSql & " , '" & "012" & "'"
                gSql = gSql & " , '" & "400" & "'"
                gSql = gSql & " , '" & CD_ITEM & "'"
                gSql = gSql & " , '" & NO_LOT & "'"
                gSql = gSql & " , '" & DT_LIMIT & "'"
                gSql = gSql & " , '" & NO_EMP & "'"
                gSql = gSql & " , '" & "N" & "'"
                gSql = gSql & " ,  " & QT_IO & " "
                gSql = gSql & " ,  " & QT_IO_SUM & " "
                gSql = gSql & " ,  TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS') "
                gSql = gSql & " , '" & NO_IO_MGMT & "'"
                gSql = gSql & " ,  " & NO_IOLINE_MGMT & " "
                gSql = gSql & " ,  " & NO_IOLINE_MGMT2 & " "
                gSql = gSql & ")"
            Else
                If ("" & .Fields("YN_ERP").Value) <> "Y" Then
                    gSql = "UPDATE " & gTBLleave & " SET QT_IO=" & QT_IO
                    gSql = gSql & "   , CD_ITEM         = '" & CD_ITEM & "'" & vbNewLine
                    gSql = gSql & "   , NO_LOT          = '" & NO_LOT & "'" & vbNewLine
                    gSql = gSql & "   , DT_LIMIT        = '" & DT_LIMIT & "'" & vbNewLine
                    gSql = gSql & "   , NO_EMP          = '" & NO_EMP & "'" & vbNewLine
                    gSql = gSql & "   , QT_IO_SUM       =  " & QT_IO_SUM & " " & vbNewLine
                    gSql = gSql & "   , DTS_INSERT      = TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS')" & vbNewLine
                    gSql = gSql & "   , NO_IO_MGM       = '" & NO_IO_MGMT & "'" & vbNewLine
                    gSql = gSql & "   , NO_IOLINE_MGMT  = '" & NO_IOLINE_MGMT & "'" & vbNewLine
                    gSql = gSql & "   , NO_IOLINE_MGMT2 = '" & NO_IOLINE_MGMT2 & "'" & vbNewLine
                    gSql = gSql & " WHERE CD_COMPANY='" & CD_COMPANY & "' AND CD_PLANT='" & CD_PLANT & "' AND CD_BIZAREA='" & CD_BIZAREA & "' AND NO_IO='" & NO_IO & "'"
                    gSql = gSql & "   AND NO_IOLINE=" & NO_IOLINE & " AND NO_IOLINE2=" & NO_IOLINE2
                End If
            End If
            .Close
            
            cfSave = cDb.cfExecute(gSql)
        End If
    End With
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False
    
End Function

Public Function cfDeleteAll() As Boolean

    gSql = "DELETE FROM " & gTBLleave & " WHERE CD_COMPANY='" & gCompany & "' AND CD_PLANT='" & gAreaCd & "' AND DT_IO='" & DT_IO & "'" & vbNewLine & _
           "   AND (YN_ERP IS NULL OR YN_ERP<>'Y')"
    cfDeleteAll = cDb.cfExecute(gSql)
    
End Function
