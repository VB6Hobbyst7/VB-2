VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMstMach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarmachcd As String '지역 복사
Private mvarmachnm As String '지역 복사
Private mvardutycd As String '지역 복사
Private mvardelfg As Integer '지역 복사
Public Property Let delfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.delfg = 5
    mvardelfg = vData
End Property


Public Property Get delfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.delfg
    delfg = mvardelfg
End Property



Public Property Let dutycd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.dutycd = 5
    mvardutycd = vData
End Property


Public Property Get dutycd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.dutycd
    dutycd = mvardutycd
End Property



Public Property Let machnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.machnm = 5
    mvarmachnm = vData
End Property


Public Property Get machnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.machnm
    machnm = mvarmachnm
End Property



Public Property Let machcd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.machcd = 5
    mvarmachcd = vData
End Property


Public Property Get machcd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.machcd
    machcd = mvarmachcd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "select * from mstMACH where machcd = '" & machcd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("machcd").Value = machcd
            End If
            .Fields("machnm").Value = machnm
            .Fields("dutycd").Value = dutycd
            .Fields("delfg").Value = delfg
            .Update
            .Close
        End If
    End With
    cfSave = True
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical

End Function

Public Function cfDelete(ByVal brCd As String) As Boolean
Dim sReturn As Boolean

    sReturn = True
    ' 운영내역자료 확인
    gSql = "select machcd from operL where machcd = '" & brCd & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            sReturn = .EOF
            .Close
        End If
    End With
    
    If sReturn Then
        ' 운영내역자료 확인
        gSql = "select machcd from machSTK where machcd = '" & brCd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                sReturn = .EOF
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        gSql = "delete from mstMACH where machcd = '" & brCd & "'"
    Else
        gSql = "update mstMACH set delfg = 1 where machcd = '" & brCd & "'"
    End If
    cfDelete = cDb.cfExecute(gSql)
    
End Function

Public Function cfList(Optional ByVal brDel As Byte = gAllData) As ADODB.Recordset

    gSql = "select a.*, b.dutynm from mstMACH a, mstDUTY b where a.dutycd *= b.dutycd"
    If brDel <> gAllData Then
        gSql = gSql & " and delfg = " & brDel
    End If
    gSql = gSql & " order by machcd"
    Set cfList = cDb.cfRecordSet(gSql)
    
End Function
