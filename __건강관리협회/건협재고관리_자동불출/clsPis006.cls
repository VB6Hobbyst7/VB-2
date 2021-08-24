VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarreasoncd As String '지역 복사
Private mvarreasonnm As String '지역 복사
Private mvarkindfg As String '지역 복사

Public Property Let kindfg(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.kindfg = 5
    mvarkindfg = vData
End Property


Public Property Get kindfg() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.kindfg
    kindfg = mvarkindfg
End Property



Public Property Let reasonnm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reasonnm = 5
    mvarreasonnm = vData
End Property


Public Property Get reasonnm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reasonnm
    reasonnm = mvarreasonnm
End Property



Public Property Let reasoncd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reasoncd = 5
    mvarreasoncd = vData
End Property


Public Property Get reasoncd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reasoncd
    reasoncd = mvarreasoncd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS006 WHERE REASONCD='" & reasoncd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("reasoncd").Value = reasoncd
            End If
            .Fields("reasonnm").Value = reasonnm
            .Fields("kindfg").Value = kindfg
            .Update
            .Close
        End If
    End With
    cfSave = True
    Exit Function
    
errSave:
    cfSave = False
    MsgBox Err.Description, vbCritical
    
End Function

Public Function cfDelete() As Boolean
Dim sReturn As Boolean

    sReturn = True
    gSql = "SELECT COUNT(*) AS CNT FROM S2PIS202 WHERE REASONCD='" & reasoncd & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Not (.Fields("CNT").Value > 0)
            End If
            .Close
        End If
    End With

    If sReturn Then
        gSql = "SELECT COUNT(*) AS CNT FROM S2PIS302 WHERE REASONCD='" & reasoncd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    sReturn = Not (.Fields("CNT").Value > 0)
                End If
                .Close
            End If
        End With
    End If

    If sReturn Then
        gSql = "SELECT COUNT(*) AS CNT FROM S2PIS301 WHERE REASONCD='" & reasoncd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    sReturn = Not (.Fields("CNT").Value > 0)
                End If
                .Close
            End If
        End With
    End If

    If sReturn Then
        gSql = "SELECT COUNT(*) AS CNT FROM S2PIS303 WHERE REASONCD='" & reasoncd & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    sReturn = Not (.Fields("CNT").Value > 0)
                End If
                .Close
            End If
        End With
    End If
    
    If sReturn Then
        gSql = "DELETE FROM S2PIS006 WHERE REASONCD='" & reasoncd & "'"
        cfDelete = cDb.cfExecute(gSql)
    Else
        cfDelete = True
    End If
    
End Function

Public Function cfList(Optional ByVal brSort As Boolean = True, Optional ByVal brKind As String) As ADODB.Recordset

    gSql = "SELECT * FROM S2PIS006 "
    If Len(brKind) > 0 Then
        gSql = gSql & " WHERE KINDFG IN ('" & brKind & "','0')"
    End If
    gSql = gSql & " ORDER BY " & IIf(brSort, "REASONCD", "REASONNM")
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function
