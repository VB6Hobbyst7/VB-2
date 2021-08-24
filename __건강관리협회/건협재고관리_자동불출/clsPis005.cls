VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPis005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvaropercd As String '지역 복사
Private mvaropernm As String '지역 복사
Public Property Let opernm(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opernm = 5
    mvaropernm = vData
End Property


Public Property Get opernm() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opernm
    opernm = mvaropernm
End Property



Public Property Let opercd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercd = 5
    mvaropercd = vData
End Property


Public Property Get opercd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercd
    opercd = mvaropercd
End Property

Public Function cfSave() As Boolean

    On Error GoTo errSave
    gSql = "SELECT * FROM S2PIS005 WHERE OPERCD='" & opercd & "'"
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("opercd").Value = opercd
            End If
            .Fields("opernm").Value = opernm
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
    gSql = "SELECT COUNT(*) AS CNT FROM S2PIS102 WHERE OPERCD='" & opercd & "' GROUP BY OPERCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Not (.Fields("CNT").Value > 0)
            End If
            .Close
        End If
    End With

    If sReturn Then
        gSql = "SELECT COUNT(*) AS CNT FROM S2PIS302 WHERE OPERCD='" & opercd & "' GROUP BY OPERCD"
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
        gSql = "SELECT COUNT(*) AS CNT FROM S2PIS103 WHERE OPERCD='" & opercd & "' GROUP BY OPERCD"
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
        gSql = "DELETE FROM S2PIS005 WHERE OPERCD='" & opercd & "'"
        cfDelete = cDb.cfExecute(gSql)
    Else
        cfDelete = True
    End If
    
End Function

Public Function cfList(Optional ByVal brSort As Boolean = True) As ADODB.Recordset

    gSql = "SELECT * FROM S2PIS005 ORDER BY " & IIf(brSort, "OPERCD", "OPERNM")
    Set cfList = cDb.cfRecordSet(gSql)
        
End Function

