VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'속성값을 가질 지역 변수
Private mvarmachcd As String '지역 복사
Private mvaroperdt As String '지역 복사
Private mvaroperseq As Integer '지역 복사
Private mvaropercd As Integer '지역 복사
Private mvarendfg As Integer '지역 복사
Private mvarreason As String '지역 복사
Private mvarusercd As String '지역 복사
Private mvarwrtdt As String '지역 복사
Private mvarmoddt As String '지역 복사
Private mvaropercnt As Integer '지역 복사
Public Property Let opercnt(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercnt = 5
    mvaropercnt = vData
End Property


Public Property Get opercnt() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercnt
    opercnt = mvaropercnt
End Property




Public Property Let moddt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.moddt = 5
    mvarmoddt = vData
End Property


Public Property Get moddt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.moddt
    moddt = mvarmoddt
End Property



Public Property Let wrtdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.wrtdt = 5
    mvarwrtdt = vData
End Property


Public Property Get wrtdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.wrtdt
    wrtdt = mvarwrtdt
End Property



Public Property Let usercd(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.usercd = 5
    mvarusercd = vData
End Property


Public Property Get usercd() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.usercd
    usercd = mvarusercd
End Property



Public Property Let reason(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.reason = 5
    mvarreason = vData
End Property


Public Property Get reason() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.reason
    reason = mvarreason
End Property



Public Property Let endfg(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.endfg = 5
    mvarendfg = vData
End Property


Public Property Get endfg() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.endfg
    endfg = mvarendfg
End Property



Public Property Let opercd(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.opercd = 5
    mvaropercd = vData
End Property


Public Property Get opercd() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.opercd
    opercd = mvaropercd
End Property



Public Property Let operseq(ByVal vData As Integer)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.operseq = 5
    mvaroperseq = vData
End Property


Public Property Get operseq() As Integer
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.operseq
    operseq = mvaroperseq
End Property



Public Property Let operdt(ByVal vData As String)
'대입문의 왼쪽에서 속성에 값을 할당할 때 사용됩니다.
'Syntax: X.operdt = 5
    mvaroperdt = vData
End Property


Public Property Get operdt() As String
'대입문의 오른쪽에서 속성값을 검색할 때 사용됩니다.
'Syntax: Debug.Print X.operdt
    operdt = mvaroperdt
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

Private Function cfLastSeq() As Integer
Dim sReturn As Integer

    gSql = "select max(operseq) as lastseq from operL with (nolock) where machcd = '" & machcd & "'" & _
           "   and operdt = '" & operdt & "' group by machcd, operdt"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("lastseq").Value) + 1
            Else
                sReturn = 1
            End If
            .Close
        End If
    End With
    
    cfLastSeq = sReturn
            
End Function

Public Function cfSave() As Boolean
Dim sDate As String, sReturn As Boolean

    On Error GoTo errSave
    If operseq = 0 Then operseq = cfLastSeq
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd Hh:Nn:Ss")
    sReturn = True
    
    gSql = "select * from operL where machcd = '" & machcd & "'" & " and operdt = '" & operdt & "' and operseq = " & operseq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If .EOF Then
                .AddNew
                .Fields("machcd").Value = machcd
                .Fields("operdt").Value = operdt
                .Fields("operseq").Value = operseq
                .Fields("wrtdt").Value = sDate
                .Fields("endfg").Value = 0
            Else
                sReturn = cfStkRmdProcess(.Fields("machcd").Value, .Fields("operdt").Value, .Fields("opercd").Value, .Fields("opercnt").Value, False)
            End If
            .Fields("opercd").Value = opercd
            .Fields("opercnt").Value = opercnt
            .Fields("reason").Value = reason
            .Fields("usercd").Value = usercd
            .Fields("moddt").Value = sDate
            .Update
            .Close
            
            sReturn = cfStkRmdProcess(machcd, operdt, opercd, opercnt, True)
        End If
    End With
    cfSave = sReturn
    Exit Function
    
errSave:
    MsgBox Err.Description, vbCritical
    cfSave = False
    
End Function

Public Function cfDelete(ByVal brMach As String, ByVal brDt As String, ByVal brSeq As Integer) As Boolean
Dim sReturn As Boolean

    gSql = "select * from operL with (nolock) where machcd = '" & brMach & "'" & " and operdt = '" & brDt & "' and operseq = " & brSeq
    With cDb.cfRecordSet(gSql, True)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = cfStkRmdProcess(.Fields("machcd").Value, .Fields("operdt").Value, .Fields("opercd").Value, .Fields("opercnt").Value, False)
            End If
            .Close
        End If
    End With

    If sReturn Then
        gSql = "delete from operL where machcd = '" & brMach & "'" & " and operdt = '" & brDt & "' and operseq = " & brSeq
        sReturn = cDb.cfExecute(gSql)
    End If
    
    cfDelete = sReturn
    
End Function

Private Function cfStkRmdProcess(ByVal brMach As String, ByVal brDt As String, ByVal brOper As Integer, ByVal brCnt As Integer, ByVal brJob As Boolean) As Boolean
Dim cRmd As clsStkRmd
Dim sReturn As Boolean, sUseQty As Single, sStkCd As Long

    sReturn = True
    Set cRmd = New clsStkRmd
    
    gSql = "select stkcd, qty from machSTK where machcd = '" & brMach & "' and opercd = " & brOper
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                sStkCd = Val("" & .Fields("stkcd").Value)
                sUseQty = (Val("" & .Fields("qty").Value) * brCnt) * IIf(brJob, 1, -1)
                
                sReturn = cRmd.cfRmdProcess(brDt, sStkCd, sUseQty, False)
                If sReturn Then
                    .MoveNext
                Else
                    .MoveLast
                End If
            Wend
            .Close
        End If
    End With

    cfStkRmdProcess = sReturn

End Function
