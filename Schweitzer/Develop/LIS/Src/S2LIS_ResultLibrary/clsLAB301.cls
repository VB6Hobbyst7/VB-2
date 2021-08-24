VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LAB301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Departments"
Attribute VB_Ext_KEY = "Member0" ,"Departments"
Option Explicit

Private mCol As Collection

Public Function Add(WorkDt As String, WorkCd As String, _
                    WorkSeq As String, WorkTm As String, WorkId As String, _
                    WorkArea As String, AccDt As String, AccSeq As String, StatFg As String, _
                    RecordCount As Long, LAB301 As LAB301, Optional sKey As String, _
                    Optional RcvDt As String, Optional RcvTm As String) As clsWSBuild
'
Dim objNewMember As clsWSBuild
   Set objNewMember = New clsWSBuild
    '
   With objNewMember
      .WorkDt = WorkDt
      .WorkCd = WorkCd
      .WorkSeq = WorkSeq
      .WorkTm = WorkTm
      .WorkId = WorkId
      .WorkArea = WorkArea
      .AccDt = AccDt
      .AccSeq = AccSeq
      .StatFg = StatFg
      .RcvDt = RcvDt
      .RcvTm = RcvTm
      .RecordCount = RecordCount
      .Key = sKey
   End With
   '
   Set objNewMember.LAB301 = LAB301
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
   '
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As clsWSBuild
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
