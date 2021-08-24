VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcolItems As Collection
Private mvarCount As Integer '지역 복사

Public Property Get Count() As Integer
    Count = mcolItems.Count
End Property

Public Property Get Item(ByVal KeyIndex As Variant) As clsItem
    Set Item = mcolItems(KeyIndex)
End Property

Private Sub Class_Initialize()
   Set mcolItems = New Collection
   mvarCount = 0
End Sub

Public Function Add(ByVal clsKey As Variant, ByVal clsData As clsItem)
   
   Dim newItem As New clsItem
   
   With newItem
      .TestCd = clsData.TestCd
      .ApplyDt = clsData.ApplyDt
      .TestNm = clsData.TestNm
      .AbbrNm5 = clsData.AbbrNm5
      .AbbrNm10 = clsData.AbbrNm10
      .WorkArea = clsData.WorkArea
      .RstType = clsData.RstType
      .TestDiv = clsData.TestDiv
      .RptSeq = clsData.RptSeq
      .PanelFg = clsData.PanelFg
      .TxtType = clsData.TxtType
      .RstDiv = clsData.RstDiv
      .OutLabCd = clsData.OutLabCd
      .GrpFg = clsData.GrpFg
      .WorkUnit = clsData.WorkUnit
      .AttrCd = clsData.AttrCd
      .ExpDt = clsData.ExpDt
      .MatchCd = clsData.MatchCd
      .DetailFg = clsData.DetailFg
      .ItemSeq = clsData.ItemSeq
      .GroupCd = clsData.GroupCd
   End With
   
   mcolItems.Add newItem, clsKey
   Set Add = newItem
   mvarCount = mvarCount + 1
   
End Function

Public Function Update(ByVal clsKey As Variant, ByVal clsData As clsItem)
   
   With mcolItems.Item(clsKey)
      .TestCd = clsData.TestCd
      .ApplyDt = clsData.ApplyDt
      .TestNm = clsData.TestNm
      .AbbrNm5 = clsData.AbbrNm5
      .AbbrNm10 = clsData.AbbrNm10
      .WorkArea = clsData.WorkArea
      .RstType = clsData.RstType
      .TestDiv = clsData.TestDiv
      .RptSeq = clsData.RptSeq
      .PanelFg = clsData.PanelFg
      .TxtType = clsData.TxtType
      .RstDiv = clsData.RstDiv
      .OutLabCd = clsData.OutLabCd
      .GrpFg = clsData.GrpFg
      .WorkUnit = clsData.WorkUnit
      .AttrCd = clsData.AttrCd
      .ExpDt = clsData.ExpDt
      .MatchCd = clsData.MatchCd
      .DetailFg = clsData.DetailFg
      .ItemSeq = clsData.ItemSeq
      .GroupCd = clsData.GroupCd
   End With

End Function

Public Function Remove(ByVal clsKey As Variant)
    mcolItems.Remove clsKey
End Function

Public Sub Clear()
   Set mcolItems = Nothing
   Set mcolItems = New Collection
   mvarCount = 0
End Sub

