VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpecimens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolSpecimens As Collection
Private mvarCount As Integer '지역 복사

Public Property Get Count() As Integer
    Count = mcolSpecimens.Count
End Property

Public Property Get Specimen(ByVal KeyIndex As Variant) As clsSpecimen
    Set Specimen = mcolSpecimens(KeyIndex)
End Property

Private Sub Class_Initialize()
   Set mcolSpecimens = New Collection
   mvarCount = 0
End Sub

Public Function Add(ByVal clsKey As Variant, ByVal clsData As clsSpecimen)
   
   Dim newSpecimen As New clsSpecimen
   
   With newSpecimen
      .TestCd = clsData.TestCd
      .SpcCd = clsData.SpcCd
      .Seq = clsData.Seq
      .ApplyDt = clsData.ApplyDt
      .SpcGrpCd = clsData.SpcGrpCd
      .RstUnit = clsData.RstUnit
      .RndFg = clsData.RndFg
      .StatFg = clsData.StatFg
      .StatFlags = clsData.StatFlags
      .AvalVal = clsData.AvalVal
      .LabelCnt = clsData.LabelCnt
      .PanicFg = clsData.PanicFg
      .PanicFrVal = clsData.PanicFrVal
      .PanicToVal = clsData.PanicToVal
      .DeltaFg = clsData.DeltaFg
      .DeltaVal1 = clsData.DeltaVal1
      .DeltaVal2 = clsData.DeltaVal2
      .TestCost = clsData.TestCost
      .StoreCd = clsData.StoreCd
      .TatAvg = clsData.TatAvg
      .SpcQty = clsData.SpcQty
      .SpcUnit = clsData.SpcUnit
      .ExpDt = clsData.ExpDt
      .TATS = clsData.TATS
      .ArletFg = clsData.ArletFg
      .ArletFrVal = clsData.ArletFrVal
      .ArletToVal = clsData.ArletToVal
   End With
   
   mcolSpecimens.Add newSpecimen, clsKey
   Set Add = newSpecimen
   mvarCount = mvarCount + 1
   
End Function

Public Sub Clear()
   Set mcolSpecimens = Nothing
   Set mcolSpecimens = New Collection
   mvarCount = 0
End Sub

Public Function Update(ByVal clsKey As Variant, ByVal clsData As clsSpecimen)
   
   With mcolSpecimens.Item(clsKey)
      .TestCd = clsData.TestCd
      .SpcCd = clsData.SpcCd
      .Seq = clsData.Seq
      .ApplyDt = clsData.ApplyDt
      .SpcGrpCd = clsData.SpcGrpCd
      .RstUnit = clsData.RstUnit
      .RndFg = clsData.RndFg
      .StatFg = clsData.StatFg
      .StatFlags = clsData.StatFlags
      .AvalVal = clsData.AvalVal
      .PanicFg = clsData.PanicFg
      .PanicFrVal = clsData.PanicFrVal
      .PanicToVal = clsData.PanicToVal
      .DeltaFg = clsData.DeltaFg
      .DeltaVal1 = clsData.DeltaVal1
      .DeltaVal2 = clsData.DeltaVal2
      .TestCost = clsData.TestCost
      .StoreCd = clsData.StoreCd
      .TatAvg = clsData.TatAvg
      .SpcQty = clsData.SpcQty
      .SpcUnit = clsData.SpcUnit
      .ExpDt = clsData.ExpDt
   End With

End Function
   
Public Function Remove(ByVal clsKey As Variant)
    mcolSpecimens.Remove clsKey
End Function
