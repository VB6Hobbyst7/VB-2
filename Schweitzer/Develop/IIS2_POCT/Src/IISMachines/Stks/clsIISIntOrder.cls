VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISIntOrder.cls
'   작성자  : 이상대
'   내  용  : Coulter Stks 장비
'   작성일  : 2004-10-06
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBlkCnt As Long     'Block Count, Block Number Count

Public Property Get BlkCnt() As Long
    BlkCnt = mBlkCnt
End Property

Public Property Let BlkCnt(ByVal vData As Long)
    mBlkCnt = vData
End Property
