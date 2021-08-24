VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsS2DLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'*--------------------------------------------------------------------------------------------------
'*기능 : 리스트 뷰에 정보를 뿌려준다.

'*Coding By Legends
'*Coding Date 2000.10



Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long





Private WithEvents MyForm As frmDLP001
Attribute MyForm.VB_VarHelpID = -1

'속성값을 가질 지역 변수
Private mvarCode As String
Private mvarSqlStmt As String
Private mvarTag As String
Private mvarBackColor As Long
Private mvarCaption As String
Private mvarWidth As Long
Private mvarHeight As Long
Private mvarHeadName As String
Private mvarColSize As String
Private mvarSelectedString As String
Private dctColSize As New Scripting.Dictionary
Dim lngMeWidth As Long
Dim lngMeHeight As Long

Public Event SendCode(ByVal SelString As String)

Public Property Let ColSize(ByVal FieldPosition As String, ByVal vData As String)
'
    If dctColSize.Exists(FieldPosition) = False Then
        dctColSize.Add FieldPosition, vData
    End If
'
End Property

Public Property Get SelectedString() As String
'
    SelectedString = mvarSelectedString
    
End Property


Public Property Let HeadName(ByVal vData As String)
'
    mvarHeadName = vData
    
End Property


Public Property Get HeadName() As String
'
    HeadName = mvarHeadName
    
End Property

Public Property Let Width(ByVal vData As Long)
'
    mvarWidth = vData
    If mvarWidth < lngMeWidth Then mvarWidth = lngMeWidth
    
End Property


Public Property Get Width() As Long
'
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
'
    mvarHeight = vData
    If mvarHeight < lngMeHeight Then mvarHeight = lngMeHeight

End Property


Public Property Get Height() As Long
'
    Height = mvarHeight
End Property

Public Property Let Caption(ByVal vData As String)
'
    mvarCaption = vData

End Property


Public Property Get Caption() As String
'
    Caption = mvarCaption

End Property
Public Property Let BackColor(ByVal vData As Long)
'
    mvarBackColor = vData

End Property


Public Property Get BackColor() As Long
'
    BackColor = mvarBackColor

End Property

Public Property Let Tag(ByVal vData As String)
'
    mvarTag = vData

End Property

Public Property Get Tag() As String
'
    Tag = mvarTag

End Property

Public Property Let Code(ByVal vData As String)
'
    mvarCode = vData

End Property


Public Property Get Code() As String
'
    Code = mvarCode

End Property

Public Property Let SqlStmt(ByVal vData As String)
'
    mvarSqlStmt = vData

End Property

Public Sub Clear()
    '
    mvarCaption = "코드찾기"
    mvarBackColor = &HEEEEEE
    lngMeHeight = 4095
    mvarHeight = 4095
    lngMeWidth = 3585
    mvarWidth = 3585
    mvarSelectedString = ""
    mvarHeadName = ""
    mvarTag = ""
    dctColSize.RemoveAll
    '
End Sub

Public Function ListPop(Optional ByVal pSqlStmt As String = "", _
                        Optional ByVal FormTop As Long = 0, _
                        Optional ByVal FormLeft As Long = 0, _
                        Optional ByRef objDic As clsDictionary)
                        
    Dim p As POINTAPI
    
    Call GetCursorPos(p)
    FormTop = IIf(FormTop = 0, p.Y * Screen.TwipsPerPixelY, FormTop)
    FormLeft = IIf(FormLeft = 0, p.X * Screen.TwipsPerPixelX, FormLeft)
    
    
    Set MyForm = frmDLP001
    
    With frmDLP001
        .SqlStmt = IIf(pSqlStmt = "", mvarSqlStmt, pSqlStmt)
        .HeadName = mvarHeadName
        Set .ColSize = dctColSize
        If Not IsMissing(objDic) Then
            Set .DicObj = objDic
        End If
        .Top = FormTop: .Left = FormLeft
        .ZOrder 0
        .Caption = mvarCaption
        .BackColor = mvarBackColor
        
        .fraDLP.BackColor = mvarBackColor
        .optCode(0).BackColor = mvarBackColor
        .optCode(1).BackColor = mvarBackColor
        '
        If mvarWidth > lngMeWidth Then
            .Width = mvarWidth
            .lvwCodeList.Width = .lvwCodeList.Width + (mvarWidth - lngMeWidth)
            .fraDLP.Width = .lvwCodeList.Width + (mvarWidth - lngMeWidth)
            .txtSearch.Width = .txtSearch.Width + (mvarWidth - lngMeWidth)
        End If
        '
        If mvarHeight > lngMeHeight Then
            .Height = mvarHeight
            .lvwCodeList.Height = .lvwCodeList.Height + (mvarHeight - lngMeHeight)
            .fraDLP.Top = .fraDLP.Top + (mvarHeight - lngMeHeight)
        End If
        
        '위치조정--- 혹시 화면 밖을 나갈까봐서...
        If .Left < 0 Then .Left = .Width
        If .Top < 0 Then .Top = .Height
        If (.Left + .Width) > Screen.Width Then .Left = .Left - .Width
        If (.Top + .Height) > Screen.Height Then .Top = .Top - .Height
        
        .WindowState = vbNormal
        .Show 1
    End With
    
End Function

Private Sub Class_Initialize()
    
    mvarCaption = "코드찾기"
    mvarBackColor = &HEEEEEE
    lngMeHeight = 4095
    mvarHeight = 4095
    lngMeWidth = 3585
    mvarWidth = 3585
    mvarSelectedString = ""
    mvarHeadName = ""
    mvarTag = ""
    dctColSize.RemoveAll
    
End Sub

Private Sub Class_Terminate()
    '
    
    '
End Sub

Private Sub MyForm_ListSelected(ByVal strSelList As String)
    mvarSelectedString = strSelList
    Unload MyForm
    Set MyForm = Nothing
    RaiseEvent SendCode(strSelList)
End Sub
