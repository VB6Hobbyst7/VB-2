VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIIS600 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Manager"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3945
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imlTree 
      Left            =   3285
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIIS600.frx":0000
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIIS600.frx":27B2
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIIS600.frx":4F64
            Key             =   "NonSelect"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIIS600.frx":5DB6
            Key             =   "Select"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   8940
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   15769
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTree"
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmIIS600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS600.frm (우리LIS랑 조인할때 사용)
'   작성자  :
'   내  용  : 마스터 트리폼
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Load()
    With frmIIS600
        .Top = 0: .Left = 0
        .Height = mdiIISMain.ScaleHeight: .Width = 4035
    End With
    
    '   - 모니터의 해상도가 변해도 항상 폼의 ScaleHeight에 맞도록 수정
    tvwMenu.Height = frmIIS600.ScaleHeight
    
    Call ShowTreeItem
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = "Manager"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (frmIIS609 Is Nothing) Then Unload frmIIS609
    If Not (frmIIS610 Is Nothing) Then Unload frmIIS610
    If Not (frmIIS611 Is Nothing) Then Unload frmIIS611
    If Not (frmIIS612 Is Nothing) Then Unload frmIIS612
    If Not (frmIIS618 Is Nothing) Then Unload frmIIS618
    Set frmIIS600 = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    Call ShowForm(Node.Key)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Manager 트리의 메뉴항목을 표시
'   인수 :
'       1.pKey : 메뉴항목의 Key
'-----------------------------------------------------------------------------'
Public Sub ShowTreeItem()
    Dim objHop  As clsIISHopMenu
    Dim strKey  As String
    
    '   - 병원별로 장비기준 검사명 입력폼을 설정할수 있도록 수정
    '## 병원별 장비기준 검사명 입력폼의 키를 설정
    Select Case PROJECTCODE
        Case "A001"     '## 성모자애병원
            strKey = "IIS618"
        Case Else
            strKey = "IIS611"
    End Select
    
    Set objHop = New clsIISHopMenu
    With tvwMenu
        .Nodes.Clear
        
        '## 검사장비 관리
        If objHop.Menus("NODE3").Visible Then
            .Nodes.Add , , "NODE3", "검사장비관리 관련", 1, 2
            If objHop.Menus("IIS609").Visible Then .Nodes.Add "NODE3", tvwChild, "IIS609", "검사장비 관리", 3, 4
            If objHop.Menus("IIS610").Visible Then .Nodes.Add "NODE3", tvwChild, "IIS610", "검사장비 통신설정", 3, 4
            If objHop.Menus("IIS611").Visible Then .Nodes.Add "NODE3", tvwChild, strKey, "장비별 검사항목 설정", 3, 4
            If objHop.Menus("IIS612").Visible Then .Nodes.Add "NODE3", tvwChild, "IIS612", "사용장비 선택", 3, 4
            .Nodes(.Nodes.Count).EnsureVisible
        End If
    End With
    Set objHop = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Manager 트리의 메뉴항목의 폼을표시
'   인수 :
'       1.pKey : 메뉴항목의 Key
'-----------------------------------------------------------------------------'
Public Sub ShowForm(ByVal pKey As String)
    Screen.MousePointer = vbHourglass
    
    Select Case pKey
        Case "IIS609"       '## 검사장비 마스터
            frmIIS609.Show
            frmIIS609.ZOrder 0
        Case "IIS610"       '## 검사장비 통신설정
            frmIIS610.Show
            frmIIS610.ZOrder 0
        Case "IIS611"       '## 장비별 검사항목(NEW)
            frmIIS611.Show
            frmIIS611.ZOrder 0
        Case "IIS618"       '## 장비별 검사항목(OLD)
            frmIIS618.Show
            frmIIS618.ZOrder 0
        Case "IIS612"       '## 검사장비 선택
            frmIIS612.Show
            frmIIS612.ZOrder 0
    End Select
    tvwMenu.Nodes(pKey).Selected = True
    
    Screen.MousePointer = vbDefault
End Sub
