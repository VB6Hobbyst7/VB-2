VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl urcTreeList 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin Threed.SSPanel lblTitle 
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   767
      _Version        =   262144
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin MSComctlLib.TreeView trvList 
      Height          =   1965
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3466
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "urcTreeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum titAlignmentP
    ssLeftTop = 0
    ssLeftMiddle = 1
    ssLeftBottom = 2
    ssRightTop = 3
    ssRightMiddle = 4
    ssRightBottom = 5
    ssCenterTop = 6
    ssCenterMiddle = 7
    ssCenterBottom = 8
End Enum

Public Enum titFont3DP
    ssNoneFont3D = 0
    ssRaisedLight = 1
    ssRaisedHeavy = 2
    ssInsetLight = 3
    ssInsetHeavy = 4
    ssDropShadow = 5
End Enum

Public Enum titBevelInnerP
    ssNoneBevel = 0
    ssInsetBevel = 1
    ssRaisedBevel = 2
End Enum

Public Enum titBevelOuterP
    ssNoneBevel = 0
    ssInsetBevel = 1
    ssRaisedBevel = 2
End Enum

Public Enum AppearanceP
    ccFlat = 0
    cc3D = 1
End Enum

Public Enum BorderStyleP
    None = 0
    singleBorder = 1
End Enum

Dim m_Border As BorderStyleP
Dim m_Appearance As AppearanceP

Event Click()
Event DblClick()
Event Expand(Node As MSComctlLib.Node)
Event Collapse(Node As MSComctlLib.Node)

Public Property Get titCaption() As String

    titCaption = lblTitle.Caption
    
End Property

Public Property Let titCaption(ByVal New_titCaption As String)

    lblTitle.Caption = New_titCaption
    PropertyChanged "titCaption"
    
End Property

Public Property Get titForeColor() As OLE_COLOR

    titForeColor = lblTitle.ForeColor
    
End Property

Public Property Let titForeColor(ByVal New_titForeColor As OLE_COLOR)

    lblTitle.ForeColor() = New_titForeColor
    PropertyChanged "titForeColor"
    
End Property

Public Property Get titFont3D() As titFont3DP

    titFont3D = lblTitle.Font3D
    
End Property

Public Property Let titFont3D(ByVal New_titFont3D As titFont3DP)

    lblTitle.Font3D() = New_titFont3D
    PropertyChanged "titFont3D"
    
End Property

Public Property Get titBackColor() As OLE_COLOR

    titBackColor = lblTitle.BackColor
    
End Property

Public Property Let titBackColor(ByVal New_titBackColor As OLE_COLOR)

    lblTitle.BackColor() = New_titBackColorr
    PropertyChanged "titBackColor"
    
End Property

Public Property Get titVisible() As Boolean

    titVisible = lblTitle.Visible

End Property

Public Property Let titVisible(ByVal New_titVisible As Boolean)

    lblTitle.Visible() = New_titVisible
    PropertyChanged "titVisible"
    
    Call UserControl_Resize

End Property

Public Property Get titFont() As Font

    Set titFont = lblTitle.Font

End Property

Public Property Let titFont(ByVal New_titFont As Font)

    Set lblTitle.Font = New_titFont
    PropertyChanged "titFont"

End Property

Public Property Get titAlignment() As titAlignmentP

    titAlignment = lblTitle.Alignment

End Property

Public Property Let titAlignment(ByVal New_titAlignment As titAlignmentP)

    lblTitle.Alignment = New_titAlignment
    PropertyChanged "titAlignment"

End Property

Public Property Get titBevelInner() As titBevelInnerP

    titBevelInner = lblTitle.BevelInner

End Property

Public Property Let titBevelInner(ByVal New_titBevelInner As titBevelInnerP)

    lblTitle.BevelInner = New_titBevelInner
    PropertyChanged "titBevelInner"

End Property

Public Property Get titBevelWidth() As Integer

    titBevelWidth = lblTitle.BevelWidth

End Property

Public Property Let titBevelWidth(ByVal New_titBevelWidth As Integer)

    lblTitle.BevelWidth = New_titBevelWidth
    PropertyChanged "titBevelWidth"

End Property

Public Property Get titBorderWidth() As Integer

    titBorderWidth = lblTitle.BorderWidth

End Property

Public Property Let titBorderWidth(ByVal New_titBorderWidth As Integer)

    lblTitle.BorderWidth = New_titBorderWidth
    PropertyChanged "titBorderWidth"

End Property

Public Property Get titBevelOuter() As titBevelOuterP

    titBevelOuter = lblTitle.BevelOuter

End Property

Public Property Let titBevelOuter(ByVal New_titBevelOuter As titBevelOuterP)

    lblTitle.BevelOuter = New_titBevelOuter
    PropertyChanged "titBevelOuter"

End Property

Public Property Get BorderStyle() As BorderStyleP

    BorderStyle = trvList.BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleP)

    trvList.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"

End Property

Public Property Get Appearance() As AppearanceP

    Appearance = trvList.Appearance

End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceP)

    trvList.Appearance = New_Appearance
    PropertyChanged "Appearance"

End Property

Private Sub trvList_Click()

    RaiseEvent Click

End Sub

Private Sub trvList_Collapse(ByVal Node As MSComctlLib.Node)

    RaiseEvent Collapse(Node)

End Sub

Private Sub trvList_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub trvList_Expand(ByVal Node As MSComctlLib.Node)

    RaiseEvent Expand(Node)

End Sub

Private Sub UserControl_Initialize()

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblTitle.BackColor = PropBag.ReadProperty("titBackColor", &H8000000F)
    lblTitle.ForeColor = PropBag.ReadProperty("titForeColor", &H80000012)
    lblTitle.Visible = PropBag.ReadProperty("titVisible", True)
    lblTitle.Alignment = PropBag.ReadProperty("titAlignment", 7)
    lblTitle.Font3D = PropBag.ReadProperty("titFont3D", 0)
    lblTitle.BevelInner = PropBag.ReadProperty("titBevelInner", 0)
    lblTitle.BevelOuter = PropBag.ReadProperty("titBevelOuter", 2)
    lblTitle.BevelWidth = PropBag.ReadProperty("titBevelWidth", 1)
    lblTitle.BorderWidth = PropBag.ReadProperty("titBorderWidth", 3)
    lblTitle.Caption = PropBag.ReadProperty("titCaption", "titTitle")
    Set lblTitle.Font = PropBag.ReadProperty("titFont", Ambient.Font)
    
    trvList.Appearance = PropBag.ReadProperty("Appearance", 1)
    trvList.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)

End Sub

Private Sub UserControl_Resize()

    lblTitle.Top = 0
    lblTitle.Left = 0
    lblTitle.Width = UserControl.Width
    
    trvList.Left = 0
    trvList.Width = UserControl.Width
    
    If lblTitle.Visible Then
        trvList.Top = lblTitle.Height + 20
        trvList.Height = UserControl.Height - lblTitle.Height - 20
    Else
        trvList.Top = 0
        trvList.Height = UserControl.Height
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("titBackColor", lblTitle.BackColor, &H8000000F)
    Call PropBag.WriteProperty("titForeColor", lblTitle.ForeColor, &H80000012)
    Call PropBag.WriteProperty("titVisible", lblTitle.Visible, True)
    Call PropBag.WriteProperty("titAlignment", lblTitle.Alignment, 7)
    Call PropBag.WriteProperty("titFont3D", lblTitle.Font3D, 0)
    Call PropBag.WriteProperty("titBevelInner", lblTitle.BevelInner, 0)
    Call PropBag.WriteProperty("titBevelOuter", lblTitle.BevelOuter, 2)
    Call PropBag.WriteProperty("titBevelWidth", lblTitle.BevelWidth, 1)
    Call PropBag.WriteProperty("titBorderWidth", lblTitle.BorderWidth, 3)
    Call PropBag.WriteProperty("titCaption", lblTitle.Caption, "titTitle")
    
    Call PropBag.WriteProperty("Appearance", trvList.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", trvList.BorderStyle, 0)

End Sub
