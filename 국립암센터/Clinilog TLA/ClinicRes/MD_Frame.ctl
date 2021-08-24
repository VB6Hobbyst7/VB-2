VERSION 5.00
Begin VB.UserControl MDFrame 
   Appearance      =   0  'Æò¸é
   CanGetFocus     =   0   'False
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   ControlContainer=   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   2535
   ToolboxBitmap   =   "MD_Frame.ctx":0000
End
Attribute VB_Name = "MDFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Win32 edge draw consts
Private Const BF_BOTTOM = &H8
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_TOP = &H2
Private Const BF_LEFT = &H1
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)

' 3 Win32 functions draw all the graphics we need
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal EDGE As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long

' define UDT rectangle for Win32 calls
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' set up an automatic EDGE selection system ENUMERATED VALUES
Enum EdgeTypes
    None = 0
    Border = 9&
    Etch = 6&
    RaiseLight = 4&
    RaiseHeavy = 5&
    SunkenLight = 2&
    SunkenHeavy = 10&
End Enum

' set up an automatic size facility
' unfortunately -- vb requires us to NAME each enumed value
Enum EdgeSizeTypes
    e1 = 1
    e2
    e3
    e4
    e5
    e6
    e7
    e8
    e9
    e10
End Enum

'Default Property Values:  what a new UC starts off with
Const m_def_EdgeOUTER = Etch
Const m_def_EdgeINNER = None
Const m_def_EdgeSpacing = e3

'Property Variables:  (internal MEMBER variables)
Dim m_EdgeINNER As Long
Dim m_EdgeOUTER As Long
Dim m_EdgeSpacing As Long


' We give the end-user of this control some Click Events just in case
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."

Private Sub UserControl_Paint()
'============================   ALL DRAWING DONE HERE
Dim di As Long
Dim rc As RECT

    ' clear the control canvas
    UserControl.Cls
    
    ' get size of control rectangle from Windows
    di = GetClientRect(hwnd, rc)
    ' win32 calls to draw the OUTER Edge using current rectangle
    di = DrawEdge(UserControl.hDC, rc, Me.EdgeOUTER, BF_TOPLEFT)
    di = DrawEdge(UserControl.hDC, rc, Me.EdgeOUTER, BF_BOTTOMRIGHT)
    
    
    ' make rectangle smaller by inner spacing property
    di = InflateRect(rc, -Me.EdgeSpacing, -Me.EdgeSpacing)
    ' win32 calls to draw the INNER Edge using current rectangle
    di = DrawEdge(UserControl.hDC, rc, Me.EdgeINNER, BF_TOPLEFT)
    di = DrawEdge(UserControl.hDC, rc, Me.EdgeINNER, BF_BOTTOMRIGHT)
End Sub


'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
'---------------------------------------------------
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'-------------------------------------------------------------------------
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    UserControl.Refresh
End Property


Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub


'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Refresh
End Property


Public Property Get EdgeINNER() As EdgeTypes
Attribute EdgeINNER.VB_Description = "Sets/returns INNER edge selection."
Attribute EdgeINNER.VB_ProcData.VB_Invoke_Property = ";Appearance"
'============================================
    EdgeINNER = m_EdgeINNER
End Property
Public Property Let EdgeINNER(ByVal New_EdgeINNER As EdgeTypes)
'=================================================================
    m_EdgeINNER = New_EdgeINNER
    PropertyChanged "EdgeINNER"
    UserControl.Refresh
End Property


Public Property Get EdgeOUTER() As EdgeTypes
Attribute EdgeOUTER.VB_Description = "Sets/returns OUTER edge selection."
Attribute EdgeOUTER.VB_ProcData.VB_Invoke_Property = ";Appearance"
'============================================
    EdgeOUTER = m_EdgeOUTER
End Property
Public Property Let EdgeOUTER(ByVal New_EdgeOUTER As EdgeTypes)
'==================================================================
    m_EdgeOUTER = New_EdgeOUTER
    PropertyChanged "EdgeOUTER"
    UserControl.Refresh
End Property


Public Property Get EdgeSpacing() As EdgeSizeTypes
Attribute EdgeSpacing.VB_Description = "Sets/returns SPACE between INNER and OUTER Edges in pixels."
Attribute EdgeSpacing.VB_ProcData.VB_Invoke_Property = ";Appearance"
'===============================================
    EdgeSpacing = m_EdgeSpacing
End Property
Public Property Let EdgeSpacing(ByVal New_EdgeSpacing As EdgeSizeTypes)
'===================================================================
    m_EdgeSpacing = New_EdgeSpacing
    PropertyChanged "EdgeSpacing"
    UserControl.Refresh
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'------------------------------------------
    m_EdgeINNER = m_def_EdgeINNER
    m_EdgeOUTER = m_def_EdgeOUTER
    m_EdgeSpacing = m_def_EdgeSpacing
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'--------------------------------------------------------------------
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_EdgeINNER = PropBag.ReadProperty("EdgeINNER", m_def_EdgeINNER)
    m_EdgeOUTER = PropBag.ReadProperty("EdgeOUTER", m_def_EdgeOUTER)
    m_EdgeSpacing = PropBag.ReadProperty("EdgeSpacing", m_def_EdgeSpacing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("EdgeINNER", m_EdgeINNER, m_def_EdgeINNER)
    Call PropBag.WriteProperty("EdgeOUTER", m_EdgeOUTER, m_def_EdgeOUTER)
    Call PropBag.WriteProperty("EdgeSpacing", m_EdgeSpacing, m_def_EdgeSpacing)
End Sub
