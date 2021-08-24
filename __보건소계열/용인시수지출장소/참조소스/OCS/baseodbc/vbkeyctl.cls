VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub GetKeyboardState Lib "User32" (lpkeystate As Any)
Private Declare Sub SetKeyboardState Lib "User32" (lpkeystate As Any)

Const VK_NUMLOCK = &H90
Const VK_CAPITAL = &H14

Dim nKeyState(128)      As Integer
Property Let CapsLock(ByVal bState As Boolean)
    
    GetKeyboardState nKeyState(0)
    
    If nKeyState(VK_CAPITAL / 2) And bState Then Exit Property
    
    nKeyState(VK_CAPITAL / 2) = IIf(bState = True, 1, 0)
    
    SetKeyboardState nKeyState(0)

End Property

Property Get CapsLock() As Boolean
    
    GetKeyboardState nKeyState(0)
    
    CapsLock = IIf(nKeyState(VK_CAPITAL / 2), True, False)

End Property

Property Let Numlock(ByVal bState As Boolean)

    GetKeyboardState nKeyState(0)
    
    If nKeyState(VK_NUMLOCK / 2) And bState Then Exit Property
    
    nKeyState(VK_NUMLOCK / 2) = IIf(bState = True, 1, 0)
    
    SetKeyboardState nKeyState(0)

End Property

Property Get Numlock() As Boolean

    GetKeyboardState nKeyState(0)
    
    Numlock = IIf(nKeyState(VK_NUMLOCK / 2), True, False)
    
End Property


