VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'## 통신기호
Public Property Get ENQ() As String
    ENQ = Chr(mENQ)
End Property

Public Property Get ACK() As String
    ACK = Chr(mACK)
End Property

Public Property Get STX() As String
    STX = Chr(mSTX)
End Property

Public Property Get ETB() As String
    ETB = Chr(mETB)
End Property

Public Property Get ETX() As String
    ETX = Chr(mETX)
End Property

Public Property Get EOT() As String
    EOT = Chr(mEOT)
End Property

Public Property Get NAK() As String
    NAK = Chr(mNAK)
End Property

Public Property Get SOH() As String
    SOH = Chr(mSOH)
End Property

Public Property Get DLE() As String
    DLE = Chr(mDLE)
End Property

Public Property Get SYN() As String
    SYN = Chr(mSYN)
End Property


