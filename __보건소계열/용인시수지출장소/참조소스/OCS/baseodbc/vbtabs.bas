Attribute VB_Name = "VbTabConstant"
'*********************************************
' TABVBX10.BAS
'
' Copyright (C) 1994 - FarPoint Technologies
' All Rights Reserved.
'*********************************************

' AlignPictureH
Global Const TABVBX_ALIGNPICT_H_LEFT = 0
Global Const TABVBX_ALIGNPICT_H_CENTER = 1
Global Const TABVBX_ALIGNPICT_H_RIGHT = 2
Global Const TABVBX_ALIGNPICT_H_LEFTTXT = 3
Global Const TABVBX_ALIGNPICT_H_RIGHTTXT = 4

' AlignPictureV
Global Const TABVBX_ALIGNPICT_V_TOP = 0
Global Const TABVBX_ALIGNPICT_V_CENTER = 1
Global Const TABVBX_ALIGNPICT_V_BOTTOM = 2

' AlignTextH
Global Const TABVBX_ALIGNTEXT_H_LEFT = 0
Global Const TABVBX_ALIGNTEXT_H_CENTER = 1
Global Const TABVBX_ALIGNTEXT_H_RIGHT = 2

' AlignTextV
Global Const TABVBX_ALIGNTEXT_V_TOP = 0
Global Const TABVBX_ALIGNTEXT_V_CENTER = 1
Global Const TABVBX_ALIGNTEXT_V_BOTTOM = 2

' ApplyTo
Global Const TABVBX_APPLYTO_DEFAULT = 0
Global Const TABVBX_APPLYTO_ACTIVE = 1
Global Const TABVBX_APPLYTO_TAB = 2

' BorderAlignTextH
Global Const TABVBX_BORDERALIGNTEXT_H_LEFT = 0
Global Const TABVBX_BORDERALIGNTEXT_H_CENTER = 1
Global Const TABVBX_BORDERALIGNTEXT_H_RIGHT = 2
Global Const TABVBX_BORDERALIGNTEXT_H_TOP = 0
Global Const TABVBX_BORDERALIGNTEXT_H_BOTTOM = 2

' BorderTextOrientation
Global Const TABVBX_BORDERORIENT_TOP = 0
Global Const TABVBX_BORDERORIENT_RIGHT = 1
Global Const TABVBX_BORDERORIENT_BOTTOM = 2
Global Const TABVBX_BORDERORIENT_LEFT = 3

' FrameThreeDStyle
Global Const TABVBX_FRAME3D_STY_NONE = 0
Global Const TABVBX_FRAME3D_STY_LOW = 1
Global Const TABVBX_FRAME3D_STY_RAISED = 2
Global Const TABVBX_FRAME3D_STY_GRPLOW = 3
Global Const TABVBX_FRAME3D_STY_GRPRAISED = 4

' Orientation
Global Const TABVBX_ORIENTATION_TOP = 0
Global Const TABVBX_ORIENTATION_RIGHT = 1
Global Const TABVBX_ORIENTATION_BOTTOM = 2
Global Const TABVBX_ORIENTATION_LEFT = 3

' TabShape
Global Const TABVBX_SHAPE_CORNERED = 0
Global Const TABVBX_SHAPE_SLANTED = 1
Global Const TABVBX_SHAPE_ROUNDED = 2

' TabState
Global Const TABVBX_STATE_ENABLED = 0
Global Const TABVBX_STATE_HIDE = 1
Global Const TABVBX_STATE_DISABLED_TEXT = 2
Global Const TABVBX_STATE_DISABLED_NOTEXT = 3

' TextRotation
Global Const TABVBX_TEXT_ROTATE_0 = 0
Global Const TABVBX_TEXT_ROTATE_90 = 1
Global Const TABVBX_TEXT_ROTATE_180 = 2
Global Const TABVBX_TEXT_ROTATE_270 = 3

' ThreeDStyle
Global Const TABVBX_3DSTYLE_NONE = 0
Global Const TABVBX_3DSTYLE_LOW = 1
Global Const TABVBX_3DSTYLE_RAISED = 2
Global Const TABVBX_3DSTYLE_GRPLOW = 3
Global Const TABVBX_3DSTYLE_GRPRAISED = 4

' ThreeDText
Global Const TABVBX_3DTEXT_NONE = 0
Global Const TABVBX_3DTEXT_RSD_LIGHT = 1
Global Const TABVBX_3DTEXT_LOW_LIGHT = 2
Global Const TABVBX_3DTEXT_RSD_HEAVY = 3
Global Const TABVBX_3DTEXT_LOW_HEAVY = 4

Declare Sub TabAssignChild Lib "fptab10.vbx" (hCtlTab As Control, hCtlChild As Control, ByVal TabIndex As Integer, ByVal xPos As Long, ByVal yPos As Long)
Declare Sub TabRemoveChild Lib "fptab10.vbx" (hCtlTab As Control, hCtlChild As Control)
