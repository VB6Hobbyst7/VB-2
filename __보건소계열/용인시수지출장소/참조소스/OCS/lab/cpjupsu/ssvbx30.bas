Attribute VB_Name = "mdSpr30"

'----------------------------------------------------------
' Spread 3.0
' File: SSVBX.BAS
'
' Copyright (C) 1998 FarPoint Technologies.
' All rights reserved.
'
'----------------------------------------------------------

' *************************  SpreadSheet Settings *************************

' Action property settings
Global Const SS_ACTION_ACTIVE_CELL = 0
Global Const SS_ACTION_GOTO_CELL = 1
Global Const SS_ACTION_SELECT_BLOCK = 2
Global Const SS_ACTION_CLEAR = 3
Global Const SS_ACTION_DELETE_COL = 4
Global Const SS_ACTION_DELETE_ROW = 5
Global Const SS_ACTION_INSERT_COL = 6
Global Const SS_ACTION_INSERT_ROW = 7
Global Const SS_ACTION_LOAD_SPREAD_SHEET = 8
Global Const SS_ACTION_SAVE_ALL = 9
Global Const SS_ACTION_SAVE_VALUES = 10
Global Const SS_ACTION_RECALC = 11
Global Const SS_ACTION_CLEAR_TEXT = 12
Global Const SS_ACTION_PRINT = 13
Global Const SS_ACTION_DESELECT_BLOCK = 14
Global Const SS_ACTION_DSAVE = 15
Global Const SS_ACTION_SET_CELL_BORDER = 16
Global Const SS_ACTION_ADD_MULTISELBLOCK = 17
Global Const SS_ACTION_GET_MULTI_SELECTION = 18
Global Const SS_ACTION_COPY_RANGE = 19
Global Const SS_ACTION_MOVE_RANGE = 20
Global Const SS_ACTION_SWAP_RANGE = 21
Global Const SS_ACTION_CLIPBOARD_COPY = 22
Global Const SS_ACTION_CLIPBOARD_CUT = 23
Global Const SS_ACTION_CLIPBOARD_PASTE = 24
Global Const SS_ACTION_SORT = 25
Global Const SS_ACTION_COMBO_CLEAR = 26
Global Const SS_ACTION_COMBO_REMOVE = 27
Global Const SS_ACTION_RESET = 28
Global Const SS_ACTION_SEL_MODE_CLEAR = 29
Global Const SS_ACTION_VMODE_REFRESH = 30
Global Const SS_ACTION_SMARTPRINT = 32

' Appearance property settings
Global Const SS_APPEARANCE_FLAT = 0
Global Const SS_APPEARANCE_3D = 1
Global Const SS_APPEARANCE_3DWITHBORDER = 2

' BackColorStyle property settings
Global Const SS_BACKCOLORSTYLE_OVERGRID = 0
Global Const SS_BACKCOLORSTYLE_UNDERGRID = 1
Global Const SS_BACKCOLORSTYLE_OVERHORZGRIDONLY = 2
Global Const SS_BACKCOLORSTYLE_OVERVERTGRIDONLY = 3

' ButtonDrawMode property settings
Global Const SS_BDM_ALWAYS = 0
Global Const SS_BDM_CURRENT_CELL = 1
Global Const SS_BDM_CURRENT_COLUMN = 2
Global Const SS_BDM_CURRENT_ROW = 4
Global Const SS_BDM_ALWAYS_BUTTON = 8
Global Const SS_BDM_ALWAYS_COMBO = 16

' CellBorderStyle property settings
Global Const SS_BORDER_STYLE_DEFAULT = 0
Global Const SS_BORDER_STYLE_SOLID = 1
Global Const SS_BORDER_STYLE_DASH = 2
Global Const SS_BORDER_STYLE_DOT = 3
Global Const SS_BORDER_STYLE_DASH_DOT = 4
Global Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Global Const SS_BORDER_STYLE_BLANK = 6
Global Const SS_BORDER_STYLE_FINE_SOLID = 11
Global Const SS_BORDER_STYLE_FINE_DASH = 12
Global Const SS_BORDER_STYLE_FINE_DOT = 13
Global Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Global Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' CellBorderType property settings
Global Const SS_BORDER_TYPE_NONE = 0
Global Const SS_BORDER_TYPE_OUTLINE = 16
Global Const SS_BORDER_TYPE_LEFT = 1
Global Const SS_BORDER_TYPE_RIGHT = 2
Global Const SS_BORDER_TYPE_TOP = 4
Global Const SS_BORDER_TYPE_BOTTOM = 8

' CellType property settings
Global Const SS_CELL_TYPE_DATE = 0
Global Const SS_CELL_TYPE_EDIT = 1
Global Const SS_CELL_TYPE_FLOAT = 2
Global Const SS_CELL_TYPE_INTEGER = 3
Global Const SS_CELL_TYPE_PIC = 4
Global Const SS_CELL_TYPE_STATIC_TEXT = 5
Global Const SS_CELL_TYPE_TIME = 6
Global Const SS_CELL_TYPE_BUTTON = 7
Global Const SS_CELL_TYPE_COMBOBOX = 8
Global Const SS_CELL_TYPE_PICTURE = 9
Global Const SS_CELL_TYPE_CHECKBOX = 10
Global Const SS_CELL_TYPE_OWNER_DRAWN = 11

' ClipboardOptions property settings
Global Const SS_CLIP_NOHEADERS = 0
Global Const SS_CLIP_COPYROWHEADERS = 1
Global Const SS_CLIP_PASTEROWHEADERS = 2
Global Const SS_CLIP_COPYCOLHEADERS = 4
Global Const SS_CLIP_PASTECOLHEADERS = 8
Global Const SS_CLIP_COPYPASTEALLHEADERS = 15

' ColHeaderDisplay and RowHeaderDisplay property settings
Global Const SS_HEADER_BLANK = 0
Global Const SS_HEADER_NUMBERS = 1
Global Const SS_HEADER_LETTERS = 2

' CursorStyle property settings
Global Const SS_CURSOR_STYLE_USER_DEFINED = 0
Global Const SS_CURSOR_STYLE_DEFAULT = 1
Global Const SS_CURSOR_STYLE_ARROW = 2
Global Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Global Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Global Const SS_CURSOR_TYPE_DEFAULT = 0
Global Const SS_CURSOR_TYPE_COLRESIZE = 1
Global Const SS_CURSOR_TYPE_ROWRESIZE = 2
Global Const SS_CURSOR_TYPE_BUTTON = 3
Global Const SS_CURSOR_TYPE_GRAYAREA = 4
Global Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Global Const SS_CURSOR_TYPE_COLHEADER = 6
Global Const SS_CURSOR_TYPE_ROWHEADER = 7
Global Const SS_CURSOR_TYPE_DRAGDROPAREA = 8
Global Const SS_CURSOR_TYPE_DRAGDROP = 9

' DAutoSize property settings
Global Const SS_AUTOSIZE_NO = 0
Global Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Global Const SS_AUTOSIZE_BEST_GUESS = 2

' EditEnterAction property settings
Global Const SS_CELL_EDITMODE_EXIT_NONE = 0
Global Const SS_CELL_EDITMODE_EXIT_UP = 1
Global Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Global Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Global Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Global Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Global Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Global Const SS_CELL_EDITMODE_EXIT_SAME = 7
Global Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' OperationMode property settings
Global Const SS_OP_MODE_NORMAL = 0
Global Const SS_OP_MODE_READONLY = 1
Global Const SS_OP_MODE_ROWMODE = 2
Global Const SS_OP_MODE_SINGLE_SELECT = 3
Global Const SS_OP_MODE_MULTI_SELECT = 4
Global Const SS_OP_MODE_EXT_SELECT = 5

' Position property settings
Global Const SS_POSITION_UPPER_LEFT = 0
Global Const SS_POSITION_UPPER_CENTER = 1
Global Const SS_POSITION_UPPER_RIGHT = 2
Global Const SS_POSITION_CENTER_LEFT = 3
Global Const SS_POSITION_CENTER_CENTER = 4
Global Const SS_POSITION_CENTER_RIGHT = 5
Global Const SS_POSITION_BOTTOM_LEFT = 6
Global Const SS_POSITION_BOTTOM_CENTER = 7
Global Const SS_POSITION_BOTTOM_RIGHT = 8

' PrintOrientation property settings
Global Const SS_PRINTORIENT_DEFAULT = 0
Global Const SS_PRINTORIENT_PORTRAIT = 1
Global Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintType property settings
Global Const SS_PRINT_ALL = 0
Global Const SS_PRINT_CELL_RANGE = 1
Global Const SS_PRINT_CURRENT_PAGE = 2
Global Const SS_PRINT_PAGE_RANGE = 3

' ScrollBars property settings
Global Const SS_SCROLLBAR_NONE = 0
Global Const SS_SCROLLBAR_H_ONLY = 1
Global Const SS_SCROLLBAR_V_ONLY = 2
Global Const SS_SCROLLBAR_BOTH = 3

' ScrollBarTrack property settings
Global Const SS_SCROLLBARTRACK_OFF = 0
Global Const SS_SCROLLBARTRACK_VERTICAL = 1
Global Const SS_SCROLLBARTRACK_HORIZONTAL = 2
Global Const SS_SCROLLBARTRACK_BOTH = 3

' SelBackColor property settings
Global Const SPREAD_COLOR_NONE = &H8000000B

' SelectBlockOptions property settings
Global Const SS_SELBLOCKOPT_COLS = 1
Global Const SS_SELBLOCKOPT_ROWS = 2
Global Const SS_SELBLOCKOPT_BLOCKS = 4
Global Const SS_SELBLOCKOPT_ALL = 8

' SortBy property settings
Global Const SS_SORT_BY_ROW = 0
Global Const SS_SORT_BY_COL = 1

' SortKeyOrder property settings
Global Const SS_SORT_ORDER_NONE = 0
Global Const SS_SORT_ORDER_ASCENDING = 1
Global Const SS_SORT_ORDER_DESCENDING = 2

' TextTip property settings
Global Const SS_TEXTTIP_OFF = 0
Global Const SS_TEXTTIP_FIXED = 1
Global Const SS_TEXTTIP_FLOATING = 2
Global Const SS_TEXTTIP_FIXEDFOCUSONLY = 3
Global Const SS_TEXTTIP_FLOATINGFOCUSONLY = 4

' TypeButtonAlign property settings
Global Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Global Const SS_CELL_BUTTON_ALIGN_TOP = 1
Global Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Global Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' TypeButtonType property settings
Global Const SS_CELL_BUTTON_NORMAL = 0
Global Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeCheckTextAlign property settings
Global Const SS_CHECKBOX_TEXT_LEFT = 0
Global Const SS_CHECKBOX_TEXT_RIGHT = 1

' TypeCheckType property settings
Global Const SS_CHECKBOX_NORMAL = 0
Global Const SS_CHECKBOX_THREE_STATE = 1

' TypeDateFormat property settings
Global Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Global Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Global Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Global Const SS_CELL_DATE_FORMAT_YYMMDD = 3

' TypeEditCharCase property settings
Global Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Global Const SS_CELL_EDIT_CASE_NO_CASE = 1
Global Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Global Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Global Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Global Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Global Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeHAlign property settings
Global Const SS_CELL_H_ALIGN_LEFT = 0
Global Const SS_CELL_H_ALIGN_RIGHT = 1
Global Const SS_CELL_H_ALIGN_CENTER = 2

' TypeTextAlignVert property settings
Global Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Global Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Global Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTime24Hour property settings
Global Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Global Const SS_CELL_TIME_24_HOUR_CLOCK = 1

' TypeVAlign property settings
Global Const SS_CELL_V_ALIGN_TOP = 0
Global Const SS_CELL_V_ALIGN_BOTTOM = 1
Global Const SS_CELL_V_ALIGN_VCENTER = 2

' UnitType property settings
Global Const SS_CELL_UNIT_NORMAL = 0
Global Const SS_CELL_UNIT_VGA = 1
Global Const SS_CELL_UNIT_TWIPS = 2

' UserResize property settings
Global Const SS_USER_RESIZE_COL = 1
Global Const SS_USER_RESIZE_ROW = 2

' UserResizeCol and UserResizeRow property settings
Global Const SS_USER_RESIZE_DEFAULT = 0
Global Const SS_USER_RESIZE_ON = 1
Global Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Global Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Global Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Global Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

' ActionKey function settings
Global Const SS_KBA_CLEAR = 0
Global Const SS_KBA_CURRENT = 1
Global Const SS_KBA_POPUP = 2

' AddCustomFunctionExt method Flags parameter settings
Global Const SS_CUSTFUNC_WANTCELLREF = 1
Global Const SS_CUSTFUNC_WANTRANGEREF = 2

' CFGetParamInfo method Type parameter settings
Global Const SS_VALUE_TYPE_LONG = 0
Global Const SS_VALUE_TYPE_DOUBLE = 1
Global Const SS_VALUE_TYPE_STR = 2
Global Const SS_VALUE_TYPE_CELL = 3
Global Const SS_VALUE_TYPE_RANGE = 4

' CFGetParamInfo method Status parameter settings
Global Const SS_VALUE_STATUS_OK = 0
Global Const SS_VALUE_STATUS_ERROR = 1
Global Const SS_VALUE_STATUS_EMPTY = 2

' GetRefStyle/SetRefStyle methods return values/parameter settings
Global Const SS_REFSTYLE_DEFAULT = 0
Global Const SS_REFSTYLE_A1 = 1
Global Const SS_REFSTYLE_R1C1 = 2

' PrintOptions method PageOrder parameter settings
Global Const SS_PAGEORDER_AUTO = 0
Global Const SS_PAGEORDER_DOWNTHENOVER = 1
Global Const SS_PAGEORDER_OVERTHENDOWN = 2

' TextTipFetch method MultiLine parameter settings
Global Const SS_TT_MULTILINE_SINGLE = 0
Global Const SS_TT_MULTILINE_MULTI = 1
Global Const SS_TT_MULTILINE_AUTO = 2

' *************************  PrintPreview Settings *************************

' GrayAreaMarginType property values
Global Const SPV_GRAYAREAMARGINTYPE_SCALED = 0
Global Const SPV_GRAYAREAMARGINTYPE_ACTUAL = 1

' MousePointer property values
Global Const SPV_MOUSEPOINTER_DEFAULT = 0
Global Const SPV_MOUSEPOINTER_ARROW = 1
Global Const SPV_MOUSEPOINTER_CROSS = 2
Global Const SPV_MOUSEPOINTER_I_BEAM = 3
Global Const SPV_MOUSEPOINTER_ICON = 4
Global Const SPV_MOUSEPOINTER_SIZE = 5
Global Const SPV_MOUSEPOINTER_SIZE_NE_SW = 6
Global Const SPV_MOUSEPOINTER_SIZE_N_S = 7
Global Const SPV_MOUSEPOINTER_SIZE_NW_SE = 8
Global Const SPV_MOUSEPOINTER_SIZE_W_E = 9
Global Const SPV_MOUSEPOINTER_UP_ARROW = 10
Global Const SPV_MOUSEPOINTER_HOURGLASS = 11
Global Const SPV_MOUSEPOINTER_NO_DROP = 12

' PageViewType property values
Global Const SPV_PAGEVIEWTYPE_WHOLE_PAGE = 0
Global Const SPV_PAGEVIEWTYPE_NORMAL_SIZE = 1
Global Const SPV_PAGEVIEWTYPE_PERCENTAGE = 2
Global Const SPV_PAGEVIEWTYPE_PAGE_WIDTH = 3
Global Const SPV_PAGEVIEWTYPE_PAGE_HEIGHT = 4
Global Const SPV_PAGEVIEWTYPE_MULTIPLE_PAGES = 5

' ScrollBarH property values
Global Const SPV_SCROLLBARH_SHOW = 0
Global Const SPV_SCROLLBARH_AUTO = 1
Global Const SPV_SCROLLBARH_HIDE = 2

' ScrollBarV property values
Global Const SPV_SCROLLBARV_SHOW = 0
Global Const SPV_SCROLLBARV_AUTO = 1
Global Const SPV_SCROLLBARV_HIDE = 2

' ZoomState property values
Global Const SPV_ZOOMSTATE_INDETERMINATE = 0
Global Const SPV_ZOOMSTATE_IN = 1
Global Const SPV_ZOOMSTATE_OUT = 2
Global Const SPV_ZOOMSTATE_SWITCH = 3

' Function prototypes
'Declare Function SpreadAddCustomFunction Lib "SPRVBX30.VBX" (hCtl As Control, ByVal lpszFunctionName As String, ByVal nParameterCnt As Integer) As Integer
'Declare Function SpreadAddCustomFunctionExt Lib "SPRVBX30.VBX" (hCtl As Control, ByVal lpszFunctionName As String, ByVal nMinParamCnt As Integer, ByVal nMaxParamCnt As Integer, ByVal Flags As Long) As Integer
'Declare Function SpreadEnumCustomFunction Lib "SPRVBX30.VBX" (SS As Control, ByVal PrevFuncName As String, FuncName As String) As Integer
'Declare Sub SpreadCFGetCellParam Lib "SPRVBX30.VBX" (hCtl As Control, ByVal dParam As Integer, Col As Long, Row As Long)
'Declare Function SpreadCFGetDoubleParam Lib "SPRVBX30.VBX" (hCtl As Control, ByVal dParam As Integer) As Double
'Declare Function SpreadCFGetLongParam Lib "SPRVBX30.VBX" (hCtl As Control, ByVal dParam As Integer) As Long
'Declare Function SpreadGetOddEvenRowColor Lib "SPRVBX30.VBX" (SS As Control, lpclrBackOdd As Long, lpclrForeOdd As Long, lpclrBackEven As Long, lpclrForeEven As Long) As Integer
'Declare Function SpreadCFGetParamInfo Lib "SPRVBX30.VBX" (hCtl As Control, ByVal dParam As Integer, wType As Integer, wStatus As Integer) As Integer
'Declare Sub SpreadCFGetRangeParam Lib "SPRVBX30.VBX" (hCtl As Control, ByVal dParam As Integer, Col As Long, Row As Long, Col2 As Long, Row2 As Long)
'Declare Function SpreadCFGetStringParam Lib "SPRVBX30.VBX" (hCtl As Control, ByVal dParam As Integer) As String
'Declare Sub SpreadCFSetResult Lib "SPRVBX30.VBX" (hCtl As Control, Var As Variant)
'Declare Function SpreadColNumberToLetter Lib "SPRVBX30.VBX" (ByVal lHeaderNumber As Long) As String
'Declare Sub SpreadColWidthToTwips Lib "SPRVBX30.VBX" (SS As Control, ByVal fColWidth As Single, lpTwips As Long)
'Declare Function SpreadGetActionKey Lib "SPRVBX30.VBX" (SS As Control, ByVal wAction As Integer, ByVal fShift As Integer, ByVal fCtrl As Integer, ByVal wKey As Integer) As Integer
'Declare Function SpreadGetArray Lib "SPRVBX30.VBX" (SS As Control, ByVal ColLeft As Long, ByVal RowTop As Long, hAD() As Variant) As Integer
'Declare Sub SpreadGetBottomRightCell Lib "SPRVBX30.VBX" (SS As Control, lpCol As Long, lpRow As Long)
'Declare Function SpreadGetCellDirtyFlag Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long) As Integer
'Declare Sub SpreadGetCellFromScreenCoord Lib "SPRVBX30.VBX" (SS As Control, lpCol As Long, lpRow As Long, ByVal x As Long, ByVal y As Long)
'Declare Function SpreadGetCellPos Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, lpx As Long, lpy As Long, lpWidth As Long, lpHeight As Long) As Integer
'Declare Sub SpreadGetClientArea Lib "SPRVBX30.VBX" (SS As Control, lplWidth As Long, lplHeight As Long)
'Declare Function SpreadGetColItemData Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long) As Long
'Declare Function SpreadGetCustomFunction Lib "SPRVBX30.VBX" (SS As Control, ByVal FuncName As String, MinArgs As Integer, MaxArgs As Integer, Flags As Long) As Integer
'Declare Function SpreadGetCustomName Lib "SPRVBX30.VBX" (SS As Control, ByVal sName As String) As String
'Declare Function SpreadGetDataFillData Lib "SPRVBX30.VBX" (SS As Control, Var As Variant, ByVal VType As Integer) As Integer
'Declare Sub SpreadGetFirstValidCell Lib "SPRVBX30.VBX" (SS As Control, lpCol As Long, lpRow As Long)
'Declare Function SpreadGetFloat Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, dfValue As Double) As Integer
'Declare Function SpreadGetFormulaSync Lib "SPRVBX30.VBX" (SS As Control) As Integer
'Declare Function SpreadGetInteger Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, lValue As Long) As Integer
'Declare Function SpreadGetItemData Lib "SPRVBX30.VBX" (SS As Control) As Long
'Declare Function SpreadGetIteration Lib "SPRVBX30.VBX" (SS As Control, MaxIterations As Integer, MaxChange As Double) As Integer
'Declare Sub SpreadGetLastValidCell Lib "SPRVBX30.VBX" (SS As Control, lpCol As Long, lpRow As Long)
'Declare Function SpreadGetMultiSelItem Lib "SPRVBX30.VBX" (SS As Control, ByVal SelPrev As Long) As Long
'Declare Function SpreadGetNextPageBreakCol Lib "SPRVBX30.VBX" (SS As Control, ByVal PrevCol As Long) As Long
'Declare Function SpreadGetNextPageBreakRow Lib "SPRVBX30.VBX" (SS As Control, ByVal PrevRow As Long) As Long
'Declare Function SpreadGetPrintOptions Lib "SPRVBX30.VBX" (SS As Control, SmartPrint As Integer, PageOrder As Integer, FirstPageNumber As Long) As Integer
'Declare Function SpreadGetPrintPageCount Lib "SPRVBX30.VBX" (SS As Control) As Long
'Declare Function SpreadGetRefStyle Lib "SPRVBX30.VBX" (SS As Control) As Integer
'Declare Function SpreadGetRowItemData Lib "SPRVBX30.VBX" (SS As Control, ByVal Row As Long) As Long
'Declare Function SpreadGetText Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, Var As Variant) As Integer
'Declare Function SpreadGetTextTipAppearance Lib "SPRVBX30.VBX" (SS As Control, TipFontName As String, TipFontSize As Integer, TipFontBold As Integer, TipFontItalic As Integer, TipBackColor As Long, TipForeColor As Long) As Integer
'Declare Function SpreadGetTwoDigitYearMax Lib "SPRVBX30.VBX" (SS As Control) As Integer
'Declare Function SpreadIsCellSelected Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long) As Integer
'Declare Function SpreadIsFormulaValid Lib "SPRVBX30.VBX" (SS As Control, hszFormula As String) As Integer
'Declare Function SpreadIsVisible Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal Partial As Integer) As Integer
'Declare Function SpreadLoadFromFile Lib "SPRVBX30.VBX" (SS As Control, ByVal FileName As String) As Integer
'Declare Function SpreadLoadTabFile Lib "SPRVBX30.VBX" (SS As Control, ByVal FileName As String) As Integer
'Declare Function SpreadQueryCustomName Lib "SPRVBX30.VBX" (SS As Control, ByVal sPrevName As String) As String
'Declare Function SpreadReCalcCell Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long) As Integer
'Declare Function SpreadRemoveCustomFunction Lib "SPRVBX30.VBX" (SS As Control, ByVal FuncName As String) As Integer
'Declare Sub SpreadRowHeightToTwips Lib "SPRVBX30.VBX" (SS As Control, ByVal Row As Long, ByVal fRowHeight As Single, lpTwips As Long)
'Declare Function SpreadSaveTabFile Lib "SPRVBX30.VBX" (SS As Control, ByVal FileName As String) As Integer
'Declare Function SpreadSaveToFile Lib "SPRVBX30.VBX" (SS As Control, ByVal FileName As String, ByVal DataOnly As Integer) As Integer
'Declare Function SpreadSetActionKey Lib "SPRVBX30.VBX" (SS As Control, ByVal wAction As Integer, ByVal lpfShift As Integer, ByVal lpfCtrl As Integer, ByVal lpwKey As Integer) As Integer
'Declare Function SpreadSetArray Lib "SPRVBX30.VBX" (SS As Control, ByVal ColLeft As Long, ByVal RowTop As Long, hAD() As Variant) As Integer
'Declare Sub SpreadSetCalText Lib "SPRVBX30.VBX" (ByVal ShortDays As String, ByVal LongDays As String, ByVal ShortMonths As String, ByVal LongMonths As String, ByVal OkText As String, ByVal CancelText As String)
'Declare Function SpreadSetCellDirtyFlag Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal Dirty As Integer) As Integer
'Declare Sub SpreadSetColItemData Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal lpVar As Long)
'Declare Function SpreadSetCustomName Lib "SPRVBX30.VBX" (SS As Control, ByVal sName As String, ByVal sValue As String) As Integer
'Declare Function SpreadSetDataFillData Lib "SPRVBX30.VBX" (SS As Control, Var As Variant) As Integer
'Declare Function SpreadSetFloat Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal dfValue As Double) As Integer
'Declare Sub SpreadSetFormulaSync Lib "SPRVBX30.VBX" (SS As Control, ByVal Sync As Integer)
'Declare Function SpreadSetInteger Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, ByVal lValue As Long) As Integer
'Declare Sub SpreadSetItemData Lib "SPRVBX30.VBX" (SS As Control, ByVal lpVar As Long)
'Declare Sub SpreadSetIteration Lib "SPRVBX30.VBX" (SS As Control, ByVal Iteration As Integer, ByVal MaxIterations As Integer, ByVal MaxChange As Double)
'Declare Function SpreadSetOddEvenRowColor Lib "SPRVBX30.VBX" (SS As Control, ByVal clrBackOdd As Long, ByVal clrForeOdd As Long, ByVal clrBackEven As Long, ByVal clrForeEven As Long) As Integer
'Declare Function SpreadSetPrintOptions Lib "SPRVBX30.VBX" (SS As Control, ByVal SmartPrint As Integer, ByVal PageOrder As Integer, ByVal FirstPageNumber As Long) As Integer
'Declare Sub SpreadSetRefStyle Lib "SPRVBX30.VBX" (SS As Control, ByVal RefStyle As Integer)
'Declare Sub SpreadSetRowItemData Lib "SPRVBX30.VBX" (SS As Control, ByVal Row As Long, ByVal lpVar As Long)
'Declare Sub SpreadSetText Lib "SPRVBX30.VBX" (SS As Control, ByVal Col As Long, ByVal Row As Long, lpVar As Variant)
'Declare Function SpreadSetTextTipAppearance Lib "SPRVBX30.VBX" (SS As Control, TipFontName As String, ByVal TipFontSize As Integer, ByVal TipFontBold As Integer, ByVal TipFontItalic As Integer, ByVal TipBackColor As Long, ByVal TipForeColor As Long) As Integer
'Declare Function SpreadSetTwoDigitYearMax Lib "SPRVBX30.VBX" (SS As Control, ByVal TwoDigitYearMax As Integer) As Integer
'Declare Sub SpreadTwipsToColWidth Lib "SPRVBX30.VBX" (SS As Control, ByVal Twips As Long, fColWidth As Single)
'Declare Sub SpreadTwipsToRowHeight Lib "SPRVBX30.VBX" (SS As Control, ByVal Row As Long, ByVal Twips As Long, fRowHeight As Single)
