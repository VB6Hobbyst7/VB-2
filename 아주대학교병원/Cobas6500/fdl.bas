Attribute VB_Name = "fdl"
'/* ----------------------- usrinc/fbuf.h --------------------- */
'/*                                                             */
'/*              Copyright (c) 2000 Tmax Soft Co., Ltd          */
'/*                   All Rights Reserved                       */
'/*                                                             */
'/* ----------------------------------------------------------- */

Const SDL_CHAR = 1
Const SDL_SHORT = 2
Const SDL_INT = 3
Const SDL_LONG = 4
Const SDL_FLOAT = 5
Const SDL_DOUBLE = 6
Const SDL_STRING = 7
Const SDL_CARRAY = 8

'/* field types */
Const FB_CHAR = SDL_CHAR
Const FB_SHORT = SDL_SHORT
Const FB_INT = SDL_INT
Const FB_LONG = SDL_LONG
Const FB_FLOAT = SDL_FLOAT
Const FB_DOUBLE = SDL_DOUBLE
Const FB_STRING = SDL_STRING
Const FB_CARRAY = SDL_CARRAY

Const BADFLDKEY = 0
Const FIRSTFLDKEY = 0

'/* ----- fb op mode ----- */
Const FBMOVE_MODE = 1
Const FBCOPY_MODE = 2
Const FBCOMP_MODE = 3
Const FBCONCAT_MODE = 4
Const FBJOIN_MODE = 5
Const FBOJOIN_MODE = 6
Const FBUPDATE_MODE = 7

'/* ------- fberror code ----- */
Const FBEBADFB = 3
Const FBEINVAL = 4
Const FBELIMIT = 5
Const FBENOENT = 6
Const FBEOS = 7
Const FBEBADFLD = 8
Const FBEPROTO = 9
Const FBENOSPACE = 10
Const FBEMALLOC = 11
Const FBESYSTEM = 12
Const FBETYPE = 13
Const FBEMATCH = 14
Const FBEBADSTRUCT = 15
Const FBEMAXNO = 19


Declare Function getfberrno Lib "TMAX4GL.DLL" () As Long
Declare Function fbget Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, pbuffer As Any, Fieldlen As Long) As Long
Declare Function fbput Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, pbuffer As Any, ByVal Fieldlen As Long) As Long
Declare Function fbinsert Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long, pbuffer As Any, ByVal Fieldlen As Long) As Long
Declare Function fbupdate Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long, pbuffer As Any, ByVal Fieldlen As Long) As Long
Declare Function fbdelete Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long) As Long
Declare Function fbgetval Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long, Fieldlen As Long) As Long
Declare Function fbgetnth Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, pbuffer As Any, ByVal Fieldlen As Long) As Long
Declare Function fbkeyoccur Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long) As Long
Declare Function fbgetf Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, pbuffer As Any, Fieldlen As Long, Pos As Long) As Long

Declare Function fbdelall Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long) As Long
Declare Function fbfldcount Lib "TMAX4GL.DLL" (ByVal pFBUF As Long) As Long
Declare Function fbispres Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long) As Long
Declare Function fbgetvals Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long) As Long
Declare Function fbgetvali Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long) As Long

Declare Function fbtypecvt Lib "TMAX4GL.DLL" (tolen As Long, ByVal totype As Long, fromval As Any, ByVal fromtype As Long, ByVal fromlen As Long) As Long
Declare Function fbputt Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, pbuffer As Any, ByVal Fieldlen As Long, ByVal ftype As Long) As Long
Declare Function fbgetvalt Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long, Fieldlen As Long, ByVal totype As Long) As Long
Declare Function fbgetntht Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, pbuffer As Any, ByVal Fieldlen As Long, ByVal fromtype As Long) As Long

Declare Function fbget_fldkey Lib "TMAX4GL.DLL" (ByVal Fname As String) As Long
Declare Function fbget_fldname Lib "TMAX4GL.DLL" (ByVal fieldid As Long) As Long
Declare Function fbget_fldno Lib "TMAX4GL.DLL" (ByVal fieldid As Long) As Long
Declare Function fbget_fldtype Lib "TMAX4GL.DLL" (ByVal fieldid As Long) As Long
Declare Function fbget_strfldtype Lib "TMAX4GL.DLL" (ByVal fieldid As Long) As Long
Declare Function fbmake_fldkey Lib "TMAX4GL.DLL" (ByVal ftype As Long, ByVal no As Long) As Long
Declare Function fbnmkey_unload Lib "TMAX4GL.DLL" () As Long
Declare Function fbkeynm_unload Lib "TMAX4GL.DLL" () As Long

Declare Function fbisfbuf Lib "TMAX4GL.DLL" (ByVal pFBUF As Long) As Long
Declare Function fbcalcsize Lib "TMAX4GL.DLL" (ByVal count As Long, ByVal datalen As Long) As Long
Declare Function fbinit Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal buflen As Long) As Long
Declare Function fballoc Lib "TMAX4GL.DLL" (ByVal count As Long, ByVal buflen As Long) As Long
Declare Function fbfree Lib "TMAX4GL.DLL" (ByVal pFBUF As Long) As Long
Declare Function fbget_fbsize Lib "TMAX4GL.DLL" (ByVal pFBUF As Long) As Long
Declare Function fbget_unused Lib "TMAX4GL.DLL" (ByVal pFBUF As Long) As Long
Declare Function fbget_used Lib "TMAX4GL.DLL" (ByVal pFBUF As Long) As Long
Declare Function fbrealloc Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal count As Long, ByVal nlen As Long) As Long

Declare Function fbbufop Lib "TMAX4GL.DLL" (ByVal pdFBUF As Long, ByVal psFBUF As Long, ByVal mode As Long) As Long
Declare Function fbbufop_proj Lib "TMAX4GL.DLL" (ByVal pdFBUF As Long, ByVal psFBUF As Long, fieldid As Long) As Long

Declare Function fbchg_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal nth As Long, pbuffer As Any, ByVal Fieldlen As Long) As Long
Declare Function fbdelall_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, fieldid As Long) As Long
Declare Function fbgetval_last_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, fieldocc As Long, Fieldlen As Long) As Long
Declare Function fbget_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long, pbuffer As Any, maxlen As Long) As Long
Declare Function fbgetalloc_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long, extralen As Long) As Long
Declare Function fbgetlast_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, fieldocc As Long, pbuffer As Any, maxlen As Long) As Long
Declare Function fbnext_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, fieldid As Long, nth As Long, pbuffer As Any, Fieldlen As Long) As Long
Declare Function fbgetvals_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long) As Long
Declare Function fbgetvall_tu Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long) As Long
Declare Function fbchg_tut Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal nth As Long, pbuffer As Any, ByVal Fieldlen As Long, ByVal ftype As Long) As Long
Declare Function fbget_tut Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal nth As Long, pbuffer As Any, Fieldlen As Long, ByVal ftype As Long) As Long
Declare Function fbgetalloc_tut Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal nth As Long, ByVal totype As Long, extralen As Long) As Long
Declare Function fbgetlen Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, ByVal fieldid As Long, ByVal fieldocc As Long) As Long

Declare Function fbftos Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, cstruct As Long, stname As Long) As Long
Declare Function fbstof Lib "TMAX4GL.DLL" (ByVal pFBUF As Long, cstruct As Long, ByVal mode As Long, stname As Long) As Long
Declare Function fbsnull Lib "TMAX4GL.DLL" (cstruct As Long, cname As Long, ByVal fieldocc As Long, stname As Long) As Long
Declare Function fbstinit Lib "TMAX4GL.DLL" (cstruct As Long, stname As Long) As Long
Declare Function fbstelinit Lib "TMAX4GL.DLL" (cstruct As Long, cname As Long, stname As Long) As Long

Declare Function fbstrerror Lib "TMAX4GL.DLL" (ByVal err_no As Long) As Long
