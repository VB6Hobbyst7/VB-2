VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSADDSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const CONST_PHERE = "1"
Const CONST_HOSFG = "1"
Const CONST_STSCD = "0"

Private mvarBldSrc      As String
Private mvarBldYY       As String
Private mvarBldNo       As String
Private mvarCompocd     As String
Private mvarVolumn      As String
Private mvarAbo         As String
Private mvarRh          As String
Private mvarPtId        As String
Private mvarExcuteID    As String
Private mvarAvailable   As String
Private mvarExpdt       As String
Private mvarRealDt      As String
Private mvarRealTm      As String
Private mvarDonorID     As String
Private mvarDonorAccdt  As String

Public Property Let Bldsrc(ByVal vData As String)
    mvarBldSrc = vData
End Property
Public Property Let Bldyy(ByVal vData As String)
    mvarBldYY = vData
End Property
Public Property Let Bldno(ByVal vData As String)
    mvarBldNo = vData
End Property
Public Property Let CompoCd(ByVal vData As String)
    mvarCompocd = vData
End Property
Public Property Let volumn(ByVal vData As String)
    mvarVolumn = vData
End Property
Public Property Let ABO(ByVal vData As String)
    mvarAbo = vData
End Property
Public Property Let RH(ByVal vData As String)
    mvarRh = vData
End Property
Public Property Let PtId(ByVal vData As String)
    mvarPtId = vData
End Property
Public Property Let ExcuteID(ByVal vData As String)
    mvarExcuteID = vData
End Property
Public Property Let available(ByVal vData As String)
    mvarAvailable = vData
End Property
Public Property Let expdt(ByVal vData As String)
    mvarExpdt = vData
End Property
Public Property Let RealDt(ByVal vData As String)
    mvarRealDt = vData
End Property
Public Property Let RealTm(ByVal vData As String)
    mvarRealTm = vData
End Property
Public Property Let Donorid(ByVal vData As String)
    mvarDonorID = vData
End Property
Public Property Let donoraccdt(ByVal vData As String)
    mvarDonorAccdt = vData
End Property
Public Function Get_PheresisCompoNm() As Recordset

'pheresis헌혈등록시 필요한 pheresis 검체만 Load하기위한 SQL문장
    Dim sSql As String
        
    sSql = " SELECT compocd,abbrnm,keepday" & _
           " FROM " & T_BBS006 & _
           " WHERE " & DBW("pherefg=", CONST_PHERE) & _
           " AND (expdt='' or expdt is null)"
    Set Get_PheresisCompoNm = New Recordset
    Get_PheresisCompoNm.Open sSql, dbconn
    
End Function

Public Function SetPheresisInsert401() As String

'----------------
'혈액입고내역작성
'----------------
    SetPheresisInsert401 = " insert into " & T_BBS401 & "(" & _
                    " bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg,pherefg,coldt,coltm,colid,available," & _
                    " expdt,exptm,entdt,enttm,entid,centercd,localcd,stscd,hosfg,splitoutfg,splitinfg," & _
                    " realexpdt,realexptm,expid,exprcvid,expbilldiv,exprsnrmk)" & _
                    "values (" & _
                    DBV("bldsrc          ", mvarBldSrc, 1) & DBV("bldyy     ", mvarBldYY, 1) & DBV("bldno  ", mvarBldNo, 1) & DBV("compocd", mvarCompocd, 1) & _
                    DBV("volumn              ", mvarVolumn, 1) & DBV("abo     ", mvarAbo, 1) & DBV("rh    ", mvarRh, 1) & DBV("ptid", mvarPtId, 1) & _
                    DBV("reserved               ", "0", 1) & DBV("autofg      ", "0", 1) & DBV("pherefg", CONST_PHERE, 1) & _
                    DBV("coldt           ", mvarRealDt, 1) & DBV("coltm", mvarRealTm, 1) & DBV("colid", mvarExcuteID, 1) & _
                    DBV("available    ", mvarAvailable, 1) & DBV("expdt ", mvarExpdt, 1) & DBV("exptm", mvarRealTm, 1) & _
                    DBV("entdt            ", mvarRealDt, 1) & DBV("enttm ", mvarRealTm, 1) & DBV("entid", mvarExcuteID, 1) & _
                    DBV("centercd", ObjSysInfo.buildingcd, 1) & DBV("localcd      ", "", 1) & DBV("stscd", CONST_STSCD, 1) & _
                    DBV("hosfg          ", CONST_HOSFG, 1) & DBV("splitoutfg   ", "", 1) & DBV("splitinfg", "", 1) & _
                    DBV("realexpdt               ", "", 1) & DBV("realexptm    ", "", 1) & DBV("expid", 0, 1) & _
                    DBV("exprcvid                 ", 0, 1) & DBV("expbilldiv   ", "", 1) & DBV("exprsnrmk", "") & ")"
End Function

Public Function SetPhereUpdate602() As String

'----------------
'헌혈접수내역수정
'----------------
    SetPhereUpdate602 = " update " & T_BBS602 & _
                        " set    " & DBW("donationdt", mvarRealDt, 3) & _
                                     DBW("bldsrc    ", mvarBldSrc, 3) & _
                                     DBW("bldyy     ", mvarBldYY, 3) & _
                                     DBW("bldno     ", mvarBldNo, 3) & _
                                     DBW("compocd   ", mvarCompocd, 3) & _
                                     DBW("volumn    ", mvarVolumn, 3) & _
                                     DBW("entfg     ", "1", 3) & _
                                     DBW("cancelfg  ", "0", 2) & _
                    " WHERE  " & _
                                     DBW("donorid", mvarDonorID, 2) & _
                           " AND " & DBW("donoraccdt", mvarDonorAccdt, 2)

End Function


Public Function SetPheresisUpdate603() As String

'----------------
'헌혈판정내역수정
'----------------
    SetPheresisUpdate603 = " update " & T_BBS603 & _
                           " set    " & DBW("stscd", DonorStatus.stsDonation, 2) & _
                           " WHERE  " & _
                                     DBW("donorid   ", mvarDonorID, 2) & _
                           " AND " & DBW("donoraccdt", mvarDonorAccdt, 2)
End Function

Public Function SetPheresisUpdate411() As String

'----------------
'헌혈증서내역수정
'----------------
    SetPheresisUpdate411 = " update " & T_BBS411 & _
                           " set  " & DBW("usedt", mvarRealDt, 3) & _
                                      DBW("useid", mvarExcuteID, 2) & _
                           " WHERE" & _
                                     DBW("bldsrc", mvarBldSrc, 2) & _
                           " AND " & DBW("bldyy", mvarBldYY, 2) & _
                           " AND " & DBW("bldno", mvarBldNo, 2)
End Function

