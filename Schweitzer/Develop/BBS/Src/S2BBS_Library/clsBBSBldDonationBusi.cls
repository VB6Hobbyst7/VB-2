VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSBldDonationBusi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public mDonorID As String
Public mDonorNm As String
Public mDOB     As String
Public mSEX     As String
Public mABO     As String
Public Mcnt     As String
Public mTotVol  As String
Public mSSN     As String

Public Function DonorFind(ByVal pDonorNm As String) As Boolean
    Dim objSearch   As clsBBSSQLStatement
    Dim Rs          As Recordset
    Dim strAge      As String
    Dim itmX        As Object
    Dim i           As Long
    
    Set objSearch = New clsBBSSQLStatement
    
    Set Rs = New Recordset
    Rs.Open objSearch.GetDonorMst(, pDonorNm), DBConn
    
    With Rs
        If .RecordCount < 1 Then
            DonorFind = False
        ElseIf .RecordCount = 1 Then
            '헌혈자에 대한 기본 정보를 보여준다.
            mDonorID = Rs.Fields("donorid").Value & ""
            mDonorNm = Rs.Fields("donornm").Value & ""
            If Trim(Rs.Fields("dob").Value & "") <> "" Then
                strAge = DateDiff("yyyy", Format(Rs.Fields("dob").Value & "", "####-##-##"), GetSystemDate)
                mDOB = Format(Rs.Fields("dob").Value & "", "####-##-##")
            End If
            mSEX = Rs.Fields("sex").Value & "" & "/" & strAge
            mABO = Rs.Fields("abo").Value & "" & Rs.Fields("rh").Value & ""
            Mcnt = Rs.Fields("cnt").Value & ""
            mTotVol = Rs.Fields("totvol").Value & ""

            If Mid(Format(Rs.Fields("ssn").Value & "", "######-#######"), 8, 1) = "1" Or Mid(Format(Rs.Fields("ssn").Value & "", "######-#######"), 8, 1) = "2" Then
                mSSN = Format(Rs.Fields("ssn").Value & "", "######-#######")
                mSSN = Replace(mSSN, "-", "")
            Else
                mSSN = Format(Rs.Fields("ssn").Value & "", "######-#######")
                mSSN = Replace(mSSN, "-", "")
            End If
            DonorFind = True
        Else
            With frmDonorFind
                .lvwPtList.ListItems.Clear
                For i = 1 To Rs.RecordCount
                    Set itmX = .lvwPtList.ListItems.Add(, , Rs.Fields("donorid").Value & "")
                    itmX.SubItems(1) = Rs.Fields("donornm").Value & ""
                    itmX.SubItems(2) = Format(Rs.Fields("ssn").Value & "", "######-#######")
                    itmX.SubItems(3) = Format(Rs.Fields("dob").Value & "", "####-##-##")
                    itmX.SubItems(4) = Rs.Fields("sex").Value & ""
                    itmX.SubItems(5) = Rs.Fields("abo").Value & "" & Rs.Fields("rh").Value & ""
                    itmX.SubItems(6) = Rs.Fields("cnt").Value & ""
                    itmX.SubItems(7) = Rs.Fields("totvol").Value & ""
                    
                    Rs.MoveNext
                Next i
                
                .Hide
                .Show vbModal
                
                If .isSelect = True Then
                    mDonorID = .Donorid
                    mDonorNm = .donornm
                    mDOB = .dob
                    mSEX = .sex
                    mABO = .ABO
                    Mcnt = .cnt
                    mTotVol = .totvol
                    mSSN = .ssn
                    DonorFind = True
                Else
                    DonorFind = False
                End If
                
            End With
            
            Set frmDonorFind = Nothing
            
        End If
    End With
    Set Rs = Nothing
    Set objSearch = Nothing
End Function

'Public Function DonorBldList(ByRef objList As Object, ByVal pDonorId As String, ByVal pDonorDt As String) As Boolean
'    Dim Rs          As RECORDSET
'    Dim objSearch   As clsBBSSQLStatement
'
'    Set objSearch = New clsBBSSQLStatement
'
'    Set Rs = OpenRecordSet(objSearch.GetDonorBldList(pDonorId, pDonorDt))
'
'    objList.Clear
'    With Rs
'        If .RecordCount < 1 Then
'            DonorBldList = False
'        Else
'            While Not .EOF
'                objList.AddItem Format("" & Rs.Fields("bldsrc").Value & "-" & _
'                                     Rs.Fields("bldyy").Value & "-" & _
'                                     Rs.Fields("bldno").Value, "!" & String(16, "@")) & _
'                                     Rs.Fields("componm").Value
'                .MoveNext
'            Wend
'        End If
'    End With
'
'    Rs.RsClose
'    Set Rs = Nothing
'    Set objSearch = Nothing
'End Function


