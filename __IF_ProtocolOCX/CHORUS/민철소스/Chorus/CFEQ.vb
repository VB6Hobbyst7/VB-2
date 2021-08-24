Imports ifcm
Imports System.IO

Public Class CFEQ
    Inherits IFBASE0101.CFBASE

    Private msWkBuf As String = ""
    Private msRcvBuf As String = ""
    Private msSvrBuf As String = ""

    Private msRcvState As String = ""
    Private miLen As Integer = 0
    Private miCnt As Integer = 0

    Dim miFileNo As Integer = FreeFile()
    Dim msRstFile As String = "Chorus_Result.txt"

    Public Overrides Sub sbPhaseCfg_Protocol(ByVal rswkbuf As String)
        Dim sFn As String = "Sub sbPhaseCfg_Protocol"

        Try
            With New IFAscii
                Dim sWkDat As String = ""

                msWkBuf = rswkbuf

                For i As Integer = 0 To msWkBuf.Length - 1
                    sWkDat = msWkBuf.Substring(i, 1)

                    If msRcvState = "" Then
                        Select Case sWkDat
                            Case .gsSTX
                                msSvrBuf = ""
                                msRcvState = "R1"
                        End Select

                    ElseIf msRcvState = "R1" Then
                        miLen = Asc(sWkDat)
                        msRcvState = "R2"

                    ElseIf msRcvState = "R2" Then
                        msRcvBuf += sWkDat
                        miCnt += 1

                        If miLen + 1 = miCnt Then
                            sbEdit_Data()

                            MyBase.EqOutput(New IFAscii().gsSTX + New IFAscii().gsSOH + New IFAscii().gsEOT + New IFAscii().gsENQ)

                            msRcvBuf = ""
                            msRcvState = ""
                            miLen = 0
                            miCnt = 0
                        End If
                    End If
                Next
            End With

        Catch ex As Exception
            MyBase.Eq_SetErr(sFn + ":" + "CFEQ - " + ex.Message)

        End Try
    End Sub

    Private Sub sbEdit_Data()
        Dim sFn As String = "Sub sbEdit_Data"

        Try
            With New IFAscii

                Dim sID As String = ""
                Dim sEqSeq As String = ""
                Dim sRack As String = ""
                Dim sPos As String = ""
                Dim sTIFCd As String = ""
                Dim sTIFRstCd As String = ""
                Dim sRst1 As String = ""
                Dim sRst2 As String = ""
                Dim sUnit As String = ""
                Dim sTRst1 As String = ""
                Dim sTRst2 As String = ""
                Dim sTUnit As String = ""
                Dim sTFlag As String = ""
                Dim iRstCnt As Integer = 0
                Dim sCtrID As String = ""

                If msRcvBuf = "" Or msRcvBuf = New IFAscii().gsENQ + "CD" Or msRcvBuf.Length <= 3 Then
                    Exit Sub
                End If

                Dim aData() As String = msRcvBuf.Split(Convert.ToChar(0))
                Dim aData2() As String
                Dim iCnt As Integer = 0

                For i As Integer = 0 To aData.Length - 1
                    If aData(i) <> "" Then
                        ReDim Preserve aData2(iCnt)
                        aData2(iCnt) = aData(i)
                        iCnt += 1
                    End If
                Next

                sID = aData2(0).Substring(1, 9).Trim()

                sTIFCd = aData2(1).Trim()

                sRst2 = aData2(2).Substring(0, 1).Trim()
                sRst1 = aData2(2).Substring(1, aData2(2).Length - 1).Trim()

                If aData2.Length > 4 Then
                    sUnit = aData2(3).Trim()
                End If

                If sRst1.StartsWith(".") Then
                    sRst1 = "0" + sRst1
                End If

                iRstCnt += 1

                sTIFRstCd += sTIFCd + .gsFLD
                sTRst1 += sRst1 + .gsFLD
                sTRst2 += sRst2 + .gsFLD
                sTUnit += sUnit + .gsFLD
                sTFlag += "" + .gsFLD

                If sCtrID = "" Then
                    MyBase.Eq_ReceiveResult(sID, sRack, sPos, sEqSeq, iRstCnt, sTIFRstCd, sTRst1, sTRst2, sTUnit, sTFlag)
                Else
                    MyBase.Eq_ReceiveResult(sID, sRack, sPos, sEqSeq, iRstCnt, sTIFRstCd, sTRst1, sTRst2, sTUnit, sTFlag, sCtrID)
                End If

                sID = "" : sRack = "" : sPos = "" : sEqSeq = "" : iRstCnt = 0 : sTIFRstCd = "" : sTRst1 = "" : sTRst2 = "" : sTUnit = "" : sTFlag = ""

            End With

        Catch ex As Exception
            MyBase.Eq_SetErr(sFn + ":" + "CFEQ - " + ex.Message)

        End Try
    End Sub

End Class
