Attribute VB_Name = "ModuleDB"
Option Explicit

Public db           As Database
Public dbcomm       As Database
Public tbcomm       As Recordset
Public dbcode       As Database
Public tbcode       As Recordset
Public dbcm         As Database
Public tbcm         As Recordset
Public dbrp         As Database


Sub CreateOrOpen_db(strmmdd As String)

    Dim db As Database
    Dim newtd As New TableDef
    Dim newtd2 As New TableDef
    Dim newindex As New Index
    Dim newindex2 As New Index
    Dim seqnoNF As New Field
    Dim slipnoNF As New Field
    Dim regchkNF As New Field
    Dim seqno2NF As New Field
    Dim tcodeNF As New Field
    Dim tresultNF As New Field
    Dim ObjectNF As New Field
    Dim i As Integer
    
    
    'filename = App.Path
    If Right(Filename, 1) <> "\" Then
        Filename = Filename & "\"
    End If
    
    If ifFileExists(Filename & "comm\" & strmmdd + ".mdb") Then
        Set db = OpenDatabase(Filename & "comm\" & strmmdd + ".mdb")
        Set identb = db.OpenRecordset("sp_identify")
        Set resulttb = db.OpenRecordset("sp_result")
        identb.Close
        resulttb.Close
        db.Close
        Exit Sub
    End If
        
        
Set db = CreateDatabase(Filename & "comm\" & strmmdd + ".mdb", dbLangGeneral)
    
    Set newtd = db.CreateTableDef("sp_identify")
            
            Set seqnoNF = newtd.CreateField("Seq_No", dbText, 4)
            newtd.Fields.Append seqnoNF
            
            Set slipnoNF = newtd.CreateField("Slip_No", dbText, 30)
            newtd.Fields.Append slipnoNF
            
            Set regchkNF = newtd.CreateField("ChkResult", dbText, 5)
            newtd.Fields.Append regchkNF
                                    
'''            If FieldAddIdenTBFlag <> 0 Then
'''                For i = 1 To FieldAddIdenTBFlag
'''                    'Set ObjectNF = IdTBNField(i)
'''                    Set ObjectNF = newtd.CreateField(IdTBFieldName, dbText, IdTBFieldDig)
'''                    newtd.Fields.Append ObjectNF
'''                Next
'''            End If
                                    
            Set newindex = newtd.CreateIndex("Primarykey")
            Set seqnoNF = newindex.CreateField("Seq_No")
            newindex.Primary = True
            newindex.Fields.Append seqnoNF
            newtd.Indexes.Append newindex

            Set newindex2 = newtd.CreateIndex("slip_no")
            Set slipnoNF = newindex.CreateField("slip_no")
            newindex2.Fields.Append slipnoNF
            newtd.Indexes.Append newindex2

    db.TableDefs.Append newtd

    
    Set newtd2 = db.CreateTableDef("sp_result")
    
            Set seqno2NF = newtd2.CreateField("Seq_No", dbText, 4)
            newtd2.Fields.Append seqno2NF
            
            Set tcodeNF = newtd2.CreateField("TestCode", dbText, 4)
            newtd2.Fields.Append tcodeNF
            
            Set tresultNF = newtd2.CreateField("TestResult", dbText, 15)
            newtd2.Fields.Append tresultNF
            
            Set newindex2 = newtd2.CreateIndex("Primarykey")
            
            Set seqno2NF = newindex2.CreateField("Seq_No")
            Set tcodeNF = newindex2.CreateField("TestCode")
            newindex2.Primary = True
            newindex2.Fields.Append seqno2NF
            newindex2.Fields.Append tcodeNF
            newtd2.Indexes.Append newindex2
            
            Set newindex = newtd2.CreateIndex("Seq_No")
            Set seqnoNF = newindex2.CreateField("seq_no")
            newindex.Fields.Append seqnoNF
            newtd2.Indexes.Append newindex
            
    db.TableDefs.Append newtd2
             
db.Close

End Sub


