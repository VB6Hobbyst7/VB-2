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

Public SemiDb As String
Sub Create_Code_DB()


    Dim db          As Database
    Dim Tb          As New TableDef
    Dim idx         As New Index
    Dim Fd(1 To 3)  As New Field
 
    
    filename = App.Path
    If Right(filename, 1) <> "\" Then
        filename = filename & "\"
    End If
    
    If ifFileExists(filename & "clinic\setcode.mdb") Then Exit Sub
        
    Set db = CreateDatabase(filename & "clinic\setcode.mdb", dbLangGeneral)
    
    Set Tb = db.CreateTableDef("cdtable")
            
            Set Fd(1) = Tb.CreateField("eqipNo", dbText, 2)
            Tb.Fields.Append Fd(1)
            
            Set Fd(2) = Tb.CreateField("Name", dbText, 10)
            Tb.Fields.Append Fd(2)
            
            Set Fd(3) = Tb.CreateField("Code", dbText, 20)
            Tb.Fields.Append Fd(3)
            
                                   
            Set idx = Tb.CreateIndex("Primarykey")
            Set Fd(1) = idx.CreateField("eqipNo")
            idx.Primary = True
            idx.Fields.Append Fd(1)
            Tb.Indexes.Append idx

    db.TableDefs.Append Tb
             
db.Close

End Sub

Sub CreateOrOpen_db(strmmdd As String)

    Dim db As Database
    Dim newTb1, newTb2, newTb3  As New TableDef
    Dim newIdx1, newIdx2, newIdx3   As New Index
    Dim newFd(1 To 7)   As New Field
    
    Dim seqnoNF As New Field
    Dim slipnoNF As New Field
    Dim regchkNF As New Field
    Dim seqno2NF As New Field
    Dim tcodeNF As New Field
    Dim tresultNF As New Field
    Dim ObjectNF As New Field
    Dim i As Integer
    
    'filename = App.Path
    If Right(filename, 1) <> "\" Then
        filename = filename & "\"
    End If
    
    If ifFileExists(filename & "comm\" & strmmdd + ".mdb") Then
        Set db = OpenDatabase(filename & "comm\" & strmmdd + ".mdb")
        Set identb = db.OpenRecordset("sp_identify")
        Set resulttb = db.OpenRecordset("sp_result")
        identb.Close
        resulttb.Close
        db.Close
        Exit Sub
    End If
        
    Set db = CreateDatabase(filename & "comm\" & strmmdd + ".mdb", dbLangGeneral)
    
    Set newTb1 = db.CreateTableDef("sp_identify")
            
            Set newFd(1) = newTb1.CreateField("Seq_No", dbText, 4)
            newTb1.Fields.Append newFd(1)
            
            Set newFd(2) = newTb1.CreateField("Slip_No", dbText, 30)
            newTb1.Fields.Append newFd(2)
            
            Set newFd(3) = newTb1.CreateField("ChkResult", dbText, 5)
            newTb1.Fields.Append newFd(3)
                                    
            Set newIdx1 = newTb1.CreateIndex("Primarykey")
            Set newFd(1) = newIdx1.CreateField("Seq_No")
            newIdx1.Primary = True
            newIdx1.Fields.Append newFd(1)
            newTb1.Indexes.Append newIdx1

            Set newIdx2 = newTb1.CreateIndex("slip_no")
            Set newFd(2) = newIdx2.CreateField("slip_no")
            newIdx2.Fields.Append newFd(2)
            newTb1.Indexes.Append newIdx2

    db.TableDefs.Append newTb1

    
    Set newTb2 = db.CreateTableDef("sp_result")
    
            Set newFd(1) = newTb2.CreateField("Seq_No", dbText, 4)
            newTb2.Fields.Append newFd(1)
            
            Set newFd(2) = newTb2.CreateField("TestCode", dbText, 4)
            newTb2.Fields.Append newFd(2)
            
            Set newFd(3) = newTb2.CreateField("TestResult", dbText, 15)
            newTb2.Fields.Append newFd(3)
            
            Set newIdx2 = newTb2.CreateIndex("Primarykey")
            
            Set newFd(1) = newIdx2.CreateField("Seq_No")
            Set newFd(2) = newIdx2.CreateField("TestCode")
            newIdx2.Primary = True
            newIdx2.Fields.Append newFd(1)
            newIdx2.Fields.Append newFd(2)
            newTb2.Indexes.Append newIdx2
            
            Set newIdx1 = newTb2.CreateIndex("Seq_No")
            Set newFd(1) = newIdx1.CreateField("seq_no")
            newIdx1.Fields.Append newFd(1)
            newTb2.Indexes.Append newIdx1
            
    db.TableDefs.Append newTb2
             
    Set newTb3 = db.CreateTableDef("Temp_Tb")
    
            Set newFd(1) = newTb3.CreateField("RecDate", dbText, 8)
            newTb3.Fields.Append newFd(1)
            
            Set newFd(2) = newTb3.CreateField("SampleNo", dbText, 20)
            newTb3.Fields.Append newFd(2)
            
            Set newFd(3) = newTb3.CreateField("TestCode", dbText, 4)
            newTb3.Fields.Append newFd(3)
            
            Set newFd(4) = newTb3.CreateField("TestResult", dbText, 15)
            newTb3.Fields.Append newFd(4)
            
            Set newIdx1 = newTb3.CreateIndex("Primarykey")
            
            Set newFd(1) = newIdx1.CreateField("RecDate")
            Set newFd(2) = newIdx1.CreateField("SampleNo")
            Set newFd(3) = newIdx1.CreateField("TestCode")
            newIdx1.Primary = True
            newIdx1.Fields.Append newFd(1)
            newIdx1.Fields.Append newFd(2)
            newTb3.Indexes.Append newIdx1
            
    db.TableDefs.Append newTb3

db.Close

End Sub


