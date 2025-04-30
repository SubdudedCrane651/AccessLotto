Option Compare Database

Public Sub DoJson()
    Call Lotto649_To_Json
    Call LottoMax_To_Json
    Call Grande_Vie_To_Json
    Call Tout_ou_Rien_To_Json
    Call DoAllLotto_Json
    
    Dim i
    For i = 1 To 3 ' Loop 3 times.
        Beep ' Sound a tone.
    Next i
    
    MsgBox ("Completed")
End Sub

Public Sub Lotto649_To_Json()
    
    Dim dbs As Database, rst As DAO.Recordset

    Set dbs = CurrentDb
    'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")
    
    'Command0.Caption = "Please Wait..."
    
    'On Error Resume Next

    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from 649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7 from 649Drawings ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\Lotto649.json", True, True)
    Fileout.WriteLine "["
    
    Dim lngCount As Long
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P1" & Chr(34) & ": " & Chr(34) & rst("P1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P2" & Chr(34) & ": " & Chr(34) & rst("P2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P3" & Chr(34) & ": " & Chr(34) & rst("P3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P4" & Chr(34) & ": " & Chr(34) & rst("P4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P5" & Chr(34) & ": " & Chr(34) & rst("P5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P6" & Chr(34) & ": " & Chr(34) & rst("P6") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P7" & Chr(34) & ": " & Chr(34) & rst("P7") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "]"
    rst.Close
    rst1.Close
    dbs.Close
    Fileout.Close
    Close #1
End Sub

Public Sub LottoMax_To_Json()
    
    Dim dbs As Database, rst As DAO.Recordset

    Set dbs = CurrentDb
    'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")
    
    'Command0.Caption = "Please Wait..."
    
    'On Error Resume Next

    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from LottoMax")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7,P8 from LottoMax ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\LottoMax.json", True, True)
    Fileout.WriteLine "["
    
    Dim lngCount As Long
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P1" & Chr(34) & ": " & Chr(34) & rst("P1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P2" & Chr(34) & ": " & Chr(34) & rst("P2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P3" & Chr(34) & ": " & Chr(34) & rst("P3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P4" & Chr(34) & ": " & Chr(34) & rst("P4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P5" & Chr(34) & ": " & Chr(34) & rst("P5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P6" & Chr(34) & ": " & Chr(34) & rst("P6") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P7" & Chr(34) & ": " & Chr(34) & rst("P7") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P8" & Chr(34) & ": " & Chr(34) & rst("P8") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "]"
    rst.Close
    rst1.Close
    dbs.Close
    Fileout.Close
    Close #1
End Sub

Public Sub Grande_Vie_To_Json()
    
    Dim dbs As Database, rst As DAO.Recordset

    Set dbs = CurrentDb
    'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")
    
    'Command0.Caption = "Please Wait..."
    
    'On Error Resume Next

    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from Grande_Vie")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,gn from Grande_Vie ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\Grande_Vie.json", True, True)
    Fileout.WriteLine "["
    
    Dim lngCount As Long
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p1" & Chr(34) & ": " & Chr(34) & rst("p1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p2" & Chr(34) & ": " & Chr(34) & rst("p2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p3" & Chr(34) & ": " & Chr(34) & rst("p3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p4" & Chr(34) & ": " & Chr(34) & rst("p4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p5" & Chr(34) & ": " & Chr(34) & rst("p5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "gn" & Chr(34) & ": " & Chr(34) & rst("gn") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "]"
    rst.Close
    rst1.Close
    dbs.Close
    Fileout.Close
    Close #1
End Sub

Public Sub Tout_ou_Rien_To_Json()
    
    Dim dbs As Database, rst As DAO.Recordset

    Set dbs = CurrentDb
    'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")
    
    'Command0.Caption = "Please Wait..."
    
    'On Error Resume Next

    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from ToutouRien")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12 from ToutouRien ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\ToutouRien.json", True, True)
    Fileout.WriteLine "["
    
    Dim lngCount As Long
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p1" & Chr(34) & ": " & Chr(34) & rst("p1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p2" & Chr(34) & ": " & Chr(34) & rst("p2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p3" & Chr(34) & ": " & Chr(34) & rst("p3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p4" & Chr(34) & ": " & Chr(34) & rst("p4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p5" & Chr(34) & ": " & Chr(34) & rst("p5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p6" & Chr(34) & ": " & Chr(34) & rst("p6") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p7" & Chr(34) & ": " & Chr(34) & rst("p7") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p8" & Chr(34) & ": " & Chr(34) & rst("p8") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p9" & Chr(34) & ": " & Chr(34) & rst("p9") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p10" & Chr(34) & ": " & Chr(34) & rst("p10") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p11" & Chr(34) & ": " & Chr(34) & rst("p11") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p12" & Chr(34) & ": " & Chr(34) & rst("p12") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "]"
    rst.Close
    rst1.Close
    dbs.Close
    Fileout.Close
    Close #1
End Sub

Public Function writeOut(cText As String, file As String) As Integer
     On Error GoTo errHandler
    Dim fsT As Object
    Dim tFilePath As String
    
    tFilePath = file
    
    'Create Stream object
    Set fsT = CreateObject("ADODB.Stream")
    
    'Specify stream type - we want To save text/string data.
    fsT.Type = 2
    
    'Specify charset For the source text data.
    fsT.Charset = "utf-8"
    
    'Open the stream And write binary data To the object
    fsT.Open
    fsT.writetext cText
    
    'Save binary data To disk
    fsT.SaveToFile tFilePath, 2
    
    GoTo finish
    
errHandler:
    MsgBox (Err.Description)
    writeOut = 0
    Exit Function
    
finish:
    writeOut = 1
End Function


Public Sub DoAllLotto_Json()
    
    Dim dbs As Database, rst As DAO.Recordset

    Set dbs = CurrentDb
    'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")
    
    'Command0.Caption = "Please Wait..."
    
    'On Error Resume Next

    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from 649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7 from 649Drawings ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\Lotto.json", True, True)
    Fileout.WriteLine "{""Lotto649"":["
    
    Dim lngCount As Long
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P1" & Chr(34) & ": " & Chr(34) & rst("P1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P2" & Chr(34) & ": " & Chr(34) & rst("P2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P3" & Chr(34) & ": " & Chr(34) & rst("P3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P4" & Chr(34) & ": " & Chr(34) & rst("P4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P5" & Chr(34) & ": " & Chr(34) & rst("P5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P6" & Chr(34) & ": " & Chr(34) & rst("P6") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P7" & Chr(34) & ": " & Chr(34) & rst("P7") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "],"
    
    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from LottoMax")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7,P8 from LottoMax ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Fileout.WriteLine """LottoMax"":["
    
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P1" & Chr(34) & ": " & Chr(34) & rst("P1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P2" & Chr(34) & ": " & Chr(34) & rst("P2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P3" & Chr(34) & ": " & Chr(34) & rst("P3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P4" & Chr(34) & ": " & Chr(34) & rst("P4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P5" & Chr(34) & ": " & Chr(34) & rst("P5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P6" & Chr(34) & ": " & Chr(34) & rst("P6") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P7" & Chr(34) & ": " & Chr(34) & rst("P7") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "P8" & Chr(34) & ": " & Chr(34) & rst("P8") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "],"

    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from Grande_Vie")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,gn from Grande_Vie ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Fileout.WriteLine """Grande_Vie"":["
    
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p1" & Chr(34) & ": " & Chr(34) & rst("p1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p2" & Chr(34) & ": " & Chr(34) & rst("p2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p3" & Chr(34) & ": " & Chr(34) & rst("p3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p4" & Chr(34) & ": " & Chr(34) & rst("p4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p5" & Chr(34) & ": " & Chr(34) & rst("p5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "gn" & Chr(34) & ": " & Chr(34) & rst("gn") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "],"


    Set rst1 = dbs.OpenRecordset("SELECT count(ID) As cnt from ToutouRien")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12 from ToutouRien ORDER BY DrawDate ASC")
    'Dim myFile As String
    'myFile = "C:\Users\rchrd\Documents\Richard\Lotto649.json"
    
    'Open myFile For Append As #1
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Fileout.WriteLine """ToutouRien"":["
    
    lngCount = rst1.cnt
    
    count = 0
    
    Do While Not rst.EOF
        Fileout.WriteLine "{"
        Fileout.WriteLine Chr(34) & "Drawdate" & Chr(34) & ": " & Chr(34) & rst("DrawDate") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p1" & Chr(34) & ": " & Chr(34) & rst("p1") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p2" & Chr(34) & ": " & Chr(34) & rst("p2") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p3" & Chr(34) & ": " & Chr(34) & rst("p3") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p4" & Chr(34) & ": " & Chr(34) & rst("p4") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p5" & Chr(34) & ": " & Chr(34) & rst("p5") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p6" & Chr(34) & ": " & Chr(34) & rst("p6") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p7" & Chr(34) & ": " & Chr(34) & rst("p7") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p8" & Chr(34) & ": " & Chr(34) & rst("p8") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p9" & Chr(34) & ": " & Chr(34) & rst("p9") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p10" & Chr(34) & ": " & Chr(34) & rst("p10") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p11" & Chr(34) & ": " & Chr(34) & rst("p11") & Chr(34) & ","
        Fileout.WriteLine Chr(34) & "p12" & Chr(34) & ": " & Chr(34) & rst("p12") & Chr(34)
        count = count + 1
        If count = lngCount Then
            Fileout.WriteLine "}"
        Else
            Fileout.WriteLine "},"
        End If
        rst.MoveNext

    Loop
    Fileout.Write "]}"
    rst.Close
    rst1.Close
    dbs.Close
    Fileout.Close
    Close #1

End Sub