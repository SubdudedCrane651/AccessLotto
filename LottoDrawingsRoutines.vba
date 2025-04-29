Option Compare Database

Public P1, P2, P3, P4, P5, P6, P7, P8, GNumber, checkcount, Bonus, count1, count2, Index, Index2, LottoMaxIndex3(50), Index3(49), Index4(49), LottoMaxDrawings(50), Drawings(49), Drawingsjdp(49), BonusNumbersjdp(49), LottoMaxBonusNumbers(50), BonusNumbers(49), Index7(49), DrawingLottoMax(49), BonusNumberLottoMax(50), LottoMaxDateDifference(50), DateDifference(49), LottoMaxdatevar(50), datevar(49) As Integer
    
    Public DrawingsTout(24), Index3Tout(24), GN(49) As Integer
    Public DrawDate As Date
    
    Public StringVar As String
    
    Public pick6(6), pick7(7), pickjdp(6), pickjdpTout(12), pickGN(1)
   

Public Function Find649HotNumbersOrderByDrawings() As String

   Dim StringVar As String
   
 Find649HotNumbersOrderByNumber = ""
    
    Do649HotNumbersCalc
         
     'DoCreateReport
     'HotNumberReport
        
        For count1 = 1 To 49
        Drawings(count1) = Drawings(count1) + BonusNumbers(count1)
        StrSQL = "UPDATE 649_Hot_Numbers SET 649_Hot_Number=" & CStr(Index3(count1)) & ",Drawings=" & CStr(Drawings(count1)) _
        & ",Bonus=" & CStr(GN(count1)) _
        & ",Date_Differnce=" & CStr(datevar(count1)) _
        & " WHERE id=" & count1 & ";"
        'StrSQL = "INSERT INTO 649_Hot_Numbers (649_Hot_Number) VALUES (" & CStr(Count1) & ");"
            'StringVar = StringVar + CStr(Count1) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
            
                    DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True
        Next count1
        
        DoCmd.OpenReport "649 Hot Numbers Order by Drawings", acViewPreview
  
 End Function
 
 Public Function FindGrandeVieHotNumbersOrderByDrawings() As String

   Dim StringVar As String
   
 FindGrandeVieHotNumbersOrderByDrawings = ""
    
    DoGrandVieHotNumbersCalc
         
     'DoCreateReport
     'HotNumberReport
        
        For count1 = 1 To 49
        Drawings(count1) = Drawings(count1)
        StrSQL = "UPDATE Grande_Vie_Hot_Numbers SET Grande_Vie_Hot_Number=" & CStr(Index3(count1)) & ",Drawings=" & CStr(Drawings(count1)) _
        & ",GN=" & CStr(GN(count1)) _
        & ",Date_Differnce=" & CStr(datevar(count1)) _
        & " WHERE id=" & count1 & ";"
        'StrSQL = "INSERT INTO 649_Hot_Numbers (649_Hot_Number) VALUES (" & CStr(Count1) & ");"
            'StringVar = StringVar + CStr(Count1) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
            
                    DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True
        Next count1
        
        DoCmd.OpenReport "Grande Vie Hot Numbers", acViewPreview
  
 End Function
 
 Public Function Find649HotNumbersOrderByNumbers() As String

Dim StringVar As String
   
 Find649HotNumbersOrderByNumber = ""
    
    Do649HotNumbersCalc
         
     'DoCreateReport
     'HotNumberReport
     
     For Index2 = 1 To (UBound(Drawings) - 1)
            For Index = 1 To (UBound(Drawings) - 1)
                If Drawings(Index) > Drawings(Index + 1) Then
                    temp = Drawings(Index)
                    Drawings(Index) = Drawings(Index + 1)
                    Drawings(Index + 1) = temp
                    temp2 = Index3(Index)
                    Index3(Index) = Index3(Index + 1)
                    Index3(Index + 1) = temp2
                    temp3 = BonusNumbers(Index)
                    BonusNumbers(Index) = BonusNumbers(Index + 1)
                    BonusNumbers(Index + 1) = temp3
                    temp4 = datevar(Index)
                    datevar(Index) = datevar(Index + 1)
                    datevar(Index + 1) = temp4
                End If
            Next
    Next
        
        For count1 = 1 To 49
        Drawings(count1) = Drawings(count1) + BonusNumbers(count1)
        StrSQL = "UPDATE 649_Hot_Numbers SET 649_Hot_Number=" & CStr(Index3(count1)) & ",Drawings=" & CStr(Drawings(count1)) _
        & ",Bonus=" & CStr(GN(count1)) _
        & ",Date_Differnce=" & CStr(datevar(count1)) _
        & " WHERE id=" & count1 & ";"
        'StrSQL = "INSERT INTO 649_Hot_Numbers (649_Hot_Number) VALUES (" & CStr(Count1) & ");"
            'StringVar = StringVar + CStr(Count1) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
            
                    DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True
        Next count1
        
        DoCmd.OpenReport "649 Hot Numbers Order by Numbers", acViewPreview
  
 End Function
 
 Public Sub Do649HotNumbersCalc()

    Dim dbs As Database, rst As DAO.Recordset
    Dim rst2 As DAO.Recordset
    
     Set dbs = CurrentDb
     'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")

     'Command0.Caption = "Please Wait..."
        
        count1 = 0
        
        For count1 = 1 To 49
            Drawings(count1) = 0
            BonusNumbers(count1) = 0
            Index3(count1) = count1
            datevar(count1) = 0
        Next

        'On Error Resume Next
        
        Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7 from 649Drawings ORDER BY DrawDate ASC")
        
    Do While Not rst.EOF
    
    For count1 = 1 To 49

                DrawDate = rst("DrawDate")
                P1 = Int(rst("P1"))
                P2 = Int(rst("P2"))
                P3 = Int(rst("P3"))
                P4 = Int(rst("P4"))
                P5 = Int(rst("P5"))
                P6 = Int(rst("P6"))
                Bonus = Int(rst("P7"))
                If P1 = count1 Or P2 = count1 Or P3 = count1 Or P4 = count1 Or P5 = count1 Or P6 = count1 Then
                    Drawings(count1) = Drawings(count1) + 1
                    datevar(count1) = DateDiff("d", DrawDate, Now())
                    'DateVar(Count1) = Date - DrawDate
                End If
                Next count1
                
                For count2 = 1 To 49
                If Bonus = count2 Then
                    GN(count2) = GN(count2) + 1
                End If
                Next count2
            rst.MoveNext
        Loop
        
        Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing
    
    End Sub
    
    Public Sub DoGrandVieHotNumbersCalc()

    Dim dbs As Database, rst As DAO.Recordset
    Dim rst2 As DAO.Recordset
    
     Set dbs = CurrentDb
     'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")

     'Command0.Caption = "Please Wait..."
        
        count1 = 0
        
        For count1 = 1 To 49
            Drawings(count1) = 0
            GN(count1) = 0
            Index3(count1) = count1
            datevar(count1) = 0
        Next

        'On Error Resume Next
        
        Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,gn from Grande_Vie ORDER BY DrawDate ASC")
        
        count3 = 0
        
        rst.MoveFirst
        
    Do While Not rst.EOF
    
    For count1 = 1 To 49

                DrawDate = rst("DrawDate")
                P1 = Int(rst("p1"))
                P2 = Int(rst("p2"))
                P3 = Int(rst("p3"))
                P4 = Int(rst("p4"))
                P5 = Int(rst("p5"))
                GNumber = Int(rst("gn"))
                If P1 = count1 Or P2 = count1 Or P3 = count1 Or P4 = count1 Or P5 = count1 Then
                
                    Drawings(count1) = Drawings(count1) + 1
                    datevar(count1) = DateDiff("d", DrawDate, Now())
                    'DateVar(Count1) = Date - DrawDate
                End If
                Next count1
                                
                For count2 = 1 To 7
                If GNumber = count2 Then
                    GN(count2) = GN(count2) + 1
                End If
                Next count2

            rst.MoveNext
        
        count3 = count3 + 1
        
        Loop
        
        Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing
    
    End Sub
    
    Public Function Do649Dates(Start As Boolean) As Date
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Date
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7 from 649Drawings ORDER BY DrawDate ASC")
    
    If Start Then
    
    rst.MoveFirst
    
    Else
    
    rst.MoveLast
    
    End If
    
    RepDate = rst!DrawDate
    
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    Do649Dates = RepDate
    
    End Function
    
    Public Function Do649DatesCount() As Long
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Long
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7 from 649Drawings ORDER BY DrawDate ASC")
    
    Do649DatesCount = rst.RecordCount
       
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    End Function
    
    Public Function DoGrandeVieDates(Start As Boolean) As Date
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Date
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,gn from Grande_Vie ORDER BY DrawDate ASC")
    
    If Start Then
    
    rst.MoveFirst
    
    Else
    
    rst.MoveLast
    
    End If
    
    RepDate = rst!DrawDate
    
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    DoGrandeVieDates = RepDate
    
    End Function
    
    Public Function DoGrandeVieDatesCount() As Long
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Long
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,gn from Grande_Vie ORDER BY DrawDate ASC")
    
    DoGrandeVieDatesCount = rst.RecordCount
       
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    End Function
    
    Public Function FindLottoMaxHotNumbersOrderByNumbers() As String

   Dim StringVar As String
   
 FindLottoMaxHotNumbersOrderByNumber = ""
    
    DoLottoMaxHotNumbersCalc
         
     'DoCreateReport
     'HotNumberReport
        
        For Index2 = 1 To (UBound(LottoMaxDrawings) - 1)
            For Index = 1 To (UBound(LottoMaxDrawings) - 1)
                If LottoMaxDrawings(Index) > LottoMaxDrawings(Index + 1) Then
                    temp = LottoMaxDrawings(Index)
                    LottoMaxDrawings(Index) = LottoMaxDrawings(Index + 1)
                    LottoMaxDrawings(Index + 1) = temp
                    temp2 = LottoMaxIndex3(Index)
                    LottoMaxIndex3(Index) = LottoMaxIndex3(Index + 1)
                    LottoMaxIndex3(Index + 1) = temp2
                    temp3 = LottoMaxBonusNumbers(Index)
                    LottoMaxBonusNumbers(Index) = LottoMaxBonusNumbers(Index + 1)
                    LottoMaxBonusNumbers(Index + 1) = temp3
                    temp4 = LottoMaxdatevar(Index)
                    LottoMaxdatevar(Index) = LottoMaxdatevar(Index + 1)
                    LottoMaxdatevar(Index + 1) = temp4
                End If
            Next
    Next
        
        For count1 = 1 To 50
        LottoMaxDrawings(count1) = LottoMaxDrawings(count1) + LottoMaxBonusNumbers(count1)
        StrSQL = "UPDATE Lotto_Max_Hot_Numbers SET Lotto_Max_Hot_Number=" & CStr(LottoMaxIndex3(count1)) & ",Drawings=" & CStr(LottoMaxDrawings(count1)) _
        & ",Bonus=" & CStr(LottoMaxBonusNumbers(count1)) _
        & ",Date_Differnce=" & CStr(LottoMaxdatevar(count1)) _
        & " WHERE id=" & count1 & ";"
        'StrSQL = "INSERT INTO 649_Hot_Numbers (649_Hot_Number) VALUES (" & CStr(Count1) & ");"
            'StringVar = StringVar + CStr(Count1) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
            
                    DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True
        Next count1
        
        DoCmd.OpenReport "Lotto Max Hot Numbers Order by Numbers", acViewPreview
  
 End Function
 
 Public Function FindLottoMaxHotNumbersOrderByDrawings() As String
 
Dim StringVar As String

StringVar = ""
   
 FindLottoMaxHotNumbersOrderByNumber = ""
    
    DoLottoMaxHotNumbersCalc
         
     'DoCreateReport
     'HotNumberReport
     
     For Index2 = 1 To (UBound(LottoMaxDrawings) - 1)
            For Index = 1 To (UBound(LottoMaxDrawings) - 1)
                If LottoMaxDrawings(Index) > LottoMaxDrawings(Index + 1) Then
                    temp = LottoMaxDrawings(Index)
                    LottoMaxDrawings(Index) = LottoMaxDrawings(Index + 1)
                    LottoMaxDrawings(Index + 1) = temp
                    temp2 = LottoMaxIndex3(Index)
                    LottoMaxIndex3(Index) = LottoMaxIndex3(Index + 1)
                    LottoMaxIndex3(Index + 1) = temp2
                    temp3 = LottoMaxBonusNumbers(Index)
                    LottoMaxBonusNumbers(Index) = LottoMaxBonusNumbers(Index + 1)
                    LottoMaxBonusNumbers(Index + 1) = temp3
                    temp4 = LottoMaxdatevar(Index)
                    LottoMaxdatevar(Index) = LottoMaxdatevar(Index + 1)
                    LottoMaxdatevar(Index + 1) = temp4
                End If
            Next
    Next
        
        For count1 = 1 To 50
        LottoMaxDrawings(count1) = LottoMaxDrawings(count1) + LottoMaxBonusNumbers(count1)
        StrSQL = "UPDATE Lotto_Max_Hot_Numbers SET Lotto_Max_Hot_Number=" & CStr(LottoMaxIndex3(count1)) & ",Drawings=" & CStr(LottoMaxDrawings(count1)) _
        & ",Bonus=" & CStr(LottoMaxBonusNumbers(count1)) _
        & ",Date_Differnce=" & CStr(LottoMaxdatevar(count1)) _
        & " WHERE id=" & count1 & ";"
        'StrSQL = "INSERT INTO 649_Hot_Numbers (649_Hot_Number) VALUES (" & CStr(Count1) & ");"
            'StringVar = StringVar + CStr(Count1) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
            
        DoCmd.SetWarnings False
        DoCmd.RunSQL StrSQL
        DoCmd.SetWarnings True
        Next count1
        
        DoCmd.OpenReport "Lotto Max Hot Numbers by Drawings", acViewPreview
  
 End Function

Public Sub DoLottoMaxHotNumbersCalc()

    Dim dbs As Database, rst As DAO.Recordset
    Dim rst2 As DAO.Recordset
    
    Set dbs = CurrentDb
     'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")

     'Command0.Caption = "Please Wait..."
        
        count1 = 0
        
        For count1 = 1 To 50
            LottoMaxDrawings(count1) = 0
            LottoMaxBonusNumbers(count1) = 0
            LottoMaxIndex3(count1) = count1
            LottoMaxdatevar(count1) = 0
        Next
      
        On Error Resume Next
        
        Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7,P8 from LottoMax ORDER BY DrawDate ASC")
        
        Do While Not rst.EOF
    
        For count1 = 1 To 50

                DrawDate = rst("DrawDate")
                P1 = Int(rst("P1"))
                P2 = Int(rst("P2"))
                P3 = Int(rst("P3"))
                P4 = Int(rst("P4"))
                P5 = Int(rst("P5"))
                P6 = Int(rst("P6"))
                P7 = Int(rst("P7"))
                Bonus = Int(rst("P8"))
                
                If P1 = count1 Or P2 = count1 Or P3 = count1 Or P4 = count1 Or P5 = count1 Or P6 = count1 Or P7 = count1 Then
                    LottoMaxDrawings(count1) = LottoMaxDrawings(count1) + 1
                    'DateDifference(Count1) = DateDiff("d", DrawDate, Now())
                    LottoMaxdatevar(count1) = Date - DrawDate
                End If
                If Bonus = count1 Then
                    LottoMaxBonusNumbers(count1) = LottoMaxBonusNumbers(count1) + 1
                End If
                Next count1
            rst.MoveNext
        Loop
        
Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing

End Sub

Public Function DoLottoMaxDates(Start As Boolean) As Date
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Date
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("LottoMax")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7,P8 from LottoMax ORDER BY DrawDate ASC")
    
    If Start Then
    
    rst.MoveFirst
    
    Else
    
    rst.MoveLast
    
    End If
    
    RepDate = rst!DrawDate
    
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    DoLottoMaxDates = RepDate
    
    End Function
    
    Public Function DoLottoMaxDatesCount() As Long
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Long
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("LottoMax")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,P1,P2,P3,P4,P5,P6,P7,P8 from LottoMax ORDER BY DrawDate ASC")
    
    DoLottoMaxDatesCount = rst.RecordCount
       
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    End Function

Public Sub HotNumberReport()

Dim rpt As Report

Dim strReportName As String

strReportName = "Report1"

DoCmd.OpenReport strReportName, acViewPreview
        Set rpt = Reports(strReportName)
        rpt.Caption = "Lotto 649 Hot Numbers by Number"
        rpt.OrderByOn = True
        rpt.OrderBy = strFieldName
        rpt.FilterOn = True
        rpt.Filter = strFilter
        rpt![txtsort] = "Lotto 649 Hot Numbers by Number"
 End Sub

Public Sub DoCreateReport()

Dim rptReport As Access.Report
Dim strReportName As String
Dim acAccessObjct As AccessObject
strReportName = "test"

For Each acAccessObjct In CurrentProject.AllReports
If acAccessObjct.Name = strReportName Then
DoCmd.DeleteObject acReport, strReportName
End If

Next acAccessObjct

Set rptReport = CreateReport

With rptReport
.Caption = "Lotto 649 Hot Numbers by Number"
End With

'rptReport![txtsort] = "test"


End Sub

Sub Load649()

Dim datevar As Date
Dim datepick As Boolean
Dim DrawDate As String
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim RepDate As Date

datepick = False

fileNum = FreeFile()
Open "C:\Users\rchrd\Documents\Richard\lotto 649 Numbers.txt" For Input As #fileNum

While Not EOF(fileNum)
    Line Input #fileNum, DataLine ' read in data 1 line at a time
    ' decide what to do with dataline,
    ' depending on what processing you need to do for each case
    On Error Resume Next
    datevar = CDate(DataLine)
    If Err.Number = 13 Then
    datevar = 2 / 22 / 1966
    Else
    datepick = True
    End If
    
If datepick Then
Line Input #fileNum, DataLine ' read in data 1 line at a time
If DataLine = "Tirage Principal" Then
Line Input #fileNum, DataLine ' read in data 1 line at a time
'add numbers to database
End If
DataLine = Replace(DataLine, " ", "")
P1 = Mid(DataLine, 1, 2)
P2 = Mid(DataLine, 3, 2)
P3 = Mid(DataLine, 5, 2)
P4 = Mid(DataLine, 7, 2)
P5 = Mid(DataLine, 9, 2)
P6 = Mid(DataLine, 11, 2)
P7 = Mid(DataLine, 14, 2)

StrSQL = "INSERT INTO 649Drawings (DrawDate,p1,p2,p3,p4,p5,p6,p7) VALUES ('" & datevar & "','" & P1 & "','" & P2 & "','" & P3 & "','" & P4 & "','" & P5 & "','" & P6 & "','" & P7 & "');"
DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True

End If
datepick = False

Wend
    
   
End Sub

Sub Load649_2()

Dim datevar As Date
Dim datepick As Boolean
Dim DrawDate As String
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim RepDate As Date

datepick = False

fileNum = FreeFile()
Open "C:\Users\rchrd\Documents\Richard\lotto 649 Numbers 2.txt" For Input As #fileNum

While Not EOF(fileNum)
    Line Input #fileNum, DataLine ' read in data 1 line at a time
    ' decide what to do with dataline,
    ' depending on what processing you need to do for each case
    On Error Resume Next
    datevar = CDate(DataLine)
    If Err.Number = 13 Then
    datevar = 2 / 22 / 1966
    Else
    datepick = True
    End If
    
If datepick Then
Line Input #fileNum, DataLine ' read in data 1 line at a time
'add numbers to database
DataLine = Replace(DataLine, " ", "")
P1 = Mid(DataLine, 1, 2)
P2 = Mid(DataLine, 3, 2)
P3 = Mid(DataLine, 5, 2)
P4 = Mid(DataLine, 7, 2)
P5 = Mid(DataLine, 9, 2)
P6 = Mid(DataLine, 11, 2)
P7 = Mid(DataLine, 14, 2)

StrSQL = "INSERT INTO 649Drawings (DrawDate,p1,p2,p3,p4,p5,p6,p7) VALUES ('" & datevar & "','" & P1 & "','" & P2 & "','" & P3 & "','" & P4 & "','" & P5 & "','" & P6 & "','" & P7 & "');"
DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True

End If
datepick = False

Wend
    
   
End Sub

Sub LoadLottoMax()

Dim datevar As Date
Dim datepick As Boolean
Dim DrawDate As String
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim RepDate As Date

datepick = False

fileNum = FreeFile()
Open "C:\Users\rchrd\Documents\Richard\Lotto Max Numbers.txt" For Input As #fileNum

While Not EOF(fileNum)
    Line Input #fileNum, DataLine ' read in data 1 line at a time
    ' decide what to do with dataline,
    ' depending on what processing you need to do for each case
    On Error Resume Next
    datevar = CDate(DataLine)
    If Err.Number = 13 Then
    datevar = 2 / 22 / 1966
    Else
    datepick = True
    End If
    
If datepick Then
Line Input #fileNum, DataLine ' read in data 1 line at a time
If DataLine = "Tirage Principal" Then
Line Input #fileNum, DataLine ' read in data 1 line at a time
'add numbers to database
End If
DataLine = Replace(DataLine, " ", "")
P1 = Mid(DataLine, 1, 2)
P2 = Mid(DataLine, 3, 2)
P3 = Mid(DataLine, 5, 2)
P4 = Mid(DataLine, 7, 2)
P5 = Mid(DataLine, 9, 2)
P6 = Mid(DataLine, 11, 2)
P7 = Mid(DataLine, 13, 2)
P8 = Mid(DataLine, 16, 2)
StrSQL = "INSERT INTO LottoMax (DrawDate,p1,p2,p3,p4,p5,p6,p7,p8) VALUES ('" & datevar & "','" & P1 & "','" & P2 & "','" & P3 & "','" & P4 & "','" & P5 & "','" & P6 & "','" & P7 & "','" & P8 & "');"
DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True

End If
datepick = False

Wend
    
   
End Sub

Sub LoadGrandeVie()

Dim datevar As Date
Dim datepick As Boolean
Dim DrawDate As String
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim RepDate As Date

datepick = False

fileNum = FreeFile()
Open "C:\Users\rchrd\Documents\Richard\Grande Vie.txt" For Input As #fileNum

While Not EOF(fileNum)
    Line Input #fileNum, DataLine ' read in data 1 line at a time
    ' decide what to do with dataline,
    ' depending on what processing you need to do for each case
    On Error Resume Next
    datevar = CDate(DataLine)
    If Err.Number = 13 Then
    datevar = 2 / 22 / 1966
    Else
    datepick = True
    End If
    
If datepick Then
Line Input #fileNum, DataLine ' read in data 1 line at a time
'add numbers to database
DataLine = Replace(DataLine, " ", "")
P1 = Mid(DataLine, 1, 2)
P2 = Mid(DataLine, 3, 2)
P3 = Mid(DataLine, 5, 2)
P4 = Mid(DataLine, 7, 2)
P5 = Mid(DataLine, 9, 2)
GN = Mid(DataLine, 12, 1)

StrSQL = "INSERT INTO Grande_Vie (DrawDate,p1,p2,p3,p4,p5,gn) VALUES ('" & datevar & "','" & P1 & "','" & P2 & "','" & P3 & "','" & P4 & "','" & P5 & "','" & GN & "');"
DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True

End If
datepick = False

Wend
    
   
End Sub

Sub LoadToutouRien()

Dim datevar As Date
Dim datepick As Boolean
Dim DrawDate As String
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim RepDate As Date

datepick = False

fileNum = FreeFile()
Open "C:\Users\rchrd\Documents\Richard\Tout ou Rien.txt" For Input As #fileNum

While Not EOF(fileNum)
    Line Input #fileNum, DataLine ' read in data 1 line at a time
    ' decide what to do with dataline,
    ' depending on what processing you need to do for each case
    On Error Resume Next
    datevar = CDate(DataLine)
    If Err.Number = 13 Then
    datevar = 2 / 22 / 1966
    Else
    datepick = True
    End If
    
If datepick Then
Line Input #fileNum, DataLine ' read in data 1 line at a time
DataLine = Replace(DataLine, " ", "")
P1 = Mid(DataLine, 1, 2)
P2 = Mid(DataLine, 3, 2)
P3 = Mid(DataLine, 5, 2)
P4 = Mid(DataLine, 7, 2)
P5 = Mid(DataLine, 9, 2)
P6 = Mid(DataLine, 11, 2)
Line Input #fileNum, DataLine ' read in data 1 line at a time
DataLine = Replace(DataLine, " ", "")
P7 = Mid(DataLine, 1, 2)
P8 = Mid(DataLine, 3, 2)
P9 = Mid(DataLine, 5, 2)
P10 = Mid(DataLine, 7, 2)
P11 = Mid(DataLine, 9, 2)
P12 = Mid(DataLine, 11, 2)

StrSQL = "INSERT INTO ToutouRien (DrawDate,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12) VALUES ('" & datevar & "','" & P1 & "','" & P2 & "','" & P3 & "','" & P4 & "','" & P5 & "','" & P6 & "','" & P7 & "','" & P8 & "','" & P9 & "','" & P10 & "','" & P11 & "','" & P12 & "');"
DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True

End If
datepick = False

Wend
    
   
End Sub


Sub Tout_ou_RienSpecial()

Dim identical As Boolean
Dim P(12) As Integer
Dim pick12(12) As Integer
Dim DrawDate As Date
Dim P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12 As Integer
Dim dbs As Database, rst As DAO.Recordset
Dim rst2 As DAO.Recordset
Set dbs = CurrentDb
Dim strPath As String

Set FSO = CreateObject("Scripting.FileSystemObject")
Dim Fileout As Object
Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\test.txt", True, True)

On Error Resume Next
                
        Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12 from ToutouRien ORDER BY DrawDate ASC")
        
10:        rst.MoveFirst
        
                   Randomize
                
                For count3 = 1 To 12
                
                Do
                    DoEvents
                    
                    RndNum = Int((24 - 1 + 1) * Rnd + 1)
                    
                    Loop Until (RndNum <> pick12(1)) And (RndNum <> pick12(2)) _
                        And (RndNum <> pick12(3)) And (RndNum <> pick12(4)) _
                        And (RndNum <> pick12(5)) And (RndNum <> pick12(6)) _
                        And (RndNum <> pick12(7)) And (RndNum <> pick12(8)) _
                        And (RndNum <> pick12(9)) And (RndNum <> pick12(10)) _
                        And (RndNum <> pick12(11)) And (RndNum <> pick12(12))
                    
                    pick12(count3) = RndNum
                    
                    Next
                    
                    For count4 = 1 To (UBound(pick12) - 1)
                    For count1 = 1 To (UBound(pick12) - 1)
                        If pick12(count1) > pick12(count1 + 1) Then
                            temp = pick12(count1)
                            pick12(count1) = pick12(count1 + 1)
                            pick12(count1 + 1) = temp
                        End If
                    Next
                Next
                
        
        Do While Not rst.EOF
        
                DrawDate = rst("DrawDate")
                P1 = Int(rst("p1"))
                P2 = Int(rst("p2"))
                P3 = Int(rst("p3"))
                P4 = Int(rst("p4"))
                P5 = Int(rst("p5"))
                P6 = Int(rst("p6"))
                P7 = Int(rst("p7"))
                P8 = Int(rst("p8"))
                P9 = Int(rst("p9"))
                P10 = Int(rst("p10"))
                P11 = Int(rst("P11"))
                P12 = Int(rst("P12"))
                
                P(1) = P1
                P(2) = P2
                P(3) = P3
                P(4) = P4
                P(5) = P5
                P(6) = P6
                P(7) = P7
                P(8) = P8
                P(9) = P9
                P(10) = P10
                P(11) = P11
                P(12) = P12

                    If (P1 = pick12(1)) And (P2 = pick12(2)) _
                        And (P3 = pick12(3)) And (P4 = pick12(4)) _
                        And (P5 = pick12(5)) And (P6 = pick12(6)) _
                        And (P7 = pick12(7)) And (P8 = pick12(8)) _
                        And (P9 = pick12(9)) And (P10 = pick12(10)) _
                        And (P11 = pick12(11)) And (P12 = pick12(12)) Then
                        
                        Fileout.Write DrawDate
                        GoTo 10
                        End If
          
            str1 = CStr(pick12(1))
            str2 = CStr(pick12(2))
            str3 = CStr(pick12(3))
            str4 = CStr(pick12(4))
            str5 = CStr(pick12(5))
            str6 = CStr(pick12(6))
            str7 = CStr(pick12(7))
            str8 = CStr(pick12(8))
            str9 = CStr(pick12(9))
            str10 = CStr(pick12(10))
            str11 = CStr(pick12(11))
            str12 = CStr(pick12(12))
            
            count2 = 0
            
            For checkcount = 1 To (UBound(pick12) - 1)
                If str1 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str2 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str3 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str4 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str5 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str6 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str7 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str8 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str9 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str10 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str11 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str12 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                
                '9 numbers and more the same
                If count2 > 8 Then
                    identical = True
                    Me.Label23.Caption = "\"
                    Fileout.Write DrawDate
                    Me.Label23.Caption = " / """
                    GoTo 10
                Else
                    identical = False
                End If
                
            Next
    rst.MoveNext
Loop

MsgBox (str1 + " " + str2 + " " + str3 + " " + str4 + " " + str5 + " " + str6 + " " + str7 + " " + str8 + " " + str9 + " " + str10 + " " + str11 + " " + str12)

Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing
FileName.Close

End Sub

Sub Lotto649Special()

Dim identical As Boolean
Dim P(6) As Integer
Dim pick6(6) As Integer
Dim DrawDate As Date
Dim P1, P2, P3, P4, P5, P6 As Integer
Dim dbs As Database, rst As DAO.Recordset
Dim rst2 As DAO.Recordset
Set dbs = CurrentDb
Dim strPath As String

Set FSO = CreateObject("Scripting.FileSystemObject")
Dim Fileout As Object
Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\test.txt", True, True)

On Error Resume Next
                
        Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6 from 649Drawings ORDER BY DrawDate ASC")
        
10:        rst.MoveFirst

Randomize
                
                For count3 = 1 To 6
                    DoEvents
                    Do
                        RndNum = Int((49 - 1 + 1) * Rnd + 1)
                    Loop Until (RndNum <> pick6(1)) And (RndNum <> pick6(2)) _
                        And (RndNum <> pick6(3)) And (RndNum <> pick6(4)) _
                        And (RndNum <> pick6(5)) And (RndNum <> pick6(6))
                        
                        pick6(count3) = RndNum
                        
                Next
                
                For count4 = 1 To (UBound(pick6) - 1)
                    For count1 = 1 To (UBound(pick6) - 1)
                        If pick6(count1) > pick6(count1 + 1) Then
                            temp = pick6(count1)
                            pick6(count1) = pick6(count1 + 1)
                            pick6(count1 + 1) = temp
                        End If
                    Next
                Next
        
        Do While Not rst.EOF
        
                DrawDate = rst("DrawDate")
                P1 = Int(rst("p1"))
                P2 = Int(rst("p2"))
                P3 = Int(rst("p3"))
                P4 = Int(rst("p4"))
                P5 = Int(rst("p5"))
                P6 = Int(rst("p6"))
                
                P(1) = P1
                P(2) = P2
                P(3) = P3
                P(4) = P4
                P(5) = P5
                P(6) = P6
                
                    If (P1 = pick6(1)) And (P2 = pick6(2)) _
                        And (P3 = pick6(3)) And (P4 = pick6(4)) _
                        And (P5 = pick6(5)) And (P6 = pick6(6)) Then
                        
                                        Fileout.Write DrawDate
                                        GoTo 10
                                        
                                        End If
           
            str1 = CStr(pick6(1))
            str2 = CStr(pick6(2))
            str3 = CStr(pick6(3))
            str4 = CStr(pick6(4))
            str5 = CStr(pick6(5))
            str6 = CStr(pick6(6))
                       
            count2 = 0
            
            For checkcount = 1 To (UBound(pick6) - 1)
                If str1 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str2 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str3 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str4 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str5 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str6 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                                
                '4 numbers and more the same
                If count2 >= 4 Then
                    identical = True
                     Fileout.Write DrawDate
                     GoTo 10
                Else
                    identical = False
                End If
                
            Next
    rst.MoveNext
Loop

MsgBox (str1 + " " + str2 + " " + str3 + " " + str4 + " " + str5 + " " + str6)

Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing
FileName.Close

End Sub

Sub LottoMaxSpecial()

Dim identical As Boolean
Dim P(7) As Integer
Dim pick7(7) As Integer
Dim DrawDate As Date
Dim P1, P2, P3, P4, P5, P6, P7 As Integer
Dim dbs As Database, rst As DAO.Recordset
Dim rst2 As DAO.Recordset
Set dbs = CurrentDb
Dim strPath As String

Set FSO = CreateObject("Scripting.FileSystemObject")
Dim Fileout As Object
Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\test.txt", True, True)

On Error Resume Next
                
        Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6,p7 from LottoMax ORDER BY DrawDate ASC")
        
10:      rst.MoveFirst
        
         Randomize
         
                For count3 = 1 To 7
                    DoEvents
                    Do
                        
                        RndNum = Int((50 - 1 + 1) * Rnd + 1)
                    Loop Until (RndNum <> pick7(1)) And (RndNum <> pick7(2)) _
                        And (RndNum <> pick7(3)) And (RndNum <> pick7(4)) _
                        And (RndNum <> pick7(5)) And (RndNum <> pick7(6)) _
                        And (RndNum <> pick7(7))
                        
                        pick7(count3) = RndNum
                        
                Next
                
                For count4 = 1 To (UBound(pick7) - 1)
                    For count1 = 1 To (UBound(pick7) - 1)
                        If pick7(count1) > pick7(count1 + 1) Then
                            temp = pick7(count1)
                            pick7(count1) = pick7(count1 + 1)
                            pick7(count1 + 1) = temp
                        End If
                    Next
                Next
        
        Do While Not rst.EOF
        
                DrawDate = rst("DrawDate")
                P1 = Int(rst("p1"))
                P2 = Int(rst("p2"))
                P3 = Int(rst("p3"))
                P4 = Int(rst("p4"))
                P5 = Int(rst("p5"))
                P6 = Int(rst("p6"))
                P7 = Int(rst("p7"))
                
                P(1) = P1
                P(2) = P2
                P(3) = P3
                P(4) = P4
                P(5) = P5
                P(6) = P6
                P(7) = P7
                
            If (P1 = pick7(1)) And (P2 = pick7(2)) _
                        And (P3 = pick7(3)) And (P4 = pick7(4)) _
                        And (P5 = pick7(5)) And (P6 = pick7(6)) _
                        And (P7 = pick7(7)) Then
                        Label23.Text = "\"
                        Fileout.Write DrawDate
                        Label23.Text = "/"
                        GoTo 10
                        End If
            
            str1 = CStr(pick7(1))
            str2 = CStr(pick7(2))
            str3 = CStr(pick7(3))
            str4 = CStr(pick7(4))
            str5 = CStr(pick7(5))
            str6 = CStr(pick7(6))
            str7 = CStr(pick7(7))
            
            count2 = 0
            
            For checkcount = 1 To (UBound(pick7) - 1)
                If str1 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str2 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str3 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str4 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str5 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str6 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str7 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                                
                '3 numbers and more the same
                If count2 >= 3 Then
                    identical = True
                    Label33.Caption = "\"
                    Fileout.Write DrawDate
                    Label33.Caption = " / """
                     GoTo 10
                Else
                    identical = False
                End If
                
            Next
    rst.MoveNext
Loop

MsgBox (str1 + " " + str2 + " " + str3 + " " + str4 + " " + str5 + " " + str6 + " " + str7)

Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing
FileName.Close

End Sub

Sub Grande_VieSpecial()

Dim identical As Boolean
Dim P(5), GN(1) As Integer
Dim pick5(5), pickGN(1) As Integer
Dim DrawDate As Date
Dim P1, P2, P3, P4, P5, GN2 As Integer
Dim dbs As Database, rst As DAO.Recordset
Dim rst2 As DAO.Recordset
Set dbs = CurrentDb
Dim strPath As String

Set FSO = CreateObject("Scripting.FileSystemObject")
Dim Fileout As Object
Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\test.txt", True, True)

On Error Resume Next
                
        Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,gn from Grande_Vie ORDER BY DrawDate ASC")
        
10:        rst.MoveFirst

Randomize
                
                For count3 = 1 To 5
                    
                    Do
                        RndNum = Int((49 - 1 + 1) * Rnd + 1)
                    Loop Until (RndNum <> pick5(1)) And (RndNum <> pick5(2)) _
                        And (RndNum <> pick5(3)) And (RndNum <> pick5(4)) _
                        And (RndNum <> pick5(5))
                        pick5(count3) = RndNum
                Next
                
                RndNum = Int((7 - 1 + 1) * Rnd + 1)
                pickGN(1) = RndNum
                
                 For count4 = 1 To (UBound(pick5) - 1)
                    For count1 = 1 To (UBound(pick5) - 1)
                        If pick5(count1) > pick5(count1 + 1) Then
                            temp = pick5(count1)
                            pick5(count1) = pick5(count1 + 1)
                            pick5(count1 + 1) = temp
                        End If
                    Next
                Next
        
        Do While Not rst.EOF
        
                DrawDate = rst("DrawDate")
                P1 = Int(rst("p1"))
                P2 = Int(rst("p2"))
                P3 = Int(rst("p3"))
                P4 = Int(rst("p4"))
                P5 = Int(rst("p5"))
                GN2 = Int(rst("gn"))
                       
                P(1) = P1
                P(2) = P2
                P(3) = P3
                P(4) = P4
                P(5) = P5
                GN(1) = GN2
                                
                If (P1 = pick5(1)) And (P2 = pick5(2)) _
                        And (P3 = pick5(3)) And (P4 = pick5(4)) _
                        And (P5 = pick5(5)) Then
                         
                        Fileout.Write DrawDate
                        GoTo 10
                        End If
      
            str1 = CStr(pick5(1))
            str2 = CStr(pick5(2))
            str3 = CStr(pick5(3))
            str4 = CStr(pick5(4))
            str5 = CStr(pick5(5))
            str6 = CStr(pickGN(1))
                        
            count2 = 0
            
            For checkcount = 1 To (5 - 1)
                If str1 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str2 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str3 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str4 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                If str5 = CStr(P(checkcount)) Then
                    count2 = count2 + 1
                End If
                'If str6 = CStr(GN(1)) Then
                    'Count2 = Count2 + 1
                'End If
                           
                '2 numbers and more the same
                If count2 >= 2 Then
                    identical = True
                    Fileout.Write DrawDate
                    GoTo 10
                Else
                    identical = False
                End If
                
            Next
    rst.MoveNext
Loop

MsgBox (str1 + " " + str2 + " " + str3 + " " + str4 + " " + str5 + " " + str6)

Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing
FileName.Close

End Sub

Sub Find649HotNumbersOrderByOccurances()

Dim Index, temp, temp1, temp2 As Integer

Do649HotNumbersCalc
        
        Dim i As Integer
        Dim j As Integer
        Dim Length As Integer
        Length = 49
        Dim Arrayx(49) As Integer
        
        For Index = 1 To 49
            Arrayx(Index) = (Drawings(Index) + BonusNumbers(Index))
        Next
        
        For i = 1 To (UBound(Drawings) - 1) 'the last element on the array does not get sorted.
            For j = 1 To (UBound(Drawings) - 1)
                If Arrayx(j) > Arrayx(j + 1) Then  ' Compare neighboring elements
                    temp = Arrayx(j)
                    Arrayx(j) = Arrayx(j + 1)
                    Arrayx(j + 1) = temp
                    temp1 = Drawings(j)
                    Drawings(j) = Drawings(j + 1)
                    Drawings(j + 1) = temp1
                    temp2 = Index3(j)
                    Index3(j) = Index3(j + 1)
                    Index3(j + 1) = temp2
                    temp3 = BonusNumbers(j)
                    BonusNumbers(j) = BonusNumbers(j + 1)
                    BonusNumbers(j + 1) = temp3
                    temp4 = DateDifference(j)
                    DateDifference(j) = DateDifference(j + 1)
                    DateDifference(j + 1) = temp4
                End If
            Next
        Next
                            ''For Count1 = 1 To 49
                ''StringVar = StringVar + CStr(Index3(Count1)) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
                ''MsgBox (StringVar)
            ''Next


End Sub

Sub FindLottoMaxHotNumbersOrderByOccurances()

Dim Index, temp, temp1, temp2 As Integer

DoLottoMaxHotNumbersCalc
        
        Dim i As Integer
        Dim j As Integer
        Dim Length As Integer
        Length = 50
        Dim Arrayx(50) As Integer
        
        For Index = 1 To 50
            Arrayx(Index) = (LottoMaxDrawings(Index) + LottoMaxBonusNumbers(Index))
        Next
        
        For i = 1 To (UBound(LottoMaxDrawings) - 1) 'the last element on the array does not get sorted.
            For j = 1 To (UBound(LottoMaxDrawings) - 1)
                If Arrayx(j) > Arrayx(j + 1) Then  ' Compare neighboring elements
                    temp = Arrayx(j)
                    Arrayx(j) = Arrayx(j + 1)
                    Arrayx(j + 1) = temp
                    temp1 = LottoMaxDrawings(j)
                    LottoMaxDrawings(j) = LottoMaxDrawings(j + 1)
                    LottoMaxDrawings(j + 1) = temp1
                    temp2 = LottoMaxIndex3(j)
                    LottoMaxIndex3(j) = LottoMaxIndex3(j + 1)
                    LottoMaxIndex3(j + 1) = temp2
                    temp3 = LottoMaxBonusNumbers(j)
                    LottoMaxBonusNumbers(j) = LottoMaxBonusNumbers(j + 1)
                    LottoMaxBonusNumbers(j + 1) = temp3
                    temp4 = LottoMaxDateDifference(j)
                    LottoMaxDateDifference(j) = LottoMaxDateDifference(j + 1)
                    LottoMaxDateDifference(j + 1) = temp4
                End If
            Next
        Next
                            ''For Count1 = 1 To 49
                ''StringVar = StringVar + CStr(Index3(Count1)) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
                ''MsgBox (StringVar)
            ''Next


End Sub

Sub FindGrandeVieHotNumbersOrderByOccurances()

Dim Index, temp, temp1, temp2 As Integer

DoGrandVieHotNumbersCalc
        
        Dim i As Integer
        Dim j As Integer
        Dim Length As Integer
        Length = 49
        Dim Arrayx(49) As Integer
        
        For Index = 1 To 49
            Arrayx(Index) = (Drawings(Index))
        Next
        
        For i = 1 To (UBound(Drawings) - 1) 'the last element on the array does not get sorted.
            For j = 1 To (UBound(Drawings) - 1)
                If Arrayx(j) > Arrayx(j + 1) Then  ' Compare neighboring elements
                    temp = Arrayx(j)
                    Arrayx(j) = Arrayx(j + 1)
                    Arrayx(j + 1) = temp
                    temp1 = Drawings(j)
                    Drawings(j) = Drawings(j + 1)
                    Drawings(j + 1) = temp1
                    temp2 = Index3(j)
                    Index3(j) = Index3(j + 1)
                    Index3(j + 1) = temp2
                    temp3 = GN(1)
                    If GN(1) <= 7 Then
                    'GN(1) = GN(J + 1)
                    GN(1) = temp3
                    End If
                    temp4 = DateDifference(j)
                    DateDifference(j) = DateDifference(j + 1)
                    DateDifference(j + 1) = temp4
                End If
            Next
        Next
                            ''For Count1 = 1 To 49
                ''StringVar = StringVar + CStr(Index3(Count1)) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
                ''MsgBox (StringVar)
            ''Next


End Sub

Sub PickAHotNumbers()
        
        Dim pickjdp(6) As Integer
        Dim RndNum As Integer
        Dim temp As Integer
        
        Find649HotNumbersOrderByOccurances
        
        Randomize
        
        For count3 = 1 To 6
                            
                    Do
                        RndNum = Int((49 - 1 + 1) * Rnd + 1)
                        DoEvents
                        Loop Until (RndNum > 31) And (Index3(RndNum) <> pickjdp(1)) And (Index3(RndNum) <> pickjdp(2)) _
                And (Index3(RndNum) <> pickjdp(3)) And (Index3(RndNum) <> pickjdp(4)) _
                And (Index3(RndNum) <> pickjdp(5)) And (Index3(RndNum) <> pickjdp(6))
            pickjdp(count3) = Index3(RndNum)
        Next
        
        For count4 = 1 To (UBound(pickjdp) - 1)
            For count1 = 1 To (UBound(pickjdp) - 1)
                If pickjdp(count1) > pickjdp(count1 + 1) Then
                    temp = pickjdp(count1)
                    pickjdp(count1) = pickjdp(count1 + 1)
                    pickjdp(count1 + 1) = temp
                End If
            Next
        Next
        
        StringVar = CStr(pickjdp(1)) + " " + CStr(pickjdp(2)) + " " + CStr(pickjdp(3)) + " " + CStr(pickjdp(4)) + " " + CStr(pickjdp(5)) + " " + CStr(pickjdp(6)) + Chr(13)
        
        MsgBox (StringVar)
End Sub

Sub PickLottoMaxHotNumbers()
        
        Dim pickjdp(7) As Integer
        Dim RndNum As Integer
        Dim temp As Integer
        
        FindLottoMaxHotNumbersOrderByOccurances
        
        Randomize
        
        For count3 = 1 To 7
                            
                    Do
                        RndNum = Int((50 - 1 + 1) * Rnd + 1)
                        DoEvents
                        Loop Until (RndNum > 31) And (LottoMaxIndex3(RndNum) <> pickjdp(1)) And (LottoMaxIndex3(RndNum) <> pickjdp(2)) _
                And (LottoMaxIndex3(RndNum) <> pickjdp(3)) And (LottoMaxIndex3(RndNum) <> pickjdp(4)) _
                And (LottoMaxIndex3(RndNum) <> pickjdp(5)) And (LottoMaxIndex3(RndNum) <> pickjdp(6)) _
                And (LottoMaxIndex3(RndNum) <> pickjdp(7))
            pickjdp(count3) = LottoMaxIndex3(RndNum)
        Next
        
        For count4 = 1 To (UBound(pickjdp) - 1)
            For count1 = 1 To (UBound(pickjdp) - 1)
                If pickjdp(count1) > pickjdp(count1 + 1) Then
                    temp = pickjdp(count1)
                    pickjdp(count1) = pickjdp(count1 + 1)
                    pickjdp(count1 + 1) = temp
                End If
            Next
        Next
        
        StringVar = CStr(pickjdp(1)) + " " + CStr(pickjdp(2)) + " " + CStr(pickjdp(3)) + " " + CStr(pickjdp(4)) + " " + CStr(pickjdp(5)) + " " + CStr(pickjdp(6)) + " " + CStr(pickjdp(7)) + Chr(13)
        
        MsgBox (StringVar)
End Sub

Sub PickGrandeVieHotNumbers()
        
        Dim pickjdp(5) As Integer
        Dim RndNum As Integer
        Dim temp As Integer
        
        FindGrandeVieHotNumbersOrderByOccurances
        
        Randomize
        
        For count3 = 1 To 5
                            
                    Do
                        RndNum = Int((49 - 1 + 1) * Rnd + 1)
                        DoEvents
                        Loop Until (RndNum > 31) And (Index3(RndNum) <> pickjdp(1)) And (Index3(RndNum) <> pickjdp(2)) _
                And (Index3(RndNum) <> pickjdp(3)) And (Index3(RndNum) <> pickjdp(4)) _
                And (Index3(RndNum) <> pickjdp(5))
            pickjdp(count3) = Index3(RndNum)
        Next
        
        RndNum = Int((7 - 1 + 1) * Rnd + 1)
                pickGN(1) = RndNum
        
        For count4 = 1 To (UBound(pickjdp) - 1)
            For count1 = 1 To (UBound(pickjdp) - 1)
                If pickjdp(count1) > pickjdp(count1 + 1) Then
                    temp = pickjdp(count1)
                    pickjdp(count1) = pickjdp(count1 + 1)
                    pickjdp(count1 + 1) = temp
                End If
            Next
        Next
        
        StringVar = CStr(pickjdp(1)) + " " + CStr(pickjdp(2)) + " " + CStr(pickjdp(3)) + " " + CStr(pickjdp(4)) + " " + CStr(pickjdp(5)) + " " + CStr(pickGN(1)) + Chr(13)
        
        MsgBox (StringVar)
End Sub

Sub PickToutouRienHotNumbers()
        
        Dim pikjdpTout(12) As Integer
        Dim RndNum As Integer
        Dim temp As Integer
        
        FindToutouRienHotNumbersOrderByOccurances
        
        Randomize
        
        For count3 = 1 To 12
                            
                    Do
                        RndNum = Int((24 - 1 + 1) * Rnd + 1)
                        DoEvents
                        Loop Until (RndNum > 6) And (Index3Tout(RndNum) <> pickjdpTout(1)) And (Index3Tout(RndNum) <> pickjdpTout(2)) _
                And (Index3Tout(RndNum) <> pickjdpTout(3)) And (Index3Tout(RndNum) <> pickjdpTout(4)) _
                And (Index3Tout(RndNum) <> pickjdpTout(5)) And (Index3Tout(RndNum) <> pickjdpTout(6)) _
                And (Index3Tout(RndNum) <> pickjdpTout(7)) And (Index3Tout(RndNum) <> pickjdpTout(8)) _
                And (Index3Tout(RndNum) <> pickjdpTout(9)) And (Index3Tout(RndNum) <> pickjdpTout(10)) _
                And (Index3Tout(RndNum) <> pickjdpTout(11)) And (Index3Tout(RndNum) <> pickjdpTout(12))
                
                pickjdpTout(count3) = Index3Tout(RndNum)
        Next
        
        For count4 = 1 To (UBound(pickjdpTout) - 1)
            For count1 = 1 To (UBound(pickjdpTout) - 1)
                If pickjdpTout(count1) > pickjdpTout(count1 + 1) Then
                    temp = pickjdpTout(count1)
                    pickjdpTout(count1) = pickjdpTout(count1 + 1)
                    pickjdpTout(count1 + 1) = temp
                End If
            Next
        Next
        
        StringVar = CStr(pickjdpTout(1)) + " " + CStr(pickjdpTout(2)) + " " + CStr(pickjdpTout(3)) + " " + CStr(pickjdpTout(4)) + " " + CStr(pickjdpTout(5)) + " " + CStr(pickjdpTout(6)) + " " + CStr(pickjdpTout(7)) + " " + CStr(pickjdpTout(8)) + " " + CStr(pickjdpTout(9)) + " " + CStr(pickjdpTout(10)) + " " + CStr(pickjdpTout(11)) + " " + CStr(pickjdpTout(12)) + Chr(13)
        
        MsgBox (StringVar)
End Sub

Sub FindToutouRienHotNumbersOrderByOccurances()

Dim Index, temp, temp1, temp2 As Integer

DoToutouRienHotNumbersCalc
        
        Dim i As Integer
        Dim j As Integer
        Dim Length As Integer
        Length = 24
        Dim Arrayx(24) As Integer
        
        For Index = 1 To 24
            Arrayx(Index) = DrawingsTout(Index)
        Next
        
        For i = 1 To (UBound(DrawingsTout) - 1) 'the last element on the array does not get sorted.
            For j = 1 To (UBound(DrawingsTout) - 1)
                If Arrayx(j) > Arrayx(j + 1) Then  ' Compare neighboring elements
                    temp = Arrayx(j)
                    Arrayx(j) = Arrayx(j + 1)
                    Arrayx(j + 1) = temp
                    temp1 = DrawingsTout(j)
                    DrawingsTout(j) = DrawingsTout(j + 1)
                    DrawingsTout(j + 1) = temp1
                    temp2 = Index3Tout(j)
                    Index3Tout(j) = Index3Tout(j + 1)
                    Index3Tout(j + 1) = temp2
                    temp4 = DateDifference(j)
                    DateDifference(j) = DateDifference(j + 1)
                    DateDifference(j + 1) = temp4
                End If
            Next
        Next
                            ''For Count1 = 1 To 49
                ''StringVar = StringVar + CStr(Index3(Count1)) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
                ''MsgBox (StringVar)
            ''Next
End Sub

Public Sub DoToutouRienHotNumbersCalc()

    Dim dbs As Database, rst As DAO.Recordset
    Dim rst2 As DAO.Recordset
    
     Set dbs = CurrentDb
     'Set dbs = DBEngine.Workspaces(0).OpenDatabase("LottoDrawings.mdb")

     'Command0.Caption = "Please Wait..."
        
        count1 = 0
        
        For count1 = 1 To 24
            DrawingsTout(count1) = 0
            Index3Tout(count1) = count1
            datevar(count1) = 0
        Next

        'On Error Resume Next
        
        Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12 from ToutouRien ORDER BY DrawDate ASC")
        
    Do While Not rst.EOF
    
    For count1 = 1 To 24

                DrawDate = rst("DrawDate")
                P1 = Int(rst("P1"))
                P2 = Int(rst("P2"))
                P3 = Int(rst("P3"))
                P4 = Int(rst("P4"))
                P5 = Int(rst("P5"))
                P6 = Int(rst("P6"))
                P7 = Int(rst("P7"))
                P8 = Int(rst("P8"))
                P9 = Int(rst("P9"))
                P10 = Int(rst("P10"))
                P11 = Int(rst("P11"))
                P12 = Int(rst("P12"))
                If P1 = count1 Or P2 = count1 Or P3 = count1 Or P4 = count1 Or P5 = count1 Or P6 = count1 Or P6 = count1 Or P7 = count1 Or P8 = count1 Or P9 = count1 Or P10 = count1 Or P11 = count1 Or P12 = count1 Then
                    DrawingsTout(count1) = DrawingsTout(count1) + 1
                    datevar(count1) = DateDiff("d", DrawDate, Now())
                    'DateVar(Count1) = Date - DrawDate
                End If
                Next count1
            rst.MoveNext
        Loop
        
        Set rst = Nothing
Set rst2 = Nothing
dbs.Close
Set dbs = Nothing
    
    End Sub
    
    Public Function FindToutouRienHotNumbersOrderByDrawings() As String

   Dim StringVar As String
   
 Find649HotNumbersOrderByNumber = ""
    
    DoToutouRienHotNumbersCalc
         
     'DoCreateReport
     'HotNumberReport
        
        For count1 = 1 To 24
        StrSQL = "UPDATE Tout_ou_Rien_Hot_Numbers SET Tout_ou_Rien_Hot_Number=" & CStr(Index3Tout(count1)) & ",Drawings=" & CStr(DrawingsTout(count1)) & ",Date_Differnce=" & CStr(datevar(count1)) & " WHERE id=" & count1 & ";"
        'StrSQL = "INSERT INTO 649_Hot_Numbers (649_Hot_Number) VALUES (" & CStr(Count1) & ");"
            'StringVar = StringVar + CStr(Count1) + ") " + CStr(Drawings(Count1)) + "  " + CStr(BonusNumbers(Count1)) + "  " + CStr(Drawings(Count1) + BonusNumbers(Count1)) + "  " + CStr(DateDifference(Count1)) + " day(s)" + Chr(13)
            
                    DoCmd.SetWarnings False
DoCmd.RunSQL StrSQL
DoCmd.SetWarnings True
        Next count1
        
        DoCmd.OpenReport "Tout ou Rien Hot Numbers Order by Drawings", acViewPreview
  
 End Function
 
 Public Function DoToutouRienDates(Start As Boolean) As Date
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Date
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12 from ToutouRien ORDER BY DrawDate ASC")
    
    If Start Then
    
    rst.MoveFirst
    
    Else
    
    rst.MoveLast
    
    End If
    
    RepDate = rst!DrawDate
    
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    DoToutouRienDates = RepDate
    
    End Function
    
    Public Function DoToutouRienDatesCount() As Long
    
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim RepDate As Long
    
    'Get the database and Recordset
    Set dbs = CurrentDb
    ''Set rst = dbs.OpenRecordset("649Drawings")
    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12 from ToutouRien ORDER BY DrawDate ASC")
    
    DoToutouRienDatesCount = rst.RecordCount
       
    Set rst = Nothing
    dbs.Close
    Set dbs = Nothing

    End Function