Sub Lotto649Special()
    Dim dbs As DAO.Database, rst As DAO.Recordset
    Dim pick6(6) As Integer, P(6) As Integer
    Dim DrawDate As Date
    Dim ChangeCaption As Integer
    Dim count2 As Integer, RndNum As Integer, temp As Integer
    Dim FSO As Object, Fileout As Object
    Dim Label35 As Object ' Ensure Label35 exists as an object for UI updates
    Dim DrawCount As Integer

    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("SELECT DrawDate, p1, p2, p3, p4, p5, p6 FROM 649Drawings ORDER BY DrawDate ASC")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\test.txt", True, True)

    Randomize

    ' Generate unique random numbers
    For count2 = 1 To 6
        DoEvents
        
        ' Implement spinner effect once per number pick
        ChangeCaption = ChangeCaption Mod 6 + 1
        Select Case ChangeCaption
            Case 1 : Me.Label35.Caption = "\"
            Case 2 : Me.Label35.Caption = "-"
            Case 3 : Me.Label35.Caption = "/"
            Case 4 : Me.Label35.Caption = "-"
            Case 5 : Me.Label35.Caption = "\"
            Case 6 : Me.Label35.Caption = "|"
        End Select

        Do
            RndNum = Int((49 - 1 + 1) * Rnd + 1)
        Loop While Not IsUnique(RndNum, pick6, count2)
        pick6(count2) = RndNum
    Next

    ' Sort numbers using Bubble Sort
    BubbleSort pick6

    DrawCount = 0 ' Initialize draw counter

    ' Scan through the records
    Do While Not rst.EOF
        DrawDate = rst("DrawDate")
        P(1) = rst("p1") : P(2) = rst("p2") : P(3) = rst("p3")
        P(4) = rst("p4") : P(5) = rst("p5") : P(6) = rst("p6")

        count2 = CountMatches(P, pick6)
        DrawCount = DrawCount + 1 ' Increment draw count

        ' Check for 3 or more occurrences and display output immediately
        If count2 >= 3 Then
            Fileout.WriteLine DrawDate
            'MsgBox "Matching draw found! Date: " & DrawDate, vbInformation, "Lotto Match"
        End If

        rst.MoveNext
    Loop
    
    Me.Label35.Caption = " "
    Call Shell("c:/Users/rchrd/Documents/Python/CallerIDGUI/.venv/Scripts/python.exe C:\Users\rchrd\Documents\Python\text2speech.py ""--lang=fr"" ""Voici les numéros gagnants de lotto 6/49""", vbNormalFocus)

    ' Show message box when no exact match is found
    MsgBox "No exact match found." & vbCrLf & _
        "Generated Numbers: " & JoinArray(pick6) & vbCrLf & _
        "Total Drawings Processed: " & DrawCount, vbInformation, "Results"

    Fileout.Close
    rst.Close
    dbs.Close

End Sub

Function IsUnique(num As Integer, arr() As Integer, count As Integer) As Boolean
    Dim i As Integer
    For i = 1 To count - 1
        If arr(i) = num Then IsUnique = False : Exit Function
    Next
    IsUnique = True
End Function

Sub BubbleSort(arr() As Integer)
    Dim i As Integer, j As Integer, temp As Integer
    For i = 1 To UBound(arr) - 1
        For j = 1 To UBound(arr) - i
            If arr(j) > arr(j + 1) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next
    Next
End Sub

Function CountMatches(arr1() As Integer, arr2() As Integer) As Integer
    Dim i As Integer, j As Integer
    CountMatches = 0
    For i = 1 To 6
        For j = 1 To 6
            If arr1(i) = arr2(j) Then CountMatches = CountMatches + 1
        Next
    Next
End Function

Function JoinArray(arr() As Integer) As String
    Dim result As String, i As Integer
    result = ""
    For i = 1 To UBound(arr)
        result = result & arr(i) & " "
    Next
    JoinArray = Trim(result)
End Function

Sub Tout_ou_RienSpecial()
    Dim dbs As DAO.Database, rst As DAO.Recordset
    Dim pick12(1 To 12) As Integer, P(1 To 12) As Integer
    Dim DrawDate As Date, ChangeCaption As Integer
    Dim count2 As Integer, RndNum As Integer
    Dim FSO As Object, Fileout As Object
    Dim Label23 As Object
    Dim DrawCount As Integer
    
    On Error Resume Next

    ' Initialize database and file system
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("SELECT DrawDate, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12 FROM ToutouRien ORDER BY DrawDate ASC")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\unique_combinations.txt", True, True)

    Randomize

    ' Generate unique numbers with spinner effect
    For count2 = 1 To 12
        DoEvents
        ' Implement spinner effect once per number pick
        ChangeCaption = ChangeCaption Mod 6 + 1
        Select Case ChangeCaption
            Case 1 : Me.Label23.Caption = "\"
            Case 2 : Me.Label23.Caption = "-"
            Case 3 : Me.Label23.Caption = "/"
            Case 4 : Me.Label23.Caption = "-"
            Case 5 : Me.Label23.Caption = "\"
            Case 6 : Me.Label23.Caption = "|"
        End Select

        Do
            RndNum = Int((25 - 1 + 1) * Rnd + 1)
        Loop While Not IsUnique(RndNum, pick12, count2)

        pick12(count2) = RndNum
    Next

    DrawCount = 0

    ' Scan database for duplicate combinations
    Do
        GenerateUniqueCombination pick12, ChangeCaption

        count2 = 0
        rst.MoveFirst ' Restart scan
        
        DrawCount = 0
        
        Do While Not rst.EOF
            LoadRecord rst, P, DrawDate
            
            DrawCount = DrawCount + 1 ' ? Increment draw count **inside the loop**

            If ArraysMatch(pick12, P) Then
                Debug.Print "Duplicate found! Generating new numbers..."
                Exit Do ' Simply restart the outer loop
            End If
            rst.MoveNext
        Loop

    Loop Until Not ArraysMatch(pick12, P) ' Ensures we retry if a match is found

    ' Write unique combination to the file
    Fileout.WriteLine Join(pick12, ", ")
    Debug.Print "Unique combination: " & Join(pick12, ", ")

    ' Announce and display final results
    Me.Label23.Caption = " "
    Call Shell("c:/Users/rchrd/Documents/Python/CallerIDGUI/.venv/Scripts/python.exe C:\Users\rchrd\Documents\Python\text2speech.py ""--lang=fr"" ""Voici les numéros gagnants de Tout ou rien""", vbNormalFocus)

    Dim pick12Strings(1 To 12) As String
    Dim i As Integer

    For i = 1 To 12
        pick12Strings(i) = CStr(pick12(i)) ' Convert each number to a string
    Next i

    MsgBox "Unique combination: " & Join(pick12Strings, " ") & vbCrLf & "Total Drawings Processed: " & DrawCount, vbInformation, "Results"

    Fileout.Close
    rst.Close
    dbs.Close

End Sub
' Generates a unique random combination of 12 numbers
Sub GenerateUniqueCombination(ByRef pick12() As Integer, ByRef ChangeCaption As Integer)
    Dim i As Integer, RndNum As Integer

    For i = 1 To 12
        Do
            RndNum = Int((24 - 1 + 1) * Rnd + 1) ' Random number between 1 and 24
        Loop Until Not IsInArray(RndNum, pick12) ' Ensure no duplicates
        pick12(i) = RndNum
        UpdateSpinner ChangeCaption
    Next i

    ' Sort the combination for consistency
    QuickSort pick12, LBound(pick12), UBound(pick12)
End Sub

' Updates the spinner to show progress
Sub UpdateSpinner(ByRef ChangeCaption As Integer)
    Dim spinner As Variant
    spinner = Array("\", "|", "/", "-")
    Label23.Caption = spinner(ChangeCaption Mod 4)
    ChangeCaption = ChangeCaption + 1
    DoEvents ' Keep UI responsive
End Sub

' Loads a record from the database into the array
Sub LoadRecord(ByRef rst As DAO.Recordset, ByRef arr() As Integer, ByRef dateOut As Date)
    Dim i As Integer
    dateOut = rst("DrawDate")
    For i = 1 To 12
        arr(i) = Nz(rst("p" & i), 0)
    Next i
End Sub

' Checks if two arrays are identical
Function ArraysMatch(arr1() As Integer, arr2() As Integer) As Boolean
    Dim i As Integer
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) <> arr2(i) Then
            ArraysMatch = False
            Exit Function
        End If
    Next i
    ArraysMatch = True
End Function

' Checks if a value exists in an array
Function IsInArray(val As Integer, ByRef arr() As Integer) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
' Counts matching numbers between two arrays
Function CountMatchingNumbers(arr1() As Integer, arr2() As Integer) As Integer
    Dim i As Integer, j As Integer, count As Integer
    count = 0
    For i = LBound(arr1) To UBound(arr1)
        For j = LBound(arr2) To UBound(arr2)
            If arr1(i) = arr2(j) Then
                count = count + 1
                Exit For ' Stop further checks once a match is found
            End If
        Next j
    Next i
    CountMatchingNumbers = count
End Function

' Sorts an array using QuickSort
Sub QuickSort(ByRef arr() As Integer, ByVal first As Integer, ByVal last As Integer)
    Dim low As Integer, high As Integer, pivot As Integer, temp As Integer
    low = first
    high = last
    pivot = arr((first + last) \ 2)

    Do While low <= high
        Do While arr(low) < pivot : low = low + 1 : Loop
        Do While arr(high) > pivot : high = high - 1 : Loop

        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub



Sub Grande_VieSpecial()
    
    Dim identical As Boolean
    Dim P(5), GN(1) As Integer
    Dim pick5(5), pickGN(1) As Integer
    Dim DrawDate As Date
    Dim P1, P2, P3, P4, P5, count1, count3, count4, GN2, RndNum, temp As Integer
    Dim str1, str2, str3, str4, str5, str6, str7, str8, str9, str10, str11, str12 As String
    Dim ChangeCaption As Integer
    Dim dbs As Database, rst As DAO.Recordset
    Dim rst2 As DAO.Recordset
    Set dbs = CurrentDb
    Dim strPath As String
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\test.txt", True, True)
    
    On Error Resume Next

    Set rst = dbs.OpenRecordset("SELECT DrawDate,p1,p2,p3,p4,p5,gn from Grande_Vie ORDER BY DrawDate ASC")
    ChangeCaption = 1

    10 : rst.MoveFirst
    
    Randomize

    For count3 = 1 To 5

        DoEvents


        Do

            If ChangeCaption = 1 Then
                Label25.Caption = "\"
            ElseIf ChangeCaption = 2 Then
                Label25.Caption = "-"
            ElseIf ChangeCaption = 3 Then
                Label25.Caption = "/"
            ElseIf ChangeCaption = 4 Then
                Label25.Caption = "-"
            ElseIf ChangeCaption = 5 Then
                Label25.Caption = "\"
            ElseIf ChangeCaption = 6 Then
                Label25.Caption = "|"
                ChangeCaption = 0
            End If

            ChangeCaption = ChangeCaption + 1

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
            Goto 10
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

            '4 numbers and more the same
            If count2 >= 4 Then
                identical = True
                Fileout.Write DrawDate
                Goto 10
            Else
                identical = False
            End If

        Next
        rst.MoveNext
    Loop
    
    Label25.Caption = " "
    Call Shell("c:/Users/rchrd/Documents/Python/CallerIDGUI/.venv/Scripts/python.exe C:\Users\rchrd\Documents\Python\text2speech.py ""--lang=fr"" ""Voici les numéros gagnants de la Grande Vie""", vbNormalFocus)
    MsgBox(str1 + " " + str2 + " " + str3 + " " + str4 + " " + str5 + " " + str6)
    
    Set rst = Nothing
    Set rst2 = Nothing
    dbs.Close
    Set dbs = Nothing
    Fileout.Close
    
End Sub

Sub LottoMaxSpecial()
    Dim dbs As DAO.Database, rst As DAO.Recordset
    Dim pick7(7) As Integer, P(7) As Integer
    Dim DrawDate As Date, ChangeCaption As Integer
    Dim count2 As Integer, RndNum As Integer, temp As Integer
    Dim FSO As Object, Fileout As Object
    Dim Label33 As Object
    Dim DrawCount As Integer

    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("SELECT DrawDate, p1, p2, p3, p4, p5, p6, p7 FROM LottoMax ORDER BY DrawDate ASC")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Fileout = FSO.CreateTextFile("C:\Users\rchrd\Documents\Richard\test.txt", True, True)

    Randomize

    ' Generate unique numbers with spinner effect
    For count2 = 1 To 7
        DoEvents
        ' Implement spinner effect once per number pick
        ChangeCaption = ChangeCaption Mod 6 + 1
        Select Case ChangeCaption
            Case 1 : Me.Label33.Caption = "\"
            Case 2 : Me.Label33.Caption = "-"
            Case 3 : Me.Label33.Caption = "/"
            Case 4 : Me.Label33.Caption = "-"
            Case 5 : Me.Label33.Caption = "\"
            Case 6 : Me.Label33.Caption = "|"
        End Select

        Do
            RndNum = Int((50 - 1 + 1) * Rnd + 1)
        Loop While Not IsUnique(RndNum, pick7, count2)

        pick7(count2) = RndNum
    Next

    ' Sort numbers
    BubbleSort pick7

    DrawCount = 0

    Do While Not rst.EOF
        DrawDate = rst("DrawDate")
        P(1) = rst("p1") : P(2) = rst("p2") : P(3) = rst("p3")
        P(4) = rst("p4") : P(5) = rst("p5") : P(6) = rst("p6") : P(7) = rst("p7")

        count2 = CountMatches(P, pick7)
        DrawCount = DrawCount + 1

        If count2 >= 3 Then
            Fileout.WriteLine DrawDate
        End If

        rst.MoveNext
    Loop
    
    Me.Label33.Caption = " "
    Call Shell("c:/Users/rchrd/Documents/Python/CallerIDGUI/.venv/Scripts/python.exe C:\Users\rchrd\Documents\Python\text2speech.py ""--lang=fr"" ""Voici les numéros gagnants de LottoMax""", vbNormalFocus)

    MsgBox "No exact match found." & vbCrLf & _
        "Generated Numbers: " & JoinArray(pick7) & vbCrLf & _
        "Total Drawings Processed: " & DrawCount, vbInformation, "Results"

    Fileout.Close
    rst.Close
    dbs.Close

End Sub