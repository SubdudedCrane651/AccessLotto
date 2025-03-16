Option Compare Database
Option Explicit

'Import the bas file from https://github.com/VBA-tools/VBA-JSON?form=MG0AV3
'Add reference to Microsoft Scripting Runtime

Sub PickLottoNumbers()
    Dim url As String
    Dim data As Collection
    Dim numbers As Collection
    Dim lottoType As Integer
    Dim retries As Integer
    Dim matchCount As Integer
    
    ' Get lotto type from user
    lottoType = InputBox("Enter Lotto Type: (1=Lotto 6/49, 2=LottoMax, 3=Grande Vie, 4=Tout Ou Rien)")
    Select Case lottoType
        Case 1
            url = "https://richard-perreault.com/Documents/Lotto649.json"
            Debug.Print "Lotto 6/49 Selected"
        Case 2
            url = "https://richard-perreault.com/Documents/LottoMax.json"
            Debug.Print "LottoMax Selected"
        Case 3
            url = "https://richard-perreault.com/Documents/Grande_Vie.json"
            Debug.Print "Grande Vie Selected"
        Case 4
            url = "https://richard-perreault.com/Documents/ToutouRien.json"
            Debug.Print "Tout ou Rien Selected"
        Case Else
            MsgBox "Invalid input. Please enter a number between 1 and 4."
            Exit Sub
    End Select
    
    ' Fetch JSON data from URL
    Set data = GetJSONData(url)
    
    Retry :
    retries = retries + 1
    Debug.Print "Retry #" & retries
    
    ' Generate lotto numbers based on lotto type
    Dim maxNumbers As Integer, maxValue As Integer, generateGrandNumber As Boolean
    Select Case lottoType
        Case 1 ' Lotto 6/49
            maxNumbers = 6
            maxValue = 49
            generateGrandNumber = False
        Case 2 ' LottoMax
            maxNumbers = 7
            maxValue = 50
            generateGrandNumber = False
        Case 3 ' Grande Vie
            maxNumbers = 5
            maxValue = 49
            generateGrandNumber = True ' Grande Vie includes a Grand Number (1 to 7)
        Case 4 ' Tout ou Rien
            maxNumbers = 12
            maxValue = 24
            generateGrandNumber = False
    End Select
    
    ' Generate lotto numbers
    Set numbers = GenerateLottoNumbers(maxNumbers, maxValue, generateGrandNumber)
    
    ' Compare the numbers with existing data
    matchCount = CountMatches(numbers, data, lottoType)
    If matchCount >= 8 Then
        Debug.Print "High number of matches found! Restarting process..."
        Goto Retry
    End If
    
    ' Output the unique numbers
    Dim output As String
    Dim num As Variant
    output = ""
    For Each num In numbers
        output = output & num & " "
    Next num
    
    ' Trim extra space and display the formatted output
    MsgBox "Unique Lotto Numbers: " & Trim(output)
End Sub

' Generates a unique random combination of numbers with optional Grand Number
Function GenerateLottoNumbers(totalNumbers As Integer, maxValue As Integer, includeGrandNumber As Boolean) As Collection
    Dim numbers As New Collection
    Dim rndNumber As Integer
    Dim i As Integer
    
    ' Generate unique numbers for the main draw
    For i = 1 To totalNumbers
        Do
            rndNumber = Int(maxValue * Rnd + 1) ' Random number between 1 and maxValue
        Loop Until Not IsInCollection(numbers, rndNumber)
        numbers.Add rndNumber
    Next i
    
    ' Add Grand Number if required
    If includeGrandNumber Then
        Do
            rndNumber = Int(7 * Rnd + 1) ' Random number between 1 and 7 (Grand Number)
        Loop Until Not IsInCollection(numbers, rndNumber)
        numbers.Add rndNumber ' Add Grand Number as the last number
    End If
    
    ' Sort the main numbers (but keep the Grand Number last)
    Dim sortedNumbers As Collection
    Dim tempNumbers As New Collection
    Dim lastNumber As Integer
    
    ' Remove the last number (Grand Number) temporarily if applicable
    If includeGrandNumber Then
        lastNumber = numbers(numbers.count)
        numbers.Remove numbers.count
    End If
    
    ' Sort the main numbers
    Set sortedNumbers = SortNumbers(numbers)
    
    ' Restore the Grand Number as the last element
    If includeGrandNumber Then
        sortedNumbers.Add lastNumber
    End If
    
    ' Return the sorted numbers
    Set GenerateLottoNumbers = sortedNumbers
End Function

' Check if value exists in the collection
Function IsInCollection(coll As Collection, val As Integer) As Boolean
    Dim item As Variant
    On Error Resume Next
    For Each item In coll
        If item = val Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    IsInCollection = False
End Function

' Sort the lotto numbers
Function SortNumbers(numbers As Collection) As Collection
    Dim i As Integer, j As Integer
    Dim temp As Integer
    Dim tempArray() As Integer
    ReDim tempArray(1 To numbers.count)
    
    ' Transfer collection to array
    For i = 1 To numbers.count
        tempArray(i) = numbers(i)
    Next i
    
    ' Perform bubble sort
    For i = LBound(tempArray) To UBound(tempArray) - 1
        For j = i + 1 To UBound(tempArray)
            If tempArray(i) > tempArray(j) Then
                temp = tempArray(i)
                tempArray(i) = tempArray(j)
                tempArray(j) = temp
            End If
        Next j
    Next i
    
    ' Transfer back to a collection
    Dim sortedNumbers As New Collection
    For i = LBound(tempArray) To UBound(tempArray)
        sortedNumbers.Add tempArray(i)
    Next i
    
    ' Return the sorted collection
    Set SortNumbers = sortedNumbers
End Function

' Counts matching numbers between two collections
Function CountMatches(numbers As Collection, data As Collection, lottoType As Integer) As Integer
    Dim matchCount As Integer
    matchCount = 0
    
    Dim pan As Object
    Dim drawnNumbers As Collection
    Dim i As Integer
    
    ' Loop through the data to compare
    For Each pan In data
        ' Extract lotto numbers based on the type of lotto
        Set drawnNumbers = New Collection
        Select Case lottoType
            Case 1 ' Lotto 6/49
                For i = 1 To 6
                    drawnNumbers.Add pan("P" & i)
                Next i
            Case 2 ' LottoMax
                For i = 1 To 7
                    drawnNumbers.Add pan("P" & i)
                Next i
            Case 3 ' Grande Vie
                For i = 1 To 5
                    drawnNumbers.Add pan("p" & i)
                Next i
                drawnNumbers.Add pan("gn") ' Add Grand Number
            Case 4 ' Tout ou Rien
                For i = 1 To 12
                    drawnNumbers.Add pan("p" & i)
                Next i
        End Select
        
        ' Compare the collections
        Dim num As Variant
        Dim drawNum As Variant
        For Each num In numbers
            For Each drawNum In drawnNumbers
                If num = drawNum Then
                    matchCount = matchCount + 1
                End If
            Next drawNum
        Next num
        
        ' Check for duplicate or high match count
        If matchCount >= 8 Then
            CountMatches = matchCount
            Exit Function
        End If
    Next pan
    
    CountMatches = 0
End Function

' Fetch JSON data from the URL
Function GetJSONData(url As String) As Collection
    Dim http As Object
    Dim JSON As Object
    Dim results As New Collection
    
    ' Create the HTTP object for request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send
    
    ' Parse JSON response using JsonConverter
    Set JSON = JsonConverter.ParseJson(http.responseText)
    
    ' Loop through JSON and add records to the collection
    Dim item As Object
    For Each item In JSON
        results.Add item
    Next item
    
    ' Return the parsed JSON data
    Set GetJSONData = results
End Function