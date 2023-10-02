Attribute VB_Name = "ImportModule"

Sub Import()

    Dim fileData, textArray, line As Object
    
    fileData = readData(buildPath("/quotes/test.txt"))
    textArray = Split(fileData, vbCrLf)
    Dim i As Integer, location As String
    For i = 0 To UBound(textArray)
        location = findLocation(textArray(i), Range("AE3"))
        If (Trim(location) <> "") Then
            writeLine textArray(i), Range(location)
        End If
    Next i

End Sub

Sub Export()
    Dim path As String, valueRange As Range
    path = buildPath("/quotes/test.txt")
    ' valueRange = findRange(Range("AE3")) 
    
    writeData path, getParamString(findRange(Range("AE3")))
    

End Sub

    
Function findRange(startRange As Range) As Range
    Dim failSafe As Integer
    Dim currentRange As Range
    
    failSafe = 0
    Set currentRange = startRange
    
    Do While failSafe < 100
        If IsEmpty(currentRange) Then
            Set findRange = Range(startRange, currentRange)
            Exit Function
        End If
        
        Set currentRange = currentRange.offset(1, 0)
        failSafe = failSafe + 1
    Loop
End Function

'functions for reading export data

    Function findLocation(line, startRange As Range)
        Dim failSafe As Integer
        Dim currentRange As Range
        Dim name As String
        Dim textArr
        textArr = Split(line, vbtab)
        if Ubound(textArr) > -1 Then
            name = Trim(textArr(0))
            failSafe = 0
            Set currentRange = startRange
            
            Do While failSafe < 100
                If IsEmpty(currentRange) Then
                    findLocation = ""
                    Exit Function
                ElseIf name = Trim(currentRange.Value) Then
                    findLocation = currentRange.Offset(0, 1).Value
                    Exit Function
                End If
                
                Set currentRange = currentRange.offset(1, 0)
                failSafe = failSafe + 1
            Loop
        End if
        
    End Function

    Function writeLine(line, location As Range)
        If (validData(line)) Then
            Dim textArr
            textArr = Split(line, vbtab)
            If (textArr(0) <> "PRICEEACH") Then
                location.value = textArr(1)
            End If
        End If
    End Function

    Function validData(data) As Boolean
        If (UBound(Split(data, vbtab)) > 0) Then
            validData = True
        End If
    End Function

'functions for gettting export data
    Function getParamString(valueRange As Range) As String
        Dim cell As Range
        
        Dim paramString As String
        paramString = ""

        For Each cell In valueRange
            If IsEmpty(cell) Then
                Exit For
            Else
                paramString = paramString + getValue(cell)
            End If
        Next cell
        getParamString = paramString

    End Function

    Function getValue(cell As Range) As String
        Dim prop As String, value As String, location As String
        
        prop = cell.value
        value = cell.offset(0, 2).value
        getValue = prop & vbtab & value & vbCrLf
    End Function

'functions for editing text files

    Function buildPath(path) As String
        buildPath = ThisWorkbook.path & path
    End Function

    Function writeData(path, newData)
        Dim fileNumber As Integer
    
        
        fileNumber = FreeFile
        Open path For Output As fileNumber

            Print #fileNumber, newData

        Close fileNumber
    End Function

    Function readData(path) As String
        Dim fileNumber As Integer
        Dim textData As String
        fileNumber = FreeFile
        Open path For Input As fileNumber
        textData = Input$(LOF(fileNumber), fileNumber)
        Close fileNumber
        readData = textData
    End Function
