Dim responses(100) As String
Dim regionSplit() As String
Dim tempResponse As String
Dim xmlhttp As Object
Dim result As String
Dim Rows() As String
Dim Labels() As String

Dim cols() As String
Public Function httpclient(region As String, CurrencyID, Optional ByVal isManagedDisk As Boolean)
Dim xmlhttp As New XMLHTTP60
Dim myurl As String
If (isManagedDisk) Then
myurl = "https://vmsizecdn.azureedge.net/api/values/csv/mdisks?seed=20&region=" & region & "&currency=" & CurrencyID

Else
myurl = "https://vmsizecdn.azureedge.net/api/values/csv?seed=20&minCores=" & mincores & "&minRam=" & minram & "&region=" & region & "&currency=" & CurrencyID

End If


xmlhttp.Open "GET", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send ""
httpclient = xmlhttp.responseText
End Function

Private Function OnTimeOutMessage()
   
    'MsgBox ("Server error: request time-out")
End Function

Function addResponse(response As String, region As String, CurrencyID As String, Optional ByVal managedDisk As Boolean = False)

Dim i As Integer
    Dim searchString As String
    If (managedDisk) Then
    searchString = "MDISK" & LCase(region)
    Else
    searchString = LCase(region)
    End If
    
For i = LBound(responses) To UBound(responses)
    'Find Empty response
    If (responses(i) = "") Then
        responses(i) = searchString & LCase(CurrencyID) & "*" & response
        Exit For
    End If
Next i

End Function
Function findResponse(region As String, CurrencyID As String, Optional ByVal managedDisk As Boolean = False)
Dim i As Integer

For i = LBound(responses) To UBound(responses)
    'Find Empty response
    If (responses(i) = "") Then
        ' No region match get region
        tempResponse = httpclient(region, CurrencyID, managedDisk)
        ok = addResponse(tempResponse, region, CurrencyID, managedDisk)
        findResponse = tempResponse
        Exit For
        
    End If
    regionSplit() = Split(responses(i), "*")
    Dim searchString As String
    If (managedDisk) Then
    searchString = "MDISK" & LCase(region)
    Else
    searchString = LCase(region)
    End If
    
    
    If (regionSplit(0) = searchString & LCase(CurrencyID)) Then
    ' Found region
    findResponse = regionSplit(1)
    Exit For
    End If
Next i


End Function


Function getVM(mincores As Integer, minram As Integer, reservedInstanceYears As Integer, region As String, CurrencyID As String, Optional ByVal excludeKeywords As String = "", Optional ByVal includeKeywords As String = "")
result = findResponse(region, CurrencyID)
Rows() = Split(result, vbCrLf)
For i = LBound(Rows) + 1 To UBound(Rows)
    cols() = Split(Rows(i), ";")
    If (cols(1) >= mincores And cols(2) >= minram And cols(4) = reservedInstanceYears And searchKeywords(cols(0), excludeKeywords) = False And includeKeywords = "") Then
        If (cols(13) = "True") Then
            Application.Caller.Font.ColorIndex = 3
        Else
            Application.Caller.Font.ColorIndex = 1
        End If
        getVM = cols(0)
        Exit For
    ElseIf (cols(1) >= mincores And cols(2) >= minram And cols(4) = reservedInstanceYears And searchKeywords(cols(0), excludeKeywords) = False And includeKeywords <> "" And searchKeywords(cols(0), includeKeywords) = True) Then
        If (cols(13) = "True") Then
            Application.Caller.Font.ColorIndex = 3
        Else
            Application.Caller.Font.ColorIndex = 1
        End If
        getVM = cols(0)
        Exit For
    End If
Next i
End Function

Function getManagedDisk(minSize As Integer, region As String, CurrencyID As String, Optional ByVal excludeKeywords As String = "", Optional ByVal includeKeywords As String = "")
result = findResponse(region, CurrencyID, True)
Rows() = Split(result, vbCrLf)
For i = LBound(Rows) + 1 To UBound(Rows)
    cols() = Split(Rows(i), ";")
    ' nothing in includeKeywords
    If (cols(1) >= minSize And searchKeywords(cols(0), excludeKeywords) = False And includeKeywords = "") Then
        If (cols(12) = "True") Then
            Application.Caller.Font.ColorIndex = 3
        Else
            Application.Caller.Font.ColorIndex = 1
        End If
        getManagedDisk = cols(0)
        Exit For
    ' something in includeKeywords
    ElseIf (cols(1) >= minSize And searchKeywords(cols(0), excludeKeywords) = False And includeKeywords <> "" And searchKeywords(cols(0), includeKeywords) = True) Then
        If (cols(12) = "True") Then
            Application.Caller.Font.ColorIndex = 3
        Else
            Application.Caller.Font.ColorIndex = 1
        End If
        getManagedDisk = cols(0)
        Exit For
    End If
Next i


End Function

Function getManagedDiskPriceMonth(name As String, region As String, CurrencyID As String)
    result = findResponse(region, CurrencyID, True)
    Rows() = Split(result, vbCrLf)
    For i = LBound(Rows) + 1 To UBound(Rows)
        cols() = Split(Rows(i), ";")
        If (cols(0) = name) Then
                If (cols(12) = "True") Then
            Application.Caller.Font.ColorIndex = 3
        Else
            Application.Caller.Font.ColorIndex = 1
        End If
            getManagedDiskPriceMonth = Val(cols(4))
            Exit For
        End If
    Next i
End Function

Function getVMPriceHour(name As String, reservedInstanceYears As Integer, region As String, CurrencyID As String)
    result = findResponse(region, CurrencyID)
    Rows() = Split(result, vbCrLf)
    For i = LBound(Rows) + 1 To UBound(Rows)
        cols() = Split(Rows(i), ";")
        If (cols(0) = name And cols(4) = reservedInstanceYears) Then
        If (cols(13) = "True") Then
            Application.Caller.Font.ColorIndex = 3
        Else
            Application.Caller.Font.ColorIndex = 1
        End If
            getVMPriceHour = Val(cols(6))
            Exit For
        End If
    Next i
End Function

Function getVMData(name As String, region As String, CurrencyID As String, ParamName As String)
result = findResponse(region, CurrencyID)
'Find the param
Rows() = Split(result, vbCrLf)
Labels() = Split(Rows(0), ";")
Do

For e = LBound(Labels) To UBound(Labels)
If (LCase(Labels(e)) = LCase(ParamName)) Then
' Search through the VM's
    For i = LBound(Rows) + 1 To UBound(Rows)
        cols() = Split(Rows(i), ";")
        If (cols(0) = name) Then
        getVMData = cols(e)
        Exit Do
            
        End If
    Next i

End If
Next e
Loop While False

End Function
Function getMDiskData(name As String, region As String, CurrencyID As String, ParamName As String)
result = findResponse(region, CurrencyID, True)
'Find the param
Rows() = Split(result, vbCrLf)
Labels() = Split(Rows(0), ";")
Do

For e = LBound(Labels) To UBound(Labels)
If (LCase(Labels(e)) = LCase(ParamName)) Then
' Search through the VM's
    For i = LBound(Rows) + 1 To UBound(Rows)
        cols() = Split(Rows(i), ";")
        If (cols(0) = name) Then
        getMDiskData = cols(e)
        Exit Do
            
        End If
    Next i

End If
Next e
Loop While False

End Function

Function searchKeywords(name As String, wordlist As String)
Dim words() As String
words() = Split(wordlist, ";")
For i = LBound(words) To UBound(words)
If (InStr(name, words(i)) > 0) Then
searchKeywords = True
Exit For
End If
If (i = UBound(words)) Then
searchKeywords = False

End If

Next i

End Function



