Dim responses(100) As String
Dim regionSplit() As String
Dim tempResponse As String
Dim xmlhttp As Object
Dim result As String
Dim Rows() As String
Dim Labels() As String

Dim cols() As String
Public Function httpclient(region As String, xCurrency, Optional ByVal isManagedDisk As Boolean)
Dim xmlhttp As New XMLHTTP60
Dim myurl As String
If (isManagedDisk) Then
myurl = "http://vmsizecdn.azureedge.net/api/values/csv/mdisks?seed=12&region=" & region & "&currency=" & xCurrency

Else
myurl = "http://vmsizecdn.azureedge.net/api/values/csv?seed=12&minCores=" & mincores & "&minRam=" & minram & "&region=" & region & "&currency=" & xCurrency

End If


xmlhttp.Open "GET", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send ""
httpclient = xmlhttp.responseText
End Function

Private Function OnTimeOutMessage()
   
    'MsgBox ("Server error: request time-out")
End Function

Function addResponse(response As String, region As String, xCurrency As String, Optional ByVal managedDisk As Boolean = False)

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
        responses(i) = searchString & LCase(xCurrency) & "*" & response
        Exit For
    End If
Next i

End Function
Function findResponse(region As String, xCurrency As String, Optional ByVal managedDisk As Boolean = False)
Dim i As Integer

For i = LBound(responses) To UBound(responses)
    'Find Empty response
    If (responses(i) = "") Then
        ' No region match get region
        tempResponse = httpclient(region, xCurrency, managedDisk)
        ok = addResponse(tempResponse, region, xCurrency, managedDisk)
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
    
    
    If (regionSplit(0) = searchString & LCase(xCurrency)) Then
    ' Found region
    findResponse = regionSplit(1)
    Exit For
    End If
Next i


End Function


Function getVM(mincores As Integer, minram As Integer, ri As Integer, region As String, xCurrency As String, Optional ByVal ex As String = "", Optional ByVal incl As String = "")
result = findResponse(region, xCurrency)
Rows() = Split(result, vbCrLf)
For i = LBound(Rows) + 1 To UBound(Rows)
    cols() = Split(Rows(i), ";")
    If (cols(1) >= mincores And cols(2) >= minram And cols(4) = ri And searchKeywords(cols(0), ex) = False And incl = "") Then
        getVM = cols(0)
        Exit For
    ElseIf (cols(1) >= mincores And cols(2) >= minram And cols(4) = ri And searchKeywords(cols(0), ex) = False And incl <> "" And searchKeywords(cols(0), incl) = True) Then
        getVM = cols(0)
        Exit For
    End If
Next i
End Function

Function getManagedDisk(minSize As Integer, region As String, xCurrency As String, Optional ByVal ex As String = "", Optional ByVal incl As String = "")
result = findResponse(region, xCurrency, True)
Rows() = Split(result, vbCrLf)
For i = LBound(Rows) + 1 To UBound(Rows)
    cols() = Split(Rows(i), ";")
    ' nothing in incl
    If (cols(1) >= minSize And searchKeywords(cols(0), ex) = False And incl = "") Then
        getManagedDisk = cols(0)
        Exit For
    ' something in incl
    ElseIf (cols(1) >= minSize And searchKeywords(cols(0), ex) = False And incl <> "" And searchKeywords(cols(0), incl) = True) Then
        getManagedDisk = cols(0)
        Exit For
    End If
Next i


End Function

Function getManagedDiskPriceMonth(name As String, region As String, xCurrency As String)
    result = findResponse(region, xCurrency, True)
    Rows() = Split(result, vbCrLf)
    For i = LBound(Rows) + 1 To UBound(Rows)
        cols() = Split(Rows(i), ";")
        If (cols(0) = name) Then
            getManagedDiskPriceMonth = Val(cols(4))
            Exit For
        End If
    Next i
End Function

Function getVMPriceHour(name As String, ri As Integer, region As String, xCurrency As String)
    result = findResponse(region, xCurrency)
    Rows() = Split(result, vbCrLf)
    For i = LBound(Rows) + 1 To UBound(Rows)
        cols() = Split(Rows(i), ";")
        If (cols(0) = name And cols(4) = ri) Then
            getVMPriceHour = Val(cols(6))
            Exit For
        End If
    Next i
End Function

Function getVMData(name As String, region As String, xCurrency As String, ParamName As String)
result = findResponse(region, xCurrency)
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
Function getMDiskData(name As String, region As String, xCurrency As String, ParamName As String)
result = findResponse(region, xCurrency, True)
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
