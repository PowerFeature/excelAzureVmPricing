Dim responses(10) As String
Dim regionSplit() As String
Dim tempResponse As String
Dim xmlhttp As Object
Dim result As String
Dim Rows() As String
Dim cols() As String



Public Function httpclient(mincores As Integer, minram As Integer, region As String)
Dim xmlhttp As New XMLHTTP60
Dim myurl As String
myurl = "http://vmsize.azurewebsites.net/api/values/csv?minCores=" & mincores & "&minRam=" & minram & "&ri=" & ri & "&region=" & region
xmlhttp.Open "GET", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send ""
httpclient = xmlhttp.responseText
End Function

Function addResponse(response As String, region As String)
For i = LBound(responses) To UBound(responses)
    'Find Empty response
    If (responses(i) = "") Then
        responses(i) = region & "*" & response
        Exit For
    End If
Next i

End Function
Function findResponse(region As String)

For i = LBound(responses) To UBound(responses)
    'Find Empty response
    If (responses(i) = "") Then
        ' No region match get region
        tempResponse = httpclient(0, 0, region)
        ok = addResponse(tempResponse, region)
        findResponse = tempResponse
        Exit For
        
    End If
    regionSplit() = Split(responses(i), "*")
    If (regionSplit(0) = region) Then
    ' Found region
    findResponse = regionSplit(1)
    Exit For
    End If
Next i


End Function


Function getVM(mincores As Integer, minram As Integer, ri As Integer, region As String)
result = findResponse(region)
Rows() = Split(result, "#")
For i = LBound(Rows) + 1 To UBound(Rows)
    cols() = Split(Rows(i), ";")
    If (cols(1) >= mincores And cols(2) >= minram And cols(4) = ri) Then
        getVM = cols(0)
        Exit For
    End If
Next i

End Function

Function getVMPriceHour(name As String, ri As Integer, region As String)
    result = findResponse(region)
    Rows() = Split(result, "#")
    For i = LBound(Rows) + 1 To UBound(Rows)
        cols() = Split(Rows(i), ";")
        If (cols(0) = name And cols(4) = ri) Then
            getVMPriceHour = Val(cols(6))
            Exit For
        End If
    Next i
End Function
