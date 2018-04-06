Dim responses(100) As String
Dim regionSplit() As String
Dim tempResponse As String
Dim xmlhttp As Object
Dim result As String
Dim Rows() As String
Dim Labels() As String

Dim cols() As String
Public Function httpclient(mincores As Integer, minram As Integer, region As String, xCurrency)
Dim xmlhttp As New XMLHTTP60
'xmlhttp.setTimeouts 10000, 10000, 10000, 10000
'xmlhttp.OnTimeOut = OnTimeOutMessage 'callback function
Dim myurl As String
myurl = "http://vmsizecdn.azureedge.net/api/values/csv?test=3242&minCores=" & mincores & "&minRam=" & minram & "&region=" & region & "&currency=" & xCurrency
xmlhttp.Open "GET", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send ""
httpclient = xmlhttp.responseText
End Function
Private Function OnTimeOutMessage()
   
    'MsgBox ("Server error: request time-out")
End Function

Function addResponse(response As String, region As String, xCurrency As String)

Dim i As Integer

For i = LBound(responses) To UBound(responses)
    'Find Empty response
    If (responses(i) = "") Then
        responses(i) = LCase(region) & LCase(xCurrency) & "*" & response
        Exit For
    End If
Next i

End Function
Function findResponse(region As String, xCurrency As String)
Dim i As Integer

For i = LBound(responses) To UBound(responses)
    'Find Empty response
    If (responses(i) = "") Then
        ' No region match get region
        tempResponse = httpclient(0, 0, region, xCurrency)
        ok = addResponse(tempResponse, region, xCurrency)
        findResponse = tempResponse
        Exit For
        
    End If
    regionSplit() = Split(responses(i), "*")
    If (regionSplit(0) = LCase(region) & LCase(xCurrency)) Then
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
Function getVMPriceHourWin(name As String, ri As Integer, region As String, xCurrency As String)
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
