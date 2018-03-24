Attribute VB_Name = "VM_Prices"
Dim result As String
Dim query As String
Dim rows() As String
Dim vms() As vm
Dim cols() As String
Dim xmlhttp As Object

Function getResult(mincores As Integer, minram As Integer, ri As Integer, region As String)

result = httpclient(mincores, minram, ri, region)

rows() = Split(result, "#")

For i = LBound(rows) To UBound(rows)
'cols = Split(rows(i))

'vms(i).Name = cols(1)
'vms(i).Cores = cols(2)
'vms(i).Ram = cols(3)
'vms(i).DiskSize = cols(4)
'vms(i).HourPrice = cols(6)
'vms(i).MonthPrice = cols(7)
'vms(i).YearPrice = cols(8)
Next i




End Function
Function getVM(mincores As Integer, minram As Integer, ri As Integer, region As String)
' This could be improved
'MsgBox (result)
If (result = "") Then
' Get new data
ok = getResult(0, 0, ri, region)

End If
rows() = Split(result, "#")

For i = LBound(rows) + 1 To UBound(rows)

cols() = Split(rows(i), ";")
If (cols(1) >= mincores And cols(2) >= minram And cols(4) = ri) Then
getVM = cols(0)
Exit For
End If
Next i
End Function
Function getVMPriceHour(mincores As Integer, minram As Integer, ri As Integer, region As String)
' This could be improved
'MsgBox (result)
If (result = "") Then
' Get new data
ok = getResult(0, 0, ri, region)

End If
rows() = Split(result, "#")

For i = LBound(rows) + 1 To UBound(rows)

cols() = Split(rows(i), ";")
If (cols(1) >= mincores And cols(2) >= minram And cols(4) = ri) Then


getVMPriceHour = Val(cols(6))
Exit For
End If
Next i
End Function


Public Function httpclient(mincores As Integer, minram As Integer, ri As Integer, region As String)

Dim xmlhttp As New XMLHTTP60
Dim myurl As String
myurl = "http://vmsize.azurewebsites.net/api/values/csv?minCores=" & mincores & "&minRam=" & minram & "&ri=" & ri & "&region=" & region
xmlhttp.Open "GET", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send ""
httpclient = xmlhttp.responseText

End Function

Function processCSV(csvInput As String)


End Function



