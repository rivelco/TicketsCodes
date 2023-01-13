' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub pag_rapidos()

Dim ultLinea As Long
Dim cant As Variant
Dim nombre As Variant

nombre = Sheets("Pagar rápidos").Cells(5, 2)
codigo = Sheets("Pagar rápidos").Cells(5, 3)
cant = Sheets("Pagar rápidos").Cells(5, 4)
cost = Sheets("Pagar rápidos").Cells(5, 5)

If IsEmpty(nombre) Or IsEmpty(codigo) Or IsEmpty(cant) Or IsEmpty(cost) Then
    MsgBox "Faltan campos por completar. No hice nada."
    Exit Sub
End If

ws_output = "Pagos rápidos"
Sheets(ws_output).Unprotect Password:=""
ultLinea = Sheets(ws_output).Range("I" & Rows.Count).End(xlUp).Row + 1
Sheets(ws_output).Cells(ultLinea, 10).Value = Now
Sheets(ws_output).Cells(ultLinea, 11).Value = nombre
Sheets(ws_output).Cells(ultLinea, 9).Value = codigo
Sheets(ws_output).Cells(ultLinea, 12).Value = cant
Sheets(ws_output).Cells(ultLinea, 13).Value = cant * cost
Sheets(ws_output).Protect Password:="", AllowFiltering:=True

Sheets("Pagar rápidos").Range("B5:E5").ClearContents
ActiveWorkbook.Save
End Sub

