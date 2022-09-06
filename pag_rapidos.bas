Attribute VB_Name = "Módulo1"
' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub pag_rapidos()

Dim ultLinea As Long
Dim cant As Variant
Dim nombre As Variant

nombre = Sheets("Pagar rápidos").Cells(5, 2)
codigo = Sheets("Pagar rápidos").Cells(5, 3)
cant = Sheets("Pagar rápidos").Cells(5, 4)

ws_output = "Pagos rápidos"
Sheets(ws_output).Unprotect Password:=""
ultLinea = Sheets(ws_output).Range("I" & Rows.Count).End(xlUp).Row + 1
Sheets(ws_output).Cells(ultLinea, 10).Value = Now
Sheets(ws_output).Cells(ultLinea, 11).Value = nombre
Sheets(ws_output).Cells(ultLinea, 9).Value = codigo
Sheets(ws_output).Cells(ultLinea, 12).Value = cant
Sheets(ws_output).Cells(ultLinea, 13).Value = 1000
Sheets(ws_output).Protect Password:="", AllowFiltering:=True

Sheets("Pagar rápidos").Range("B5:D5").ClearContents
ActiveWorkbook.Save
End Sub

