Attribute VB_Name = "Módulo3"
' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub reg_venta_rapidos()
ws_output = "Venta rápidos"
Sheets(ws_output).Unprotect Password:=""
next_row = Sheets(ws_output).Range("A" & Rows.Count).End(xlUp).Offset(1).Row
Sheets(ws_output).Cells(next_row, 1).Value = Now
Sheets(ws_output).Cells(next_row, 2).Value = Range("rap_venta_nom").Value
Sheets(ws_output).Cells(next_row, 3).Value = Range("rap_venta_cant").Value
Sheets(ws_output).Protect Password:="", AllowFiltering:=True

Dim cont As Long
Dim ultLinea As Long
Dim cant As Variant
Dim nombre As Variant

Sheets("Info rápidos").Unprotect Password:=""
ultLinea = Sheets("Info rápidos").Range("A" & Rows.Count).End(xlUp).Row
nombre = Sheets("Vender").Cells(6, 2)
cant = Sheets("Vender").Cells(6, 3)
For cont = 1 To ultLinea
    current = Sheets("Info rápidos").Cells(cont, 1)
    If nombre = current Then
        anterior = Sheets("Info rápidos").Cells(cont, 4)
        Sheets("Info rápidos").Cells(cont, 4) = anterior - cant
    End If
Next cont
Sheets("Info rápidos").Protect Password:="", AllowFiltering:=True

Range("rap_venta_nom").ClearContents
Range("rap_venta_cant").Clear
ActiveWorkbook.Save
End Sub
