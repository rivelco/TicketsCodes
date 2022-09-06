Attribute VB_Name = "Módulo2"
' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub reg_rapidos()

Dim cont As Long
Dim ultLinea As Long
Dim cant As Variant
Dim nombre As Variant

Sheets("Info rápidos").Unprotect Password:=""
ultLinea = Sheets("Info rápidos").Range("A" & Rows.Count).End(xlUp).Row
nombre = Sheets("Registro rápidos").Cells(5, 2)
codigo = Sheets("Registro rápidos").Cells(5, 3)
cant = Sheets("Registro rápidos").Cells(5, 4)
For cont = 1 To ultLinea
    current = Sheets("Info rápidos").Cells(cont, 1)
    If nombre = current Then
        anterior = Sheets("Info rápidos").Cells(cont, 4)
        Sheets("Info rápidos").Cells(cont, 4) = anterior + cant
        anterior = Sheets("Info rápidos").Cells(cont, 3)
        Sheets("Info rápidos").Cells(cont, 3) = anterior + cant
    End If
Next cont
Sheets("Info rápidos").Protect Password:="", AllowFiltering:=True

ws_output = "Pagos rápidos"
Sheets(ws_output).Unprotect Password:=""
ultLinea = Sheets(ws_output).Range("A" & Rows.Count).End(xlUp).Row + 1
Sheets(ws_output).Cells(ultLinea, 1).Value = Now
Sheets(ws_output).Cells(ultLinea, 2).Value = nombre
Sheets(ws_output).Cells(ultLinea, 3).Value = codigo
Sheets(ws_output).Cells(ultLinea, 4).Value = cant
Sheets(ws_output).Cells(ultLinea, 5).Value = 1000
Sheets(ws_output).Protect Password:="", AllowFiltering:=True

Sheets("Registro rápidos").Range("B5:D5").ClearContents
ActiveWorkbook.Save
End Sub
