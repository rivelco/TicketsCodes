Attribute VB_Name = "Módulo5"
' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub reg_venta_loteria()
ws_output = "Venta lotería"
Sheets(ws_output).Unprotect Password:=""
next_row = Sheets(ws_output).Range("A" & Rows.Count).End(xlUp).Offset(1).Row
Sheets(ws_output).Cells(next_row, 1).Value = Now
Sheets(ws_output).Cells(next_row, 2).Value = Range("lot_venta_nom").Value
Sheets(ws_output).Cells(next_row, 3).Value = Range("lot_venta_cant").Value
Sheets(ws_output).Protect Password:="", AllowFiltering:=True

Dim cont As Long
Dim ultLinea As Long
Dim cant As Variant
Dim nombre As Variant

Sheets("Info lotería").Unprotect Password:=""
ultLinea = Sheets("Info lotería").Range("A" & Rows.Count).End(xlUp).Row
nombre = Sheets("Vender").Cells(15, 2)
cant = Sheets("Vender").Cells(15, 3)
For cont = 1 To ultLinea
    current = Sheets("Info lotería").Cells(cont, 1)
    If nombre = current Then
        anterior = Sheets("Info lotería").Cells(cont, 4)
        Sheets("Info lotería").Cells(cont, 4) = anterior - cant
    End If
Next cont
Sheets("Info lotería").Protect Password:="", AllowFiltering:=True

Range("lot_venta_nom").ClearContents
Range("lot_venta_cant").Clear
ActiveWorkbook.Save
End Sub
