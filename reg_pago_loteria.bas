' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub reg_pago_loteria()

Dim ultLinea As Long
Dim cant As Variant
Dim nombre As Variant
Dim numero As Variant
Dim iniciales As Variant
Dim devueltos As Variant
Dim vendidos As Variant
Dim vendidosT As Variant
Dim diferencia As Variant
Dim pagado As Variant
Dim comision As Variant

numero = Sheets("Pagar lotería").Cells(7, 9)
nombre = Sheets("Pagar lotería").Cells(5, 3)
iniciales = Sheets("Pagar lotería").Cells(7, 3)
devueltos = Sheets("Pagar lotería").Cells(9, 3)

If IsEmpty(numero) Or IsEmpty(nombre) Or IsEmpty(iniciales) Or IsEmpty(devueltos) Then
    MsgBox "Faltan campos por completar. No hice nada."
    Exit Sub
End If

vendidos = iniciales - devueltos
vendidosT = iniciales - Sheets("Pagar lotería").Cells(11, 3)
diferencia = vendidosT - vendidos
pagado = Sheets("Pagar lotería").Cells(17, 3)
comision = Sheets("Pagar lotería").Cells(18, 3)

ultLinea = Sheets("Pagos lotería").Range("A" & Rows.Count).End(xlUp).Offset(1).Row

Sheets("Pagos lotería").Unprotect Password:=""
Sheets("Pagos lotería").Cells(ultLinea, 1) = "No"
Sheets("Pagos lotería").Cells(ultLinea, 2) = Now
Sheets("Pagos lotería").Cells(ultLinea, 3) = nombre
Sheets("Pagos lotería").Cells(ultLinea, 4) = numero
Sheets("Pagos lotería").Cells(ultLinea, 5) = iniciales
Sheets("Pagos lotería").Cells(ultLinea, 6) = vendidos
Sheets("Pagos lotería").Cells(ultLinea, 7) = devueltos
Sheets("Pagos lotería").Cells(ultLinea, 8) = diferencia
Sheets("Pagos lotería").Cells(ultLinea, 9) = pagado
Sheets("Pagos lotería").Cells(ultLinea, 10) = comision
Sheets("Pagos lotería").Protect Password:="", AllowFiltering:=True

Sheets("Info lotería").Unprotect Password:=""
ultLinea = Sheets("Info lotería").Range("A" & Rows.Count).End(xlUp).Row
For cont = 1 To ultLinea
    current = Sheets("Info lotería").Cells(cont, 1)
    If nombre = current Then
        Sheets("Info lotería").Cells(cont, 2) = 0
        Sheets("Info lotería").Cells(cont, 4) = 0
        Sheets("Info lotería").Cells(cont, 5) = 0
    End If
Next cont
Sheets("Info lotería").Protect Password:="", AllowFiltering:=True

ActiveWorkbook.Save
End Sub


