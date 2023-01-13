' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub reg_loteria()

Dim cont As Long
Dim ultLinea As Long
Dim cant As Variant
Dim nombre As Variant

Sheets("Info lotería").Unprotect Password:=""
ultLinea = Sheets("Info lotería").Range("A" & Rows.Count).End(xlUp).Row
nombre = Sheets("Registro lotería").Cells(5, 2)
numero = Sheets("Registro lotería").Cells(5, 3)
cant = Sheets("Registro lotería").Cells(5, 4)

If IsEmpty(nombre) Or IsEmpty(numero) Or IsEmpty(cant) Then
    MsgBox "Faltan campos por completar. No hice nada."
    Exit Sub
End If

For cont = 1 To ultLinea
    current = Sheets("Info lotería").Cells(cont, 1)
    If nombre = current Then
        anterior = Sheets("Info lotería").Cells(cont, 4)
        Sheets("Info lotería").Cells(cont, 2) = numero
        Sheets("Info lotería").Cells(cont, 4) = cant
        Sheets("Info lotería").Cells(cont, 5) = cant
    End If
Next cont
Sheets("Info lotería").Protect Password:="dckgzf", AllowFiltering:=True

Range("reg_nom_lot").ClearContents
Range("reg_num_lot").Clear
Range("reg_cant_lot").Clear
ActiveWorkbook.Save
End Sub

