Attribute VB_Name = "Módulo9"
' Este programa fue desarrollado por Ricardo Velázquez Contreras
' Publicación original y licencia en GitHub https://github.com/rivelco

Sub reg_venta_superior()
venta_sup = "VentaSup"
Dim i As Integer
Dim RapidosNames(1 To 9) As String
Dim RapidosCants(1 To 9) As Integer
Dim RapidosCosts(1 To 9) As Integer
Dim iSavedR As Integer

iSavedR = 0
For i = 3 To 11
    If Sheets(venta_sup).Cells(i, 4).Value > 0 Then
        iSavedR = iSavedR + 1
        RapidosNames(iSavedR) = Sheets(venta_sup).Cells(i, 3).Value
        RapidosCants(iSavedR) = Sheets(venta_sup).Cells(i, 4).Value
        RapidosCosts(iSavedR) = Sheets(venta_sup).Cells(i, 5).Value
    End If
Next i
        
Dim LoteriaNames(1 To 8) As String
Dim LoteriaCants(1 To 8) As Integer
Dim LoteriaCosts(1 To 8) As Integer
Dim iSavedL As Integer
iSavedL = 0
For i = 15 To 22
    If Sheets(venta_sup).Cells(i, 4).Value > 0 Then
        iSavedL = iSavedL + 1
        LoteriaNames(iSavedL) = Sheets(venta_sup).Cells(i, 3).Value
        LoteriaCants(iSavedL) = Sheets(venta_sup).Cells(i, 4).Value
        LoteriaCosts(iSavedL) = Sheets(venta_sup).Cells(i, 5).Value
    End If
Next i

ws_output = "Venta rápidos"
Sheets(ws_output).Unprotect Password:=""
For i = 1 To iSavedR
    next_row = Sheets(ws_output).Range("A" & Rows.Count).End(xlUp).Offset(1).Row
    Sheets(ws_output).Cells(next_row, 1).Value = Now
    Sheets(ws_output).Cells(next_row, 2).Value = RapidosNames(i)
    Sheets(ws_output).Cells(next_row, 3).Value = RapidosCants(i)
    Sheets(ws_output).Cells(next_row, 4).Value = RapidosCosts(i)
Next i
Sheets(ws_output).Protect Password:="", AllowFiltering:=True

Dim cant As Variant
Dim nombre As Variant
Sheets("Info rápidos").Unprotect Password:=""
ultLinea = Sheets("Info rápidos").Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To iSavedR
    nombre = RapidosNames(i)
    cant = RapidosCants(i)
    For cont = 1 To ultLinea
        current = Sheets("Info rápidos").Cells(cont, 1)
        If nombre = current Then
            anterior = Sheets("Info rápidos").Cells(cont, 4)
            Sheets("Info rápidos").Cells(cont, 4) = anterior - cant
        End If
    Next cont
Next i
Sheets("Info rápidos").Protect Password:="", AllowFiltering:=True

ws_output = "Venta lotería"
Sheets(ws_output).Unprotect Password:=""
For i = 1 To iSavedL
    next_row = Sheets(ws_output).Range("A" & Rows.Count).End(xlUp).Offset(1).Row
    Sheets(ws_output).Cells(next_row, 1).Value = Now
    Sheets(ws_output).Cells(next_row, 2).Value = LoteriaNames(i)
    Sheets(ws_output).Cells(next_row, 3).Value = LoteriaCants(i)
    Sheets(ws_output).Cells(next_row, 4).Value = LoteriaCosts(i)
Next i
Sheets(ws_output).Protect Password:="", AllowFiltering:=True

Sheets("Info lotería").Unprotect Password:=""
ultLinea = Sheets("Info lotería").Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To iSavedL
    nombre = LoteriaNames(i)
    cant = LoteriaCants(i)
    For cont = 1 To ultLinea
        current = Sheets("Info lotería").Cells(cont, 1)
        If nombre = current Then
            anterior = Sheets("Info lotería").Cells(cont, 4)
            Sheets("Info lotería").Cells(cont, 4) = anterior - cant
        End If
    Next cont
Next i
Sheets("Info lotería").Protect Password:="", AllowFiltering:=True

Sheets("Pagos lotería").Protect Password:="", AllowFiltering:=True
Sheets("Pagos rápidos").Protect Password:="", AllowFiltering:=True
Sheets("VentaSup").Protect Password:=""

Sheets(venta_sup).Range("D3:D11").ClearContents
Sheets(venta_sup).Range("D15:D22").ClearContents
Sheets(venta_sup).Range("I15").Value = 0
Sheets(venta_sup).Range("I19").Value = 0
Sheets(venta_sup).Range("I12").Value = 0
Sheets(venta_sup).Activate

ActiveWorkbook.Save

End Sub
