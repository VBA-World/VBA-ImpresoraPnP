# VBA-ImpresoraPnP
Código en VBA para imprimir en una impresora fiscal venezolana PnP

Se puede imprimir facturas, notas de crédito, reportes Z y demás desde cualquier programa compatible con código VBA (e.g Excel, Word, PowerPoint, Access o inclusive AutoCad)

# Modo de uso
ImpresoraPnP es un mólulo de clase que debe ser agregado a su proyecto VBA y luego escribir las rutinas como más le convenga

```VB.net
Public Sub SimularImpresiónDeFactura()
  Dim EquipoFiscal As New ImpresoraPnP
  EquipoFiscal.Definir(Puerto:=1,Tipo:=PF300,ManejaGaveta:=False)
  
  With EquipoFiscal
    .Factura(Nombre:="Pedro Pérez", RIF:="YXXXXXXXXX")
    Dim i As Long
    For i=1 To 3
      .FacturaArtículo(Descripción:="Artículo " & i, Precio:=10 * i)
    Next i
    .FacturaTotalizar(PieDePágina:="Gracias por su compra")
  End With
End Sub
```
