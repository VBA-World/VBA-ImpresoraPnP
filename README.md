# VBA-ImpresoraPnP
Código en VBA para imprimir en una impresora fiscal venezolana PnP

Se puede imprimir facturas, notas de crédito, reportes Z y demás desde cualquier programa compatible con código VBA (e.g Excel, Word, PowerPoint, Access o inclusive AutoCad)

# Modo de uso
ImpresoraPnP es un mólulo de clase que debe ser agregado a su proyecto VBA y luego escribir las rutinas como más le convenga

```VB.net
Public Sub imprimirFactura()
  Dim equipoFiscal As New ImpresoraPnP
End Sub
```
