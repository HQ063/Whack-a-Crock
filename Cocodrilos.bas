Attribute VB_Name = "Cocodrilos"
Option Explicit
Dim ultimo
Dim repite As Boolean
Dim numerocasa
Public nivel
Public segundos As String
Public minutos As Long
Public totalsegundos As Long

Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function creanumerocasa()
ultimo = Rnd(0)
numerocasa = Rand(1, 5)
If (numerocasa = ultimo & repite = True) Then
numerocasa = Rand(1, 5)
Else
repite = False
End If
If (numerocasa = ultimo) Then
repite = True
End If
creanumerocasa = numerocasa

End Function
