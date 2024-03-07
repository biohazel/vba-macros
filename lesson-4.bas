Sub TestandoVariavel()
   'Dim v_numero as Integer
   Dim v_numero As Integer
   v_numero = 10
End Sub

Sub validaCPF()
   Dim nome as String
End Sub
   'Não consigo usar essa variável em outra subrotina, no end sub a variável se autodestrói na memória. Libero para que uma outra subrotina possa
   'usar o mesmo identificador. Cada conjunto Sub-End Sub é um escopo diferente.

Sub TesteEscipo() 
   Dim x As Integer, y as Integer, y as Integer
   Dim valor As Long, result As Double
   Dim x, y, z as integer 'só o z é inteiro, x e y saem como variant. VBA não permite declaração e tipagem em groupos
End Sub


'Explicitar se é pública ou privada
Option Explicit
Private x As Integer
Public v_nome as String

'Declaração como static, o nome e valor da variável não é apagado da memória ao End Sub, ou seja, na próxima execução
'se ela tinha um valor x, ela vai ser tratada partindo do ponto deste valor 
Sub testeEscopo()
   Static contador As Integer
   Dim msg As String
   contador = contador + 1
   msg = "Número de execuções: " & contador
   MsgBox msg
End Sub

