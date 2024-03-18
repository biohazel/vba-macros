'Captar inteiro fornecido pelo usuario e verificar se o valor eh maior que 20
Sub Exercicio_1()
  Dim var_1 As Integer
  var_1 = InputBox("Digite um valor inteiro: ")
  If var_1 > 20 Then
    'Exibir mensagem se valor superior a 20
    MsgBox "O número digitado: " & var_1 & " eh maior que 20"
  End If
End Sub

'Receber dois valores do usuario e verificar se a soma eh maior que 10
Sub Exercicio_2()
  Dim var_1 As Integer
  Dim var_2 As Integer
  Dim somatorio As Integer
  
  var_1 = InputBox("Digite um valor inteiro: ")
  var_2 = InputBox("Digite novamente um valor inteiro: ")
  
  somatorio = var_1 + var_2
  
  If somatorio > 10 Then
    'Exibir soma se for maior que 10
    MsgBox "Somatório de: " & var_1 & " + " & var_2 & " = " & somatorio
  End If
End Sub

'Calcular a soma de dois inteiros fornecidos pelo usuario
Sub Exercicio_3()
 Dim var_1 As Integer
 Dim var_2 As Integer
 Dim somatorio As Integer
 
 var_1 = InputBox("Digite um valor inteiro: ")
 var_2 = InputBox("Digite novamente um valor inteiro.")
 somatorio = var_1 + var_2
 
 If somatorio > 20 Then
   'Somatorio maior que 20 adicionar 8
   somatorio = somatorio + 8
 Else
   'Somatorio menor que 20 subtrair 5
   somatorio = somatorio - 5
 End If
 
 MsgBox "Resultado: " & var_1 & " + " & var_2 & " = " & somatorio
End Sub


  