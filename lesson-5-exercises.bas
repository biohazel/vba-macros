Sub Exercicio_1()
  Dim var_1 As Integer
  var_1 = InputBox("Digite um valor inteiro: ")
  If var_1 > 20 Then
    MsgBox "O número digitado: " & var_1 & " eh maior que 20"
  End If
End Sub


Sub Exercicio_2()
  Dim var_1 As Integer
  Dim var_2 As Integer
  Dim somatorio As Integer
  
  var_1 = InputBox("Digite um valor inteiro: ")
  var_2 = InputBox("Digite novamente um valor inteiro: ")
  
  somatorio = var_1 + var_2
  
  If somatorio > 10 Then
    MsgBox "Somatório de: " & var_1 & " + " & var_2 & " = " & somatorio
  End If
End Sub


Sub Exercicio_3()
 Dim var_1 As Integer
 Dim var_2 As Integer
 Dim somatorio As Integer
 
 var_1 = InputBox("Digite um valor inteiro: ")
 var_2 = InputBox("Digite novamente um valor inteiro.")
 somatorio = var_1 + var_2
 
 If somatorio > 20 Then
   somatorio = somatorio + 8
 Else
   somatorio = somatorio - 5
 End If
 
 MsgBox "Resultado: " & var_1 & " + " & var_2 & " = " & somatorio
End Sub
