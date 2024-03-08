'Calcular o produto de dois números inteiros
Sub MultiplicarCelulas()


   Dim v_a As Integer 'Declarar variáveis
   Dim v_b As Integer
   Dim resultado As Integer
   
   'Obter valores de células da planilha
   v_a = Worksheets(1).Range("D4").Value
   v_b = Worksheets(1).Range("D5").Value
   
   resultado = v_a * v_b
   
   'Mapeamento do resultado que é escrito na planilha. Atribuído ao local indicado.
   Worksheets(1).Range("D6").Value = resultado
   
   
End Sub


'Programa para calcular a média aritmética de números inteiros
Sub MediaAritmetica()
   
   'Declarar variáveis
   Dim var_a As Integer
   Dim var_b As Integer
   Dim var_c As Integer
   Dim resultado As Double
   
   'Obter valores de células da planilha
   var_a = Worksheets(1).Range("D10").Value
   var_b = Worksheets(1).Range("D11").Value
   var_c = Worksheets(1).Range("D12").Value
   
   resultado = (var_a + var_b + var_c) / 3
   
   Worksheets(1).Range("D13").Value = resultado
   
End Sub

'Exibir antecessor e sucessor de um número informado pelo usuário
Sub AntecessorSucessor()
   
   Dim numero As Integer
   Dim antecessor As Integer
   Dim sucessor As Integer
   
   numero = Worksheets(1).Range("D16").Value
   
   antecessor = numero - 1
   sucessor = numero + 1
   
   Worksheets(1).Range("D17").Value = antecessor
   Worksheets(1).Range("D18").Value = sucessor
   
End Sub

'Programa para calcular média ponderada
Sub MediaPonderada()

   Dim nota_a As Integer
   Dim nota_b As Integer
   Dim nota_c As Integer
   Dim nota_d As Integer
   Dim media_ponderada As Double
   Dim soma_pesos As Integer
   
   nota_a = Worksheets(1).Range("D21").Value
   nota_b = Worksheets(1).Range("D22").Value
   nota_c = Worksheets(1).Range("D23").Value
   nota_d = Worksheets(1).Range("D24").Value
   
   'Metodologia da média ponderada: dividir pela soma dos pesos
   soma_pesos = 1 + 2 + 3 + 4
   
   media_ponderada = ((nota_a * 1) + (nota_b * 2) + (nota_c * 3) + (nota_d * 4)) / soma_pesos
   
   Worksheets(1).Range("D25").Value = media_ponderada
   
End Sub