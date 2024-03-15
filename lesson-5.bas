Sub DemoGoTo()
    nome = InputBox("Digite o nome.")
    If nome <> "Joaquim" Then
        GoTo NomeErrado
    End If
    MsgBox
    End Sub

NomeErrado:
    MsgBox
End Sub

'Estrutura If-Then
'Time 0h - 6h 25%
'Time 6h - 12h 25%
'Time 12h - 18h 25%
'Time 18h - 24h 25%
'Estrutura de 1 dia - 100%

Sub GetHora()
    If Time < 0.5 Then
        MsgBox "Bom dia!"
    End If

If Time >= 0.5 Then
    MsgBox "Boa tarde."
    End If
End Sub

'Estrutura If-Then-Else

Sub GetHora_2()
    If Time < 0.5 Then
    MsgBox "Bom dia."
    Else
    MsgBox "Boa tarde."
    End If
End Sub

'Estrutura If-Then-Else

Sub GetHora_3()
    
    Dim msg As String
    If Time < 0.5 Then msg = "Manhã."
    If Time >= 0.5 And Time < 0.75 Then msg = "Tarde."
    If Time >= 0.75 Then msg = "Noite."
    
    MsgBox "Boa" & msg
End Sub

'If-Then-Else
'Similar ao exemplo anterior, entretanto, fazendo uso da estrutura If-Then-End If

Sub GetHora_4()
    Dim msg As String
    If Time < 0.5 Then
        msg = "Manhã."
    End If
    If Time >= 0.5 And Time < 0.75 Then
        msg = "Tarde."
    End If
    If Time >= 0.75 Then
        msg = "Noite."
    End If
    MsgBox "Boa" & msg
End Sub

'Estrutura ElseIf

Sub GetHora_5()
    Dim msg As String
    If Time < 0.5 Then
        msg = "Manhã."
    ElseIf Time >= 0.5 And Time < 0.75 Then
     msg = "Tarde."
    Else: msg = "Noite."
    End If
    MsgBox "Boa" & msg
End Sub

'Solicita valor e exibe desconto apropriado

Sub ExibeDesconto()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um valor: ")
    If quantidade > 0 Then desconto = 0.1
    If quantidade >= 25 Then desconto = 0.15
    If quantidade >= 50 Then desconto = 0.2
    If quantidade >= 75 Then desconto = 0.25
    MsgBox "Desconto:" & desconto
End Sub


'Variavel escalar só consegue guardar  1 vetor por vez
'Variaveis indexadas são vetores e matrizes (unidimensional e multidimensional)

'Cada sensor- dimensão no banco de dados (64 sensores craniais), processos
'Grupo de pesquisa de gestão em saúde

'Estrutura If-Then

'Similar ao exemplo anterior, porém utiliza a sintaxe else
'Neste caso, a rotina termina imediatamente  depois de executar as declarações quando a condição é verdadeira

Sub ExibeDesconto_2()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um valor:")
    If quantidade > 0 And quantidade < 25 Then
        desconto = 0.1
    ElseIf quantidade >= 25 And quantidade < 50 Then
        desconto = 0.15
    ElseIf quantidade >= 50 And quantidade < 75 Then
        desconto = 0.2
    ElseIf quantidade > 75 Then
        desconto = 0.25
    End If
    MsgBox "Desconto:" & desconto
End Sub

'Select-Case

Sub ExibeDesconto_3()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um valor: ")
    
    Select Case quantidade
        Case 0 To 24
         desconto = 0.1
        Case 25 To 49
         desconto = 0.15
        Case 50 To 74
         desconto = 0.2
        Case Is >= 75
        desconto = 0.25
    End Select
    MsgBox "Desconto: " & desconto
End Sub

'Exemplo com a variável inline after :

Sub ExibeDesconto_4()
    Dim quantidade As Integer
    Dim desconto As Double
    
    quantidade = InputBox("Digite um valor:")
    Select Case quantidade
      Case 0 To 24: desconto = 0.1
      Case 25 To 49: desconto = 0.15
      Case 50 To 74: desconto = 0.2
      Case Is >= 75: desconto = 0.25
    End Select
    MsgBox "Desconto:" & desconto
End Sub

'Aninhamento da estrutura Select Case

Sub VerificarCelula()
   Dim msg As String
   
   Select Case IsEmpty(ActiveCell)
     Case True
       msg = "está vazia"
     Case Else
       Select Case ActiveCell.HasFormula
         Case True
          msg = "tem uma fórmula"

   Case False
     Select Case IsNumeric(ActiveCell)
       Case True
         msg = "tem um número"
       Case Else
         msg = "tem texto"
    End Select
  End Select
End Select
  MsgBox "Célula:" & ActiveCell.Address & " " & msg
End Sub

'Next incrementa a variável contador

Sub PreencherRange()
 Dim contador As Long
 
 For contador = 0 To 19
  ActiveCell.Offset(contador, 0) = Rnd
 Next contador
End Sub

'Este código pula linhas

Sub PreencheRange_2()
  Dim contador As Long
  
  For contador = 0 To 19 Step 2
    ActiveCell.Offset(contador, 0) = Rnd
  Next contador
End Sub

'Loop

Sub ExitForDemo()
  Dim valor_max As Double
  Dim linha As Long
  
  valor_max = WorksheetFunction.Max(Range("A:A"))
  
  For linha = 1 To Rows.Count
    If (Range("A1").Offset(linha - 1, 0).Value = valor_max) Then
      Range("A1").Offset(linha - 1, 0).Activate
      MsgBox "Valor máximo é a linha: " & linha
      Exit For
    End If
  Next linha
End Sub

Sub ExitForDemo()
  Dim valor_max As Double
  Dim linha As Long
  
  valor_max = WorksheetFunction.Max(Range("A:A"))
  
  For linha = 1 To Rows.Count
    If (Range("A1").Offset(linha - 1, 0).Value = valor_max) Then
      Range("A1").Offset(linha - 1, 0).Activate
      MsgBox "Valor máximo é a linha: " & linha
      Exit For
    End If
  Next linha
End Sub

'Long é o número do Bitcoin

Sub PreencherCelulas()
  Dim linha As Long
  Dim coluna As Long
  
  For coluna = 1 To 5
    For linha = 1 To 12
      Cells(linha, coluna) = Rnd
    Next linha
  Next coluna
End Sub


Sub DoWhile_1()
  Do While ActiveCell.Value <> Empty
    ActiveCell.Value = ActiveCell.Value * 2
    ActiveCell.Offset(1, 0).Select
  Loop
End Sub

Sub DoWhile_2()
  Do
    ActiveCell.Value = ActiveCell.Value * 2
    ActiveCell.Offset(1, 0).Select
  Loop While ActiveCell.Value <> Empty
End Sub

'Do-Until executa os falsos até encontrar a verdadeira

Sub DoUntil_1()
  Do Until IsEmpty(ActiveCell.Value)
    ActiveCell.Value = ActiveCell.Value * 2
    ActiveCell.Offset(1, 0).Select
  Loop
End Sub

'Do-until embaixo

Sub DoUntil_2()
  Do
    ActiveCell.Value = ActiveCell.Value * 2
    ActiveCell.Offset(1, 0).Select
  Loop Until IsEmpty(ActiveCell.Value)
End Sub


