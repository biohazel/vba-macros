'É aqui que vamos trabalhar!

Sub PrimeiroPrograma()

'Declaração das variáveis

Dim nm_aluno As String
Dim nm_curso As String
Dim nm_faculdade As String

'Entrada de valores

nm_aluno = InputBox("Digite o nome")
nm_curso = InputBox("Digite o curso")
nm_faculdade = InputBox("Digite a faculdade")

'Não teremos processamento
'Apenas saída de variáveis

Range("A1").Value = nm_aluno
Range("A2").Value = nm_curso
Range("A3").Value = nm_faculdade

MsgBox ("Aluno: " & nm_aluno & Chr(13) & _
"Curso: " & nm_curso & Chr(13) & _
"Faculdade: " & nm_faculdade)


End Sub

'Salvar em xlsm, pasta habilitada para macros do excel

Sub MaiorMenor()
   Dim x As Integer, y As Integer, z As Integer
   Dim maior As Integer, menor As Integer
   
'Entrada de variáveis
x = InputBox("Digite o 1º Valor")
y = InputBox("Digite o 2º Valor")
z = InputBox("Digite o 3º Valor")

'Processamento para solucionar o problema

If (x > y) Then
   maior = x
   menor = y

Else
   maior = y
   menor = x
End If

If (z < menor) Then
   menor = z
End If

'Saída do resultado

MsgBox ("O maior valor eh " & maior & Chr(13) & "O menor valor eh " & menor)

End Sub

'Add new value to cell

Sub SetValor()

Worksheets("AlunoCursoFaculdade").Range("A1").Value = 1223

End Sub

'Obter número de pastas de trabalho abertas

Sub ContadorBooks()

   MsgBox Workbooks.Count
   
End Sub

'Criar botão e ativar seleção

Sub Range_A1_D6()

   Range("A1:D6").Select

End Sub

'Exercício 1-Suponha que sua posição atual na planilha 1 
'(Plan 1 é a célula A 2–. Elabore uma sub rotina a qual 
'deverá utilizar a propriedade Offset para determinar como 
'nova posição a célula C 5. Lembrando que a propriedade Offset
'é utilizada com a propriedade Range e serve para especificar um novo
'local para manipulação

Sub Offset()
   Range("A2").Offset(3, 2).Select
End Sub

'Exercício 02–Suponha que sua referência atual na planilha 1 
'(Plan 1 é a célula E 8 – Elabore uma sub rotina responsável 
'por acessar a célula B 3 utilizando Range e Offset

Sub Offset_2()
   Range("E8").Offset(-5, -3).Select
End Sub

'Exercício 03–Suponha que sua célula atual na planilha 1 
'(Plan 1 é a célula A2 – Elabore uma sub rotina que utilize 
'resize para selecionar as células A2 a C2

Sub Resize()
   Range("A2").Resize(,3).Select
End Sub

Sub testandoVBA()
    MsgBox "Testando VBA!"
End Sub

'Não pode ter espaço no nome da sub, não começar com número, não ter pontos finais, não ter caracteres especiais


Sub GetValor()
    v_x = Worksheets("Planilha1").Range("A1").Value
    MsgBox v_x
End Sub

Sub SetValor()
    Worksheets("Planilha1").Range("A1").Value = 123.45
End Sub

Sub ContadorBooks()
    MsgBox Workbooks.Count
End Sub

Sub LimpaCelula()
    Range("A1").ClearContents
End Sub

'Copia o valor da célula A1 para a célula B1
Sub CopiaCelula()
    Worksheets("Planilha1").Activate
    Range("A1").Copy Range("B1")
End Sub

'Criar novas pastas de trabalho
Sub AdicionarWorkbooks()
    Workbooks.Add
End Sub

'Seleciona células no range A1 a D6

Sub Range_A1_D6()
    Range("A1:D6").Select
End Sub




