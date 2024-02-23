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