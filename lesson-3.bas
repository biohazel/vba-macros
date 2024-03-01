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

'Propriedade Offset-define novo local, usada em conjunto com Range

Sub TesteOffset()
    Range("A1").Offset(RowOffset:=1, ColumnOffset:=1).Select
End Sub

Sub TesteOffset_1()
    Range("A1").Offset(1, 1).Select
End Sub

'Só vai saltar para a linha de baixo
Sub TesteOffset_2()
    Range("A1").Offset(1).Select
End Sub

'Pula uma coluna, mas permanece na mesma linha

Sub Teste_Offset_3()
    Range("A1").Offset(, 1).Select
End Sub

'Números negativos vão de trás pra frente e de cima para baixo

Sub TesteOffset_4()
    Range("B2").Offset(-1, -1).Select
End Sub



'Transacionar um intervalo pela planilha. Não considera a origem na resposta final, apenas a usa como referência
Sub TesteOffset_5()
    Range("A1:C3").Offset(1, 1).Select
End Sub


'Resize redimensiona, ou seja, ele considera a origem na resposta final
Sub TesteOffset_6()
    Range("A1").Resize(RowSize:=2, ColumnSize:=2).Select
End Sub

'Resize de forma mais curta
Sub TesteOffset_7()
    Range("A1").Resize(2, 2).Select
End Sub

'Passando apenas o número de linhas que quero redimensionar

Sub TesteOffset_8()
    Range("A1").Resize(2).Select
End Sub

'Passar o número de colunas que quer redimensionar, ignorando as linhas

Sub TesteOffset_9()
    Range("A1").Resize(, 2).Select
End Sub