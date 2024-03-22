Sub RemoverLinha()
 Dim WkSht As Worksheet
 For Each WkSht In ActiveWorkbook.Worksheets
  WkSht.Rows(1).Delete
 Next WkSht
End Sub

Sub AlterarSinal()
 Dim celula As Range
 For Each celula In Range("A1:E50")
  If IsNumeric(celula.Value) Then
   celula.Value = celula.Value * -1
  End If
 Next celula
End Sub

'Codeshare
'Dontpad

Sub Tabuada()
   Dim valor As Integer
   Dim contador As Integer
   
   valor = InputBox("Digite um número:")
   
   For contador = 0 To 10
     ActiveCell.Offset(contador, 0) = valor & " x " & contador & " = " & valor * contador
   Next contador
End Sub

Sub Tabuada_2()
   Dim valor As Integer
   Dim contador As Integer
   Dim msg As String
   
   valor = InputBox("Digite um número:")
   
   For contador = 0 To 10
     msg = msg & valor & " x " & contador & " = " & valor * contador & vbNewLine
   Next contador
   
   MsgBox "Tabuada do: " & valor & vbNewLine & msg
End Sub

Sub RemoverLinha()
 Dim WkSht As Worksheet
 For Each WkSht In ActiveWorkbook.Worksheets
  WkSht.Rows(1).Delete
 Next WkSht
End Sub

Sub AlterarSinal()
 Dim celula As Range
 For Each celula In Range("A1:E50")
  If IsNumeric(celula.Value) Then
   celula.Value = celula.Value * -1
  End If
 Next celula
End Sub

Sub MostraMensagem()
 MsgBox Worksheets("Planilha5").Range("A1").Value
End Sub

'Não é possível ler range, só escrever range.
Sub EscreverRange()
  Worksheets("Planilha6").Range("A1:C3").Value = 123
End Sub

Sub PropriedadeText()
  MsgBox Worksheets("Planilha7").Range("A1").Text
End Sub


