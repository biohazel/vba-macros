Utilizando comentários com VBA
Descreva bem Sub e Functions
Traduzir siglas
Descrever variáveis
Descrever solução para resolução de bugs
Use conversões

nmCliente
snm_cliente (s de stirng)
fTx_desc (f de tipo float, taxa de desconto)- mas em vba não tem floar, só tem double
é string e representa o nome do cliente
não tem salário menor pra que está cuidando da documentação do software
precisões escalares

Consertar código - documentar - liberar a release
Dim é a declaração de que estamos criando a variável dimensione na memória uma 
variável entitulada dolores as float
Em vba vc precisa atribuir a varável na linha debaixo da declaração 
Não pode iniciar variável com número em hipótese alguma, nem caracteres reservados 
O End Sub limpa as variáveis da memória RAM, todas

Um caracterer com acento tem 1,5 bytes

Option Explicit força a tipagem
Obrigatoriamente declarar a variável E o datatype a faz fortemente tipada. 

Local, global, pública, estática -> escopo de variáveis

Determina quais módulos e procedimentos podem manipular a variável. 

Static, publica e private

variável global é só dentro do mesmo módulo
variável pública você acessa de qualquer contexto, até em outro arquivo de excel
variável pública fica dentro do mesmo diretório

Constante são declaradas com const, e não Dim mais

Se colocar a variável na primeira linha do módulo ele é global praquele módulo
Variáveis sem escopo declarado é local - dentro da sub
