# Planilha-de-Horario
Uma planilha com uso simples de VBA para criar um ponto eletronico gerido pelo usuario

Esta planilha possui uma organização simples de um Mês completo. Cada dia possui 3 botões, que buscam o horário do computador do usuário e insere
na célula ao lado. 

É necessário permitir a edição e a execução de scripts pelo documento.

Além disso, a tabela calcula automaticamente a quantidade de horas trabalhadas de acordo com o ponto registrado pelo usuário, e calcula horas extras.

O arquivo Ponto.png possui uma imagem demonstrando a tabela em si.

Código em VBA de cada Botão:

Private Sub CommandButton_Click()
  Range("Celula").Value = Now()
End Sub

Arquivo de uso livre, assim como os scritps.
