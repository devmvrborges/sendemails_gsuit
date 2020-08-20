function enviarEmail() {
  var sheet = SpreadsheetApp.getActiveSheet(); //seleciona a planilha que está rodando esse script
  var startRow = 2; // Linha de inicio (desconsiderar titulos, rotulos, etc)
  var numRows = 3; // Quantidade de linhas que irá percorrer o script
  
  // Faz um filtro dos valores e popula a variavel "dataRange" com as colunas "Email", "Mensagem", "Assunto"
  var dataRange = sheet.getRange(startRow, 1, numRows, 3);
  
  // Transforma esses valores em um formato de lista para podermos utilizar um loop for
  var data = dataRange.getValues();
  
  // Percorre populano a "var i" cada conteudo da variavel "data"
  for (var i in data) {
    var row = data[i]; 
    var emailAddress = row[0]; // Valor da coluna de email
    var message = row[1]; // Valor da coluna da mensagem
    var subject = row[2]; // Valor da coluna do assunto
    
    //Finalmente envia o e-mail.
    MailApp.sendEmail(emailAddress, subject, message); 
  }
}
