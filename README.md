# Script-Google-Sheets
Código para integração de formulário HTML com google Sheets.

`
// Função principal para manipular os dados do Elementor
function doPost(e) {
  try {
    // Verifica se o objeto de evento "e" está definido
    if (typeof e !== 'undefined') {
      // Salva todos os parâmetros na variável parametros
      var parametros = e.parameter;

      // Log para verificar os parâmetros recebidos
      Logger.log('Parâmetros recebidos: ' + JSON.stringify(parametros));

      // Seleciona em qual página da planilha os dados deverão ser incluídos
      // Esta página precisa existir
      var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("nome da sua página de planilha aqui!");

      // Log para verificar a planilha selecionada
      Logger.log('Planilha selecionada: ' + (planilha ? planilha.getName() : 'Planilha não encontrada'));

      // Encontra a última linha já preenchida da página  
      var ultimaLinha = Math.max(planilha.getLastRow(), 1);
      // Especificamos em qual linha vamos salvar os dados
      // I.E. a linha que fica após a última linha já preenchida
      var linhaAtual = ultimaLinha + 1;

      // Log para verificar a última linha e linha atual
      Logger.log('Última linha: ' + ultimaLinha);
      Logger.log('Linha atual: ' + linhaAtual);

      // Verifica se a planilha foi encontrada antes de tentar inserir uma linha
      if (planilha) {
        // Insere uma linha após a última linha encontrada
        planilha.insertRowAfter(ultimaLinha);
        
        // Aqui vamos abrir aquela caixinha e separar tudo o que precisamos
        // cada um em sua variável
        var nome = parametros['Nome'];
        var cpf = parametros['Cpf'].toString();
        var empresa = parametros['Empresa'];
        var cnpj = parametros['Cnpj'];
        var celular = parametros['Celular'];
        var email = parametros['Email'];
        var cep = parametros['Cep'];
        var bairro = parametros['Bairro'];
        var rua = parametros['Rua'];
        var numero = parametros['Número'];
        var complemento = parametros['Complemento'];
        var idnome = parametros['(ID-001)NOME'];
        var idcargo = parametros['(ID-002)CARGO-OU-PROFISSÃO'];
        var idempresa = parametros['(ID-003)-EMPRESA'];
        var instagram = parametros['Instagram(EM-LINK)'];
        var whatsapp = parametros['Whatsapp-(EM-LINK)'];
        var salvarcontato = parametros['Salvar-Contato'];
        var site = parametros['Site'];
        var chavepix = parametros['Chave-Pix'];
        var destinatariopix = parametros['Destinatário-Pix'];
        var cidadecontapix = parametros['Cidade-da-Conta-Pix'];
        var outrosum = parametros['Outros'];
        var outrosdois = parametros['Outros-2'];

        // Log para verificar os valores recebidos
        Logger.log('Valores recebidos: ' + JSON.stringify({ Nome: nome, Cpf: cpf, Empresa: empresa, Cnpj: cnpj, Celular: celular, Email: email }));

        // Vamos aplicar os valores das variáveis em suas respectivas 
        // linhas e colunas, a linha é sempre a mesma: a última
        // e as colunas são especificadas após a linha atual
        planilha.getRange(linhaAtual, 1).setValue(new Date()); // Data atual
        planilha.getRange(linhaAtual, 2).setValue(nome);
        planilha.getRange(linhaAtual, 3).setNumberFormat('@').setValue(cpf);
        planilha.getRange(linhaAtual, 4).setValue(empresa);
        planilha.getRange(linhaAtual, 5).setValue(cnpj);
        planilha.getRange(linhaAtual, 6).setValue(celular);
        planilha.getRange(linhaAtual, 7).setValue(email);
        planilha.getRange(linhaAtual, 8).setValue(cep);
        planilha.getRange(linhaAtual, 9).setValue(bairro);
        planilha.getRange(linhaAtual, 10).setValue(rua);
        planilha.getRange(linhaAtual, 11).setValue(numero);
        planilha.getRange(linhaAtual, 12).setValue(complemento);
        planilha.getRange(linhaAtual, 13).setValue(idnome);
        planilha.getRange(linhaAtual, 14).setValue(idcargo);
        planilha.getRange(linhaAtual, 15).setValue(idempresa);
        planilha.getRange(linhaAtual, 16).setValue(instagram);
        planilha.getRange(linhaAtual, 17).setValue(whatsapp);
        planilha.getRange(linhaAtual, 18).setValue(salvarcontato);
        planilha.getRange(linhaAtual, 19).setValue(site);
        planilha.getRange(linhaAtual, 20).setValue(chavepix);
        planilha.getRange(linhaAtual, 21).setValue(destinatariopix);
        planilha.getRange(linhaAtual, 22).setValue(cidadecontapix);
        planilha.getRange(linhaAtual, 23).setValue(outrosum);
        planilha.getRange(linhaAtual, 24).setValue(outrosdois);
        

        // Log para verificar os valores inseridos na planilha
        Logger.log('Valores inseridos na planilha: ' + JSON.stringify([nome, cpf, empresa, cnpj, celular, email]));

        // Retorna uma mensagem de sucesso para o remetente
        return HtmlService.createHtmlOutput("Requisição recebida com sucesso!");
      } else {
        // Retorna uma mensagem de erro se a planilha não foi encontrada
        return HtmlService.createHtmlOutput("Erro: Planilha 'planilhacartaoecard' não encontrada.");
      }
    } else {
      // Retorna uma mensagem de erro se o objeto de evento não estiver definido
      return HtmlService.createHtmlOutput("Erro: Objeto de evento não definido.");
    }
  } catch (error) {
    // Retorna uma mensagem de erro se ocorrer um erro durante a execução
    return HtmlService.createHtmlOutput("Erro: " + error.message);
  }
}
`
