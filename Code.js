/**
 * =========================================================================
 * CONFIGURAÇÕES GLOBAIS DO SISTEMA
 * =========================================================================
 */
const ABA_RELATORIO = 'Relatório mensal de manutenção';
const LINHA_INICIAL = 1623; // Linha onde começam os registros reais na planilha
const ID_PLANILHA = '12SQ5q77gHH38KiAQNq05gyKy6728OKAA1WGQ008eqK4'; // ID da Planilha Google


/**
 * =========================================================================
 * WEB APP E MENU
 * =========================================================================
 */

/**
 * Função padrão do Google Apps Script chamada ao acessar a URL do Web App.
 * Retorna o arquivo HTML 'Formulario' para ser exibido no navegador.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Formulario')
    .setTitle('Emteco Motores - Registro de Manutenção');
}

/**
 * Cria um menu personalizado na planilha do Google quando ela é aberta.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Emteco')
    .addItem('Novo Registro de Manutenção', 'abrirFormulario')
    .addToUi();
}

/**
 * Abre o formulário em formato de janela modal dentro do Google Sheets.
 */
function abrirFormulario() {
  const html = HtmlService
    .createHtmlOutputFromFile('Formulario')
    .setWidth(1400)
    .setHeight(900);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Novo Registro de Manutenção'
  );
}


/**
 * =========================================================================
 * CARREGAMENTO DE DADOS E VALIDAÇÕES (SELECTS)
 * =========================================================================
 */

/**
 * Lê as validações de dados (listas suspensas) da planilha para preencher 
 * os selects do formulário HTML no carregamento da página.
 */
function obterDadosFormulario() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);

  return {
    faseFalha: obterValidacaoColuna(aba, 'F'),          // Coluna F
    modalidade: obterValidacaoColuna(aba, 'G'),         // Coluna G
    produto: obterValidacaoColuna(aba, 'I'),            // Coluna I
    qualidade: obterValidacaoColuna(aba, 'J'),          // Coluna J
    estoque: obterValidacaoColuna(aba, 'K'),            // Coluna K
    responsavel: obterValidacaoColuna(aba, 'M'),        // Coluna M
    defeitoIdentificado: obterValidacaoColuna(aba, 'Q'),// Coluna Q
    itemMovimentado: obterValidacaoColuna(aba, 'Z'),    // Coluna Z
    tipoEstoque: obterValidacaoColuna(aba, 'AA')        // Coluna AA
  };
}

/**
 * Função auxiliar que extrai os valores permitidos de uma validação de dados 
 * em uma determinada coluna da planilha.
 */
function obterValidacaoColuna(aba, coluna) {
  const linhaReferencia = 1623;
  const celula = aba.getRange(`${coluna}${linhaReferencia}`);
  const validacao = celula.getDataValidation();

  // Se a célula não tiver validação, retorna um array vazio
  if (!validacao) {
    return [];
  }

  const criterio = validacao.getCriteriaType();
  const valores = validacao.getCriteriaValues();

  // Caso 1: A validação é uma lista manual de valores digitados
  if (criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return valores[0];
  }

  // Caso 2: A validação puxa os valores de um intervalo (intervalo de células)
  if (criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    const intervalo = valores[0];
    return intervalo
      .getValues()
      .flat()
      .filter(valor => valor !== '');
  }

  return [];
}


/**
 * =========================================================================
 * MANIPULAÇÃO DE LINHAS E INSERÇÃO DE DADOS
 * =========================================================================
 */

/**
 * Localiza a próxima linha vazia a partir da LINHA_INICIAL (1623).
 */
function encontrarProximaLinha(aba) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < LINHA_INICIAL) {
    return LINHA_INICIAL;
  }

  // Pega os dados das primeiras 8 colunas a partir da linha 1623 para verificar se está vazia
  const dados = aba
    .getRange(LINHA_INICIAL, 1, ultimaLinha - LINHA_INICIAL + 1, 8)
    .getValues();

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const vazia = linha.every(valor => valor === '');

    if (vazia) {
      return LINHA_INICIAL + i; // Retorna a primeira linha totalmente vazia encontrada
    }
  }

  return ultimaLinha + 1; // Se não achar vazia no meio, retorna logo após a última
}

/**
 * Copia a formatação e as fórmulas da linha anterior para manter o padrão 
 * da planilha e limpa os campos preenchidos pelo formulário.
 */
function prepararNovaLinha(aba, linhaDestino) {
  if (linhaDestino <= LINHA_INICIAL) {
    return;
  }

  const linhaOrigem = linhaDestino - 1;
  const origem = aba.getRange(linhaOrigem, 1, 1, 28);

  // Copia apenas a formatação visual da linha de cima
  origem.copyTo(
    aba.getRange(linhaDestino, 1, 1, 28),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  // Copia as fórmulas matemáticas da linha de cima
  const formulas = origem.getFormulas();
  aba.getRange(linhaDestino, 1, 1, 28).setFormulas(formulas);

  // Colunas do formulário que devem ser limpas na nova linha para receberem os novos dados
  const colunasFormulario = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 20, 24, 25, 26, 27, 28];

  colunasFormulario.forEach(coluna => {
    aba.getRange(linhaDestino, coluna).clearContent();
  });
}

/**
 * Salva os registros enviados pelo formulário web, criando tantas linhas 
 * quantos forem os produtos informados.
 */
function salvarFormulario(dados) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);

  const quantidade = Number(dados.quantidade) || 1;
  let linhaAtual = encontrarProximaLinha(aba);

  // Loop para inserir cada produto em uma linha separada
  for (let i = 0; i < quantidade; i++) {
    prepararNovaLinha(aba, linhaAtual);
    const produto = dados.produtos[i] || {};

    // Preenche cada coluna correspondente (base 1 = Coluna da planilha)
    aba.getRange(linhaAtual, 1).setValue(dados.dataReclamacao);
    aba.getRange(linhaAtual, 2).setValue(dados.empresa);
    aba.getRange(linhaAtual, 3).setValue(dados.codigoEmteco);
    aba.getRange(linhaAtual, 4).setValue(dados.dataCompra);
    aba.getRange(linhaAtual, 5).setValue(dados.reclamacaoDefeito);
    aba.getRange(linhaAtual, 6).setValue(dados.faseFalha);
    aba.getRange(linhaAtual, 7).setValue(dados.modalidade);
    aba.getRange(linhaAtual, 8).setValue(dados.numeroTarefa);

    aba.getRange(linhaAtual, 9).setValue(produto.produto || '');
    aba.getRange(linhaAtual, 10).setValue(produto.qualidade || '');
    aba.getRange(linhaAtual, 11).setValue(produto.estoqueRetirada || '');
    aba.getRange(linhaAtual, 12).setValue(produto.codigoLote || '');

    aba.getRange(linhaAtual, 13).setValue(dados.responsavel);
    aba.getRange(linhaAtual, 14).setValue(dados.dataRecebimento);
    aba.getRange(linhaAtual, 15).setValue(dados.numeroControle);
    aba.getRange(linhaAtual, 16).setValue(dados.dataTeste);

    aba.getRange(linhaAtual, 17).setValue(produto.defeitoIdentificado || '');
    aba.getRange(linhaAtual, 20).setValue(dados.observacaoExtra);
    aba.getRange(linhaAtual, 24).setValue(dados.video);
    aba.getRange(linhaAtual, 25).setValue(dados.observacaoFornecedor);
    aba.getRange(linhaAtual, 26).setValue(dados.itemMovimentado);
    aba.getRange(linhaAtual, 27).setValue(dados.tipoEstoque);
    aba.getRange(linhaAtual, 28).setValue(dados.dataFinalizacao);

    linhaAtual++; // Vai para a próxima linha para o próximo produto
  }

  return {
    sucesso: true,
    mensagem: `${quantidade} produto(s) inserido(s) com sucesso.`,
    proximaLinha: linhaAtual
  };
}


/**
 * =========================================================================
 * BUSCA E ATUALIZAÇÃO DE TAREFAS EXISTENTES
 * =========================================================================
 */

/**
 * Busca o número da tarefa na Coluna H (índice 7) a partir da LINHA_INICIAL (1623).
 * Retorna os dados mapeados para preencher o formulário de edição.
 */
function buscarTarefa(numeroTarefa) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO); 
  
  const dados = aba.getDataRange().getValues();
  
  // Define a linha inicial de busca convertendo para o índice do array JavaScript (linha - 1)
  const indiceInicial = LINHA_INICIAL - 1;
  
  // Varre a planilha de baixo para cima ou a partir da linha 1623
  for (let i = indiceInicial; i < dados.length; i++) {
    const valTarefa = String(dados[i][7]).trim(); // Coluna H = Índice 7 (A=0, B=1, ..., H=7)
    
    // Se encontrou a tarefa informada
    if (valTarefa === String(numeroTarefa).trim()) {
      return {
        encontrado: true,
        linha: i + 1,                          // Número real da linha na planilha
        codigoEmteco: dados[i][2] || "",       // Coluna C (Índice 2)
        loteDefeito: dados[i][11] || "",      // Coluna L (Índice 11)
        dataRecebimento: dados[i][13] || "",  // Coluna N (Índice 13)
        numeroControle: dados[i][14] || "",   // Coluna O (Índice 14)
        dataTeste: dados[i][15] || "",        // Coluna P (Índice 15)
        defeitoIdentificado: dados[i][16] || "", // Coluna Q (Índice 16)
        observacaoExtra: dados[i][19] || ""   // Coluna T (Índice 19)
      };
    }
  }
  
  return { encontrado: false }; // Retorna falso se não achar a tarefa
}

/**
 * Atualiza os dados de uma tarefa existente na planilha com base nas alterações feitas na interface.
 */
function atualizarTarefasEmLote(listaAtualizacoes) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);

  listaAtualizacoes.forEach(item => {
    const linha = Number(item.linha);
    if (linha && linha >= 2) {
      if (item.produto !== undefined) aba.getRange(linha, 9).setValue(item.produto);                     // Coluna I
      if (item.qualidade !== undefined) aba.getRange(linha, 10).setValue(item.qualidade);               // Coluna J
      if (item.responsavel !== undefined) aba.getRange(linha, 13).setValue(item.responsavel);           // Coluna M
      if (item.dataRecebimento) aba.getRange(linha, 14).setValue(item.dataRecebimento);                 // Coluna N
      if (item.dataTeste) aba.getRange(linha, 16).setValue(item.dataTeste);                             // Coluna P
      if (item.defeitoIdentificado !== undefined) aba.getRange(linha, 17).setValue(item.defeitoIdentificado); // Coluna Q
      if (item.itemMovimentado !== undefined) aba.getRange(linha, 26).setValue(item.itemMovimentado);   // Coluna Z
      if (item.dataFinalizacao) aba.getRange(linha, 28).setValue(item.dataFinalizacao);                 // Coluna AB
    }
  });

  return {
    sucesso: true,
    mensagem: "Registros atualizados com sucesso!"
  };
}