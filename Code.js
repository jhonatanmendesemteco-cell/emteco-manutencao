const ABA_RELATORIO = 'Relatório mensal de manutenção'; 
const LINHA_INICIAL = 1623;

// ID da sua planilha
const ID_PLANILHA = '12SQ5q77gHH38KiAQNq05gyKy6728OKAA1WGQ008eqK4';


/**
 * Abre o formulário como página Web App
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Formulario')
    .setTitle('Emteco Motores - Registro de Manutenção');
}


/**
 * Cria o menu ao abrir a planilha
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Emteco')
    .addItem('Novo Registro de Manutenção', 'abrirFormulario')
    .addToUi();
}


/**
 * Abre o painel lateral
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
 * Retorna as opções dos validadores
 */
function obterDadosFormulario() {

  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);

  return {
    faseFalha: obterValidacaoColuna(aba, 'F'),
    modalidade: obterValidacaoColuna(aba, 'G'),
    produto: obterValidacaoColuna(aba, 'I'),
    qualidade: obterValidacaoColuna(aba, 'J'),
    estoque: obterValidacaoColuna(aba, 'K'),
    responsavel: obterValidacaoColuna(aba, 'M'),
    defeitoIdentificado: obterValidacaoColuna(aba, 'Q'),
    itemMovimentado: obterValidacaoColuna(aba, 'Z'),
    tipoEstoque: obterValidacaoColuna(aba, 'AA')
  };
}


/**
 * Lê os valores da validação de uma coluna
 */
function obterValidacaoColuna(aba, coluna) {

  const linhaReferencia = 1623;
  const celula = aba.getRange(`${coluna}${linhaReferencia}`);

  const validacao = celula.getDataValidation();

  if (!validacao) {
    return [];
  }

  const criterio = validacao.getCriteriaType();
  const valores = validacao.getCriteriaValues();

  if (
    criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST
  ) {
    return valores[0];
  }

  if (
    criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE
  ) {

    const intervalo = valores[0];

    return intervalo
      .getValues()
      .flat()
      .filter(valor => valor !== '');

  }

  return [];
}


/**
 * Localiza a próxima linha disponível
 */
function encontrarProximaLinha(aba) {

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < LINHA_INICIAL) {
    return LINHA_INICIAL;
  }

  const dados = aba
    .getRange(
      LINHA_INICIAL,
      1,
      ultimaLinha - LINHA_INICIAL + 1,
      8
    )
    .getValues();

  for (let i = 0; i < dados.length; i++) {

    const linha = dados[i];

    const vazia = linha.every(valor => valor === '');

    if (vazia) {
      return LINHA_INICIAL + i;
    }

  }

  return ultimaLinha + 1;
}


/**
 * Copia estrutura e fórmulas da linha anterior
 */
function prepararNovaLinha(aba, linhaDestino) {

  if (linhaDestino <= LINHA_INICIAL) {
    return;
  }

  const linhaOrigem = linhaDestino - 1;

  const origem = aba.getRange(linhaOrigem, 1, 1, 28);

  origem.copyTo(
    aba.getRange(linhaDestino, 1, 1, 28),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  /**
   * Copia apenas fórmulas
   */
  const formulas = origem.getFormulas();

  aba
    .getRange(linhaDestino, 1, 1, 28)
    .setFormulas(formulas);

  /**
   * Limpa os campos que serão preenchidos pelo formulário
   */

  const colunasFormulario = [
    1,  // A
    2,  // B
    3,  // C
    4,  // D
    5,  // E
    6,  // F
    7,  // G
    8,  // H
    9,  // I
    10, // J
    11, // K
    12, // L
    13, // M
    14, // N
    15, // O
    16, // P
    17, // Q
    20, // T
    24, // X
    25, // Y
    26, // Z
    27, // AA
    28  // AB
  ];

  colunasFormulario.forEach(coluna => {

    aba
      .getRange(linhaDestino, coluna)
      .clearContent();

  });

}


/**
 * Salva os registros
 */
function salvarFormulario(dados) {

  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);

  const quantidade = Number(dados.quantidade) || 1;

  let linhaAtual = encontrarProximaLinha(aba);

  for (let i = 0; i < quantidade; i++) {

    prepararNovaLinha(aba, linhaAtual);

    const produto = dados.produtos[i] || {};

    /**
     * Dados gerais
     */

    aba.getRange(linhaAtual, 1).setValue(dados.dataReclamacao);
    aba.getRange(linhaAtual, 2).setValue(dados.empresa);
    aba.getRange(linhaAtual, 3).setValue(dados.codigoEmteco);
    aba.getRange(linhaAtual, 4).setValue(dados.dataCompra);
    aba.getRange(linhaAtual, 5).setValue(dados.reclamacaoDefeito);
    aba.getRange(linhaAtual, 6).setValue(dados.faseFalha);
    aba.getRange(linhaAtual, 7).setValue(dados.modalidade);
    aba.getRange(linhaAtual, 8).setValue(dados.numeroTarefa);

    /**
     * Dados individuais do produto
     */

    aba.getRange(linhaAtual, 9).setValue(produto.produto || '');
    aba.getRange(linhaAtual, 10).setValue(produto.qualidade || '');
    aba.getRange(linhaAtual, 11).setValue(produto.estoqueRetirada || '');
    aba.getRange(linhaAtual, 12).setValue(produto.codigoLote || '');

    /**
     * Continua dados gerais
     */

    aba.getRange(linhaAtual, 13).setValue(dados.responsavel);
    aba.getRange(linhaAtual, 14).setValue(dados.dataRecebimento);
    aba.getRange(linhaAtual, 15).setValue(dados.numeroControle);
    aba.getRange(linhaAtual, 16).setValue(dados.dataTeste);

    /**
     * Defeito individual
     */

    aba.getRange(linhaAtual, 17)
      .setValue(produto.defeitoIdentificado || '');

    /**
     * Observações
     */

    aba.getRange(linhaAtual, 20)
      .setValue(dados.observacaoExtra);

    aba.getRange(linhaAtual, 24)
      .setValue(dados.video);

    aba.getRange(linhaAtual, 25)
      .setValue(dados.observacaoFornecedor);

    aba.getRange(linhaAtual, 26)
      .setValue(dados.itemMovimentado);

    aba.getRange(linhaAtual, 27)
      .setValue(dados.tipoEstoque);

    aba.getRange(linhaAtual, 28)
      .setValue(dados.dataFinalizacao);

    linhaAtual++;

  }

  return {
    sucesso: true,
    mensagem: `${quantidade} produto(s) inserido(s) com sucesso.`,
    proximaLinha: linhaAtual
  };

}