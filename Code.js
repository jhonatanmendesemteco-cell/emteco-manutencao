/**
 * =========================================================================
 * CONFIGURAÇÕES GLOBAIS DO SISTEMA
 * =========================================================================
 */
const ABA_RELATORIO = 'Relatório mensal de manutenção';
const LINHA_INICIAL = 1623;
const ID_PLANILHA = '12SQ5q77gHH38KiAQNq05gyKy6728OKAA1WGQ008eqK4';

/** ABERTURA DO SISTEMA */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Formulario')
    .setTitle('Sistema de Controle de Manutenção - Emteco Motores');
}

/** MENU DA PLANILHA */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Emteco')
    .addItem('Novo Registro de Manutenção', 'abrirFormulario')
    .addToUi();
}

function abrirFormulario() {
  const html = HtmlService.createHtmlOutputFromFile('Formulario')
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sistema de Controle de Manutenção');
}

/** NORMALIZAÇÃO DO NÚMERO DA TAREFA */
function normalizarNumeroTarefa(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  let texto = String(valor).trim();
  if (/^\d+\.0$/.test(texto)) texto = texto.replace(/\.0$/, '');
  return texto;
}

/** CARREGAMENTO DOS SELECTS */
function obterDadosFormulario() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);
  if (!aba) throw new Error('A aba "' + ABA_RELATORIO + '" não foi encontrada.');

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

function obterValidacaoColuna(aba, coluna) {
  const celula = aba.getRange(coluna + LINHA_INICIAL);
  const validacao = celula.getDataValidation();
  if (!validacao) return [];

  const criterio = validacao.getCriteriaType();
  const valores = validacao.getCriteriaValues();

  if (criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return valores[0] || [];
  }

  if (criterio === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return valores[0].getValues().flat().filter(function(valor) {
      return valor !== '';
    });
  }
  return [];
}

/** LOCALIZAÇÃO DA PRÓXIMA LINHA */
function encontrarProximaLinha(aba) {
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < LINHA_INICIAL) return LINHA_INICIAL;

  const dados = aba.getRange(LINHA_INICIAL, 1, ultimaLinha - LINHA_INICIAL + 1, 8).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (dados[i].every(function(valor) { return valor === ''; })) {
      return LINHA_INICIAL + i;
    }
  }
  return ultimaLinha + 1;
}

/** PREPARAÇÃO DE NOVA LINHA */
function prepararNovaLinha(aba, linhaDestino) {
  if (linhaDestino <= LINHA_INICIAL) return;

  const origem = aba.getRange(linhaDestino - 1, 1, 1, 28);
  origem.copyTo(
    aba.getRange(linhaDestino, 1, 1, 28),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  aba.getRange(linhaDestino, 1, 1, 28).setFormulas(origem.getFormulas());

  const colunasFormulario = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,20,24,25,26,27,28];
  colunasFormulario.forEach(function(coluna) {
    aba.getRange(linhaDestino, coluna).clearContent();
  });
}

/** ESCREVE UMA LINHA DO FORMULÁRIO */
function preencherLinhaFormulario(aba, linha, dados, produto) {
  produto = produto || {};
  aba.getRange(linha, 1).setValue(dados.dataReclamacao || '');
  aba.getRange(linha, 2).setValue(dados.empresa || '');
  aba.getRange(linha, 3).setValue(dados.codigoEmteco || '');
  aba.getRange(linha, 4).setValue(dados.dataCompra || '');
  aba.getRange(linha, 5).setValue(dados.reclamacaoDefeito || '');
  aba.getRange(linha, 6).setValue(dados.faseFalha || '');
  aba.getRange(linha, 7).setValue(dados.modalidade || '');
  aba.getRange(linha, 8).setValue(normalizarNumeroTarefa(dados.numeroTarefa));
  aba.getRange(linha, 9).setValue(produto.produto || '');
  aba.getRange(linha, 10).setValue(produto.qualidade || '');
  aba.getRange(linha, 11).setValue(produto.estoqueRetirada || '');
  aba.getRange(linha, 12).setValue(produto.codigoLote || '');
  aba.getRange(linha, 13).setValue(dados.responsavel || '');
  aba.getRange(linha, 14).setValue(dados.dataRecebimento || '');
  aba.getRange(linha, 15).setValue(dados.numeroControle || '');
  aba.getRange(linha, 16).setValue(dados.dataTeste || '');
  aba.getRange(linha, 17).setValue(produto.defeitoIdentificado || '');
  aba.getRange(linha, 20).setValue(dados.observacaoExtra || '');
  aba.getRange(linha, 24).setValue(dados.video || '');
  aba.getRange(linha, 25).setValue(dados.observacaoFornecedor || '');
  aba.getRange(linha, 26).setValue(dados.itemMovimentado || '');
  aba.getRange(linha, 27).setValue(dados.tipoEstoque || '');
  aba.getRange(linha, 28).setValue(dados.dataFinalizacao || '');
}

/** SALVAR NOVO REGISTRO */
function salvarFormulario(dados) {
  if (!dados) throw new Error('Nenhum dado foi recebido.');
  const numeroTarefa = normalizarNumeroTarefa(dados.numeroTarefa);
  if (!numeroTarefa) throw new Error('Informe o número da tarefa.');
  dados.numeroTarefa = numeroTarefa;

  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);
  if (!aba) throw new Error('A aba de manutenção não foi encontrada.');

  const produtos = Array.isArray(dados.produtos) ? dados.produtos : [];
  const quantidade = produtos.length > 0 ? produtos.length : 1;
  let linhaAtual = encontrarProximaLinha(aba);

  for (let i = 0; i < quantidade; i++) {
    prepararNovaLinha(aba, linhaAtual);
    preencherLinhaFormulario(aba, linhaAtual, dados, produtos[i] || {});
    linhaAtual++;
  }

  return {
    sucesso: true,
    mensagem: quantidade + ' produto(s) inserido(s) com sucesso.',
    proximaLinha: linhaAtual
  };
}

/** FORMATAÇÃO DE DATA PARA O HTML */
function formatarDataParaInput(valor, fuso) {
  if (!valor) return '';
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, fuso, 'yyyy-MM-dd');
  }
  const texto = String(valor).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;
  return texto;
}

/** BUSCAR TAREFA EXISTENTE */
function buscarTarefa(numeroTarefa) {
  numeroTarefa = normalizarNumeroTarefa(numeroTarefa);
  if (!numeroTarefa) throw new Error('Informe o número da tarefa.');

  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);
  if (!aba) throw new Error('A aba de manutenção não foi encontrada.');

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < LINHA_INICIAL) {
    return { encontrado: false, mensagem: 'Nenhum registro encontrado.' };
  }

  const registros = aba.getRange(
    LINHA_INICIAL, 1, ultimaLinha - LINHA_INICIAL + 1, 28
  ).getValues();

  const fuso = ss.getSpreadsheetTimeZone();
  const encontrados = [];

  registros.forEach(function(registro, indice) {
    const tarefa = normalizarNumeroTarefa(registro[7]);
    if (tarefa === numeroTarefa) {
      encontrados.push({ linha: LINHA_INICIAL + indice, registro: registro });
    }
  });

  if (encontrados.length === 0) {
    return {
      encontrado: false,
      mensagem: 'Tarefa ' + numeroTarefa + ' não encontrada.'
    };
  }

  const principal = encontrados[0].registro;
  const produtos = encontrados.map(function(item) {
    const registro = item.registro;
    return {
      linha: item.linha,
      produto: registro[8] || '',
      qualidade: registro[9] || '',
      estoqueRetirada: registro[10] || '',
      codigoLote: registro[11] || '',
      defeitoIdentificado: registro[16] || ''
    };
  });

  return {
    encontrado: true,
    numeroTarefa: numeroTarefa,
    dataReclamacao: formatarDataParaInput(principal[0], fuso),
    empresa: principal[1] || '',
    codigoEmteco: principal[2] || '',
    dataCompra: formatarDataParaInput(principal[3], fuso),
    reclamacaoDefeito: principal[4] || '',
    faseFalha: principal[5] || '',
    modalidade: principal[6] || '',
    quantidade: produtos.length,
    produtos: produtos,
    responsavel: principal[12] || '',
    dataRecebimento: formatarDataParaInput(principal[13], fuso),
    numeroControle: principal[14] || '',
    dataTeste: formatarDataParaInput(principal[15], fuso),
    observacaoExtra: principal[19] || '',
    video: principal[23] || '',
    observacaoFornecedor: principal[24] || '',
    itemMovimentado: principal[25] || '',
    tipoEstoque: principal[26] || '',
    dataFinalizacao: formatarDataParaInput(principal[27], fuso)
  };
}

/** ATUALIZAR TAREFA EXISTENTE */
function atualizarFormulario(dados) {
  if (!dados) throw new Error('Nenhum dado foi recebido.');
  const numeroTarefa = normalizarNumeroTarefa(dados.numeroTarefa);
  if (!numeroTarefa) throw new Error('Informe o número da tarefa.');
  dados.numeroTarefa = numeroTarefa;

  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);
  if (!aba) throw new Error('A aba de manutenção não foi encontrada.');

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < LINHA_INICIAL) throw new Error('Não existem registros de manutenção.');

  const registros = aba.getRange(
    LINHA_INICIAL, 1, ultimaLinha - LINHA_INICIAL + 1, 28
  ).getValues();

  const linhasExistentes = [];
  registros.forEach(function(registro, indice) {
    if (normalizarNumeroTarefa(registro[7]) === numeroTarefa) {
      linhasExistentes.push(LINHA_INICIAL + indice);
    }
  });

  if (linhasExistentes.length === 0) {
    throw new Error('Tarefa ' + numeroTarefa + ' não encontrada.');
  }

  const produtos = Array.isArray(dados.produtos) ? dados.produtos : [];
  if (produtos.length === 0) {
    throw new Error('A tarefa precisa possuir pelo menos um produto.');
  }

  const quantidadeComum = Math.min(linhasExistentes.length, produtos.length);
  for (let i = 0; i < quantidadeComum; i++) {
    preencherLinhaFormulario(aba, linhasExistentes[i], dados, produtos[i]);
  }

  if (produtos.length > linhasExistentes.length) {
    for (let i = linhasExistentes.length; i < produtos.length; i++) {
      const novaLinha = encontrarProximaLinha(aba);
      prepararNovaLinha(aba, novaLinha);
      preencherLinhaFormulario(aba, novaLinha, dados, produtos[i]);
    }
  }

  if (produtos.length < linhasExistentes.length) {
    const colunasFormulario = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,20,24,25,26,27,28];
    for (let i = produtos.length; i < linhasExistentes.length; i++) {
      const linha = linhasExistentes[i];
      colunasFormulario.forEach(function(coluna) {
        aba.getRange(linha, coluna).clearContent();
      });
    }
  }

  return {
    sucesso: true,
    mensagem: 'Tarefa ' + numeroTarefa + ' atualizada com sucesso.'
  };
}

/** COMPATIBILIDADE COM A FUNÇÃO ANTIGA */
function atualizarTarefasEmLote(listaAtualizacoes) {
  if (!Array.isArray(listaAtualizacoes)) {
    throw new Error('Lista de atualizações inválida.');
  }

  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);
  if (!aba) throw new Error('A aba de manutenção não foi encontrada.');

  listaAtualizacoes.forEach(function(item) {
    const linha = Number(item.linha);
    if (!linha || linha < LINHA_INICIAL) return;

    if (item.produto !== undefined) aba.getRange(linha, 9).setValue(item.produto);
    if (item.qualidade !== undefined) aba.getRange(linha, 10).setValue(item.qualidade);
    if (item.estoqueRetirada !== undefined) aba.getRange(linha, 11).setValue(item.estoqueRetirada);
    if (item.codigoLote !== undefined) aba.getRange(linha, 12).setValue(item.codigoLote);
    if (item.responsavel !== undefined) aba.getRange(linha, 13).setValue(item.responsavel);
    if (item.dataRecebimento !== undefined) aba.getRange(linha, 14).setValue(item.dataRecebimento);
    if (item.numeroControle !== undefined) aba.getRange(linha, 15).setValue(item.numeroControle);
    if (item.dataTeste !== undefined) aba.getRange(linha, 16).setValue(item.dataTeste);
    if (item.defeitoIdentificado !== undefined) aba.getRange(linha, 17).setValue(item.defeitoIdentificado);
    if (item.observacaoExtra !== undefined) aba.getRange(linha, 20).setValue(item.observacaoExtra);
    if (item.video !== undefined) aba.getRange(linha, 24).setValue(item.video);
    if (item.observacaoFornecedor !== undefined) aba.getRange(linha, 25).setValue(item.observacaoFornecedor);
    if (item.itemMovimentado !== undefined) aba.getRange(linha, 26).setValue(item.itemMovimentado);
    if (item.tipoEstoque !== undefined) aba.getRange(linha, 27).setValue(item.tipoEstoque);
    if (item.dataFinalizacao !== undefined) aba.getRange(linha, 28).setValue(item.dataFinalizacao);
  });

  return { sucesso: true, mensagem: 'Registros atualizados com sucesso!' };
}

/** ENVIO MANUAL DE RELATÓRIO PDF POR E-MAIL */
function enviarRelatorioEmail(numeroTarefa, destinatarios) {
  numeroTarefa = normalizarNumeroTarefa(numeroTarefa);
  destinatarios = String(destinatarios || '').trim();

  if (!numeroTarefa) throw new Error('Informe o número da tarefa.');
  if (!destinatarios) throw new Error('Informe pelo menos um destinatário.');

  const emails = destinatarios.split(',').map(function(email) {
    return email.trim();
  }).filter(function(email) {
    return email !== '';
  });

  if (emails.length === 0) throw new Error('Informe pelo menos um destinatário.');

  const emailInvalido = emails.find(function(email) {
    return !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  });
  if (emailInvalido) throw new Error('E-mail inválido: ' + emailInvalido);

  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_RELATORIO);
  if (!aba) throw new Error('A aba de manutenção não foi encontrada.');

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < LINHA_INICIAL) throw new Error('Não existem registros de manutenção.');

  const dados = aba.getRange(
    LINHA_INICIAL, 1, ultimaLinha - LINHA_INICIAL + 1, 28
  ).getValues();

  const registros = dados.filter(function(linha) {
    return normalizarNumeroTarefa(linha[7]) === numeroTarefa;
  });

  if (registros.length === 0) {
    throw new Error('Nenhum registro encontrado para a tarefa ' + numeroTarefa + '.');
  }

  const fuso = ss.getSpreadsheetTimeZone();
  function formatarValor(valor) {
    if (valor instanceof Date && !isNaN(valor.getTime())) {
      return Utilities.formatDate(valor, fuso, 'dd/MM/yyyy');
    }
    return (valor === null || valor === undefined) ? '' : String(valor);
  }

  const campos = [
    ['Data da Reclamação', 0], ['Empresa / Cliente', 1], ['Código Emteco', 2],
    ['Data da Compra', 3], ['Reclamação / Defeito Relatado', 4], ['Fase da Falha', 5],
    ['Modalidade', 6], ['Número da Tarefa', 7], ['Produto', 8], ['Qualidade', 9],
    ['Estoque Retirada', 10], ['Código / Lote', 11], ['Responsável Técnico', 12],
    ['Data de Recebimento', 13], ['Número de Controle', 14], ['Data do Teste', 15],
    ['Defeito Identificado', 16], ['Observação Extra', 19], ['Link do Vídeo', 23],
    ['Observação Fornecedor', 24], ['Item Movimentado', 25], ['Tipo de Estoque', 26],
    ['Data de Finalização', 27]
  ];

  const documento = DocumentApp.create('Relatório de Manutenção - Tarefa ' + numeroTarefa);
  const corpo = documento.getBody();

  corpo.appendParagraph('EMTECO MOTORES').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  corpo.appendParagraph('Relatório de Manutenção').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  corpo.appendParagraph('Tarefa: ' + numeroTarefa);
  corpo.appendParagraph('Quantidade de registros: ' + registros.length);
  corpo.appendParagraph('');

  registros.forEach(function(registro, indice) {
    corpo.appendParagraph('Registro ' + (indice + 1))
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);

    const tabela = corpo.appendTable();
    campos.forEach(function(campo) {
      const valor = formatarValor(registro[campo[1]]);
      if (valor !== '') {
        const linha = tabela.appendTableRow();
        linha.appendTableCell(campo[0]);
        linha.appendTableCell(valor);
      }
    });
    corpo.appendParagraph('');
  });

  documento.saveAndClose();

  try {
    const arquivoPdf = DriveApp.getFileById(documento.getId())
      .getAs(MimeType.PDF)
      .setName('Relatorio_Manutencao_Tarefa_' + numeroTarefa + '.pdf');

    MailApp.sendEmail({
      to: emails.join(','),
      subject: 'Relatório de Manutenção - Tarefa ' + numeroTarefa,
      body: 'Olá!\n\nSegue em anexo o relatório de manutenção referente à tarefa ' +
        numeroTarefa + '.\n\nAtenciosamente,\nEmteco Motores',
      attachments: [arquivoPdf],
      name: 'Emteco Motores'
    });
  } finally {
    DriveApp.getFileById(documento.getId()).setTrashed(true);
  }

  return {
    sucesso: true,
    mensagem: 'Relatório enviado com sucesso para: ' + emails.join(', ')
  };
}
