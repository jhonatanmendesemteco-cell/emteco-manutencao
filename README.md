Emteco Manutenção

Sistema desenvolvido para registrar e organizar os processos de manutenção da Emteco Motores.

A aplicação utiliza Google Apps Script integrado ao Google Sheets, permitindo registrar reclamações, produtos, defeitos e informações técnicas em um único fluxo.

Como funciona
1. Início do registro

O usuário acessa o formulário pelo Google Sheets ou pelo Web App e inicia um novo registro de manutenção.

2. Dados da reclamação

São preenchidas as informações principais, como:

Data da reclamação
Empresa
Código Emteco
Data da compra
Reclamação ou defeito
Fase da falha
Modalidade
3. Informações da manutenção

O sistema registra os dados do processo:

Número da tarefa
Responsável
Data de recebimento
Número de controle
Data de teste
Movimentação e tipo de estoque
4. Registro dos produtos

É possível adicionar múltiplos produtos no mesmo atendimento.

Para cada produto são registrados individualmente:

Produto
Qualidade
Estoque de retirada
Código ou lote
Defeito identificado
5. Observações e evidências

O formulário também permite incluir:

Observações extras
Observações do fornecedor
Links de vídeos ou evidências
Data de finalização
6. Salvamento automático

Ao finalizar, o sistema:

Localiza a próxima linha disponível na planilha
Mantém a estrutura e fórmulas existentes
Registra os dados gerais
Cria registros individuais para cada produto
Mantém todas as informações centralizadas no relatório de manutenção
Fluxo do sistema
Novo Registro
      ↓
Dados da Reclamação
      ↓
Informações da Manutenção
      ↓
Adicionar Produtos
      ↓
Observações e Evidências
      ↓
Salvar
      ↓
Google Sheets
Estrutura
Code.js
└── Backend e integração com Google Sheets

Formulario.html
└── Interface do formulário
Tecnologias
Google Apps Script
Google Sheets
HTML
CSS
JavaScript
Objetivo

Centralizar e padronizar o processo de registro de manutenção, facilitando o acompanhamento de reclamações, análise de defeitos e controle dos produtos recebidos.
