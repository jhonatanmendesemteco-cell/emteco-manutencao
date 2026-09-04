Emteco Manutenção

Sistema desenvolvido para centralizar e padronizar o registro de manutenções da Emteco Motores.

A aplicação conecta um formulário personalizado ao Google Sheets, permitindo registrar informações da reclamação, processo de manutenção, produtos analisados e resultados técnicos.

Visão geral do processo
INÍCIO
  │
  ▼
Registro da Reclamação
  │
  ▼
Informações da Manutenção
  │
  ▼
Cadastro dos Produtos
  │
  ▼
Análise dos Defeitos
  │
  ▼
Observações e Evidências
  │
  ▼
Registro no Google Sheets
Etapas do sistema
01. Registro da reclamação

Primeira etapa responsável pela identificação do atendimento.

Informações registradas:

Data da reclamação · Empresa · Código Emteco · Data da compra · Reclamação · Fase da falha · Modalidade

02. Processo de manutenção

Nesta etapa são registradas as informações internas relacionadas ao processo.

Número da tarefa · Responsável · Data de recebimento · Número de controle · Data de teste

03. Cadastro dos produtos

O sistema permite registrar múltiplos produtos dentro de uma mesma solicitação.

Cada produto possui informações individuais:

Produto · Qualidade · Estoque · Código/Lote · Defeito identificado

A quantidade de produtos é definida pelo usuário e os campos são criados dinamicamente.

04. Informações complementares

Etapa destinada ao registro de informações adicionais relacionadas à análise.

Observações extras · Observações do fornecedor · Links de evidências · Data de finalização

05. Processamento e salvamento

Após o envio do formulário, o sistema executa automaticamente:

Localização da próxima linha disponível
Identifica onde o novo registro será inserido.

Preservação da estrutura
Mantém fórmulas e formatação existentes na planilha.

Registro dos produtos
Cria registros individuais para cada produto informado.

Centralização das informações
Todos os dados são armazenados na aba Relatório mensal de manutenção.

Estrutura do projeto
emteco-manutencao
│
├── Code.js
│   └── Backend e integração com Google Sheets
│
└── Formulario.html
    └── Interface do sistema
Tecnologias utilizadas
Tecnologia	Função
Google Apps Script	Backend
Google Sheets	Banco de dados
HTML	Estrutura da interface
CSS	Estilização
JavaScript	Interações e lógica
Objetivo

Automatizar o registro de manutenções e concentrar as informações técnicas em um único processo, reduzindo preenchimentos manuais e facilitando o acompanhamento dos registros.
