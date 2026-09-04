# Emteco Manutenção

Sistema desenvolvido para centralizar e padronizar o registro de manutenções da Emteco Motores.

A aplicação conecta um formulário personalizado ao Google Sheets, permitindo registrar informações da reclamação, processo de manutenção, produtos analisados e resultados técnicos.

---

## Visão geral do processo

```text
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
```

---

## Etapas do sistema

```text
1. REGISTRO DA RECLAMAÇÃO

Primeira etapa responsável pela identificação do atendimento.

Data da reclamação
Empresa
Código Emteco
Data da compra
Reclamação
Fase da falha
Modalidade


2. PROCESSO DE MANUTENÇÃO

Informações internas relacionadas ao processo.

Número da tarefa
Responsável
Data de recebimento
Número de controle
Data de teste


3. CADASTRO DOS PRODUTOS

Permite registrar múltiplos produtos dentro da mesma solicitação.

Produto
Qualidade
Estoque
Código/Lote
Defeito identificado

A quantidade de produtos é definida pelo usuário e os campos são criados dinamicamente.


4. INFORMAÇÕES COMPLEMENTARES

Registro de informações adicionais relacionadas à análise.

Observações extras
Observações do fornecedor
Links de evidências
Data de finalização


5. PROCESSAMENTO E SALVAMENTO

Após o envio do formulário, o sistema executa:

Localização da próxima linha disponível
Preservação da estrutura da planilha
Registro individual dos produtos
Centralização das informações no Google Sheets
```

---

## Estrutura do projeto


emteco-manutencao
│
├── Code.js
│   └── Backend e integração com Google Sheets
│
└── Formulario.html
    └── Interface do sistema
```

---

## Tecnologias utilizadas

| Tecnologia | Função |
|---|---|
| Google Apps Script | Backend e automações |
| Google Sheets | Banco de dados |
| HTML | Estrutura da interface |
| CSS | Estilização |
| JavaScript | Interações e lógica |

---

## Objetivo

Automatizar o registro de manutenções e concentrar as informações técnicas em um único processo, reduzindo preenchimentos manuais e facilitando o acompanhamento dos registros.
