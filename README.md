# RPA Qualibank — Aprovação de Propostas

Automação em Python que acessa o portal **Quali** (joinbank) e apoia o processo de **Liberação de Propostas**, hoje feito manualmente por um analista/supervisor. O objetivo é **facilitar e agilizar** a triagem: o robô entra no sistema, abre cada proposta pendente, aplica a regra de valor e registra a decisão em um relatório — eliminando o trabalho repetitivo de abrir proposta por proposta só para checar o valor.

> ⚠️ **Modo atual: PRODUÇÃO.** O robô **aprova de verdade** as propostas com Valor Líquido ≤ R$ 10.000,00 (Ações → Aprovação Supervisor → Confirmar) e registra tudo em `resultado_aprovacoes.xlsx`. Não existe mais modo simulação.

## O que foi utilizado

| Item | Uso |
|---|---|
| **Python 3.14** | Linguagem da automação |
| **Playwright** | Controle do navegador (login, navegação, leitura das telas) |
| **python-dotenv** | Carrega credenciais/URL de um arquivo `.env` (fora do versionamento) |
| **openpyxl** | Geração do relatório final em Excel (`.xlsx`), com formatação e cores |
| **Git/GitHub** | Versionamento do código (branches `main` e `LAjustes`) |

## Estrutura do projeto

```
RPA_Qualibank/
├── login.py                  # Login no portal + fechamento do aviso de notificação
├── propostas.py              # Navegação, leitura das propostas, regra de decisão e relatório
├── requirements.txt          # Dependências Python
├── .env                      # Credenciais e URL (não versionado)
└── resultado_aprovacoes.xlsx  # Relatório gerado (não versionado)
```

## Regra de negócio aplicada

O critério de decisão é o **Valor Líquido** de cada proposta:

- **Valor Líquido ≤ R$ 10.000,00** → proposta *é aprovada*
- **Valor Líquido > R$ 10.000,00** → proposta é *pulada* (segue para análise manual)

## Como executar

1. Criar o arquivo `.env` na raiz do projeto:
   ```
   QUALI_URL=https://quali.joinbank.com.br/sign-in?redirectURL=%2Fmain
   QUALI_ACCESS_ID=seu-email@dominio.com
   QUALI_PASSWORD=sua-senha
   ```
2. Instalar as dependências:
   ```
   pip install -r requirements.txt
   playwright install chromium
   ```
3. Rodar o robô (aprova de verdade):
   ```
   python propostas.py
   ```
4. Abrir `resultado_aprovacoes.xlsx` — os resultados ficam na aba do ano corrente (ex.: `2026`).

## Relatório final

Planilha Excel com uma aba por ano da data de cadastro da proposta. Rodar novamente não duplica propostas já registradas (atualiza pelo Código do Contrato).

| Coluna | Origem |
|---|---|
| Código do Contrato | Seção "Proposta" |
| Nome da Pessoa | Seção "Dados Pessoais" |
| Valor Líquido | Seção "Proposta" — usado na regra de decisão |
| Data e Horário da Proposta | Seção "Log do Registro" (Data de Cadastro) |
| Data e Horário da Aprovação do Supervisor | Preenchido só quando a decisão é "Aprovado" |
| Decisão | "Aprovado" (linha verde) ou "Não Aprovado" (sem cor) |

## Fluxo do processo (BPM)

```mermaid
flowchart TD
    subgraph ROBO["Robô RPA — automático "]
        A([Início]) --> B[Login no portal Quali]
        B --> C[Fechar aviso de notificação]
        C --> D[Abrir lista de Empréstimos]
        D --> E{Há próxima<br/>proposta a processar?}
        E -- Sim --> F["Abrir proposta (Visualizar)"]
        F --> G[Ler Valor Líquido da proposta]
        G --> H{Valor Líquido<br/>≤ R$ 10.000,00?}
        H -- Sim --> I[Aprovar (Ações → Aprovação Supervisor)]
        H -- Não --> J[Pular proposta]
        I --> K[Registrar linha na planilha<br/>aba do ano — verde]
        J --> L[Registrar linha na planilha<br/>aba do ano — sem cor]
        K --> E
        L --> E
        E -- Não --> M([Salvar relatório .xlsx])
    end

    subgraph HUMANO["Supervisor — fluxo real (fora do escopo atual)"]
        N[Revisar propostas marcadas<br/>'Aprovado' no relatório]
        N --> O["Aprovar manualmente no sistema<br/>(Ações → Aprovação Supervisor)"]
    end

    M -. entrega o relatório para revisão .-> N
```

## Limitações e próximos passos

- A função `aprovar_proposta_real()` (em `propostas.py`) documenta como seria o clique real de aprovação (Ações → Aprovação Supervisor → observação "Aprovado via regra - RPA"), mas **nunca foi executada nem testada** — falta inclusive mapear o botão final de confirmação.
- `MAX_PROPOSTAS` controla quantas propostas são processadas por execução (ajustável em `propostas.py`).

## Segurança

- Credenciais ficam apenas no `.env`, que **não é versionado** (`.gitignore`).
- O relatório gerado (`resultado_aprovacoes.xlsx`) contém dados reais de clientes e também não é versionado.
