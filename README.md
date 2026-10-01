# Organizador de Vagas

Script em Python que transforma os e-mails de alerta de vagas (LinkedIn e Jobbol) em uma planilha organizada, filtrada e pronta para acompanhar candidaturas.

Resolve um problema simples: os alertas chegam em dezenas de e-mails, com vagas repetidas, de áreas e níveis variados, e fica difícil saber o que vale a pena e onde já me candidatei.

## O que ele faz

- Lê a caixa de entrada do Gmail em **modo somente leitura** (não altera nem marca nada como lido)
- Extrai **todas as vagas** de cada e-mail: cargo, empresa, local, salário (quando informado) e link
- Ignora e-mails que não são vagas (confirmações de candidatura, pesquisas, novidades)
- Remove vagas repetidas pelo número da vaga e mantém só as da área de desenvolvimento
- Dá uma **nota de compatibilidade** por palavras-chave (ex.: front-end, React, júnior, remoto) e descarta vagas acima do nível
- Gera o `vagas.xlsx` com filtros, links clicáveis, coluna **Remoto** e uma coluna de **Status** (Nova, Salva, Candidatado, Entrevista, Oferta, Descartada)
- Mantém o seu Status e as suas Notas quando o script é executado de novo
- Cria uma aba **Resumo** com fórmulas que contam as vagas por status

## Tecnologias

Python 3, `imaplib` e `email` (leitura do Gmail), `openpyxl` (planilha Excel) e `python-dotenv` (variáveis de ambiente).

## Como usar

1. Clone o repositório e instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

2. Ative a verificação em duas etapas na conta Google e crie uma **senha de app** em <https://myaccount.google.com/apppasswords>.

3. Copie o arquivo `.env.example` para `.env` e preencha com o seu e-mail e a senha de app:

   ```
   GMAIL_USER=seuemail@gmail.com
   GMAIL_APP_PASSWORD=sua-senha-de-app
   ```

4. Feche o `vagas.xlsx` se estiver aberto no Excel e execute:

   ```bash
   python organizador_vagas.py
   ```

## Personalização

No topo do script dá para ajustar:

- `PESOS`: palavras-chave e pontos que compõem a nota
- `REGEX_DEV`: quais cargos entram na planilha
- `IGNORAR_ASSUNTO`: tipos de e-mail que devem ser ignorados
- `DIAS` e `LIMITE_NOTA`: período de busca e nota mínima

## Segurança e privacidade

- A senha de app fica somente no arquivo `.env`, que está no `.gitignore` e **nunca deve ser publicado**
- O script apenas **lê** e-mails e não interage com a conta do LinkedIn, para respeitar os termos de uso da plataforma
- `vagas.xlsx` contém dados pessoais e também está no `.gitignore`

## Próximos passos

- Avaliação de compatibilidade com IA (comparar cada vaga com o meu perfil)
- Execução automática diária
- Registro automático das candidaturas enviadas
