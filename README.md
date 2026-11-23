# 📧 Automação de Envio de Pedidos de Compra via E-mail

Projeto desenvolvido em Python para automatizar o processo de extração de pedidos de compra no SAP, geração de arquivos PDF e preparação de e-mails pelo Microsoft Outlook.

A solução foi desenvolvida com o objetivo de reduzir atividades operacionais repetitivas, otimizar o fluxo de comunicação com fornecedores e diminuir o tempo gasto pelos compradores no envio individual de pedidos de compra.

## 🎯 Objetivo do Projeto

Em processos de compras, o envio de pedidos aos fornecedores envolve diversas etapas manuais:

- Consultar os pedidos de compra no SAP.
- Localizar e acessar cada pedido individualmente.
- Gerar e salvar os documentos em PDF.
- Identificar os setores responsáveis.
- Localizar os destinatários dos e-mails.
- Anexar os pedidos e preparar as mensagens individualmente.

Este projeto automatiza grande parte desse fluxo, utilizando Python para integrar diferentes ferramentas e executar as tarefas de forma sequencial.

## ⚙️ Funcionalidades

- **Leitura de planilhas Excel:** utiliza Pandas para importar os pedidos de compra e suas respectivas informações.
- **Integração com SAP GUI:** conecta-se a uma sessão ativa do SAP utilizando SAP GUI Scripting.
- **Extração automatizada de pedidos:** realiza consultas no SAP conforme o tipo de pedido de compra.
- **Geração de PDFs:** automatiza a exportação e o salvamento dos documentos em uma pasta específica.
- **Mapeamento de e-mails por setor:** consulta um arquivo JSON para identificar os destinatários correspondentes a cada setor.
- **Integração com Outlook:** prepara mensagens utilizando os dados dos pedidos e anexa os respectivos documentos PDF.
- **Logs de execução:** registra informações, etapas executadas e possíveis erros durante o processamento.
- **Barra de progresso:** utiliza TQDM para acompanhar visualmente a execução das tarefas.
- **Limpeza de arquivos temporários:** realiza a exclusão dos PDFs da pasta de processamento ao final da rotina.

## 🛠️ Tecnologias utilizadas

| Tecnologia | Aplicação |
|---|---|
| Python | Linguagem principal |
| Pandas | Manipulação de dados e planilhas |
| PyWinAuto | Automação de janelas do Windows |
| PyWin32 | Integração com SAP GUI e Outlook via COM |
| TQDM | Indicadores de progresso |
| Logging | Registro de eventos e erros |
| JSON | Armazenamento dos destinatários por setor |
| OS / Glob | Gerenciamento de arquivos e diretórios |
| Threading | Recursos de execução concorrente |

## 📁 Estrutura do Projeto

```text
envio_po_email/
│
├── pedidos_compra/
│   └── (Diretório de armazenamento temporário dos PDFs)
│
├── emails_setor.json
│
├── Follow-up Pedidos.xlsx
│
├── pedidos.XLSX
│
├── script.py
│
├── teste.py
│
├── extacao_sap.log
│
└── .gitignore
```

### Descrição dos arquivos

- `script.py`: código principal responsável pela automação.
- `emails_setor.json`: mapeamento dos destinatários de e-mail conforme o setor.
- `Follow-up Pedidos.xlsx`: planilha utilizada como fonte de dados dos pedidos liberados.
- `pedidos_compra/`: diretório utilizado para armazenar temporariamente os documentos PDF extraídos do SAP.
- `extacao_sap.log`: arquivo de registro das execuções e ocorrências.
- `teste.py`: arquivo auxiliar para testes e validações.

## 📊 Estrutura da planilha

O projeto utiliza a biblioteca Pandas para realizar a leitura da planilha Excel.

A aba utilizada no processamento é:

`po_liberados`

Principais colunas utilizadas:

| Coluna | Descrição |
|---|---|
| PO | Número do pedido de compra |
| TIPO RC | Tipo do pedido de compra |
| SETOR | Setor responsável |
| DATA DE ENTREGA | Prazo previsto para entrega |

## 📩 Configuração dos destinatários

Os endereços de e-mail são armazenados externamente no arquivo `emails_setor.json`, permitindo organizar os destinatários por setor sem precisar alterar diretamente o código Python.

Exemplo de estrutura:

```json
{
    "NUTRIÇÃO": [
        "nutricao@empresa.com",
        "responsavel@empresa.com"
    ],
    "MANUTENÇÃO": [
        "manutencao@empresa.com"
    ],
    "CENTRO CIRÚRGICO": [
        "centrocirurgico@empresa.com"
    ]
}
```

Os endereços acima são apenas exemplos ilustrativos.

## 🔄 Fluxo de execução

O funcionamento da automação segue as seguintes etapas:

1. **Leitura dos dados:** importa os pedidos de compra da planilha Excel.
2. **Conexão com SAP:** identifica e conecta-se a uma sessão ativa do SAP GUI.
3. **Identificação dos pedidos:** verifica o número e o tipo de cada pedido.
4. **Extração dos documentos:** executa as transações correspondentes no SAP e gera os arquivos PDF.
5. **Organização dos arquivos:** salva os documentos no diretório configurado.
6. **Identificação dos destinatários:** consulta o arquivo JSON para localizar os e-mails vinculados a cada setor.
7. **Preparação dos e-mails:** monta as mensagens automaticamente, adiciona informações do pedido e anexa o PDF correspondente.
8. **Finalização:** realiza a limpeza dos arquivos temporários e registra o resultado da execução nos logs.

## 📦 Instalação e configuração

### 1. Pré-requisitos

Para executar o projeto, é necessário possuir:

- Windows.
- Python instalado.
- SAP GUI instalado e configurado.
- SAP GUI Scripting habilitado e autorizado.
- Microsoft Excel.
- Microsoft Outlook Desktop.
- Acesso ao ambiente SAP e às transações utilizadas pelo projeto.

### 2. Instalação das dependências

Clone o repositório:

```bash
git clone https://github.com/SEU-USUARIO/envio-po-email.git
```

Acesse o diretório:

```bash
cd envio-po-email
```

Instale as dependências:

```bash
pip install pandas openpyxl tqdm pywinauto pywin32
```

### 3. Configuração

Antes de executar o projeto:

1. Configure o caminho de destino dos arquivos PDF no código.
2. Disponibilize a planilha Excel com a aba `po_liberados`.
3. Configure os destinatários no arquivo `emails_setor.json`.
4. Abra o SAP GUI e realize o login.
5. Verifique se o Outlook Desktop está configurado.

### 4. Execução

Execute o arquivo principal:

```bash
python script.py
```

## 📋 Monitoramento e tratamento de erros

O projeto utiliza a biblioteca `logging` para registrar as etapas de execução.

Os registros são armazenados no arquivo de log, permitindo acompanhar:

- Tentativas de conexão com o SAP.
- Pedidos processados.
- Falhas durante a extração.
- Problemas na localização dos PDFs.
- Erros na preparação dos e-mails.
- Exclusão dos arquivos temporários.

Esse mecanismo facilita a identificação de falhas e o acompanhamento do processamento.

## 🚀 Possíveis melhorias futuras

- Implementar validações completas da planilha antes de iniciar a execução.
- Criar uma interface gráfica para facilitar a utilização pelos compradores.
- Implementar relatórios de pedidos processados e pendentes.
- Adicionar mecanismos de recuperação de execução em caso de falhas.
- Implementar controle de envio para evitar duplicidade de e-mails.
- Permitir a configuração de diretórios e parâmetros sem alteração do código.
- Implementar testes automatizados para as principais funções.

## 👨‍💻 Desenvolvedor

**Guilherme Costa Barbosa**

Projeto desenvolvido com foco em automação de processos administrativos e integração entre sistemas corporativos, utilizando Python para transformar atividades operacionais repetitivas em fluxos automatizados.

