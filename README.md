# 🔄 Integração ZIG-Omie

Sistema automatizado de integração para transferência de notas fiscais entre as plataformas ZIG e Omie.

## 📋 Sumário

- [Visão Geral](#visão-geral)
- [Componentes Principais](#componentes-principais)
- [Tecnologias Utilizadas](#tecnologias-utilizadas)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Uso](#uso)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Fluxo de Execução](#fluxo-de-execução)

## 🎯 Visão Geral

O projeto automatiza a transferência de notas fiscais da plataforma ZIG para o sistema Omie, realizando a conversão de formatos e garantindo a sincronização dos dados entre as plataformas.

## 🛠️ Componentes Principais

### 1. Configuração (`Config`)
- Gerenciamento de variáveis de ambiente
- Carregamento de credenciais das APIs
- Configuração segura de tokens e chaves de acesso

### 2. Processamento de NF-e
- **`NFeProc`**: Extrator de dados XML
  - Processamento de notas fiscais
  - Extração de informações cruciais
  - Armazenamento estruturado

- **`parse_nfe_xml`**: Conversor XML para Python
  - Transformação de XMLs em dicionários
  - Validação de dados
  - Estruturação de informações

### 3. Integração com API ZIG (`fetch_invoices`)
- Requisições HTTP automatizadas
- Filtragem por período
- Tratamento de erros de API
- Validação de respostas

### 4. Geração de Arquivos
- **Excel**: 
  - Criação de planilhas formatadas
  - Cabeçalhos personalizados
  - Dados estruturados

- **JSON**:
  - Geração de arquivos com timestamp
  - Estrutura compatível com Omie
  - Backup de dados

### 5. Processamento Omie
- Conversão para formato Omie
- Preenchimento automático de campos
- Validação de dados
- Envio para API

## 🚀 Tecnologias Utilizadas

- Python 3.x
- Bibliotecas:
  - `requests`: Comunicação HTTP
  - `xml.etree.ElementTree`: Processamento XML
  - `openpyxl`: Manipulação Excel
  - `python-dotenv`: Configuração
  - `APScheduler`: Agendamento

## 💻 Instalação

```bash
# Clone o repositório
git clone https://github.com/CarlosRoGuerra/Integra-o-zig-omie
# Entre no diretório
cd zig-omie-integration

# Instale as dependências
pip install -r requirements.txt
```

## ⚙️ Configuração

1. Crie um arquivo `.env` na raiz do projeto:
```env
ZIG_TOKEN=seu_token_zig
OMIE_APP_KEY=sua_app_key_omie
OMIE_APP_SECRET=seu_app_secret_omie
```

2. Configure o agendamento em `config.py`:
```python
SCHEDULE_INTERVAL = 60  # minutos
```

## 📌 Uso

```bash
# Inicie a integração
python main.py
```

## 📂 Estrutura do Projeto

```
zig-omie-integration/
├── .env
├── main.py
└── requirements.txt
```

## 🔄 Fluxo de Execução

1. **Inicialização**
   - Carregamento de configurações
   - Validação de credenciais

2. **Busca de Notas (ZIG)**
   - Requisição de notas fiscais
   - Validação de respostas
   - Tratamento de erros

3. **Processamento**
   - Conversão XML → JSON
   - Formatação para Omie
   - Geração de arquivos

4. **Envio (Omie)**
   - Validação de dados
   - Envio para API
   - Registro de resultados

5. **Monitoramento**
   - Logs de execução
   - Tratamento de timeouts
   - Registro de erros

## ⏱️ Controle de Timeout

- Limite de 4 minutos por operação
- Tratamento automático de falhas
- Registro de operações interrompidas

## 📊 Logs e Monitoramento

- Registro detalhado de operações
- Acompanhamento de falhas
- Histórico de execuções

## 🤝 Contribuições

Contribuições são bem-vindas! Por favor, leia o guia de contribuição antes de submeter alterações.

## 📝 Licença

Este projeto está sob a licença Carlos Guerra.
