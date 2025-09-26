# Gerador de Arquivos Base e Baixa

Este projeto é uma aplicação Python com interface gráfica (usando `tkinter`) que gera arquivos Excel e SQL para operações de **Base** e **Baixa** com base em arquivos Excel de entrada. A aplicação suporta três empresas (Humana NE, SAMP, Humana SUL) e inclui validação de diretório de destino, filtragem de CPFs/CNPJs, correção automática de CNPJs com base no banco de dados, e formatação de datas compatível com o banco Firebird.

## Funcionalidades
- **Base**: Gera três arquivos Excel (`HUMANA_NE_ARQUIVO_BASE_800_ATHENA.xlsx`, `SAMP_ARQUIVO_BASE_800_ATHENA.xlsx`, `HUMANA_SUL_ARQUIVO_BASE_800_ATHENA.xlsx`) a partir de um arquivo de entrada, com campos como `TIPO`, `NR OPERAÇÃO`, `CPF / CNPJ`, `CONTA`, e `TELEFONE 1` a `6`.
- **Baixa**: Gera um arquivo Excel e um arquivo SQL para a empresa selecionada, com as seguintes funcionalidades:
  - Filtra `CPF / CNPJ` com 11 caracteres ou menos, substituindo por vazio (`""`).
  - Corrige automaticamente CNPJs com base na coluna `CGC` da tabela `CLIENTES`, usando a relação entre `CLIENTE` (de `OPERACOES`) e `CODIGO` (de `CLIENTES`).
  - Processa grandes listas de clientes em lotes para evitar limitações do Firebird (máximo de 65.535 valores na cláusula `IN`).
  - Garante compatibilidade de tipos de dados entre `CLIENTE` e `CODIGO` para evitar erros de merge.
  - Formata datas no padrão `aaaa-mm-dd` para o arquivo SQL.
  - Exibe mensagem informativa sobre o número de CNPJs corrigidos.
- Interface gráfica com seleção de arquivo de origem, tipo de arquivo, empresa, diretório de destino, e botão de ação.
- Validação de diretório de destino com mensagem de erro e seleção alternativa.
- Mensagem de sucesso para **Baixa** indicando o número de operações exportadas.

## Pré-requisitos
- Python 3.6 ou superior
- Bibliotecas Python:
  ```bash
  pip install pandas openpyxl firebirdsql
  ```
- Cliente Firebird instalado no PC de destino para operações de **Baixa** (disponível em https://firebirdsql.org/).
- Banco de dados Firebird acessível em `servidor2` com o caminho `D:\Dados_interbase\COB_DB_EXECUTIVA_ATHENA_SAUDE.FDB`.

## Instalação
1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/gerador-base-baixa.git
   cd gerador-base-baixa
   ```
2. Instale as dependências:
   ```bash
   pip install pandas openpyxl firebirdsql
   ```

## Uso
1. Execute o script:
   ```bash
   python gerar_base_baixa_gui.py
   ```
2. Na interface gráfica:
   - Clique em **Procurar** para selecionar o arquivo Excel de origem (ex.: `BASE EXECUTIVA - 11092025.xlsx` para Base ou `PAGAMENTO HUMANA NE - 01A15092025.xlsx` para Baixa).
   - Escolha o **Tipo de Arquivo** (Base ou Baixa).
   - Para **Baixa**, selecione a empresa (Humana NE, SAMP, ou Humana SUL).
   - Confirme o diretório de destino (`C:\Users\Suporte\Desktop\Mercado\Athena Saude` por padrão) ou clique em **Selecionar Destino** para escolher outro.
   - Clique em **Gerar Arquivo(s)**.
3. Verifique os arquivos gerados no diretório de destino:
   - Para **Base**: Três arquivos Excel.
   - Para **Baixa**: Um arquivo Excel e um arquivo SQL, com CNPJs corrigidos conforme a tabela `CLIENTES`.
4. Uma mensagem informará quantos CNPJs foram corrigidos (ex.: `5 CNPJs foram corrigidos com base na tabela CLIENTES.`).

## Geração do Executável
Para criar um executável independente (Windows):
1. Instale o PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Gere o executável:
   ```bash
   pyinstaller --onefile --windowed --name GeradorBaseBaixa gerar_base_baixa_gui.py
   ```
3. O executável será gerado em `dist\GeradorBaseBaixa.exe`.
4. Copie o executável para o PC de destino e instale o cliente Firebird, se necessário.

## Notas
- O arquivo de origem para **Base** deve conter colunas como `DOCUMENTO`, `CODIGO`, `CPF_CNPJ_CAEPF`, etc.
- O arquivo de origem para **Baixa** deve conter colunas como `Documento`, `Vencimento`, `CPF do titular`, etc.
- O arquivo SQL gerado para **Baixa** usa o formato de data `aaaa-mm-dd` (ex.: `2025-09-15`) compatível com o Firebird.
- CPFs/CNPJs com 11 caracteres ou menos são substituídos por `""` no arquivo de **Baixa**.
- CNPJs no arquivo de **Baixa** são corrigidos com base na coluna `CGC` da tabela `CLIENTES`, usando a relação entre `CLIENTE` (de `OPERACOES`) e `CODIGO` (de `CLIENTES`).
- Para grandes volumes de dados, a consulta à tabela `CLIENTES` é dividida em lotes de 1.000 registros para evitar erros de limite do Firebird.
- Tipos de dados de `CLIENTE` e `CODIGO` são convertidos para strings para garantir compatibilidade no merge.

## Solução de Problemas
- **Erro de conexão com o banco Firebird**: Verifique se o cliente Firebird está instalado e se o servidor (`servidor2`) está acessível.
- **Erro de limite do Firebird (`Too many values`)**: O script processa consultas em lotes de 1.000 valores para evitar esse erro. Se persistir, reduza o `batch_size` na função `create_baixa_df`.
- **Erro de merge (`object and int64 columns`)**: O script converte `CLIENTE` e `CODIGO` para strings para evitar incompatibilidade de tipos.
- **Diretório de destino inválido**: A aplicação exibirá uma mensagem de erro e permitirá selecionar um novo diretório.
- **Erro no executável**: Verifique se todas as dependências estão instaladas e se o arquivo `.spec` inclui `hiddenimports=['pandas', 'openpyxl', 'firebirdsql']` caso necessário.
- **CNPJs não corrigidos**: Confirme se as colunas `CLIENTE` (em `OPERACOES`) e `CODIGO` (em `CLIENTES`) estão preenchidas corretamente e correspondem entre si. Verifique se os valores de `NROPERACAO` no arquivo de entrada correspondem aos da tabela `OPERACOES`.

## Licença
Este projeto está licenciado sob a Licença MIT.