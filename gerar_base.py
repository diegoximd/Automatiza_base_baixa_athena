import pandas as pd
import datetime
import re


# Função para dividir telefones
def split_phones(phone_str):
    if pd.isna(phone_str):
        return ['', '', '', '', '', '']
    phones = [p.strip() for p in phone_str.split('|') if p.strip()]
    phones += [''] * (6 - len(phones))  # Preencher com vazios até 6
    return phones[:6]  # Garantir no máximo 6


# Carregar o arquivo de origem, forçando CPF_CNPJ_CAEPF como string
source_file = r'C:\Users\Suporte\Desktop\Mercado\Athena Saude\BASE EXECUTIVA - 11092025.xlsx'
df = pd.read_excel(source_file, sheet_name='Planilha1', header=0, dtype={'CPF_CNPJ_CAEPF': str})

# Limpar nomes de colunas
df.columns = df.columns.str.strip()

# Extrair data do nome do arquivo
date_match = re.search(r'(\d{8})\.xlsx$', source_file)
if date_match:
    date_str = date_match.group(1)
    remessa_date = datetime.datetime.strptime(date_str, '%d%m%Y').date()
else:
    raise ValueError("Não foi possível extrair a data do nome do arquivo.")

remessa_num = remessa_date.strftime('%y%m%d')  # ex: 250911 para 11/09/2025
excel_serial = remessa_date.strftime('%d/%m/%Y')  # Formatado como dd/mm/aaaa

# Normalizar espaços em ESTABELECIMENTO
df['ESTABELECIMENTO_norm'] = df['ESTABELECIMENTO'].str.replace(r'\s+', ' ', regex=True).str.strip()

# Definir os grupos baseados em ESTABELECIMENTO
ne_keywords = [
    'HUMANA ASSISTENCIA MEDICA - THE', 'HUMANA ASSISTENCIA MEDICA - MA',
    'HUMANA ASSISTENCIA MEDICA - PHB', 'HUMANA ASSISTENCIA MEDICA - FLORIANO',
    'HUMANA ASSISTENCIA MEDICA - PICOS', 'MEDPLAN', 'HUMANA ASSISTENCIA MEDICA - UNIHOSP',
    'HUMANA ASSISTENCIA MEDICA - NATAL', 'HUMANA ASSISTENCIA MEDICA - ONCOLIFE',
    'HUMANA SAUDE NORDESTE - CLINICA ABA TERESINA II FREI SERAFIM'
]
ne_keywords_norm = [re.sub(r'\s+', ' ', kw).strip() for kw in ne_keywords]

samp_keywords = [
    '007-00 SAMP ESPIRITO SANTO ASSISTENCIA MEDICA LTDA',
    '099-SAO BERNARDO SAUDE', '114 - CLINICA ABA SERRA'
]
samp_keywords_norm = [re.sub(r'\s+', ' ', kw).strip() for kw in samp_keywords]

sul_keywords = [
    'HUMANA SAUDE SUL LTDA – MARINGA E REGIAO', 'HUMANA SAUDE SUL LTDA – CAXIAS',
    'HUMANA SAUDE SUL LTDA – RONDON', 'HUMANA SAUDE SUL LTDA – HOSPITAL RONDON',
    'HUMANA SAUDE SUL LTDA – CAXIAS – MEDICINA OCUPACIONAL'
]
sul_keywords_norm = [re.sub(r'\s+', ' ', kw).strip() for kw in sul_keywords]

# Filtrar os DataFrames
df_ne = df[df['ESTABELECIMENTO_norm'].str.contains('|'.join(ne_keywords_norm), case=False, na=False)]
df_samp = df[df['ESTABELECIMENTO_norm'].str.contains('|'.join(samp_keywords_norm), case=False, na=False)]
df_sul = df[df['ESTABELECIMENTO_norm'].str.contains('|'.join(sul_keywords_norm), case=False, na=False)]


# Função para criar o DataFrame no formato do modelo
def create_model_df(input_df):
    # Cabeçalhos do modelo (a partir de row3)
    model_columns = [
        'TIPO', 'NR OPERAÇÃO', 'NOME OPERAÇÃO', 'AGENCIA', 'CONTA', 'PRODUTO', 'MODALIDADE',
        'DT. ATUALIZADO', 'DT. VENCIMENTO', 'VALOR OPERAÇÃO', 'VALOR VENCIDO', 'VALOR IOF',
        'COND. NEGOCIAIS', 'FORMA ATUALIZAÇÃO', 'GARANTIAS', 'NR. IDENTIDADE', 'CPF / CNPJ',
        'MCI', 'NR FICHA', 'NOME DO CLIENTE', 'ENDEREÇO', 'NUMERO', 'BAIRRO', 'CEP', 'CIDADE',
        'UF', 'TELEFONE 1', 'TELEFONE 2', 'TELEFONE 3', 'TELEFONE 4', 'TELEFONE 5', 'TELEFONE 6',
        'DATA NASCIMENTO', 'NATURALIDADE', 'SEXO', 'ESTADO CIVIL', 'NOME DO PAI', 'NOME DA MÃE',
        'NOME AVALISTA 1', 'CPF/CNPJ AVALISTA 1', 'ENDEREÇO AVALISTA 1', 'BAIRRO AVALISTA 1',
        'CEP AVALISTA 1', 'CIDADE AVALISTA 1', 'UF AVALISTA 1', 'TELEFONE 1 AVALISTA 1',
        'TELEFONE 2 AVALISTA 1', 'NOME AVALISTA 2', 'CPF/CNPJ AVALISTA 2', 'ENDEREÇO AVALISTA 2',
        'BAIRRO AVALISTA 2', 'CEP AVALISTA 2', 'CIDADE AVALISTA 2', 'UF AVALISTA 2',
        'TELEFONE 1 AVALISTA 2', 'TELEFONE 2 AVALISTA 2', 'NOME AVALISTA 3', 'CPF/CNPJ AVALISTA 3',
        'ENDEREÇO AVALISTA 3', 'BAIRRO AVALISTA 3', 'CEP AVALISTA 3', 'CIDADE AVALISTA 3',
        'UF AVALISTA 3', 'TELEFONE 1 AVALISTA 3', 'TELEFONE 2 AVALISTA 3', 'NOME AVALISTA 4',
        'CPF/CNPJ AVALISTA 4', 'ENDEREÇO AVALISTA 4', 'BAIRRO AVALISTA 4', 'CEP AVALISTA 4',
        'CIDADE AVALISTA 4', 'UF AVALISTA 4', 'TELEFONE 1 AVALISTA 4', 'TELEFONE 2 AVALISTA 4',
        'NOME AVALISTA 5', 'CPF/CNPJ AVALISTA 5', 'ENDEREÇO AVALISTA 5', 'BAIRRO AVALISTA 5',
        'CEP AVALISTA 5', 'CIDADE AVALISTA 5', 'UF AVALISTA 5', 'TELEFONE 1 AVALISTA 5',
        'TELEFONE 2 AVALISTA 5', 'NOME AVALISTA 6', 'CPF/CNPJ AVALISTA 6', 'ENDEREÇO AVALISTA 6',
        'BAIRRO AVALISTA 6', 'CEP AVALISTA 6', 'CIDADE AVALISTA 6', 'UF AVALISTA 6',
        'TELEFONE 1 AVALISTA 6', 'TELEFONE 2 AVALISTA 6', 'PROFISSÃO', 'NOME LOCAL DE TRABALHO',
        'ENDEREÇO LOCAL DE TRABALHO', 'BAIRRO LOCAL DE TRABALHO', 'CEP LOCAL DE TRABALHO',
        'CIDADE LOCAL DE TRABALHO', 'UF LOCAL DE TRABALHO', 'TELEFONE 1 LOCAL DE TRABALHO',
        'TELEFONE 2 LOCAL DE TRABALHO', 'REFERENCIA PESSOAL', 'TELEFONE 1 REFERENCIA',
        'TELEFONE 2 REFERENCIA', 'REFERENCIA PESSOAL 2', 'TELEFONE 1 REFERENCIA 2',
        'TELEFONE 2 REFERENCIA 2', 'REFERENCIA PESSOAL 3', 'TELEFONE 1 REFERENCIA 3',
        'TELEFONE 2 REFERENCIA 3', 'SPC/SERASA', 'E-MAIL', 'E-MAIL1', 'DT. EMISSÃO',
        'VALOR PROTESTO', 'OBS. OPERAÇÃO', 'OBS. CLIENTE', 'DT. FIMTERCERIZAÇÃO', 'VALOR JUROS',
        'COD_CLASSIFICACAO_CLIENTE', 'COD_CLASSIFICACAO_OPERACAO', 'DATA ASSINATURA DO CONTRATO',
        'SCORE', 'SCORE INTERNO', 'RENDA'
    ]

    # Criar DataFrame vazio com essas colunas
    output_df = pd.DataFrame(columns=model_columns)

    # Mapear as colunas conforme DE > PARA
    if not input_df.empty:
        output_df['NR OPERAÇÃO'] = input_df['DOCUMENTO'].astype(str)
        output_df['NOME OPERAÇÃO'] = input_df['PLANO']
        output_df['AGENCIA'] = input_df['LOCAL_PAGAMENTO']
        output_df['CONTA'] = input_df['CODIGO'].astype(str)  # Usar CODIGO do arquivo de origem como string
        output_df['PRODUTO'] = input_df['TIPO PLANO']
        output_df['MODALIDADE'] = input_df['MATRIZ DE OFERTA']
        output_df['VALOR OPERAÇÃO'] = input_df['VALOR_TOTAL']
        output_df['CPF / CNPJ'] = input_df['CPF_CNPJ_CAEPF'].astype(str)  # Manter formato original como string
        output_df['NOME DO CLIENTE'] = input_df['TITULAR']
        output_df['ENDEREÇO'] = input_df['ENDERECO']
        output_df['NUMERO'] = input_df['NUMERO']
        output_df['BAIRRO'] = input_df['BAIRRO']
        output_df['CEP'] = input_df['CEP']
        output_df['CIDADE'] = input_df['CIDADE']
        output_df['UF'] = input_df['ESTADO']
        output_df['E-MAIL'] = input_df['EMAIL']
        output_df['COND. NEGOCIAIS'] = input_df['SITUACAO']
        output_df['GARANTIAS'] = input_df['NATUREZA_CONTRATO']
        output_df['FORMA ATUALIZAÇÃO'] = input_df['ESTABELECIMENTO']
        output_df['OBS. OPERAÇÃO'] = input_df['STATUS']

        # Formatar DT. VENCIMENTO
        def format_date(x):
            if pd.isna(x):
                return ''
            if isinstance(x, datetime.datetime):
                return x.strftime('%d/%m/%Y')
            else:
                # Se for serial number, converter
                try:
                    excel_epoch = datetime.date(1899, 12, 30)
                    date_val = excel_epoch + datetime.timedelta(days=int(x))
                    return date_val.strftime('%d/%m/%Y')
                except:
                    return str(x)

        output_df['DT. VENCIMENTO'] = input_df['VENCIMENTO'].apply(format_date)

        # Preencher telefones
        phones = input_df['TELEFONE'].apply(split_phones)
        output_df['TELEFONE 1'] = phones.apply(lambda x: x[0])
        output_df['TELEFONE 2'] = phones.apply(lambda x: x[1])
        output_df['TELEFONE 3'] = phones.apply(lambda x: x[2])
        output_df['TELEFONE 4'] = phones.apply(lambda x: x[3])
        output_df['TELEFONE 5'] = phones.apply(lambda x: x[4])
        output_df['TELEFONE 6'] = phones.apply(lambda x: x[5])

        # Garantir que o campo TIPO seja preenchido com '1' como string
        output_df['TIPO'] = '1'  # Atribuir '1' diretamente
        output_df['TIPO'] = output_df['TIPO'].astype(str).fillna('1')  # Garantir que não haja NaN

    return output_df


# Criar DataFrames para cada saída
df_ne_model = create_model_df(df_ne)
df_samp_model = create_model_df(df_samp)
df_sul_model = create_model_df(df_sul)


# Função para escrever o arquivo no formato do modelo
def write_to_excel(output_file, model_df, company_code):
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
        # Escrever row1 (cabeçalho inicial)
        initial_header = pd.DataFrame(
            [['Dt. Remessa', 'Número da Remessa', 'Código da Empresa', 'Código de Evento Ref. A Atualização',
              'Retomar/Liquidar Operacao não Presentes'] + [''] * (len(model_df.columns) - 5)])
        initial_header.to_excel(writer, sheet_name='Modelo_Excel_Incluir_Clientes_I', index=False, header=False,
                                startrow=0)

        # Escrever row2
        numero_remessa = company_code + remessa_num
        row2_data = [excel_serial, numero_remessa, company_code, '', 'RETOMAR'] + [''] * (len(model_df.columns) - 5)
        pd.DataFrame([row2_data]).to_excel(writer, sheet_name='Modelo_Excel_Incluir_Clientes_I', index=False,
                                           header=False, startrow=1)

        # Escrever row3 (cabeçalhos dos dados)
        pd.DataFrame([model_df.columns]).to_excel(writer, sheet_name='Modelo_Excel_Incluir_Clientes_I', index=False,
                                                  header=False, startrow=2)

        # Escrever os dados a partir de row4
        model_df.to_excel(writer, sheet_name='Modelo_Excel_Incluir_Clientes_I', index=False, header=False, startrow=3)


# Escrever os arquivos com caminhos absolutos
write_to_excel(r'C:\Users\Suporte\Desktop\Mercado\Athena Saude\HUMANA_NE_ARQUIVO_BASE_800_ATHENA.xlsx', df_ne_model,
               '2002')
write_to_excel(r'C:\Users\Suporte\Desktop\Mercado\Athena Saude\SAMP_ARQUIVO_BASE_800_ATHENA.xlsx', df_samp_model,
               '2003')
write_to_excel(r'C:\Users\Suporte\Desktop\Mercado\Athena Saude\HUMANA_SUL_ARQUIVO_BASE_800_ATHENA.xlsx', df_sul_model,
               '2004')

print("Arquivos gerados com sucesso!")