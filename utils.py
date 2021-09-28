
from pandas import read_excel
from statistics import mode
from re import sub, findall
from pandas import concat, isna
from numpy import nan
import warnings
import base64

def fns(file: str, skip: int, range_cols: str):
	
	# Função trata o arquivo retornado pelo sistema do FNS

	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file, 
			skiprows=skip,
			usecols = range_cols
		)
	except:
		return print('erro ao ler o arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index') # exclui linhas e colunas a mais
	exclude_line = [df[i].count() for i in df.columns] 
	value_max = max(exclude_line)
	value_min = mode(exclude_line) # a maioria das colunas tem x linhas
	exclude_values = list(range(value_min+1, value_max+1)) # recorto nessas x linhas
	print('removendo '+str(len(exclude_values))+' linhas.')
	if value_max == 2:
		exclude_values = [1]
	else:
		exclude_values = exclude_values

	df = df.drop(exclude_values) # exclui linha sobressalentes

	list_integers = [
		'Nº OB', 'Banco OB', 'Agência OB', 'Conta OB', 'Processo'
	]

	for i in list_integers:
		df[i] = [int(n) for n in list(df[i])]

	list_values = ['Valor Total', 'Desconto', 'Valor Líquido']
	for i in list_values:
		df[i] = [''.join(findall('[\d\,]+', n)) for n in df[i]]
		df[i] = [sub(',', '.', n) for n in df[i]]
		df[i] = df[i].astype(float)

	if 'Nº Proposta' in df.columns:
		df['Nº Proposta'] = [int(i) for i in df['Nº Proposta']]

	return df

def valida_cnpj(cnpj):
	if len(cnpj) in 19:
		cnpj = cnpj[0:18]
		return cnpj
	else:
		return cnpj

def sigef(file: str, skip: int, range_cols: str):
	
	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip-2,
			usecols = range_cols
		)
	except:
		return print('erro ao ler arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')

	print('analisando credores...')
	dfx = df[
		(df.count(axis=1).isin([2])) &
		(df['Unnamed: 1'] != '210901 FES/Unidade Central - 21901 FES - Unidade Central')
	].filter(items=['Unnamed: 3'])

	df = df[
		df.count(axis=1).isin([12])
	]

	df.columns = [
		'PP', 'TIPO', 'OB', 'FONTE', 'DATA_PGO', 
		'NOTA_EMPENHO', 'ND', 'NL', 'DATA_NL', 
		'CE', 'N_PROCESSO', 'VALOR'
	]

	print('corrigindo número de processos (válido a partir de 2019)')
	df['N_PROCESSO'] = [sub('\D', '/', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [findall('[\d]+', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = ['/'.join(i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [sub('/21$', '/2021', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [sub('/20$', '/2020', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [sub('/19$', '/2019', i) for i in df['N_PROCESSO']]
	df['VALOR'] = df['VALOR'].astype(float)
	df['N_PROCESSO'] = df['N_PROCESSO'].astype(str)

	tabela = concat([df, dfx], axis=1)

	print('separando info credores...')
	tabela['Unnamed: 3'] = tabela['Unnamed: 3'].ffill()

	tabela = tabela.dropna(axis='index')

	tabela = tabela.rename(columns={'Unnamed: 3': 'CREDOR'})

	tabela['CREDOR'] = [i.split(' ', 1) for i in tabela['CREDOR']]
	tabela['CREDOR_CPFCNPJ'] = [i[0] for i in tabela['CREDOR']]
	tabela['CREDOR_NOME'] = [i[1] for i in tabela['CREDOR']]

	tabela = tabela.drop(columns=['CREDOR'])
	 
	return tabela

def valida_ob(ob: str, ano: str):
	ob = str(ob)
	if ob[0:4] == '2021' and len(ob) == 10:
		return sub(ob[0:4], ano+'OB', ob)
	else:
		return ob

def extrato(file: str, skip: int, range_cols: str):
	print('lendo arquivo...')

	with warnings.catch_warnings(record=True):
		warnings.simplefilter("always")
		df = read_excel(
			io=file, 
			skiprows=skip,
			usecols = range_cols
		)
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index') # exclui linhas e colunas a mais
	
	exclude_line = [df[i].count() for i in df.columns] 
	value_max = max(exclude_line)
	value_min = mode(exclude_line)
	exclude_values = list(range(value_min+1, value_max+1))
	print('removendo '+str(len(exclude_values))+' linhas.')
	df = df.drop(exclude_values) # exclui linha sobressalentes

	df.columns = [
		'DATA', 'OBS', 'DATA_BALANCETE', 'AGENCIA',
		'LOTE', 'N_DOCUMENTO', 'COD_HISTORICO', 'HISTORICO',
		'VALOR', 'DESCRICAO'
	]

	list_integers = [
		'AGENCIA', 'LOTE', 'N_DOCUMENTO', 'COD_HISTORICO'
	]

	for i in list_integers:
		df[i] = [int(n) for n in list(df[i])]

	df = df.drop(['OBS'], axis=1)
	df['N_DOCUMENTO'] = df['N_DOCUMENTO'].astype(str)
	df['N_DOCUMENTO'] = [valida_ob(i, ano='2021') for i in df['N_DOCUMENTO']]
	df['VALOR'] = [''.join(findall('[\d\,]+', i)) for i in df['VALOR']]
	df['VALOR'] = [sub(',', '.', i) for i in df['VALOR']]
	df['VALOR'] = df['VALOR'].astype(float)
	df['VALOR'] = ['{:.2f}'.format(i) for i in df['VALOR']]
	df['VALOR'] = df['VALOR'].astype(float)

	return df

def sigef2(file: str, skip: int, range_cols: str):

	colnames = [
		'NUMERO', 'DATA', 'DOM_BANC_ORIGEM', 
		'SITUACAO', 'PREP_PAG', 'FONTE',
		'FAVORECIDO', 'DOM_BANC_DEST', 'VALOR' 
	]

	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip-2,
			usecols = range_cols
		)
	except:
		return print('erro ao ler arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')
	
	info_ob = df[df['Unnamed: 0'].str.contains('OB', na=False)]
	info_ob = info_ob.drop(['Unnamed: 3', 'Unnamed: 5', 'Unnamed: 6'], axis=1)
	
	desc = df[df['Unnamed: 1'].str.contains('PP', na=False)]
	desc = desc.drop(['Unnamed: 0'], axis=1)
	
	tabela = concat([info_ob, desc], ignore_index=True, axis='columns')
	tabela[0] = tabela[0].ffill()
	tabela[1] = tabela[1].ffill()
	tabela[2] = tabela[2].ffill()
	tabela[3] = tabela[3].ffill() # completa informações
	
	tabela = tabela.dropna(axis='index', thresh=5) # remove informações até 5 não nas
	tabela = tabela.dropna(how='all', axis='columns')
	tabela.columns = colnames

	tabela['FAVORECIDO'] = [i.split(' ', 1) for i in tabela['FAVORECIDO']]
	tabela['FAVORECIDO_CPFCNPJ'] = [i[0] for i in tabela['FAVORECIDO']]
	tabela['FAVORECIDO_NOME'] = [i[1] for i in tabela['FAVORECIDO']]
	tabela = tabela.drop(columns=['FAVORECIDO'])
	tabela['VALOR'] = tabela['VALOR'].astype(float)

	return tabela

def sigef3(file: str, skip: int, range_cols: str):

	colnames = [
		'EMPENHADO', 'LIQUIDADO', 'RETIDO', 'A LIQUIDAR',
		'PAGO', 'A PAGAR', 'N_EMPENHO', 'N_PRE_EMPENHO', 'GESTAO', 'SUBACAO', 
		'FONTE', 'NATUREZA', 'CREDOR_CPFCNPJ', 'CREDOR_NOME',
	]

	reoder_colnames = [
		'N_EMPENHO', 'N_PRE_EMPENHO', 'GESTAO', 'SUBACAO', 'Nome Subação',
		'FONTE', 'NATUREZA', 'CREDOR_CPFCNPJ', 'CREDOR_NOME',
		'EMPENHADO', 'LIQUIDADO', 'RETIDO', 'A LIQUIDAR', 'PAGO', 'A PAGAR'
	]

	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip-2,
			usecols=range_cols
		)
	except:
		return print('erro ao ler arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')

	df = df[:-1]

	df['Unnamed: 2'] = [str(i).split(' / ') for i in df['Unnamed: 2']]
	df['N_EMPENHO'] = [i[0] for i in df['Unnamed: 2']]
	df['N_PRE_EMPENHO'] = [i[1] if len(i) == 2 else None for i in df['Unnamed: 2']]
	df = df.drop(columns=['Unnamed: 2'])

	df['Unnamed: 4'] = [str(i).split(' ') for i in df['Unnamed: 4']]
	df['GESTAO'] = [i[0] for i in df['Unnamed: 4']]
	df['SUBACAO'] = [i[1] if len(i) > 1 else None for i in df['Unnamed: 4']]
	df['FONTE'] = [i[2] if len(i) > 2 else None for i in df['Unnamed: 4']]
	df['NATUREZA'] = [i[3] if len(i) > 3 else None for i in df['Unnamed: 4']]
	df = df.drop(columns=['Unnamed: 4'])

	df['Unnamed: 5'] = [str(i).split(' ', 1) for i in df['Unnamed: 5']]
	df['CREDOR_CPFCNPJ'] = [i[0] for i in df['Unnamed: 5']]
	df['CREDOR_NOME'] = [i[1] if len(i) == 2 else None for i in df['Unnamed: 5']]
	df = df.drop(columns=['Unnamed: 5'])

	list_values = [
		'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 
		'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11'
	]

	for i in list_values:
		df[i] = [''.join(findall('[0-9|,]+', n)) for n in df[i]]
		df[i] = [sub(',', '.', n) for n in df[i]]
		df[i] = [sub('', '0.00', n) if n == '' else n for n in df[i]]

	df['Unnamed: 6'] = df['Unnamed: 6'].astype(float)
	df['Unnamed: 7'] = df['Unnamed: 7'].astype(float)
	df['Unnamed: 8'] = df['Unnamed: 8'].astype(float)
	df['Unnamed: 9'] = df['Unnamed: 9'].astype(float)
	df['Unnamed: 10'] = df['Unnamed: 10'].astype(float)
	df['Unnamed: 11'] = df['Unnamed: 11'].astype(float) 

	df.columns = colnames
	df = df[~isna(df['CREDOR_NOME'])]

	subacao = read_excel('subacao_complemento.xls', skiprows=12, usecols='B:D')
	subacao = subacao.dropna(how='all', axis='columns')
	subacao = subacao.dropna(how='all', axis='index')
	subacao = subacao[:-2]
	subacao['Código'] = subacao['Código'].astype(int)

	df['SUBACAO'] = df['SUBACAO'].astype(int)

	df = df.merge(subacao, how='left', left_on='SUBACAO', right_on='Código')
	df = df.drop(columns=['Código'])

	df = df.reindex(reoder_colnames, axis=1)

	return df

def sigef4(file: str, skip: int, range_cols: str):

	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip-1,
			usecols=range_cols
		)
	except:
		return print('erro ao ler arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')
	df = df.dropna(thresh=3, axis='index')
	
	df['Saldo'] = [abs(i) for i in df['Saldo']]
	df = df.dropna(how='all', axis='columns')

	df.columns = [
		'DATA', 'UNIDADE GESTORA', 
		'GESTAO', 'DOCUMENTO', 'EVENTO',
		'MOVIMENTO', 'TIPO_MOVIMENTO', 'SALDO', 'TIPO_MOVIMENTO_SALDO'
	]

	return df

def sigef5(file: str, skip: int, range_cols: str):

	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip-1,
			usecols=range_cols
		)
	except:
		return print('erro ao ler arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')
	df['Unnamed: 1'] = df['Unnamed: 1'].ffill()
	df['FONTE'] = [i if i[0:2] == '0.' else float('nan') for i in df['Unnamed: 1']]
	df['CATEGORIA'] = [i if i[0:2] != '0.' else 'NaN' for i in df['Unnamed: 1']]
	df['FONTE'] = df['FONTE'].ffill()
	df = df[~df['Unnamed: 2'].isna()]
	df['ATUALIZADO'] = df['Unnamed: 7'].astype(str) + df['Unnamed: 8'].astype(str)
	df['INDISPONIVEL'] = df['Unnamed: 9'].astype(str) + df['Unnamed: 10'].astype(str)
	df['PRE_EMPENHADO'] = df['Unnamed: 11'].astype(str) + df['Unnamed: 12'].astype(str)
	df['EMPENHADO'] = df['Unnamed: 13'].astype(str) + df['Unnamed: 14'].astype(str)
	df['DISPONIVEL'] = df['Unnamed: 15'].astype(str) + df['Unnamed: 16'].astype(str)
	df['LIQUIDADO'] = df['Unnamed: 17'].astype(str) + df['Unnamed: 18'].astype(str)
	df['PAGO'] = df['Unnamed: 19'].astype(str) + df['Unnamed: 20'].astype(str) 
	df['A_LIQUIDAR'] =  df['Unnamed: 21'].astype(str) + df['Unnamed: 22'].astype(str) 
	df['A_PAGAR'] =  df['Unnamed: 23'].astype(str) + df['Unnamed: 24'].astype(str)

	df['COD_FONTE'] = [i.split(' ', 1)[0] for i in df['FONTE']]
	df['DESC_FONTE'] = [i.split(' ', 1)[1] for i in df['FONTE']]

	# df2 = df[df['COD_FONTE'] == '0.1.14.000000']
	# print(df2)
	# df = df[df['COD_FONTE'] != '0.1.14.000000']

	df['COD_CATEGORIA_CONTA'] = [i.split(' ', 1)[0] for i in df['Unnamed: 2']]
	df['DESC_CATEGORIA_CONTA'] = [i.split(' ', 1)[1] for i in df['Unnamed: 2']]

	# print(df.iloc[1])
	df = df.drop(columns=[
		'Unnamed: 1', 'Unnamed: 7', 
		'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10',
		'Unnamed: 11', 'Unnamed: 12', 
		'Unnamed: 13', 'Unnamed: 14',
		'Unnamed: 15', 'Unnamed: 16', 
		'Unnamed: 17', 'Unnamed: 18',
		'Unnamed: 19', 'Unnamed: 20', 
		'Unnamed: 21', 'Unnamed: 22',
		'Unnamed: 23', 'Unnamed: 24', 
		'FONTE', 'Unnamed: 2'
	])

	df.columns = [
	 	'DOTACAO_INICIAL', 'CATEGORIA', 'ATUALIZADO', 'INDISPONIBILIDADES',
	 	'PRE_EMPENHADO', 'EMPENHADO', 'DISPONIVEL', 'LIQUIDADO', 'PAGO',
	 	'A LIQUIDAR', 'A PAGAR', 'COD_FONTE', 'DESC_FONTE', 'COD_CATEGORIA_CONTA', 'DESC_CATEGORIA_CONTA'
	]

	reoder_colnames = [
		'COD_FONTE', 'DESC_FONTE', 'CATEGORIA','COD_CATEGORIA_CONTA', 
		'DESC_CATEGORIA_CONTA', 'DOTACAO_INICIAL', 
		'ATUALIZADO', 'INDISPONIBILIDADES', 
	 	'PRE_EMPENHADO', 'EMPENHADO', 'DISPONIVEL', 'LIQUIDADO', 'PAGO',
	 	'A LIQUIDAR', 'A PAGAR'
	]

	df = df.reindex(reoder_colnames, axis=1)

	for i in df.columns:
		df[i] = [sub('nan','', i) for i in df[i].astype(str)]

	values = [
		'DOTACAO_INICIAL', 'ATUALIZADO', 'INDISPONIBILIDADES', 
		'PRE_EMPENHADO', 'EMPENHADO', 'DISPONIVEL', 'LIQUIDADO', 
		'PAGO', 'A LIQUIDAR', 'A PAGAR'
	]

	for i in values:
		df[i] = [sub('\.', '', i) for i in df[i]]
		df[i] = [sub('\,', '.', i) for i in df[i]]
		df[i] = df[i].replace('', '0.00')
		df[i] = df[i].astype(float)

	return df

def sigef6(file: str, skip: int, range_cols: str):

	colnames = [
		'DOCUMENTO', 'DATA', 'DATA_LANCAMENTO', 
		'CELULA_ORCAMENTARIA', 'CREDOR', 'EMPENHO' 
	]

	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip,
			usecols = range_cols
		)
	except:
		return print('erro ao ler arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')
	df = df.dropna(thresh=2, axis='index')
	
	df.columns = colnames

	df['SUBACAO'] = [i.split(' ')[1] for i in df['CELULA_ORCAMENTARIA']]
	df['FONTE'] = [i.split(' ')[2] for i in df['CELULA_ORCAMENTARIA']]
	df['NATUREZA'] = [i.split(' ')[3] for i in df['CELULA_ORCAMENTARIA']]

	df['VALOR_EMPENHO'] = [i.split(' ')[0] for i in df['EMPENHO']]
	df['VALOR_EMPENHO'] = [sub('\.', '', i) for i in df['VALOR_EMPENHO']]
	df['VALOR_EMPENHO'] = [sub('\,', '.', i) for i in df['VALOR_EMPENHO']]
	df['VALOR_EMPENHO'] = df['VALOR_EMPENHO'].astype(float)
	df['MOV_EMPENHO'] = [i.split(' ')[1] for i in df['EMPENHO']]

	subacao = read_excel('subacao_complemento.xls', skiprows=12, usecols='B:D')
	subacao = subacao.dropna(how='all', axis='columns')
	subacao = subacao.dropna(how='all', axis='index')
	subacao = subacao[:-2]
	subacao['Código'] = subacao['Código'].astype(int)

	df['SUBACAO'] = df['SUBACAO'].astype(int)

	df = df.merge(subacao, how='left', left_on='SUBACAO', right_on='Código')

	df['CREDOR_CPFCNPJ'] = [i.split(' ')[0] for i in df['CREDOR']]
	df['CREDOR_NOME'] = [' '.join(i.split(' ')[1:]) for i in df['CREDOR']]

	df = df.drop(columns=['CELULA_ORCAMENTARIA', 'Código', 'EMPENHO', 'CREDOR'])

	return df

def sigef7(file: str, skip: int, range_cols: str):

	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip,
			usecols = range_cols
		)
	except:
		return print('erro ao ler arquivo...')
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')

	print(df)

	df = df[df['Unnamed: 2'].isin(['Número', 'Ordem Bancária', 'Observação'])]
	df = df.filter(items=['Unnamed: 2', 'Unnamed: 9'])
	df = df.pivot(columns='Unnamed: 2', values='Unnamed: 9')

	numero = df['Número'].dropna().reset_index(drop=True)
	ob = df['Ordem Bancária'].dropna().reset_index(drop=True)
	obs = df['Observação'].dropna().reset_index(drop=True)

	if len(numero) == len(obs):
		tabela = concat([numero, obs], axis=1)
		tabela.columns = ['NUMERO', 'OBSERVACAO']
	elif len(ob) == len(obs): 
		tabela = concat([ob, obs], axis=1)
		tabela.columns = ['OB', 'OBSERVACAO']
	elif len(ob) == len(numero):
		tabela = concat([numero, ob, obs], axis=1)
		tabela.columns = ['NUMERO', 'OB', 'OBSERVACAO']
	else:
		return None

	return tabela

def export_data(data):
	print('exportando arquivo...')
	try:
		arquivo = data.to_csv(
			sep=';',
			decimal=',',
			float_format='%.2f',
			index=False
		).encode('utf-8-sig')
		print('arquivo exportado com sucesso!')
		return arquivo
	except:
		pass
		return print('erro ao exportar arquivo...')
