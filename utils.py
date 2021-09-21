
from pandas import read_excel
from statistics import mode
from re import sub, findall
from pandas import concat, isna
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
	value_min = mode(exclude_line)
	exclude_values = list(range(value_min+1, value_max+1))
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

	colnames2 = [
		'NUMERO', 'DATA', 'DOM_BANC_ORIGEM', 
		'PREP_PAG', 'FONTE',
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
	desc = desc.drop(['Unnamed: 0', 'Unnamed: 7'], axis=1)

	tabela = concat([info_ob, desc], axis='columns')
	tabela['Unnamed: 0'] = tabela['Unnamed: 0'].ffill()
	tabela['Unnamed: 1'] = tabela['Unnamed: 1'].ffill()
	tabela['Unnamed: 2'] = tabela['Unnamed: 2'].ffill()
	tabela['Unnamed: 7'] = tabela['Unnamed: 7'].ffill()
	tabela = tabela.dropna(subset=colnames2)

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