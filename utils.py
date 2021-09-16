
from pandas import read_excel
from statistics import mode
from re import sub, findall
from pandas import concat
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
	if len(cnpj) == 19:
		cnpj = cnpj[0:18]
		return cnpj
	else:
		return cnpj

def sigef(file: str, skip: int, range_cols: str):
	
	print('lendo arquivo...')
	try:
		df = read_excel(
			io=file,  
			skiprows=skip-1,
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

	tabela['CREDOR_CPFCNPJ'] = [''.join(findall('[\d\.\/\-]+', i)) for i in tabela['CREDOR']]
	tabela['CREDOR_CPFCNPJ'] = [valida_cnpj(i) for i in tabela['CREDOR_CPFCNPJ']]
	tabela['CREDOR_NOME'] = [' '.join(findall('[A-Z|ÇÃÁÂÁÀÉÊÍÓÔÚ]+', i)) for i in tabela['CREDOR']]

	tabela = tabela.drop(columns=['CREDOR'])
	 
	return tabela

def valida_ob(ob: str, ano: str):
	ob = str(ob)
	if ob[0:4] == '2021':
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