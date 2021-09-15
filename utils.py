
from pandas import read_excel
from statistics import mode
from re import sub, findall
from pandas import concat
import warnings

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

	return df

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
		'NOTA_EMPENHO', 'ND', 'NL', 'Data NL', 
		'CE', 'N_PROCESSO', 'VALOR'
	]

	print('corrigindo número de processos (válido a partir de 2019)')
	df['N_PROCESSO'] = [sub('\D', '/', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [findall('[\d]+', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = ['/'.join(i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [sub('/21$', '/2021', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [sub('/20$', '/2020', i) for i in df['N_PROCESSO']]
	df['N_PROCESSO'] = [sub('/19$', '/2019', i) for i in df['N_PROCESSO']]
	# df['N_PROCESSO'] = [findall('[\d|\/]+', i)[0] for i in df['N_PROCESSO']]
	df['VALOR'] = df['VALOR'].astype(float)
	df['N_PROCESSO'] = df['N_PROCESSO'].astype(str)

	tabela = concat([df, dfx], axis=1)

	print('finalizando tratamento...')
	tabela['Unnamed: 3'] = tabela['Unnamed: 3'].ffill()

	tabela = tabela.dropna(axis='index')

	tabela = tabela.rename(columns={'Unnamed: 3': 'CREDOR'})

	return tabela

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

	df = df.drop(['OBS'], axis=1)

	df['HISTORICO'] = [i.strip() for i in df['HISTORICO']]

	return df

def export_data(data, output_name: str):
	print('exportando arquivo...')
	try:
		data.to_csv( 
			path_or_buf=output_name+'.csv', 
			sep=";", 
			decimal=',',
			index=False, 
			encoding='utf-8-sig'
		)
		return print('arquivo exportado com sucesso!')
	except:
		return print('erro ao exportar arquivo...')