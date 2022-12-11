
from pandas import read_excel, ExcelWriter
from re import sub, findall, match
from pandas import concat, isna
from numpy import nan
from io import BytesIO
import warnings
import base64

def fns(file: str, skip: int):
	
	# Função trata o arquivo retornado pelo sistema do FNS

	df = read_excel(
		io = file, 
		skiprows = skip - 1,
		usecols = 'B:X'
	)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index') 
	df = df.dropna(thresh=5, axis='index')

	for i in ['Valor Total', 'Desconto', 'Valor Líquido']:

		df[i] = [float(sub('\,', '.', sub('\.', '', j))) for j in df[i]]

	return df.reset_index(drop=True)

def valida_cnpj(cnpj):
	
	if len(cnpj) == 19:
		
		return cnpj[0:18]
	
	else:
		
		return cnpj

def processo(text):

	try:

		proc = findall('\d{4,8}/\d{2,4}', text)[0]

		return proc.upper()

	except:

		return 'Processo não identificado'

def pagamento(file: str, skip: int):
	
	df = read_excel(
		io = file,
		skiprows = skip - 2,
		usecols = 'B:M'
	)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')

	## separando credores

	credores = df[(df.count(axis=1).isin([2])) & (df['Unnamed: 1'] != '210901 FES/Unidade Central - 21901 FES - Unidade Central')].filter(items=['Unnamed: 3'])

	df = df.dropna(thresh = 11, axis = 'index')

	df.columns = [
		'PreparacaoPagamento', 'Tipo', 'OrdemBancaria', 'Fonte', 'DataPagamento', 
		'NotaEmpenho', 'NaturezaDespesa', 'NotaLiquidacao', 'DataNotaLiquidacao', 
		'DespesaCertificada', 'Processo', 'Valor'
	]

	## padronização número de processo

	df['Processo'] = df['Processo'].apply(processo)
	df['Valor'] = df['Valor'].astype(float)
	df['Processo'] = df['Processo'].astype(str)

	## unindo credores ao valores

	tabela = concat([df, credores], axis=1)

	tabela['Unnamed: 3'] = tabela['Unnamed: 3'].ffill()

	tabela = tabela.dropna(axis='index')

	tabela = tabela.rename(columns={'Unnamed: 3': 'Credor'})

	tabela['CredorCpfCnpj'] = [i.split(' ', 1)[0] for i in tabela['Credor']]
	tabela['CredorNome'] = [i.split(' ', 1)[1] for i in tabela['Credor']]

	tabela = tabela.drop(columns=['Credor'])

	## melhorando para cálculos

	tabela = tabela.reindex([
		'PreparacaoPagamento', 'OrdemBancaria', 'NotaEmpenho', 'Processo', 
		'CredorCpfCnpj', 'CredorNome',
		'NotaLiquidacao', 'DespesaCertificada', 
		'Tipo', 'Fonte', 'NaturezaDespesa',
		'DataNotaLiquidacao', 'DataPagamento',
		'Valor'
	], axis = 'columns')

	return tabela.reset_index(drop=True)

def valida_numero_obpp(number: str):
	
	if number[0:4] in ['2022', '2021', '2020', '2019'] and len(number) == 10:
		
		return sub(number[0:4], number[0:4] + 'OB', number)
	
	elif number[0:4] in ['2022', '2021', '2020', '2019'] and len(number) == 14:
		
		return sub(number, number[0:4] + 'PP0' + number[5:10], number)

	else:

		return number

def extrato(file: str, skip: int):

	with warnings.catch_warnings(record=True):
		
		warnings.simplefilter("always")
		
		df = read_excel(
			io = file, 
			skiprows = skip,
			usecols = 'A:J'
		)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index') 
	df = df.dropna(thresh=5, axis='index')
		
	df.columns = [
		'Data', 'Observacao', 'DataBalancete', 'AgenciaOrigem',
		'Lote', 'NumeroDocumento', 'CodigoHistorico', 'Historico',
		'Valor', 'TipoMovimento'
	]

	df['NumeroDocumento'] = [valida_numero_obpp(number = str(int(i))) for i in df['NumeroDocumento']]

	df = df.drop(columns = ['Observacao', 'DataBalancete'])

	df = df.reindex([
		'NumeroDocumento', 'Data', 'AgenciaOrigem', 'Lote', 'CodigoHistorico', 'Historico', 'TipoMovimento', 'Valor'
	], axis = 'columns')

	df['Valor'] = [float(sub('\,', '.', sub('\.|\*|\ ', '', i))) for i in df['Valor']]

	return df.reset_index(drop=True)

def listar_ordem(file: str, skip: int):

	df = read_excel(
		io = file,  
		skiprows = skip - 2,
		usecols = 'A:H'
	)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')
	
	info_ob = df[df['Unnamed: 0'].str.contains('OB', na=False)]

	info_ob = info_ob.drop(['Unnamed: 3', 'Unnamed: 5', 'Unnamed: 6'], axis=1)
	
	info_ob['Unnamed: 2'] = [i[-7:] for i in info_ob['Unnamed: 2']]

	info_pp = df[df['Unnamed: 1'].str.contains('PP', na=False)]

	info_pp = info_pp.drop(['Unnamed: 0'], axis=1)
	
	tabela = concat([info_ob, info_pp], ignore_index=True, axis='columns')

	tabela[0] = tabela[0].ffill()
	tabela[1] = tabela[1].ffill()
	tabela[2] = tabela[2].ffill()
	tabela[3] = tabela[3].ffill() # completa informações
	
	tabela = tabela.dropna(thresh=5, how = 'index')

	tabela.columns = [
		'OrdemBancaria', 'DataOrdem', 'Conta', 'SituacaoOB',
		'PreparacaoPagamento', 'Fonte', 'Credor', 'ContaCredor',
		'Valor', 'SN'
	]

	tabela['CredorCpfCnpj'] = [i.split(' ', 1)[0] for i in tabela['Credor']]
	tabela['CredorNome'] = [i.split(' ', 1)[1] for i in tabela['Credor']]
	
	tabela['Valor'] = tabela['Valor'].astype(float)

	tabela = tabela.drop(columns = ['Credor', 'SN'])

	tabela = tabela.reindex([
		'OrdemBancaria', 'PreparacaoPagamento', 'Conta', 'Fonte', 
		'SituacaoOB', 'DataOrdem',
		'CredorCpfCnpj', 'CredorNome', 'ContaCredor',
		'Valor'
	], axis = 'columns')

	return tabela.reset_index(drop=True)

def check_pre_empenho(x):

	if len(x) > 1:

		return x[1]

	else:

		return None

def nota_empenho_celula(file: str, skip: int):

	df = read_excel(
		io = file, 
		skiprows = skip - 1, 
		usecols = 'A:L', 
		header = None
	)
	
	df = df.dropna(how = 'all', axis = 'index')
	df = df.dropna(how = 'all', axis = 'columns')
	
	df[0] = df[0].ffill()
	
	df[2] = df[2].astype(str)
	
	df = df[df[2].str.contains('NE')]
	
	for j in [6, 7, 8, 9, 10, 11]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', i)))) for i in df[j]]
	
	df['NotaEmpenho'] = [i.split(' / ')[0] for i in df[2]]

	df['NotaPreEmpenho'] = [check_pre_empenho(i.split(' / ')) for i in df[2]]

	df['Subacao'] = [i.split(' ')[1] for i in df[4]]
	df['Fonte'] = [i.split(' ')[2] for i in df[4]]
	df['Natureza'] = [i.split(' ')[3] for i in df[4]]
	
	df['CpfCnpj'] = [i.split(' ', 1)[0] for i in df[5]]
	df['Credor'] = [i.split(' ', 1)[1] for i in df[5]]
	
	df = df.drop(columns = [1, 2, 4, 5])
	
	df.columns = [
		'Subfuncao', 'Empenhado', 'Liquidado', 
		'Retido', 'ALiquidar', 'Pago', 'APagar', 
		'NotaEmpenho', 'NotaPreEmpenho', 'Subacao', 'Fonte', 'Natureza',
		'CpfCnpj', 'Credor'
	]

	subacao = read_excel('files/Relatorio_30052022092044.xls', skiprows=12, usecols='B:F', dtype=str)
	
	subacao = subacao.dropna(how='all', axis='columns')
	
	subacao = subacao.dropna(how='all', axis='index')
	
	subacao.columns = ['Codigo', 'SubacaoNome', 'Acao']

	df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	
	df = df.drop(columns=['Codigo'])
	
	df = df.reindex(
		[
			'Subfuncao', 'NotaEmpenho', 'NotaPreEmpenho', 'Acao', 'Subacao', 'SubacaoNome', 'Fonte', 
			'Natureza', 'CpfCnpj', 'Credor', 'Empenhado',
			'Liquidado', 'Retido', 'ALiquidar', 'Pago', 'APagar'
		], axis = 'columns'
	)

	return df

def nota_empenho_celula2(file: str, skip: int):

	df = read_excel(
		io = file, 
		skiprows = skip - 1, 
		usecols = 'A:O', 
		header = None
	)
	
	df = df.dropna(how = 'all', axis = 'index')
	df = df.dropna(how = 'all', axis = 'columns')
	
	df[0] = df[0].ffill()
	
	df[2] = df[2].astype(str)
	
	df = df[df[2].str.contains('NE')]
	
	for j in [6, 7, 8, 10, 12, 14]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', i)))) for i in df[j]]
	
	df['NotaEmpenho'] = [i.split(' / ')[0] for i in df[2]]

	df['PiCodigo'] = [i.split(' ')[1] for i in df[4]]
	df['Fonte'] = [i.split(' ')[2] for i in df[4]]
	df['Natureza'] = [i.split(' ')[3] for i in df[4]]
	
	df['CpfCnpj'] = [i.split(' ', 1)[0] for i in df[5]]
	df['Credor'] = [i.split(' ', 1)[1] for i in df[5]]
	
	df = df.drop(columns = [1, 2, 4, 5, 11])
	
	df.columns = [
		'Subfuncao', 'Empenhado', 'Liquidado', 
		'Retido', 'ALiquidar', 'Pago', 'APagar', 
		'NotaEmpenho', 'Subacao', 'Fonte', 'Natureza',
		'CpfCnpj', 'Credor'
	]

	subacao = read_excel('files/Relatorio_07122022152748.xls', skiprows=12, usecols='B:F', dtype=str)
	
	subacao = subacao.dropna(how='all', axis='columns')
	
	subacao = subacao.dropna(how='all', axis='index')
	
	subacao.columns = ['Codigo', 'SubacaoNome', 'Acao']

	df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	
	df = df.drop(columns=['Codigo'])
	
	df = df.reindex(
		[
			'Subfuncao', 'NotaEmpenho', 'Acao', 'Subacao', 'SubacaoNome', 'Fonte', 
			'Natureza', 'CpfCnpj', 'Credor', 'Empenhado',
			'Liquidado', 'Retido', 'ALiquidar', 'Pago', 'APagar'
		], axis = 'columns'
	)

	return df

def nota_empenho_celula3(file: str, skip: int):

	df = read_excel(
		io = file, 
		skiprows = skip - 1, 
		usecols = 'A:L', 
		header = None
	)
	
	df = df.dropna(how = 'all', axis = 'index')
	df = df.dropna(how = 'all', axis = 'columns')
	
	df[0] = df[0].ffill()
	
	df[2] = df[2].astype(str)
	
	df = df[df[2].str.contains('NE')]
	
	for j in [6, 7, 8, 9, 10, 11]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', i)))) for i in df[j]]
	
	df['NotaEmpenho'] = [i.split(' / ')[0] for i in df[2]]

	df['Subacao'] = [i.split(' ')[1] for i in df[4]]
	df['Fonte'] = [i.split(' ')[2] for i in df[4]]
	df['Natureza'] = [i.split(' ')[3] for i in df[4]]
	
	df['CpfCnpj'] = [i.split(' ', 1)[0] for i in df[5]]
	df['Credor'] = [i.split(' ', 1)[1] for i in df[5]]
	
	df = df.drop(columns = [1, 2, 4, 5])
	
	df.columns = [
		'Subfuncao', 'Empenhado', 'Liquidado', 
		'Retido', 'ALiquidar', 'Pago', 'APagar', 
		'NotaEmpenho', 'Subacao', 'Fonte', 'Natureza',
		'CpfCnpj', 'Credor'
	]

	subacao = read_excel('files/Relatorio_30052022092044.xls', skiprows=12, usecols='B:F', dtype=str)
	
	subacao = subacao.dropna(how='all', axis='columns')
	
	subacao = subacao.dropna(how='all', axis='index')
	
	subacao.columns = ['Codigo', 'SubacaoNome', 'Acao']

	df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	
	df = df.drop(columns=['Codigo'])
	
	df = df.reindex(
		[
			'Subfuncao', 'NotaEmpenho', 'Acao', 'Subacao', 'SubacaoNome', 'Fonte', 
			'Natureza', 'CpfCnpj', 'Credor', 'Empenhado',
			'Liquidado', 'Retido', 'ALiquidar', 'Pago', 'APagar'
		], axis = 'columns'
	)

	return df

def competencia(text):

	try:

		comp = findall('[A-Za-zÇç]{3,9}/\d{2,4}', text)[0]

		return comp.lower()

	except:

		if any(i in text.split(' ') for i in ['Única', 'UNICA', 'Unica', 'Única', 'Única;', 'UNICA;', 'Unica;', 'Única;']):

			return 'Parcela Única'

		else:

			return 'Competência não identificada'

def contrato(text):

	try:

		text = [i for i in text.split(' ') if i not in ['', ' ']]
	
		filter_list_text = [i for i in text if 'CT' in i]
	
		if len(filter_list_text) > 0:
	
			return ' '.join(['CT', str(text[text.index(filter_list_text[0]) + 1]).replace(';', '')])
	
		else:
	
			return 'Contrato não identificado'

	except:

		return 'Sem observação'

def observacoes(file: str, skip: int):

	df = read_excel(
		io = file,  
		skiprows = skip - 2,
		usecols = 'C:J'
	)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')

	# pivotagem

	df = df[df['Unnamed: 2'].isin(['Número', 'Ordem Bancária', 'Observação'])]
	df = df.filter(items=['Unnamed: 2', 'Unnamed: 9'])
	df = df.pivot(columns='Unnamed: 2', values='Unnamed: 9')

	df.columns = ['PreparacaoPagamento', 'Observacao', 'OrdemBancaria']

	df = df.reindex(['PreparacaoPagamento', 'OrdemBancaria', 'Observacao'], axis = 'columns')

	df['PreparacaoPagamento'] = df['PreparacaoPagamento'].ffill()
	df['OrdemBancaria'] = df['OrdemBancaria'].ffill(limit = 2)

	df = df[~isna(df['Observacao'])]

	df['Contrato'] = df['Observacao'].apply(contrato)

	df['Competencia'] = df['Observacao'].apply(competencia)

	df['Processo'] = df['Observacao'].apply(processo)

	return df

def situacao_pp(file: str, skip: int):

	data = read_excel(
		io = file,  
		skiprows = skip - 1,
		usecols = 'B:N',
		header = None
	)

	data[1] = data[1].astype(str)

	data = data[data[1].str.contains('PP')]

	data = data.dropna(how='all', axis='index')
	data = data.dropna(how='all', axis='columns')
	data = data.dropna(thresh=2, axis='index')

	data.columns = ['PreparacaoPagamento', 'OrdemBancaria', 'Favorecido', 'NotaEmpenho', 'DetalheNE', 'Valor', 'SituacaoPP']

	return data

def nivel(text):

	if match('(\d\ )+', text[0:2]):

		return 'Nível 1'

	elif match('(\d\d\ )+', text[0:3]):

		return 'Nivel 2'

	elif match('(\d\d\.\d\d\ )+', text[0:6]):

		return 'Nivel 3'

	elif match('(\d\d\.\d\d\.\d\d\ )+', text[0:9]):

		return 'Nivel 4'

def money(text):

	return float(sub('\,', '.', sub('\.', '', str(text))))

def orc(file: str, skip: int):
	
	df = read_excel(
		io = file, 
		skiprows = skip - 1,
		usecols = 'B:Z',
		header = None
	)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index') 

	df[1] = df[1].ffill()

	df['subacao'] = [i if match('(\d\d)+', i[0:2]) else 'Vazio' for i in df[1]]
	df['fonte'] = [i if match('(\d\.)+', i[0:2]) else 'Vazio' for i in df[1]]

	df['subacao'] = df['subacao'].replace('Vazio', nan)
	df['fonte'] = df['fonte'].replace('Vazio', nan)

	df['subacao'] = df['subacao'].ffill()
	df['fonte'] = df['fonte'].ffill()

	df = df.drop(columns = [1])

	df = df.dropna(subset=[2])

	df = df.filter(['subacao', 'fonte', 2, 6, 7, 8, 9, 10, 11, 12])

	df['nivel'] = df[2].apply(nivel)

	df[7] = df[7].replace(nan, '')

	df[8] = df[8].replace(nan, '')

	df[9] = df[9].replace(nan, '')

	df[10] = df[10].replace(nan, '')

	df[11] = df[11].replace(nan, '')

	df[12] = df[12].replace(nan, '')

	df[7] = df[7] + df[8]

	df[9] = df[9] + df[10]

	df[11] = df[11] + df[12]

	df[['CodigoSubacao', 'NomeSubacao']] = df['subacao'].str.split(' ', 1, expand = True)

	df[['Fonte', 'NomeFonte']] = df['fonte'].str.split(' ', 1, expand = True)

	df = df.drop(columns = [8, 10, 12, 'subacao', 'fonte', 'NomeFonte'])

	df.columns = ['Natureza', 'Dotacao', 'Atualizado', 'Indisponivel', 'PreEmpenhado', 'NivelNatureza', 'Subacao', 'NomeSubacao', 'Fonte']

	df = df.reindex(
		['Fonte', 'Subacao', 'NomeSubacao', 'Natureza', 'NivelNatureza', 'Dotacao', 'Atualizado', 'Indisponivel', 'PreEmpenhado'], 
		axis = 'columns'
	)

	df['Dotacao'] = df['Dotacao'].replace([nan, ''], 0)

	df['Atualizado'] = df['Atualizado'].replace([nan, ''], 0)

	df['Indisponivel'] = df['Indisponivel'].replace([nan, ''], 0)

	df['PreEmpenhado'] = df['PreEmpenhado'].replace([nan, ''], 0)

	df['Dotacao'] = df['Dotacao'].apply(money)
	df['Atualizado'] = df['Atualizado'].apply(money)
	df['Indisponivel'] = df['Indisponivel'].apply(money)
	df['PreEmpenhado'] = df['PreEmpenhado'].apply(money)

	return df

def simplifica_evento(x):

	eventos = {
		'RC08-Emissão de Pré-Empenho da Despesa': 'Emissão',
		'RC08-Anulação de Pré-Empenho da Despesa': 'Anulação',
		'RC08-Reforço de Pré-Empenho da Despesa': 'Reforço',
		'RC24 - Anulação de Pré-Empenho de Emenda Parlamentar': 'Anulação (EP)',
		'RC24 - Liberação da Emenda por Pré-Empenho.': 'Liberação (EP)'
	}

	return eventos.get(x)

def listar_pre_empenho(file, skip):

	df = read_excel(
		io = file,  
		skiprows = skip - 1,
		usecols = 'B:K',
		header = None
	)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')

	df = df[~isna(df[3])]

	df['Eventos'] = [simplifica_evento(i) for i in df[6]]

	df = df.drop(columns = [5, 6, 7])
	
	df.columns = ['NotaPreEmpenho', 'Data', 'Processo', 'Valor', 'Eventos']

	df = df.reindex(
		[
			'NotaPreEmpenho', 'Data', 'Eventos', 'Processo', 'Valor' 
		], axis = 'columns'
	)

	return df

def nota_pre_empenho_celula(file: str, skip: int):

	df = read_excel(
		io = file,  
		skiprows = skip - 1,
		usecols = 'B:N',
		header = None
	)

	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')

	df = df[~df[1].astype(str).str.contains('2022NE')]

	df = df.dropna(thresh=7, axis='index')

	df['Subacao'] = [i.split(' ')[1] for i in df[4]]
	df['Fonte'] = [i.split(' ')[2] for i in df[4]]
	df['Natureza'] = [i.split(' ')[3] for i in df[4]]

	for j in [6, 7, 8, 10, 12]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', i)))) for i in df[j]]

	df['Liquidado'] = df[8] - df[12]

	df = df.dropna(how='all', axis='columns')

	df = df.drop(columns = [4])

	df.columns = [
		'DataEmissao', 'NotaPreEmpenho',
		'PreEmpenhoOriginal', 'PreEmpenhoAtual', 'Empenhado', 'AEmpenhar', 'ALiquidar', 
		'Subacao', 'Fonte', 'Natureza', 'Liquidado'
	]

	subacao = read_excel('files/Relatorio_30052022092044.xls', skiprows=12, usecols='B:F', dtype=str)
	
	subacao = subacao.dropna(how='all', axis='columns')
	
	subacao = subacao.dropna(how='all', axis='index')
	
	subacao.columns = ['Codigo', 'SubacaoNome', 'Acao']

	df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	
	df = df.drop(columns=['Codigo'])
	
	df = df.reindex(
		[
			'DataEmissao', 'NotaPreEmpenho', 
			'Acao', 'Subacao', 'SubacaoNome', 'Fonte', 'Natureza', 
			'PreEmpenhoOriginal', 'PreEmpenhoAtual', 'Empenhado', 'AEmpenhar', 'ALiquidar'
		], axis = 'columns'
	)

	return df

def classifica_fonte(x):

	if x[0:6] in ['0.1.11', '0.3.11']:

		return 'Convênio Federal'

	elif x[0:6] in ['5.1.21', '5.3.21']:

		return 'Contrapartida de Convênio Federal'

	elif x[0:6] in ['0.3.08']:

		return 'Superávit Federal'

	elif x[0:6] in ['0.1.08']:

		return 'Corrente Federal'

	elif x[0:6] in [
		'0.1.21', '0.1.39', '0.1.14', '0.1.22', '0.1.36', '0.1.33',
		'0.3.21', '0.3.39', '0.3.14', '0.3.22', '0.3.36', '0.3.33'
	]:

		return 'Tesouro Estadual'

	elif x[0:6] in ['0.1.16', '0.1.34', '0.3.16', '0.3.34']:

		return 'Doações'

	elif x[0:6] in ['9.9.99']:

		return 'Obrigações e Consignações'

def deta_conta(file: str, skip: int):

	df = read_excel(io = file, skiprows = skip - 1, usecols = 'B:F', header = None)
	
	df = df.dropna(how='all', axis='columns')
	df = df.dropna(how='all', axis='index')
	
	df['Conta'] = [i.split(' ')[2] for i in df[1]]
	df['Fonte'] = [i.split(' ')[3] for i in df[1]]
	
	df = df.drop(columns = [1, 2, 4])
	
	df['TipoRecurso'] = df['Fonte'].apply(classifica_fonte)
	
	df = df.reindex(
		['Conta', 'Fonte', 'TipoRecurso', 5],
		axis = 'columns'
	)
	
	df = df[~isna(df[5])]
	
	df.columns = ['Conta', 'Fonte', 'TipoRecurso', 'Saldo em Conta']

	return df

def export_excel(data):

	output = BytesIO()
	
	writer = ExcelWriter(output)
	
	data.to_excel(writer, index=False, sheet_name='Plan1')
	
	writer.save()
	
	processed_data = output.getvalue()
	
	return processed_data
