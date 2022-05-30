
from pandas import read_excel, ExcelWriter
from re import sub, findall, match
from pandas import concat, isna
from numpy import nan
from io import BytesIO
import warnings
import base64

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

	df.to_excel('teste.xlsx')

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

	df.to_excel('Orcamento, Atualizado, Indisponivel e Pre Empenhado.xlsx', index = False)

	return df


print(
	orc(
		file = 'Imprimir Execução Orçamentária29042022014642.xls',
		skip = 19
	)
)	
