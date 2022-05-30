from pandas import read_excel
from re import sub

def nota_empenho_celula_rpp(file: str, skip: int):

	df = read_excel(
		io = file, 
		skiprows = skip - 1, 
		usecols = 'C:Y', 
		header = None
	)
	
	df = df.dropna(how = 'all', axis = 'index')
	df = df.dropna(how = 'all', axis = 'columns')
		
	df[2] = df[2].ffill()

	df = df[df[6].astype(str).str.contains('NE')]
	
	df[2] = df[2].fillna('122 Administração Geral')

	for j in [15, 17, 21, 23]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', str(i))))) for i in df[j]]

	df['Subacao'] = [i.split(' ')[1] for i in df[9]]
	df['Fonte'] = [i.split(' ')[2] for i in df[9]]
	df['Natureza'] = [i.split(' ')[3] for i in df[9]]
	
	df['CpfCnpj'] = [i.split(' ', 1)[0] for i in df[12]]
	df['Credor'] = [i.split(' ', 1)[1] for i in df[12]]

	df = df.drop(columns = [5, 9, 12])

	df = df.dropna(how = 'all', axis = 'columns')

	df.columns = [
		'Subfuncao', 'NotaEmpenho', 
		'DataReferencia', 'DataLancamento', 
		'Inscrito', 'Pago', 
		'Cancelado', 'Retido', 'APagar',
		'Subacao', 'Fonte', 'Natureza',
		'CpfCnpj', 'Credor'
	]

	# subacao = read_excel('files/Relatorio_31012022112606.xls', skiprows=12, usecols='B:D')
	# 
	# subacao = subacao.dropna(how='all', axis='columns')
	# 
	# subacao = subacao.dropna(how='all', axis='index')
	# 
	# subacao.columns = ['Codigo', 'SubacaoNome']

	# df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	# 
	# df = df.drop(columns=['Codigo'])
	
	df = df.reindex(
		[
			'Subfuncao', 'NotaEmpenho', 'Subacao', 'Fonte',
			'Natureza', 'CpfCnpj', 'Credor',
			'Inscrito', 'Pago', 'Cancelado', 'Retido', 'APagar'
		], axis = 'columns'
	)

	return df

def nota_empenho_celula_rpnp(file: str, skip: int, bi: str):

	if bi == '6':

		cols = 'C:AB'

	else:

		cols = 'C:Z'

	df = read_excel(
		io = file, 
		skiprows = skip - 1, 
		usecols = cols, 
		header = None
	)
	
	df = df.dropna(how = 'all', axis = 'index')
	df = df.dropna(how = 'all', axis = 'columns')
		
	df[2] = df[2].ffill()

	df = df[df[6].astype(str).str.contains('NE')]
	
	df[2] = df[2].fillna('122 Administração Geral')

	df = df.dropna(how = 'all', axis = 'columns')

	print(df)

	for j in [14, 16, 18, 20, 22, 24]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', str(i))))) for i in df[j]]

	df['Subacao'] = [i.split(' ')[1] for i in df[9]]
	df['Fonte'] = [i.split(' ')[2] for i in df[9]]
	df['Natureza'] = [i.split(' ')[3] for i in df[9]]
	
	df['CpfCnpj'] = [i.split(' ', 1)[0] for i in df[12]]
	df['Credor'] = [i.split(' ', 1)[1] for i in df[12]]

	df = df.drop(columns = [5, 9, 12])
 
	if bi == '6':

		df.columns = [
			'Subfuncao', 'NotaEmpenho', 
			'DataReferencia', 'DataLancamento', 
			'Inscrito', 'Cancelado', 'Pago', 'ALiquidar', 'LiquidadoAPagar', 'Retido', 'APagar',
			'Subacao', 'Fonte', 'Natureza',
			'CpfCnpj', 'Credor'
		]

	else:

		df.columns = [
			'Subfuncao', 'NotaEmpenho', 
			'DataReferencia', 'DataLancamento', 
			'Inscrito', 'Cancelado', 'Pago', 'ALiquidar', 'LiquidadoAPagar', 'Retido', 'APagar',
			'Subacao', 'Fonte', 'Natureza',
			'CpfCnpj', 'Credor'
		]

	# subacao = read_excel('files/Relatorio_31012022112606.xls', skiprows=12, usecols='B:D')
	# 
	# subacao = subacao.dropna(how='all', axis='columns')
	# 
	# subacao = subacao.dropna(how='all', axis='index')
	# 
	# subacao.columns = ['Codigo', 'SubacaoNome']

	# df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	# 

	if bi == '6':

		df = df.reindex(
			[
				'Subfuncao', 'NotaEmpenho', 'Subacao', 'Fonte', 
				'Natureza', 'CpfCnpj', 'Credor',
				'Inscrito', 'Cancelado', 'Pago', 'ALiquidar', 'LiquidadoAPagar', 'Retido', 'APagar'
			], axis = 'columns'
		)

	else:
		
		df = df.drop(columns=['Retido', 'APagar'])

		df = df.reindex(
			[
				'Subfuncao', 'NotaEmpenho', 'Subacao', 'Fonte', 
				'Natureza', 'CpfCnpj', 'Credor',
				'Inscrito', 'Pago', 'ALiquidar', 'LiquidadoAPagar'
			], axis = 'columns'
		)

	return df

# while True:
# 
# 	numero = input('numero: ')
# 
# 	# x = nota_empenho_celula_rpp(file = 'raps/' + numero + 'b - processado.xls', skip = 18)
# 
# 	# print(
# 	# 	x
# 	# )
# 
# 	# x.to_excel(numero + 'bproc.xlsx', index = False)
# 	
# 	y = nota_empenho_celula_rpnp(file = 'raps/' + numero + 'b - nao processado.xls', skip = 18, bi = numero)
# 	
# 	print(
# 		y
# 	)
# 	
# 	y.to_excel(numero + 'bnproc.xlsx', index = False)

def nota_empenho_celula_rpp(file: str, skip: int):

	df = read_excel(
		io = file, 
		skiprows = skip,
		usecols = 'F:Y', 
		header = None
	)
	
	df = df.dropna(how = 'all', axis = 'index')
	df = df.dropna(how = 'all', axis = 'columns')
	df = df.dropna(thresh=6, axis = 'index')
	
	for j in [15, 17, 19, 20, 21, 23]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', str(i))))) for i in df[j]]

	df['Subacao'] = [i.split(' ')[1] for i in df[9]]
	df['Fonte'] = [i.split(' ')[2] for i in df[9]]
	df['Natureza'] = [i.split(' ')[3] for i in df[9]]
	
	df['CpfCnpj'] = [i.split(' ', 1)[0] for i in df[12]]
	df['Credor'] = [i.split(' ', 1)[1] for i in df[12]]

	df = df.drop(columns = [5, 9, 12])

	df = df.dropna(how = 'all', axis = 'columns')

	df.columns = [
		'NotaEmpenho', 
		'DataReferencia', 'DataLancamento', 
		'Inscrito', 'Pago', 
		'Cancelado', 'Retido', 'APagar',
		'Subacao', 'Fonte', 'Natureza',
		'CpfCnpj', 'Credor'
	]

	subacao = read_excel('files/Relatorio_31012022112606.xls', skiprows=12, usecols='B:D')
	
	subacao = subacao.dropna(how='all', axis='columns')
	
	subacao = subacao.dropna(how='all', axis='index')
	
	subacao.columns = ['Codigo', 'SubacaoNome']

	df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	
	df = df.drop(columns=['Codigo'])
	
	df = df.reindex(
		[
			'NotaEmpenho', 'Subacao', 'SubacaoNome', 'Fonte',
			'Natureza', 'CpfCnpj', 'Credor',
			'Inscrito', 'Pago', 'Cancelado', 'Retido', 'APagar'
		], axis = 'columns'
	)

	df.to_excel('teste.xlsx', index=False)

	return df


# print(nota_empenho_celula_rpp(file = 'Imprimir Nota Empenho Célula Assíncrono21032022180635.xls', skip=18))

def nota_empenho_celula_rpnp(file: str, skip: int):

	df = read_excel(
		io = file, 
		skiprows = skip, 
		usecols = 'F:AB', 
		header = None
	)
	
	df = df.dropna(how = 'all', axis = 'index')
	df = df.dropna(how = 'all', axis = 'columns')
	df = df.dropna(thresh=6, axis = 'index')

	for j in [14, 16, 18, 20, 22, 24]:
	
		df[j] = [float(sub(' ', '0', sub('\,', '.', sub('[A-Z]|\.', '', str(i))))) for i in df[j]]

	df['Subacao'] = [i.split(' ')[1] for i in df[9]]
	df['Fonte'] = [i.split(' ')[2] for i in df[9]]
	df['Natureza'] = [i.split(' ')[3] for i in df[9]]
	
	df['CpfCnpj'] = [i.split(' ', 1)[0] for i in df[12]]
	df['Credor'] = [i.split(' ', 1)[1] for i in df[12]]

	df = df.drop(columns = [5, 9, 12])

	print(df.columns)

	df.columns = [
		'NotaEmpenho', 
		'DataReferencia', 'DataLancamento', 
		'Inscrito', 'Cancelado', 'Pago', 'ALiquidar', 'LiquidadoAPagar', 'Retido',
		'Subacao', 'Fonte', 'Natureza',
		'CpfCnpj', 'Credor'
	]

	df.to_excel('teste.xlsx', index =False)

	# subacao = read_excel('files/Relatorio_31012022112606.xls', skiprows=12, usecols='B:D')
	# 
	# subacao = subacao.dropna(how='all', axis='columns')
	# 
	# subacao = subacao.dropna(how='all', axis='index')
	# 
	# subacao.columns = ['Codigo', 'SubacaoNome']

	# df = df.merge(subacao, how='left', left_on='Subacao', right_on='Codigo')
	# 

	df = df.reindex(
		[
			'NotaEmpenho', 'Subacao', 'Fonte', 
			'Natureza', 'CpfCnpj', 'Credor',
			'Inscrito', 'Cancelado', 'Pago', 'ALiquidar', 'LiquidadoAPagar', 'Retido', 'APagar'
		], axis = 'columns'
	)

	df.to_excel('teste.xlsx', index=False)

	return df

print(nota_empenho_celula_rpnp(file = 'Relatorio_21032022180803.xls', skip = 18))