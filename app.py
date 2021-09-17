
import streamlit as st
import utils
from pandas import read_csv

icon = 'https://bluefocus.com.br/sites/default/files/styles/medium/public/icon-financeiro.png'

st.set_page_config(
	layout='wide', 
	page_title='Tratamento das Fontes de Dados',
	page_icon=icon,
	initial_sidebar_state='collapsed'
)

st.write('''
# Tratamento das Fontes de Dados - SES/MA
''')

c1, c2, c3, c4 = st.columns(4)
with c1:
	type_problem = st.selectbox(
			label='Fonte de dados:',
			options=['FNS - Investimento', 'Extrato Bancário', 'SIGEF - Listar Ordem', 'SIGEF - PP', 'SIGEF - Execução Financeira']
		)
with c2:
	info_skip = st.number_input(label = 'Linhas para pular:', value=0)
with c3:
	info_range1 = st.text_input(label='Coluna Inicial:', help='ex: A ou a')
with c4:
	info_range2 = st.text_input(label='Coluna Final:', help='ex: B ou b')

c5, c6 = st.columns(2)

with c5:
	file = st.file_uploader('Navegar pelo Computador:', ['xlsx', 'xls'])
with c6:
	st.write('''
	#### Aplicação para limpeza de dados das principais bases de dados usadas pelo Controle Financeiro da Secretaria de Estado da Saúde do Maranhão.\n
	''')

try:
	info_range = info_range1.upper()+':'+info_range2.upper()
except:
	pass

def create_data():
	if type_problem == 'FNS - Investimento':
		st.warning(
			'Selecione uma linha antes do cabeçalho da planilha (onde ficam os nomes das colunas).'
		)
		tabela = utils.fns(
			file = file,
			skip = info_skip,
			range_cols = info_range
		)
	elif type_problem == 'SIGEF - PP':
		st.warning(
			'Selecione a linha onde apareça o nome do primeiro credor.'
		)
		tabela = utils.sigef(
			file=file,
			skip=info_skip,
			range_cols=info_range
		)
	elif type_problem == 'SIGEF - Listar Ordem':
		st.warning(
			'Selecione a linha onde apareça a primeira ordem bancária.'
		)
		tabela = utils.sigef2(
			file=file,
			skip=info_skip,
			range_cols=info_range
		)
	elif type_problem == 'SIGEF - Execução Financeira':
		st.warning(
			'Selecione a primeira linha onde apareça o número do empenho e a partir da coluna C.'
		)
		tabela = utils.sigef3(
			file=file,
			skip=info_skip,
			range_cols=info_range
		)
	elif type_problem == 'Extrato Bancário':
		st.warning(
			'Selecione uma linha antes do cabeçalho da planilha (onde ficam os nomes das colunas)'
		)
		tabela = utils.extrato(
			file=file,
			skip=info_skip,
			range_cols=info_range
		)
	return tabela

try:
	if st.button('Visualizar planilha'):
		try:
			tabela = create_data()
			st.success('Limpeza feita com sucesso!')
		except:
			st.error('Insira um arquivo excel ou defina corretamente as características da planilha!')
			st.stop()
		try:
			exportable = utils.export_data(data=tabela)
		except:
			st.error('Sem arquivo para exportar...')
			st.stop() 
		try:
			with st.spinner('Tratando informações...'):
				st.write(tabela)
		except:
			st.error('Não é possível exibir a planilha!')
			st.stop()
		st.download_button(
				'Exportar planilha', 
				data=exportable,
				file_name=type_problem+'.csv',
				mime='text/csv'
			)
	else:
		st.stop()	
except:
	st.stop()

