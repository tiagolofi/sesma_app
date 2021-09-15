
import streamlit as st
import utils
from pandas import read_csv

icon = 'https://bluefocus.com.br/sites/default/files/styles/medium/public/icon-financeiro.png'

tabela = None

st.set_page_config(
	layout='wide', 
	page_title='Tratamento das Fontes de Dados Fiscais',
	page_icon=icon,
	initial_sidebar_state='collapsed'
)

st.write('''
# Tratamento das Fontes de Receitas - SES/MA
''')

c1, c2, c3 = st.columns(3)
with c1:
	type_problem = st.selectbox(
			label='Fonte de dados:',
			options=['FNS', 'Extrato Bancário', 'SIGEF']
		)
with c2:
	info_skip = st.number_input(label = 'Linhas iniciais para pular:', value=0)
with c3:
	info_range = st.text_input(label='Colunas para recortar:', help='e.g: B:M')

file = st.file_uploader('Navegar pelo Computador:', ['xlsx', 'xls'])

def create_data():
	if type_problem == 'FNS':
		tabela = utils.fns(
			file = file,
			skip = info_skip,
			range_cols = info_range
		)
	elif type_problem == 'SIGEF':
		tabela = utils.sigef(
			file=file,
			skip=info_skip,
			range_cols=info_range
		)
	elif type_problem == 'Extrato Bancário':
		tabela = utils.extrato(
			file=file,
			skip=info_skip,
			range_cols=info_range
		)
	return tabela

if st.button('Visualizar planilha'):
	try:
		tabela = create_data()
	except:
		pass 
	st.write(tabela)
else:	
	st.write('')

st.download_button(
	'Exportar planilha', 
	data=utils.export_data(data=tabela),
	file_name=type_problem+'corrigido.csv',
	mime='text/csv'
)