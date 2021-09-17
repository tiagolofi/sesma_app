
import streamlit as st
import utils
from pandas import read_csv

icon = 'https://bluefocus.com.br/sites/default/files/styles/medium/public/icon-financeiro.png'

st.set_page_config(
	layout='wide', 
	page_title='Tratamento das Fontes de Dados Fiscais',
	page_icon=icon,
	initial_sidebar_state='collapsed'
)

st.write('''
# Tratamento das Fontes de Receitas - SES/MA
''')

c1, c2, c3, c4 = st.columns(4)
with c1:
	type_problem = st.selectbox(
			label='Fonte de dados:',
			options=['FNS', 'Extrato Bancário', 'SIGEF']
		)
with c2:
	info_skip = st.number_input(label = 'Linhas iniciais para pular:', value=0)
with c3:
	info_range1 = st.text_input(label='Colunas Inicial:', help='ex: A ou a')
with c4:
	info_range2 = st.text_input(label='Coluna Final:', help='ex: B ou b')

file = st.file_uploader('Navegar pelo Computador:', ['xlsx', 'xls'])

try:
	info_range = info_range1.upper()+':'+info_range2.upper()
except:
	pass

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
		try:
			exportable = utils.export_data(data=tabela)
		except:
			st.write('Não há planilha para exportar')
	except:
		st.write('Erro ao definir características da planilha!') 
	try:
		st.write(tabela)
	except:
		st.write('Não é possível exibir a planilha')
	st.download_button(
			'Exportar planilha', 
			data=exportable,
			file_name=type_problem+'.csv',
			mime='text/csv'
		)
else:	
	st.write('')

