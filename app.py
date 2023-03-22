
import streamlit as st
from utils import * # fns, pagamento, extrato, listar_ordem, nota_empenho_celula, observacoes, situacao_pp, orc, listar_pre_empenho, nota_pre_empenho_celula, deta_conta, export_excel
from datetime import datetime

icon = 'https://bluefocus.com.br/sites/default/files/styles/medium/public/icon-financeiro.png'

st.set_page_config(
	page_title='Tratamento das Fontes de Dados',
	layout = 'wide',
	page_icon = icon,
	initial_sidebar_state = 'collapsed' 
)

css = """

<style>
	
	@import url('https://fonts.googleapis.com/css2?family=Poppins&display=swap');
	
	@font-face{
		font-family: 'Poppins', sans-serif;
	}

	html, body, [class*="css"] {
		font-family: 'Poppins', sans-serif;
	}

	.block-container {
		background-image: linear-gradient(to left, white, whitesmoke);
	}

</style>

"""

st.markdown(css, unsafe_allow_html=True)


c1, c2 = st.columns([3, 1])

with c1:

	st.write('''# **Tratamento das Fontes de Dados - SES/MA**''')

	st.write('''##### Versão 1.7.6''')

with c2:

	st.image('img/logo_ses.png')

c3, c4, c5 = st.columns(3)

with c3:

	type_problem = st.selectbox(
			label='Fonte de Informação',
			options=[
				'FNS', 'Extrato Bancário', 'Listar Ordem Bancária',
				'Imprimir Pagamento Efetuado', 'Imprimir Preparação Pagamento',
				'Listar Preparação Pagamento', 
				'Imprimir Nota Empenho Célula', 'Imprimir Nota Empenho Célula (2019)', 'Imprimir Nota Empenho Célula (2020)',
				'Imprimir Execução Orçamentária', 'Listar Pré-Empenho',
				'Imprimir Nota Pré-Empenho Célula', 'Detalhar Conta 8.2.1.7.2.01',
				'Imprimir Liquidação Credor', 'Imprimir Despesa Certificada Situação'
			]
		)

with c4:

	info_skip = st.number_input(label = 'Linhas para pular:', min_value = 0)

with c5:

	file = st.file_uploader('Navegar pelo Computador:', ['xlsx', 'xls'])

st.sidebar.write('''**Instruções de Linhas**''')

st.sidebar.write(
	'''
	FNS - 8 Linhas;\n
	Extrato Bancário - 2 Linhas;\n
	Listar Ordem Bancária - primeira ordem bancária;\n
	Imprimir Pagamento Efetuado - primeiro nome de credor;\n
	Imprimir Preparação Pagamento - primeira ordem bancária;\n
	Listar Preparação Pagamento - primeiro número de preparação de pagamento;\n
	Imprimir Nota Empenho Célula - primeiro nome de subfunção (Agrupamento Nível 1 deve ser "Subfunção");\n
	Imprimir Execução Orçamentária - primeio código de subação (19);\n
	Listar Pré Empenho - primeira nota de pré-empenho;\n
	Imprimir Nota Pré-Empenho Célula - primeira linha da UG;\n
	Detalhar Conta 8.2.1.7.2.01 - primeiro número de conta/fonte;\n
	Imprimir Liquidação Credor (NLs) - primeira nota de empenho;\n
	Imprimir Despesa Certificada Situação (Processos) - primeira linha da certificação
	'''
)

if type_problem == 'FNS' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = fns(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
	
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Extrato Bancário' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = extrato(file = file, skip = info_skip)
	
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Listar Ordem Bancária' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = listar_ordem(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Pagamento Efetuado' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = pagamento(file = file, skip = info_skip)
	
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Preparação Pagamento' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = observacoes(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Listar Preparação Pagamento' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = situacao_pp(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Nota Empenho Célula' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = nota_empenho_celula(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Nota Empenho Célula (2019)' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = nota_empenho_celula2(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Nota Empenho Célula (2020)' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = nota_empenho_celula3(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Execução Orçamentária' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = orc(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Listar Pré-Empenho' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = listar_pre_empenho(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Nota Pré-Empenho Célula' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = nota_pre_empenho_celula(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:
	
			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Detalhar Conta 8.2.1.7.2.01' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = deta_conta(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:

			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Liquidação Credor' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = liquidacao_credor(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:

			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')

elif type_problem == 'Imprimir Despesa Certificada Situação' and file != None:

	visualizar = st.button('Visualizar Planilha')

	if visualizar:

		try:
	
			data = despesa_certificada_situacao(file = file, skip = info_skip)
	
			st.dataframe(data)
	
			st.success('Arquivo lido com sucesso!')
			
			st.download_button(
				label = 'Baixar Planilha',
				data = export_excel(data = data),
				file_name = type_problem + ' ' + str(int(datetime.now().timestamp())) + '.xlsx'
			)
	
		except:

			st.error('Erro ao tentar ler o arquivo, verifique a quantidade de linhas a pular.')
