
# import utils
# 
# print(
# 	utils.fns(
# 		file='C:/Users/usuario/Documents/Ambientes/SES/sesma/PlanilhaDetalhada4.xls',
# 		skip=7,
# 		range_cols='B:X'
# 	)
# )

import pandas

pp = pandas.read_csv('SIGEF - PP.csv', sep = ';')
execucao = pandas.read_csv('SIGEF - Execução Financeira.csv', sep = ';')
ordem = pandas.read_csv('SIGEF - Listar Ordem.csv', sep = ';')

print(pp.columns)
print(execucao.columns)
print(ordem.columns)

tabela = pp.merge(ordem, how='left', left_on='PP', right_on='PREP_PAG')
tabela = tabela.merge(execucao, how='left', left_on='NOTA_EMPENHO', right_on='N_EMPENHO')

print(tabela)

tabela.to_excel('monstrao.xlsx')