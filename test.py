
import utils

# tabela = utils.fns(
# 	file = 'PlanilhaDetalhada3.xls',
# 	skip = 7,
# 	range_cols = 'B:X'
# )


# utils.export_data(data=tabela, output_name='qualquer')

tabela = utils.sigef(
	file='Imprimir Pagamento Efetuado13092021150639.xls',
	skip=21,
	range_cols='B:M'
)

print(tabela)

# tabela = extrato(
# 	file='Extrato3846684557.xlsx',
# 	skip=2,
# 	range_cols='A:J'
# )

# print(utils.valida_ob(ob='0000002021054715', ano='2021'))