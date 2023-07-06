import dateutil.utils
import pandas as pd
import fdb

dst_path = r'MTK:C:/Microsys/MsysIndustrial/Dados/MSYSDADOS.FDB'
excel_path = r'C:/Users/Gabriel/Desktop/boleto-gerar-cr.xlsx'
excel_path1 = r'C:/Users/Gabriel/Desktop/boleto-emitir.xlsx'
excel_path2 = r'C:/Users/Gabriel/Desktop/pix-emitir.xlsx'
a = dateutil.utils.today()

TABLE_NAME = 'RECEBER_TITULOS'
TABLE_NAME1 = 'PEDIDOS_VENDAS'
TABLE_NAME2 = 'CLIENTES'

SELECT = 'select REC_CLI_CODIGO, REC_PEDIDO, REC_STL_CODIGO, REC_VALOR from %s ' \
         'WHERE REC_STL_CODIGO = 5' % (TABLE_NAME)

SELECT1 = 'select PDV_CLI_CODIGO, PDV_NUMERO, PDV_FECHADO, PDV_PSI_CODIGO, PDV_TIPOPAGAMENTO,' \
          'PDV_VALORPRODUTOS from %s' % (TABLE_NAME1)

SELECT2 = 'select CLI_CODIGO, CLI_NOME from %s' % (TABLE_NAME2)
SELECT3 = 'select CLI_CODIGO, CLI_NOME from %s' % (TABLE_NAME2)

con = fdb.connect(dsn=dst_path, user='SYSDBA', password='masterkey', charset='UTF8')
cur = con.cursor()
####################################
cur.execute(SELECT)
table_rows = cur.fetchall()
####################################
cur.execute(SELECT1)
table_rows1 = cur.fetchall()
####################################
cur.execute(SELECT2)
table_rows2 = cur.fetchall()
####################################
cur.execute(SELECT3)
table_rows3 = cur.fetchall()
####################################
df = pd.DataFrame(table_rows)
####################################
dfx = pd.DataFrame(table_rows1)
df_pix = dfx
df_pix = df_pix[df_pix[4] == 'D']
df_pix = df_pix[df_pix[3] == 'AA']
dfx = dfx[dfx[4] == 'B']
dfx = dfx[dfx[3] == 'AA']
dfy = pd.DataFrame(table_rows2)
dfz = pd.DataFrame(table_rows3)
####################################
m = pd.merge(df,dfy, how='inner', on=0)
m = m.drop(columns=[0,2])
m = m.rename(columns={'1_x':'PEDIDO',3:'VALOR', '1_y':'NOME'})
####################################
n = pd.merge(dfx,dfz, how='inner', on=0)
n = n.drop(columns=[0,2,3,4])
n = n.rename(columns={'1_x':'PEDIDO',5:'VALOR', '1_y':'NOME'})
####################################
o = pd.merge(df_pix,dfz, how='inner', on=0)
o = o.drop(columns=[0,2,3,4])
o = o.rename(columns={'1_x':'PEDIDO',5:'VALOR', '1_y':'NOME'})
####################################
m.to_excel(excel_path1, index=False)
n.to_excel(excel_path, index=False)
o.to_excel(excel_path2, index=False)




