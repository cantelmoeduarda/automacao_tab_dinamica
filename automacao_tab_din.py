import pandas as pd
import xlsxwriter

# Ler o arquivo Excel
df = pd.read_excel('C:\\Users\\X15848318638\\OneDrive - CAMG\\Área de Trabalho\\Apuração do faturamento 202408.xlsx', sheet_name='Apuração do Faturamento')

# Filtrar o DataFrame para incluir apenas 'Pedido Original' na coluna 'Prioridade'
df['Prioridade'] = df['Prioridade'].str.strip()
df_pedido_original = df[df['Prioridade'] == 'Pedido Original']
df_pedido_original = df_pedido_original.rename(columns={'ID Pedido': 'Pedidos Originais'})

# Criar a tabela dinâmica
pivot_table = df_pedido_original.pivot_table(
    index='Órgão/Entidade',
    values='Pedidos Originais',
    aggfunc=pd.Series.nunique,
    margins=True,
    margins_name='Total Geral',
    dropna=True
)

# Garantir que o nome do índice esteja definido
if pivot_table.index.name is None:
    pivot_table.index.name = 'Órgão/Entidade'

# Criar um objeto ExcelWriter
excel_file = 'pivot_table.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

# Escrever a tabela dinâmica no Excel sem cabeçalhos e índices
pivot_table.to_excel(writer, sheet_name='PivotTable', startrow=1, header=False)

# Obter os objetos workbook e worksheet
workbook = writer.book
worksheet = writer.sheets['PivotTable']

# Definir formatos
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 0
})

cell_format = workbook.add_format({
    'border': 0
})

number_format = workbook.add_format({
    'num_format': '#,##0',
    'border': 0
})

total_format = workbook.add_format({
    'bold': True,
    'border': 0,
    'fg_color': '#D7E4BC',
    'num_format': '#,##0'
})

# Definir a largura das colunas
worksheet.set_column(0, 0, 70)  # Coluna "Órgão/Entidade"
worksheet.set_column(1, 1, 15)  # Coluna "Pedidos Originais"

# Escrever os cabeçalhos com formatação
worksheet.write(0, 0, pivot_table.index.name, header_format)
worksheet.write(0, 1, pivot_table.columns[0], header_format)

# Obter as dimensões do DataFrame
(max_row, max_col) = pivot_table.shape

# Escrever os dados e aplicar formatação
for row_num in range(max_row):
    index_value = pivot_table.index[row_num]
    if index_value == 'Total Geral':
        index_fmt = total_format
        data_fmt = total_format
    else:
        index_fmt = cell_format
        data_fmt = number_format
    worksheet.write(row_num + 1, 0, index_value, index_fmt)
    data_value = pivot_table.iloc[row_num, 0]
    worksheet.write(row_num + 1, 1, data_value, data_fmt)

# Salvar e fechar o arquivo Excel
writer.close()
