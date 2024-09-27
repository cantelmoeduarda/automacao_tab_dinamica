import pandas as pd
import xlsxwriter

# Ler o arquivo Excel
df = pd.read_excel('C:\\Users\\X15848318638\\OneDrive - CAMG\\Área de Trabalho\\Apuração do faturamento 202408.xlsx', sheet_name='Apuração do Faturamento')

# Filtrar o DataFrame para incluir apenas 'Pedido Original' na coluna 'Prioridade'
df['Prioridade'] = df['Prioridade'].str.strip()
df_pedido_original = df[df['Prioridade'] == 'Pedido Original']
df_pedido_original = df_pedido_original.rename(columns={'ID Pedido':'Pedidos Originais'})

# Criar a tabela dinâmica
pivot_table = df_pedido_original.pivot_table(
    index='Órgão/Entidade',
    values='Pedidos Originais',
    aggfunc=pd.Series.nunique,
    margins=True,
    margins_name='Total Geral',
    dropna=True
)

# Criar um objeto ExcelWriter
excel_file = 'pivot_table.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

# Escrever a tabela dinâmica no Excel
pivot_table.to_excel(writer, sheet_name='PivotTable')

# Obter os objetos workbook e worksheet do xlsxwriter
workbook = writer.book
worksheet = writer.sheets['PivotTable']

# Adicionar formatação
format1 = workbook.add_format({'num_format': '#,##0'})

format_custom = workbook.add_format({
    'bold': False,             # Fonte em negrito

})

# Definir a largura e o formato das colunas
worksheet.set_column('A:A', 70,format_custom)  # Coluna "Órgão/Entidade"
worksheet.set_column('B:B', 15)  # Coluna "ID Pedido"


# Salvar e fechar o arquivo Excel
writer.close()
