#%% Importação de bibliotecas
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os

# Definir o estilo visual
sns.set(style="whitegrid")

# Caminho da área de trabalho
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

#%% Passo 2: Leitura dos Dados

# Leitura dos dados
df = pd.read_excel('C:/Users/diogo/Desktop/leitura-escrita/data/vendas_loja_online.xlsx')

# Exibir as primeiras linhas dos dados
print(df.head())

# Verificar se há valores ausentes
missing_data = df.isnull().sum()
print(f'Dados faltantes por coluna:\n{missing_data}')

# Tratar valores ausentes - Preenchendo com a média ou mediana
df['Quantidade Vendida'] = df['Quantidade Vendida'].fillna(df['Quantidade Vendida'].mean())
df['Lucro Total (R$)'] = df['Lucro Total (R$)'].fillna(df['Lucro Total (R$)'].median())

# Conversão da coluna 'Data' para o formato datetime
df['Data'] = pd.to_datetime(df['Data'])

#%% Passo 3: Análise Descritiva

# Calcular as vendas mensais por produto
vendas_mensais = df.groupby(['Data', 'Produto'])['Quantidade Vendida'].sum().reset_index()

# Estatísticas descritivas das vendas mensais por produto
estatisticas = vendas_mensais.groupby('Produto')['Quantidade Vendida'].describe()

# Produto com maior e menor volume de vendas
produto_maior_venda = vendas_mensais.groupby('Produto')['Quantidade Vendida'].sum().idxmax()
produto_menor_venda = vendas_mensais.groupby('Produto')['Quantidade Vendida'].sum().idxmin()

# Distribuição do lucro por produto e região
distribuicao_lucro_produto = df.groupby('Produto')['Lucro Total (R$)'].sum()
distribuicao_lucro_regiao = df.groupby('Região de Vendas')['Lucro Total (R$)'].sum()

# Exibir resultados da análise
print(f'Distribuição de Lucro por Produto:\n{distribuicao_lucro_produto}')
print(f'Distribuição de Lucro por Região:\n{distribuicao_lucro_regiao}')

#%% Passo 4: Tendências e Padrões Temporais

# Agrupar as vendas por mês
df['Mês'] = df['Data'].dt.to_period('M')
vendas_mensais = df.groupby('Mês')['Quantidade Vendida'].sum()

# Gráfico de linha para a evolução das vendas ao longo do tempo
plt.figure(figsize=(12, 6))
plt.plot(vendas_mensais.index.astype(str), vendas_mensais.values, marker='o', linestyle='-', color='b', label='Vendas')
plt.title('Evolução das Vendas ao Longo do Tempo', fontsize=14)
plt.xlabel('Mês', fontsize=12)
plt.ylabel('Quantidade Vendida', fontsize=12)
plt.xticks(rotation=45)
plt.grid(True)
plt.tight_layout()
plt.legend()

# Salvar gráfico na área de trabalho
plt.savefig(os.path.join(desktop_path, 'evolucao_vendas.png'))
plt.show()  # Exibe o gráfico

# Comparar o crescimento das vendas em relação ao mesmo mês do ano anterior
vendas_mensais_ano_anterior = vendas_mensais.shift(12)  # Desloca os dados para comparar com o ano anterior
crescimento_ano_anterior = ((vendas_mensais - vendas_mensais_ano_anterior) / vendas_mensais_ano_anterior) * 100

# Exibir crescimento mensal
print(crescimento_ano_anterior)

# Calcular a diferença percentual de vendas mês a mês
dif_perc_vendas = vendas_mensais.pct_change() * 100
print(dif_perc_vendas)

#%% Passo 5: Análise de Rentabilidade

# Calcular o lucro total de cada produto (Preço Unitário (R$) - Custo Total (R$)) * Quantidade Vendida
df['Lucro Total'] = (df['Preço Unitário (R$)'] - df['Custo Total (R$)']) * df['Quantidade Vendida']

# Produtos mais rentáveis (com maior lucro total)
lucro_total_produto = df.groupby('Produto')['Lucro Total'].sum()
produto_mais_rentavel = lucro_total_produto.idxmax()

# Exibir resultados
print(f'Produto mais rentável: {produto_mais_rentavel}')
print(f'Lucro total por produto:\n{lucro_total_produto}')

#%% Passo 6: Análise de Performance por Canal de Vendas

# Calcular a contribuição de cada canal de vendas para o total de vendas e lucro
contribuicao_canal_vendas = df.groupby('Canal de Vendas').agg(
    Total_Vendas=('Quantidade Vendida', 'sum'),
    Total_Lucro=('Lucro Total', 'sum')  # Usar 'Lucro Total' após o cálculo ou 'Lucro Total (R$)' se for uma coluna existente
)

# Calcular a quantidade de vendas por canal
quantidade_vendas_canal = df.groupby('Canal de Vendas')['Quantidade Vendida'].sum().reset_index()

# Exibir resultados
print(f'Contribuição dos Canais de Vendas:\n{contribuicao_canal_vendas}')
print(f'Quantidade de Vendas por Canal de Vendas:\n{quantidade_vendas_canal}')

# Gráfico de barras - Quantidade de Vendas por Canal de Vendas
plt.figure(figsize=(14, 7))
sns.barplot(x='Canal de Vendas', y='Quantidade Vendida', data=quantidade_vendas_canal, palette='coolwarm')
plt.title('Quantidade de Vendas por Canal de Vendas', fontsize=14)
plt.xlabel('Canal de Vendas', fontsize=12)
plt.ylabel('Quantidade Vendida', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()

# Adicionar os valores nas barras
for i, value in enumerate(quantidade_vendas_canal['Quantidade Vendida']):
    plt.text(i, value + 5, f'{value}', ha='center', fontsize=10)

# Salvar gráfico na área de trabalho
plt.savefig(os.path.join(desktop_path, 'quantidade_vendas_canal.png'))
plt.show()  # Exibe o gráfico

# Gráfico de pizza - Quantidade de Vendas por Canal de Vendas
plt.figure(figsize=(8, 8))
plt.pie(quantidade_vendas_canal['Quantidade Vendida'], 
        labels=quantidade_vendas_canal['Canal de Vendas'], 
        autopct='%1.1f%%', 
        startangle=140, 
        colors=sns.color_palette('Set3', len(quantidade_vendas_canal)), 
        wedgeprops={'edgecolor': 'black'})

plt.title('Distribuição da Quantidade de Vendas por Canal de Vendas', fontsize=14)
plt.axis('equal')  # Garante que o gráfico de pizza será circular
plt.tight_layout()

# Salvar gráfico na área de trabalho
plt.savefig(os.path.join(desktop_path, 'distribuicao_vendas_canal.png'))
plt.show()  # Exibe o gráfico

#%% Passo 7: Melhorias nos Gráficos

# Gráfico de dispersão - Relação entre Preço Unitário e Lucro
plt.figure(figsize=(12, 6))
sns.scatterplot(data=df, x='Preço Unitário (R$)', y='Lucro Total', hue='Produto', palette='Set1', s=100, edgecolor='black')
plt.title('Relação entre Preço Unitário e Lucro', fontsize=14)
plt.xlabel('Preço Unitário (R$)', fontsize=12)
plt.ylabel('Lucro Total (R$)', fontsize=12)
plt.legend(title='Produto', loc='upper left', bbox_to_anchor=(1, 1))
plt.tight_layout()

# Salvar gráfico na área de trabalho
plt.savefig(os.path.join(desktop_path, 'relacao_preco_lucro.png'))
plt.show()  # Exibe o gráfico

# Gráfico de barras - Distribuição do Lucro por Produto
plt.figure(figsize=(14, 7))
sns.barplot(x=distribuicao_lucro_produto.index, y=distribuicao_lucro_produto.values, palette='viridis')
plt.title('Distribuição do Lucro por Produto', fontsize=14)
plt.xlabel('Produto', fontsize=12)
plt.ylabel('Lucro Total (R$)', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()

# Adicionar os valores nas barras
for i, value in enumerate(distribuicao_lucro_produto.values):
    plt.text(i, value + 5, f'R${value:,.2f}', ha='center', fontsize=10)

# Salvar gráfico na área de trabalho
plt.savefig(os.path.join(desktop_path, 'distribuicao_lucro_produto.png'))
plt.show()  # Exibe o gráfico

# Gráfico de barras - Distribuição do Lucro por Região
plt.figure(figsize=(14, 7))
sns.barplot(x=distribuicao_lucro_regiao.index, y=distribuicao_lucro_regiao.values, palette='coolwarm')
plt.title('Distribuição do Lucro por Região', fontsize=14)
plt.xlabel('Região', fontsize=12)
plt.ylabel('Lucro Total (R$)', fontsize=12)
plt.xticks(rotation=45)
plt.tight_layout()

# Adicionar os valores nas barras
for i, value in enumerate(distribuicao_lucro_regiao.values):
    plt.text(i, value + 5, f'R${value:,.2f}', ha='center', fontsize=10)

# Salvar gráfico na área de trabalho
plt.savefig(os.path.join(desktop_path, 'distribuicao_lucro_regiao.png'))
plt.show()  # Exibe o gráfico

#%% Passo 8: Visualizar a diferença entre o preço e o custo para cada produto
df['Lucro Unitário'] = df['Preço Unitário (R$)'] - df['Custo Total (R$)']
sns.histplot(df['Lucro Unitário'], kde=True)
plt.title('Distribuição do Lucro Unitário por Produto')
plt.xlabel('Lucro Unitário (R$)')
plt.ylabel('Frequência')

# Salvar gráfico na área de trabalho
plt.savefig(os.path.join(desktop_path, 'distribuicao_lucro_unitario.png'))
plt.show()  # Exibe o gráfico

# Exibir os 10 maiores lucros negativos
print(df[df['Lucro Unitário'] < 0].sort_values('Lucro Unitário').head(10))

#%% Passo 9: Exportação dos Resultados 

# Definir o arquivo de exportação
file_path = 'resultados_analise_vendas.xlsx'

# Exportar as tabelas no Excel com formatação condicional
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    vendas_mensais.to_excel(writer, sheet_name='Vendas Mensais')
    lucro_total_produto.to_excel(writer, sheet_name='Lucro por Produto')
    distribuicao_lucro_produto.to_excel(writer, sheet_name='Distribuição Lucro Produto') 
    distribuicao_lucro_regiao.to_excel(writer, sheet_name='Distribuição Lucro Região')
    contribuicao_canal_vendas.to_excel(writer, sheet_name='Contribuição Canal de Vendas')

print(f'Análise exportada para {file_path}')

#%% Passo 10: Análise de Correlação

# Filtrar apenas as colunas numéricas
df_numerico = df.select_dtypes(include=[np.number])

# Calcular a matriz de correlação
correlacao = df_numerico.corr()

# Exibir a matriz de correlação
print(f'Matriz de Correlação:\n{correlacao}')

# Gráfico de correlação
plt.figure(figsize=(10, 8))
sns.heatmap(correlacao, annot=True, cmap='coolwarm', fmt='.2f', linewidths=0.5)
plt.title('Matriz de Correlação', fontsize=14)
plt.tight_layout()

# Salvar gráfico da matriz de correlação na área de trabalho
plt.savefig(os.path.join(desktop_path, 'matriz_correlacao.png'))
plt.show()  # Exibe o gráfico
