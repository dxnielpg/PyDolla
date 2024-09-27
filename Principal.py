import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# Carregar o arquivo Excel
file_path = 'C:/Users/Jhonathan/PycharmProjects/pythonProject/AtividadeExcel.xlsx'  # Substitua pelo caminho correto

# Carregar a aba "Dados"
df_dados = pd.read_excel(file_path, sheet_name='Dados')

# Limpar os dados, removendo linhas com NaN e redefinindo o índice
df_dados_clean = df_dados[['Unnamed: 0', 'Unnamed: 1']].dropna().reset_index(drop=True)
df_dados_clean.columns = ['MÊS', 'DOLAR']

# Converter a coluna 'MÊS' para string e 'DOLAR' para numérico
df_dados_clean['MÊS'] = pd.to_datetime(df_dados_clean['MÊS'], errors='coerce').dt.strftime('%Y-%m')
df_dados_clean['DOLAR'] = pd.to_numeric(df_dados_clean['DOLAR'], errors='coerce')

# Excluir valores nulos (após a conversão)
df_dados_clean = df_dados_clean.dropna()

# Calcular a média geral e o desvio padrão
mean_value = df_dados_clean['DOLAR'].mean()
std_value = df_dados_clean['DOLAR'].std()

# Exibir a média e o desvio padrão no console
print(f"Média: R$ {mean_value:.2f}")
print(f"Desvio Padrão: R$ {std_value:.2f}")

# Configurar o gráfico de dispersão
plt.figure(figsize=(12, 7))

# Criar gráfico de dispersão
plt.scatter(df_dados_clean['MÊS'], df_dados_clean['DOLAR'], color='cornflowerblue', label='Valores Mensais')

# Adicionar linha de média e desvio padrão
plt.axhline(label=f'Média: R$ {mean_value:.2f}')
plt.axhline(label=f'Média + Desvio Padrão: R$ {mean_value + std_value:.2f}')
plt.axhline(linewidth=1, label=f'Média - Desvio Padrão: R$ {mean_value - std_value:.2f}')

# Título do gráfico
plt.title('DOLAR - Gráfico de Dispersão com Média e Desvio Padrão', fontsize=18, fontweight='bold')

# Rótulos dos eixos
plt.xlabel('Mês', fontsize=14, labelpad=10)
plt.ylabel('Valor em R$', fontsize=14, labelpad=10)

# Rotacionar os rótulos do eixo X e ajustar espaçamento
plt.xticks(rotation=45, ha='right', fontsize=12)
plt.yticks(fontsize=12)

# Ajustar o limite superior do eixo Y
plt.ylim(0, df_dados_clean['DOLAR'].max() + 2)

# Formatação do eixo Y
formatter = ticker.FormatStrFormatter('R$ %.2f')
plt.gca().yaxis.set_major_formatter(formatter)

# Exibir a legenda
plt.legend(loc='upper left', fontsize=12)

# Ajuste automático do layout
plt.tight_layout()

# Mostrar o gráfico
plt.show()

