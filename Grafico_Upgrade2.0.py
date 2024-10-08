import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import streamlit as st
import requests
import io
import os

# Carregar as URLs das variáveis de ambiente
url_base_upgrade = os.getenv('URL_BASE_UPGRADE')
url_historicobkp = os.getenv('URL_HISTORICO_BKP')
url_reajuste = os.getenv('URL_REAJUSTE')

# Depuração: Verificar se as variáveis de ambiente estão carregadas
# Não precisa ficar visivel no comando
#st.write("URL Base Upgrade:", url_base_upgrade)
#st.write("URL Histórico BKP:", url_historicobkp)
#st.write("URL Reajuste:", url_reajuste)

# Verificar se as variáveis foram carregadas corretamente
if not all([url_base_upgrade, url_historicobkp, url_reajuste]):
    raise ValueError("Uma ou mais variáveis de ambiente não foram carregadas corretamente.")

# Função para carregar um arquivo Excel de uma URL
def load_excel_from_url(url):
    response = requests.get(url)
    response.raise_for_status()  # Levanta um erro se a requisição falhar
    return pd.read_excel(io.BytesIO(response.content), engine='openpyxl')

# Carregar as planilhas usando as URLs das variáveis de ambiente
df_upgrade = load_excel_from_url(url_base_upgrade)
df_historicobkp = load_excel_from_url(url_historicobkp)
df_reajuste = load_excel_from_url(url_reajuste)


# Contar o total de contratos
total_contratos = df_upgrade['CONTRATO_NRO'].nunique()

# Contar quantos contratos são Upsell
total_upsell = df_upgrade[df_upgrade['UPSELL'] == 'UPSELL']['CONTRATO_NRO'].nunique()

#contar quantos contratos são Moover
total_moover = df_upgrade[df_upgrade['SEGMENTO'] == 'MOOVER']['CONTRATO_NRO'].nunique()

# Calcular a diferença entre o total de contratos e o total de Upsell
diferenca = total_contratos - total_upsell - total_moover

# Verifica placas reservadas
filtro_historicobkp = ['Ajuste tarifa']
placas_reservadas = df_historicobkp[df_historicobkp['Motivo'].isin(filtro_historicobkp)].shape[0]  # Contagem total

filtro_reajuste = ['Sim']
placas_substituidas =   df_reajuste[df_reajuste['REALIZOU A SUB?'].isin(filtro_reajuste)].shape[0]  # Contagem total
reajuste_aplicado =     df_reajuste[df_reajuste['APLICAÇÃO FEITA?'].isin(filtro_reajuste)].shape[0]  # Contagem total
contratos_encerrados =  df_reajuste[df_reajuste['CONTRATO ENC?'].isin(filtro_reajuste)].shape[0]
emails_enviados =       df_reajuste[df_reajuste['Email enviado?'].isin(filtro_reajuste)].shape[0]

# 1. Contar o total de linhas na coluna B-BX onde o valor é 'B-BX'
total_bbx = df_upgrade[df_upgrade['B-BX'] == 'B-BX'].shape[0]
# 2. Filtrar os números de contrato na planilha df_upgrade onde o valor na coluna B-BX é 'B-BX'
contratos_bbx = df_upgrade[df_upgrade['B-BX'] == 'B-BX']['CONTRATO_NRO'].unique()
# 3. Verificar quantos desses contratos estão presentes na df_historicobkp com motivo 'Ajuste tarifa'
contratos_ajuste_tarifa = df_reajuste[df_reajuste['REALIZOU A SUB?'] == 'Sim']['CONTRATO'].unique()
contratos_bbx_ajuste = [contrato for contrato in contratos_bbx if contrato in contratos_ajuste_tarifa]
# Contar o total de contratos que satisfazem ambas as condições
total_ajuste_tarifa = len(contratos_bbx_ajuste)

#Grafico 4 valores dos reajustes aplicados
# Filtrar a planilha para 'APLICAÇÃO FEITA?' == 'Sim'
df_filtrado = df_reajuste[df_reajuste['APLICAÇÃO FEITA?'] == 'Sim']

#Comando teste
# Remover vírgulas e pontos e converter as colunas para float (tratando formato brasileiro)
#df_filtrado['VALOR_REAJUSTE_GRUPO_CONTRATO'] = df_filtrado['VALOR_REAJUSTE_GRUPO_CONTRATO'].replace('.', '', regex=True).replace(',', '.', regex=True).astype(float)
#df_filtrado['VALOR_REAJUSTE_GRUPO_RESERVA'] = df_filtrado['VALOR_REAJUSTE_GRUPO_RESERVA'].replace('.', '', regex=True).replace(',', '.', regex=True).astype(float)
# Remover vírgulas e converter as colunas para float
df_filtrado['VALOR_REAJUSTE_GRUPO_CONTRATO'] = df_filtrado['VALOR_REAJUSTE_GRUPO_CONTRATO'].replace(',', '', regex=True).astype(float)
df_filtrado['VALOR_REAJUSTE_GRUPO_RESERVA'] = df_filtrado['VALOR_REAJUSTE_GRUPO_RESERVA'].replace(',', '', regex=True).astype(float)

# Inicializar somas
soma_com_upgrade = 0
soma_sem_upgrade = 0

# Processar as linhas filtradas
for _, row in df_filtrado.iterrows():
    if row['CONTRATO ATIVO'] == 'COM UPGRADE':
        soma_com_upgrade += row['VALOR_REAJUSTE_GRUPO_CONTRATO']
    elif row['CONTRATO ATIVO'] == 'SEM UPGRADE' and row['FLAG_REAJUSTE_GRUPO_RESERVA'] == 'NECESSÁRIO REAJUSTE':
        soma_sem_upgrade += row['VALOR_REAJUSTE_GRUPO_RESERVA']

# -------------------------------------------------------------------------------------
# 1 Primeiro gráfico: Contratos e Upsell
chart_data = pd.DataFrame({
    'Status': ['Total de Contratos', 'Upsell', 'Moover'],
    'Quantidade': [total_contratos, total_upsell, total_moover]
})

fig, ax = plt.subplots()
bars = ax.bar(chart_data['Status'], chart_data['Quantidade'], color=['orange', 'orange', '#7f3e98'])

# Adicionar valores acima das barras
for bar in bars:
    height = bar.get_height()
    ax.annotate(f'{height}',
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 5),
                textcoords="offset points",
                ha='center', va='bottom',
                fontsize=12)

# Configurações do gráfico
ax.set_title("Contratos e Upsell")
ax.set_xlabel("Status")
ax.set_ylabel("Quantidade")
ax.yaxis.grid(True)
ax.set_ylim(0, max(chart_data['Quantidade']) * 1.2)

# Remover bordas desnecessárias
for spine in ['top', 'right']:
    ax.spines[spine].set_visible(False)

# Exibir o gráfico no Streamlit
st.pyplot(fig)

# Exibir as quantidades no painel
st.write(f"Total de Contratos: {total_contratos}")
st.write(f"Total de Contratos Upsell: {total_upsell}")
st.write(f"Total de Contratos Moover {total_moover}")
st.write(f"Diferença: {diferenca}")


# -------------------------------------------------------------------------------------
# 2 Segundo gráfico: Placas Reservadas, Substituídas e Reajustes Aplicados
chart_data_2 = pd.DataFrame({
    'Status': ['Emails Enviados','Placas Reservadas', 'Placas Substituídas', 'Reajuste Aplicado', 'Contrato Encerrado'],
    'Quantidade': [int(emails_enviados), int(placas_reservadas), int(placas_substituidas), int(reajuste_aplicado), int(contratos_encerrados)]
})

# Certifique-se de que todos os valores são inteiros
chart_data_2['Quantidade'] = chart_data_2['Quantidade'].astype(int)

fig2, ax2 = plt.subplots()
bars2 = ax2.bar(chart_data_2['Status'], chart_data_2['Quantidade'], color='#ff7f0e')

# Adicionar valores acima das barras
for bar in bars2:
    height = bar.get_height()
    ax2.annotate(f'{height}',
                 xy=(bar.get_x() + bar.get_width() / 2, height),
                 xytext=(0, 5),
                 textcoords="offset points",
                 ha='center', va='bottom',
                 fontsize=12)

# Configurações do gráfico
ax2.set_title("Placas e Reajustes")
ax2.set_xlabel("Status")
ax2.set_ylabel("Quantidade")
ax2.yaxis.grid(True)
ax2.set_ylim(0, max(chart_data_2['Quantidade']) * 1.2)

# Ajustar o tamanho da fonte das legendas do eixo X
ax2.set_xticklabels(chart_data_2['Status'], fontsize=6)  # Reduzindo a fonte para 10

# Remover bordas desnecessárias
for spine in ['top', 'right']:
    ax2.spines[spine].set_visible(False)

# Exibir o segundo gráfico no Streamlit
st.pyplot(fig2)

# Exibir as quantidades no painel
st.write(f"Total de Emails Enviados: {emails_enviados}")
st.write(f"Total de Placas Reservadas: {placas_reservadas}")
st.write(f"Total de Placas Substituídas: {placas_substituidas}")
st.write(f"Total de Reajustes Aplicados: {reajuste_aplicado}")
st.write(f"Total de Contratos Encerrados: {contratos_encerrados}")

#-------------------------------------------------------------------------
# 3 Terceiro grafico: B para BX
chart_data_3 = pd.DataFrame({
    'Status': ['Total B-BX', 'Substituidos'],
    'Quantidade': [total_bbx, total_ajuste_tarifa]
})

# Criar o gráfico com matplotlib
fig3, ax3 = plt.subplots()
bars3 = ax3.bar(chart_data_3['Status'], chart_data_3['Quantidade'], color='orange')

# Adicionar valores acima das barras
for bar in bars3:
    height = bar.get_height()
    ax3.annotate(f'{height}',
                 xy=(bar.get_x() + bar.get_width() / 2, height),
                 xytext=(0, 5),
                 textcoords="offset points",
                 ha='center', va='bottom',
                 fontsize=12)

# Configurações do gráfico
ax3.set_title("Total de B para BX")
ax3.set_xlabel("Status")
ax3.set_ylabel("Quantidade")
ax3.yaxis.grid(True)
ax3.set_ylim(0, max(chart_data_3['Quantidade']) * 1.2)

# Remover bordas desnecessárias
for spine in ['top', 'right']:
    ax3.spines[spine].set_visible(False)

# Exibir o gráfico no Streamlit
st.pyplot(fig3)

# Exibir as quantidades no painel
st.write(f"Total de B-BX: {total_bbx}")
st.write(f"Substituidos: {total_ajuste_tarifa}")
st.write("-----------------------------------------------------------")

#---------------------------------------------------------------------
# 4 Quarto Grafico: Valores dos reajustes
chart_data_4 = pd.DataFrame({
    'Status': ['COM UPGRADE', 'SEM UPGRADE'],
    'Valores (R$)': [soma_com_upgrade, soma_sem_upgrade]
})


# Criar o gráfico
fig4, ax4 = plt.subplots()
bars4 = ax4.bar(chart_data_4['Status'], chart_data_4['Valores (R$)'], color='#ff7f0e')

# Adicionar os valores acima das barras com formatação adequada
for bar in bars4:
    height = bar.get_height()
    formatted_value = f'{height:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    ax4.annotate(f'R${formatted_value}',
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 5),
                textcoords="offset points",
                ha='center', va='bottom',
                fontsize=12)

# Configurações do gráfico
ax4.set_title("Valores de Reajuste por Tipo de Contrato")
ax4.set_xlabel("Tipo de Contrato")
ax4.set_ylabel("Valor Total (R$)")
ax4.yaxis.grid(True)
ax4.set_ylim(0, max(chart_data_4['Valores (R$)']) * 1.2)

# Remover bordas desnecessárias
for spine in ['top', 'right']:
    ax4.spines[spine].set_visible(False)

# Exibir o gráfico no Streamlit
st.pyplot(fig4)

# Formatar e exibir as somas para referência com formatação individual
formatted_soma_com_upgrade = f'{soma_com_upgrade:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
formatted_soma_sem_upgrade = f'{soma_sem_upgrade:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

st.write(f"Soma total 'COM UPGRADE': R${formatted_soma_com_upgrade}")
st.write(f"Soma total 'SEM UPGRADE': R${formatted_soma_sem_upgrade}")
