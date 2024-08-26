import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import os

# Definir o caminho para o arquivo Excel
current_dir = os.path.dirname(os.path.abspath(__file__))
path_to_base_upgrade = os.path.join(current_dir, 'Base_Upgrade.xlsx')
path_to_historivco_BKP = os.path.join(current_dir, 'Histórico BKP.xlsx')
path_to_reajuste = os.path.join(current_dir, 'Tabela_Upgrade_Reajuste.xlsx')

# Carregar a planilha Base_Upgrade.xlsx
df_upgrade = pd.read_excel(path_to_base_upgrade)
df_historicobkp = pd.read_excel(path_to_historivco_BKP, sheet_name='Enviados')
df_reajuste = pd.read_excel(path_to_reajuste)

# Contar o total de contratos
total_contratos = df_upgrade['CONTRATO_NRO'].nunique()

# Contar quantos contratos são Upsell
total_upsell = df_upgrade[df_upgrade['UPSELL'] == 'UPSELL']['CONTRATO_NRO'].nunique()

# Calcular a diferença entre o total de contratos e o total de Upsell
diferenca = total_contratos - total_upsell

#verifica placas reservads
filtro_historicobkp = ['Ajuste tarifa']
placas_reservadas = df_historicobkp[df_historicobkp['Motivo'].isin(filtro_historicobkp)].nunique().shape[0]

filtro_reajuste = ['Sim']
placas_substituidas = df_reajuste[df_reajuste['REALIZOU A SUB?'].isin(filtro_reajuste)].nunique().shape[0]
reajuste_aplicado = df_reajuste[df_reajuste['APLICAÇÃO FEITA?'].isin(filtro_reajuste)].nunique().shape[0]

#-------------------------------------------------------------------------------------
# Preparar os dados para o gráfico
chart_data = pd.DataFrame({
    'Status': ['Total de Contratos', 'Upsell'],
    'Quantidade': [total_contratos, total_upsell]
})

# Criar o gráfico com matplotlib
fig, ax = plt.subplots()
bars = ax.bar(chart_data['Status'], chart_data['Quantidade'], color='orange')

# Adicionar valores acima das barras
for bar in bars:
    height = bar.get_height()
    ax.annotate(f'{height}',
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 5),  # Deslocamento em pixels
                textcoords="offset points",
                ha='center', va='bottom',
                fontsize=12)

# Configurações do gráfico
ax.set_title("Contratos e Upsell")
ax.set_xlabel("Tipo")
ax.set_ylabel("Quantidade")
ax.yaxis.grid(True) #Adiciona grade
ax.set_ylim(0, max(chart_data['Quantidade']) * 1.2 )

# Remover bordas desnecessárias
for spine in ['top', 'right']:
    ax.spines[spine].set_visible(False)

# Exibir o gráfico no Streamlit
st.pyplot(fig)

# Exibir as quantidades no painel
st.write(f"Total de Contratos: {total_contratos}")
st.write(f"Total de Contratos Upsell: {total_upsell}")
st.write(f"Diferença: {diferenca}")

# -------------------------------------------------------------------------------------
# Segundo gráfico: Placas Reservadas, Substituídas e Reajustes Aplicados
chart_data_2 = pd.DataFrame({
    'Status': ['Placas Reservadas', 'Placas Substituídas', 'Reajuste Aplicado'],
    'Quantidade': [int(placas_reservadas), int(placas_substituidas), int(reajuste_aplicado)]
})

# Certifique-se de que todos os valores são inteiros
chart_data_2['Quantidade'] = chart_data_2['Quantidade'].astype(int)

fig2, ax2 = plt.subplots()
bars2 = ax2.bar(chart_data_2['Status'], chart_data_2['Quantidade'], color='blue')

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

# Remover bordas desnecessárias
for spine in ['top', 'right']:
    ax2.spines[spine].set_visible(False)

# Exibir o segundo gráfico no Streamlit
st.pyplot(fig2)

# Exibir as quantidades no painel
st.write(f"Total de Placas Reservadas: {placas_reservadas}")
st.write(f"Total de Placas Substituídas: {placas_substituidas}")
st.write(f"Total de Reajustes Aplicados: {reajuste_aplicado}")
