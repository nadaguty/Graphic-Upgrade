import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import os
import pdfkit

# Definir o caminho para o arquivo Excel
current_dir = os.path.dirname(os.path.abspath(__file__))
path_to_base_upgrade = os.path.join(current_dir, 'Base_Upgrade.xlsx')
path_to_historivco_BKP = os.path.join(current_dir, 'Histórico BKP.xlsx')
path_to_reajuste = os.path.join(current_dir, 'Tabela_Upgrade_Reajuste.xlsx')  # Corrigido para .xlsx

# Carregar as planilhas
df_upgrade = pd.read_excel(path_to_base_upgrade)
df_historicobkp = pd.read_excel(path_to_historivco_BKP)
df_reajuste = pd.read_excel(path_to_reajuste)

# Contar o total de contratos
total_contratos = df_upgrade['CONTRATO_NRO'].nunique()

# Contar quantos contratos são Upsell
total_upsell = df_upgrade[df_upgrade['UPSELL'] == 'UPSELL']['CONTRATO_NRO'].nunique()

# Calcular a diferença entre o total de contratos e o total de Upsell
diferenca = total_contratos - total_upsell

# Verifica placas reservadas
filtro_historicobkp = ['Ajuste tarifa']
placas_reservadas = df_historicobkp[df_historicobkp['Motivo'].isin(filtro_historicobkp)].shape[0]  # Contagem total

filtro_reajuste = ['Sim']
placas_substituidas = df_reajuste[df_reajuste['REALIZOU A SUB?'].isin(filtro_reajuste)].shape[0]  # Contagem total
reajuste_aplicado = df_reajuste[df_reajuste['APLICAÇÃO FEITA?'].isin(filtro_reajuste)].shape[0]  # Contagem total
contratos_encerrados = df_reajuste[df_reajuste['CONTRATO ENC?'].isin(filtro_reajuste)].shape[0]

# 1. Contar o total de linhas na coluna B-BX onde o valor é 'B-BX'
total_bbx = df_upgrade[df_upgrade['B-BX'] == 'B-BX'].shape[0]
# 2. Filtrar os números de contrato na planilha df_upgrade onde o valor na coluna B-BX é 'B-BX'
contratos_bbx = df_upgrade[df_upgrade['B-BX'] == 'B-BX']['CONTRATO_NRO'].unique()
# 3. Verificar quantos desses contratos estão presentes na df_historicobkp com motivo 'Ajuste tarifa'
contratos_ajuste_tarifa = df_historicobkp[df_historicobkp['Motivo'] == 'Ajuste tarifa']['Contrato'].unique()
contratos_bbx_ajuste = [contrato for contrato in contratos_bbx if contrato in contratos_ajuste_tarifa]
# Contar o total de contratos que satisfazem ambas as condições
total_ajuste_tarifa = len(contratos_bbx_ajuste)

# -------------------------------------------------------------------------------------
# Primeiro gráfico: Contratos e Upsell
chart_data = pd.DataFrame({
    'Status': ['Total de Contratos', 'Upsell'],
    'Quantidade': [total_contratos, total_upsell]
})

fig, ax = plt.subplots()
bars = ax.bar(chart_data['Status'], chart_data['Quantidade'], color='orange')

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
ax.set_xlabel("Tipo")
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
st.write(f"Diferença: {diferenca}")


# -------------------------------------------------------------------------------------
# Segundo gráfico: Placas Reservadas, Substituídas e Reajustes Aplicados
chart_data_2 = pd.DataFrame({
    'Status': ['Placas Reservadas', 'Placas Substituídas', 'Reajuste Aplicado', 'Contrato Encerrado'],
    'Quantidade': [int(placas_reservadas), int(placas_substituidas), int(reajuste_aplicado), int(contratos_encerrados)]
})

# Certifique-se de que todos os valores são inteiros
chart_data_2['Quantidade'] = chart_data_2['Quantidade'].astype(int)

fig2, ax2 = plt.subplots()
bars2 = ax2.bar(chart_data_2['Status'], chart_data_2['Quantidade'], color='orange')

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
ax2.set_xticklabels(chart_data_2['Status'], fontsize=7)  # Reduzindo a fonte para 10

# Remover bordas desnecessárias
for spine in ['top', 'right']:
    ax2.spines[spine].set_visible(False)

# Exibir o segundo gráfico no Streamlit
st.pyplot(fig2)

# Exibir as quantidades no painel
st.write(f"Total de Placas Reservadas: {placas_reservadas}")
st.write(f"Total de Placas Substituídas: {placas_substituidas}")
st.write(f"Total de Reajustes Aplicados: {reajuste_aplicado}")
st.write(f"Total de Reajustes Aplicados: {contratos_encerrados}")

#-------------------------------------------------------------------------
# 4. Preparar os dados para o gráfico
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
ax3.set_xlabel("Tipo")
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


# Gerar um botão para baixar o PDF
if st.button("Gerar PDF"):
    # Salvar a página atual em HTML
    path_html = "streamlit_page.html"
    path_pdf = "streamlit_page.pdf"

    # Capturar o conteúdo da página (usando a opção de salvar o HTML)
    with open(path_html, "w") as f:
        f.write(st.get_report_ctx().get_current().output)
    
    # Converter o HTML para PDF
    pdfkit.from_file(path_html, path_pdf)

    st.success("PDF gerado com sucesso!")
    st.markdown(f"[Baixar PDF]({path_pdf})")
