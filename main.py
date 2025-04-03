import pandas as pd
import streamlit as st
import os
import matplotlib.pyplot as plt
import numpy as np
import io


def calcular_rv(tipo_doc, ponderado_media):
    if tipo_doc == "CPF":
        if 10000 < ponderado_media < 15000:
            return 360.00
        elif 15000 < ponderado_media < 30000:
            return 720.00
        elif ponderado_media >= 30000:
            return 960.00
        else:
            return 0.00
    elif tipo_doc == "CNPJ":
        if 10000 < ponderado_media < 15000:
            return 450.00
        elif 15000 < ponderado_media < 30000:
            return 900.00
        elif ponderado_media >= 30000:
            return 1200.00
        else:
            return 0.00
    return 0.00


st.title("Processamento de Planilha e Análise de RV_FORTSUN")
st.write("Faça o upload do arquivo Excel para gerar a análise.")


uploaded_file = st.file_uploader("Carregar arquivo Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    # Calcular Ponderado M1
    df["Ponderado M1"] = (df["TPV PARCELADO M1"] * 2) + (df["TPV CRÉDITO M1"] * 1) + (df["TPV DÉBITO M1"] * 0.5)

    # Calcular Ponderado M2
    df["Ponderado M2"] = (df["TPV PARCELADO M2"] * 2) + (df["TPV CRÉDITO M2"] * 1) + (df["TPV DÉBITO M2"] * 0.5)

    # Calcular Ponderado Média
    df["Ponderado Média"] = (df["Ponderado M1"] + df["Ponderado M2"]) / 2


    df["RV_FORTSUN"] = df.apply(lambda row: calcular_rv(row["TIPO_DOC"], row["Ponderado Média"]), axis=1)

    # Converter 'DATA_VENDA' para formato de data
    df["DATA_VENDAS"] = pd.to_datetime(df["DATA_VENDAS"])


    df["Ano-Mês"] = df["DATA_VENDAS"].dt.to_period("M")


    df_summary = df.groupby("Ano-Mês")["RV_FORTSUN"].sum().reset_index()


    norm = plt.Normalize(df_summary["RV_FORTSUN"].min(), df_summary["RV_FORTSUN"].max())
    colors = plt.cm.Blues(norm(df_summary["RV_FORTSUN"]))  # Usa a paleta de cores "Blues"

    # Criar o gráfico de barras com gradiente e valores fixos acima das barras
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(df_summary["Ano-Mês"].astype(str), df_summary["RV_FORTSUN"], color=colors)


    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval, f"{yval:,.2f}", ha='center', va='bottom', fontsize=10, fontweight="bold")

    ax.set_xlabel("Mês")
    ax.set_ylabel("Total de RV_FORTSUN")
    ax.set_title("Total de RV_FORTSUN por Mês (Com Gradiente)")
    ax.tick_params(axis="x", rotation=45)


    st.pyplot(fig)


    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    # Botão para baixar o arquivo corrigido
    st.download_button(
        label="Baixar Planilha Processada",
        data=buffer.getvalue(),
        file_name="New_Report_Transformed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
