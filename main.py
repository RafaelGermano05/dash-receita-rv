import pandas as pd
import streamlit as st
import os
import matplotlib.pyplot as plt
import numpy as np
import io


def calcular_rv(tipo_doc, ponderado):
    if tipo_doc == "CPF":
        if 10000 < ponderado < 15000:
            return 360.00
        elif 15000 < ponderado < 30000:
            return 720.00
        elif ponderado >= 30000:
            return 960.00
        else:
            return 0.00
    elif tipo_doc == "CNPJ":
        if 10000 < ponderado < 15000:
            return 450.00
        elif 15000 < ponderado < 30000:
            return 900.00
        elif ponderado >= 30000:
            return 1200.00
        else:
            return 0.00
    return 0.00


st.set_page_config(layout="wide")
st.title("Análise de RV_FORTSUN com Comparação por Mês e Consultor")
st.write("Faça o upload do arquivo Excel para gerar a análise.")

uploaded_file = st.file_uploader("Carregar arquivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    df["DATA_VENDAS"] = pd.to_datetime(df["DATA_VENDAS"])
    df["Ano-Mês"] = df["DATA_VENDAS"].dt.to_period("M").astype(str)

    # Calcular os ponderados
    df["Ponderado M1"] = (df["TPV PARCELADO M1"] * 2) + (df["TPV CRÉDITO M1"] * 1) + (df["TPV DÉBITO M1"] * 0.5)
    df["Ponderado M2"] = (df["TPV PARCELADO M2"] * 2) + (df["TPV CRÉDITO M2"] * 1) + (df["TPV DÉBITO M2"] * 0.5)
    df["Ponderado M1 Projetado"] = (df["Ponderado M1"] / 22) * 30 
    df["Ponderado M2 Projetado"] = (df["Ponderado M2"] / 22) * 30 
    df["Ponderado para Fevereiro"] = ((df["Ponderado M1"] + df["Ponderado M2 Projetado"]) / 2) * 0.95
    df["Ponderado Média"] = ((df["Ponderado M1"] + df["Ponderado M2"]) / 2) * 0.95
    

    # FILTRAR A PARTIR DE OUTUBRO
    df = df[df["DATA_VENDAS"] >= "2023-10-01"]

    # Gráfico 1 - Todos os meses com Ponderado Média
    df["RV_FORTSUN_Media"] = df.apply(lambda row: calcular_rv(row["TIPO_DOC"], row["Ponderado Média"]), axis=1)
    summary_media = df.groupby("Ano-Mês")["RV_FORTSUN_Media"].sum().reset_index()

    # Gráfico 2 - Fevereiro com Ponderado M1 e Projetado M2, e os outros com Ponderado Média e Março com projetado do M1


    def calc_rv_condicional(row):
        if row["Ano-Mês"].endswith("-03"):
            return calcular_rv(row["TIPO_DOC"], row["Ponderado M1 Projetado"])
        elif row["Ano-Mês"].endswith("-02"):
            return calcular_rv(row["TIPO_DOC"], row["Ponderado para Fevereiro"])
        else:
            return calcular_rv(row["TIPO_DOC"], row["Ponderado Média"])

    df["RV_FORTSUN_Condicional"] = df.apply(calc_rv_condicional, axis=1)
    summary_condicional = df.groupby("Ano-Mês")["RV_FORTSUN_Condicional"].sum().reset_index()

    # GRÁFICO 1
    st.subheader("Total de RV_FORTSUN por Mês - Baseado no Ponderado Média")
    fig1, ax1 = plt.subplots(figsize=(12, 5))
    norm1 = plt.Normalize(summary_media["RV_FORTSUN_Media"].min(), summary_media["RV_FORTSUN_Media"].max())
    colors1 = plt.cm.Blues(norm1(summary_media["RV_FORTSUN_Media"]))
    bars1 = ax1.bar(summary_media["Ano-Mês"], summary_media["RV_FORTSUN_Media"], color=colors1)
    for bar in bars1:
        yval = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2, yval, f"{yval:,.2f}", ha='center', va='bottom', fontsize=10, fontweight="bold")
    ax1.set_xlabel("Mês")
    ax1.set_ylabel("Total RV_FORTSUN")
    ax1.set_title("Total de RV_FORTSUN por Mês (Ponderado Média)")
    ax1.tick_params(axis="x", rotation=45)
    st.pyplot(fig1)

    # GRÁFICO 2
    st.subheader("Total de RV_FORTSUN por Mês - Fevereiro com Ponderado M1 e Projetado M2 E Março Ponderado M1")
    fig2, ax2 = plt.subplots(figsize=(12, 5))
    norm2 = plt.Normalize(summary_condicional["RV_FORTSUN_Condicional"].min(), summary_condicional["RV_FORTSUN_Condicional"].max())
    colors2 = plt.cm.Oranges(norm2(summary_condicional["RV_FORTSUN_Condicional"]))
    bars2 = ax2.bar(summary_condicional["Ano-Mês"], summary_condicional["RV_FORTSUN_Condicional"], color=colors2)
    for bar in bars2:
        yval = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2, yval, f"{yval:,.2f}", ha='center', va='bottom', fontsize=10, fontweight="bold")
    ax2.set_xlabel("Mês")
    ax2.set_ylabel("Total RV_FORTSUN")
    ax2.set_title("Total de RV_FORTSUN por Mês (Fevereiro com Ponderado M1 e Proj Pond M2)")
    ax2.tick_params(axis="x", rotation=45)
    st.pyplot(fig2)

    # GRÁFICO DE CONSULTOR (TOP 1 POR MÊS)
    st.subheader("Melhor Consultor por Mês (Baseando Fevereiro em Ponderado M1)")

    df["RV_Consultor"] = df.apply(calc_rv_condicional, axis=1)
    top_consultores = df.groupby(["Ano-Mês", "EXECUTIVO"])["RV_Consultor"].sum().reset_index()
    top1_por_mes = top_consultores.sort_values(["Ano-Mês", "RV_Consultor"], ascending=[True, False]).groupby("Ano-Mês").head(1)

    fig3, ax3 = plt.subplots(figsize=(15, 6))
    bars3 = ax3.bar(top1_por_mes["Ano-Mês"], top1_por_mes["RV_Consultor"], color="green")

    for i, bar in enumerate(bars3):
        yval = bar.get_height()
        consultor = top1_por_mes.iloc[i]["EXECUTIVO"]
        ax3.text(bar.get_x() + bar.get_width()/2, yval + 100, f"{consultor}\nR$ {yval:,.2f}", ha='center', va='bottom', fontsize=10, fontweight="bold")

    ax3.set_xlabel("Mês")
    ax3.set_ylabel("Valor RV_FORTSUN")
    ax3.set_title("Consultor que Mais Gerou Valor por Mês")
    ax3.tick_params(axis="x", rotation=45)
    st.pyplot(fig3)

    # Planilha para download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        label="Baixar Planilha Processada",
        data=buffer.getvalue(),
        file_name="Relatorio_RV_FORTSUN_Analise.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
