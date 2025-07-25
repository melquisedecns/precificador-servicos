import streamlit as st
import pandas as pd
import datetime
import os
from gerador_proposta_app import gerar_proposta

st.set_page_config(page_title="Precificador de Serviços", layout="wide")
st.title("📊 Precificador de Serviços Técnicos")
st.markdown("---")

# ====== 1. Premissas ======
st.header("1. Premissas")
col1, col2, col3 = st.columns(3)
with col1:
    cliente = st.text_input("Cliente")
    cidade_mobilizacao = st.text_input("Cidade de Mobilização")
with col2:
    estado = st.selectbox("Estado", ["MT", "SP", "RJ", "GO"])
    cidade = st.text_input("Cidade do Serviço")
with col3:
    servico = st.text_input("Serviço Prestado")
    data_inicio = st.date_input("Data de Início", value=datetime.date.today())

# ====== 2. Itinerário ======
st.header("2. Itinerários")

itinerarios = []
for i in range(1, 4):
    st.subheader(f"Itinerário {i}")
    c1, c2, c3 = st.columns(3)
    with c1:
        partida = st.text_input(f"Cidade de Partida {i}")
    with c2:
        chegada = st.text_input(f"Cidade de Chegada {i}")
    with c3:
        distancia = st.number_input(f"Distância Km {i}", min_value=0.0, step=0.1)
    if partida and chegada and distancia:
        itinerarios.append({"partida": partida, "chegada": chegada, "km": distancia})

# ====== 3. Técnicos ======
st.header("3. Técnicos")
usa_base = st.selectbox("Usar base de dados Equipe?", ["SIM", "NÃO"])
qtd_tecnicos = st.number_input("Quantidade de Técnicos", min_value=1, max_value=50, step=1)
salario_base = st.number_input("Salário Base R$", min_value=1000.0, step=100.0)

dados_tecnicos = []
if usa_base == "SIM":
    bd_arquivo = st.file_uploader("Importar Arquivo Base de Dados Equipe (.xlsx)", type=["xlsx"])
    if bd_arquivo:
        df_bd = pd.read_excel(bd_arquivo)
        st.dataframe(df_bd.head(10))
        dados_tecnicos = df_bd.to_dict(orient="records")
else:
    for i in range(qtd_tecnicos):
        nome = st.text_input(f"Nome Técnico {i+1}")
        if nome:
            dados_tecnicos.append({"nome": nome, "salario": salario_base})

# ====== 4. Frota ======
st.header("4. Frota")
frota_file = st.file_uploader("Importar Tabela Frota (.xlsx)", type="xlsx")
if frota_file:
    df_frota = pd.read_excel(frota_file)
    st.dataframe(df_frota.head())

# ====== 5. Custos ======
st.header("5. Custos Gerais")
alimentacao_sim = st.selectbox("Incluir Alimentação?", ["SIM", "NÃO"])
hospedagem_sim = st.selectbox("Incluir Hospedagem?", ["SIM", "NÃO"])
custo_miscelanea = st.number_input("Valor Total de Miscelâneas R$", min_value=0.0, step=10.0)
custo_locacoes = st.number_input("Valor Total de Locações R$", min_value=0.0, step=10.0)

# ====== 6. Proposta ======
st.header("6. Geração de Proposta Técnica")
logo = st.file_uploader("Importar logo (.png)", type="png")
proposta_valor = st.number_input("Valor da Proposta R$", min_value=0.0, step=100.0)

if st.button("Gerar Proposta"):
    dados = {
        "local": cidade,
        "data": data_inicio.strftime("%d/%m/%Y"),
        "codigo": "PROP-001",
        "cliente": cliente,
        "titulo": servico,
        "responsavel": "Responsável Cliente",
        "email": "email@cliente.com",
        "telefone": "(00) 00000-0000",
        "introducao": f"Apresentamos a seguir a proposta para o serviço de {servico}.",
        "representante_1": {"nome": "Melquisedec N. Soares", "email": "planejamento.msn@outlook.com", "telefone": "65993543464"},
        "representante_2": {"nome": "", "email": "", "telefone": ""},
        "resumo": "Serviço técnico conforme premissas estabelecidas.",
        "escopo": [servico],
        "materiais": ["Inclusos conforme planilha"],
        "observacoes": "Valores estimados, sujeitos à confirmação.",
        "investimento": {"Serviço": proposta_valor},
        "garantia": "90 dias após execução",
        "resp_contratada": ["Execução técnica"],
        "resp_contratante": ["Liberação de área"],
        "condicoes": {"PAGAMENTO": "30 dias", "VALIDADE": "15 dias"},
        "assinatura": "Planejamento - MSN"
    }
    path = gerar_proposta(dados)
    st.success("✅ Proposta gerada com sucesso!")
    with open(path, "rb") as f:
        st.download_button("📥 Baixar Proposta (Word)", f, file_name=os.path.basename(path))
