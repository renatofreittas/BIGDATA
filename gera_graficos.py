import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
from pathlib import Path

import plotly.express as px
import plotly.graph_objects as go

from dash import Dash, html, dcc, Input, Output
import dash_bootstrap_components as dbc

INPUT_XLSX = r"dados.xlsx"

def carregar_planilha(path):
    if not Path(path).exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {path}")
    df = pd.read_excel(path)
    return df

def normalizar_colunas(df):
    m = {
        'id': ['id','customer_id','customerid'],
        'nome': ['nome','name','cliente','customer_name'],
        'idade': ['idade','age'],
        'sexo': ['sexo','gender','genero','sex'],
        'marca': ['marca','car_make','make','brand'],
        'modelo': ['modelo','car_model','model'],
        'combustivel': ['combustivel','fuel_type','fuel','combustível'],
        'preco': ['preco','purchase_price','price','valor','preço'],
        'data_compra': ['data_compra','purchase_date','date','data de compra'],
        'km': ['km','mileage','quilometragem'],
        'servicos': ['servicos','service_count','serviços'],
        'ultima_manutencao': ['ultima_manutencao','last_service_date','last_service','ultimo_servico'],
        'cidade': ['cidade','city','location'],
        'financiado': ['financiado','finance','is_financed','financiamento'],
        'parcela_mensal': ['parcela_mensal','monthly_payment','parcela','valor_parcela']
    }
    df2 = df.copy()

    def achar(df, lista):
        for cand in lista:
            for col in df.columns:
                if ''.join(ch for ch in col.lower() if ch.isalnum()) == ''.join(ch for ch in cand.lower() if ch.isalnum()):
                    return col
        return None

    for target, candidates in m.items():
        col = achar(df2, candidates)
        df2[target] = df2[col] if col is not None else pd.NA
    return df2

def preparar_dados(df):
    if 'data_compra' in df.columns:
        df['data_compra'] = pd.to_datetime(df['data_compra'], errors='coerce').dt.date

    df['preco'] = pd.to_numeric(df.get('preco'), errors='coerce')
    df['km'] = pd.to_numeric(df.get('km'), errors='coerce')
    df['parcela_mensal'] = pd.to_numeric(df.get('parcela_mensal'), errors='coerce')

    if 'financiado' in df.columns:
        def to_bool(v):
            try:
                if pd.isna(v): return False
                if isinstance(v, bool): return v
                s = str(v).strip().lower()
                return s in ('true','1','sim','s','yes','y','financiado')
            except:
                return False
        df['financiado'] = df['financiado'].apply(to_bool).astype(bool)
    else:
        df['financiado'] = False

    df['forma_pagamento'] = df['financiado'].apply(lambda v: 'Financiado' if v else 'À vista')

    if 'data_compra' in df.columns:
        df['mes_compra'] = pd.to_datetime(df['data_compra']).dt.to_period('M')
        df['ano_compra'] = pd.to_datetime(df['data_compra']).dt.year

    return df


# ------------------------------
#   DASHBOARD INTERATIVO
# ------------------------------
def iniciar_dashboard(df):

    app = Dash(__name__, external_stylesheets=[dbc.themes.COSMO])

    app.layout = dbc.Container([
        html.H2("Dashboard de Vendas — Interativo", className="mt-3"),

        html.Hr(),

        dbc.Row([
            dbc.Col([
                html.Label("Escolha o tipo de gráfico:"),
                dcc.Dropdown(
                    id="tipo_grafico",
                    options=[
                        {"label": "Histograma de Preço", "value": "hist"},
                        {"label": "Top 10 Marcas", "value": "marcas"},
                        {"label": "Top 10 Cidades", "value": "cidades"},
                        {"label": "Compras por Mês", "value": "mes"},
                        {"label": "Financiado x À Vista", "value": "pagamento"}
                    ],
                    value="hist"
                )
            ], width=4),

            dbc.Col([
                html.Label("Filtrar cidade (opcional):"),
                dcc.Dropdown(
                    id="filtro_cidade",
                    options=[{"label": c, "value": c} for c in sorted(df["cidade"].dropna().unique())],
                    value=None,
                    placeholder="Todas",
                    clearable=True
                )
            ], width=4)
        ], className="mt-3"),

        html.Br(),

        dcc.Graph(id="grafico_principal")
    ], fluid=True)

    # ------------------------------
    # CALLBACK DOS GRÁFICOS
    # ------------------------------
    @app.callback(
        Output("grafico_principal", "figure"),
        Input("tipo_grafico", "value"),
        Input("filtro_cidade", "value")
    )
    def atualizar_grafico(tipo, cidade):

        df2 = df.copy()
        if cidade:
            df2 = df2[df2["cidade"] == cidade]

        if tipo == "hist":
            fig = px.histogram(df2, x="preco", nbins=20,
                               title="Distribuição de Preços")
            return fig

        elif tipo == "marcas":
            top = df2["marca"].value_counts().nlargest(10)
            fig = px.bar(x=top.index, y=top.values, title="Top 10 Marcas")
            return fig

        elif tipo == "cidades":
            top = df2["cidade"].value_counts().nlargest(10)
            fig = px.bar(x=top.index, y=top.values, title="Top 10 Cidades")
            return fig

        elif tipo == "mes":
            if df2["mes_compra"].notna().sum() == 0:
                return go.Figure()
            series = df2["mes_compra"].value_counts().sort_index()
            fig = px.line(x=series.index.astype(str), y=series.values,
                          markers=True, title="Compras Por Mês")
            return fig

        elif tipo == "pagamento":
            counts = df2["forma_pagamento"].value_counts()
            fig = px.pie(values=counts.values, names=counts.index,
                         title="Forma de Pagamento")
            return fig

        return go.Figure()

    app.run(debug=True, port=8050, host="127.0.0.1")


# ------------------------------
# MAIN
# ------------------------------
if __name__ == "__main__":
    df = carregar_planilha(INPUT_XLSX)
    df = normalizar_colunas(df)
    df = preparar_dados(df)

    print("Iniciando dashboard em: http://127.0.0.1:8050")
    iniciar_dashboard(df)

