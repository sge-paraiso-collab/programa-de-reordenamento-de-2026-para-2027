import pandas as pd
import math
import io
import os
import streamlit as st
import plotly.express as px

# ==========================================
# 1. FUNÇÕES DE PROCESSAMENTO E CARREGAMENTO
# ==========================================

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    txt = str(texto).upper().strip()
    subst = {
        'PARAISO DO TOCANTINS': 'PARAÍSO DO TOCANTINS',
        'PARAÍSO': 'PARAÍSO DO TOCANTINS',
        'PARAISO': 'PARAÍSO DO TOCANTINS',
        'MARIANOPOLIS DO TOCANTINS': 'MARIANÓPOLIS DO TOCANTINS',
        'DIVINOPOLIS DO TOCANTINS': 'DIVINÓPOLIS DO TOCANTINS',
        'CRISTALANDIA': 'CRISTALÂNDIA',
        'ABREULANDIA': 'ABREULÂNDIA',
        'BARROLANDIA': 'BARROLÂNDIA',
        'COLEGIO': 'COLÉGIO',
        'DAVID BARBOA ROLINS': 'DAVID BARBOSA ROLINS',
        'COLÉGIO OTACÍLIO MARQUES ROSAL': 'COLÉGIO MILITAR DO ESTADO DO TOCANTINS - OTACÍLIO MARQUES ROSAL',
        'ESCOLA ESTADUAL DE TEMPO INTEGRAL': 'ESCOLA DE TEMPO INTEGRAL',
        'DONA CÃNDIDA': 'DONA CANDIDA',
        'KUBITSCHEK DE OLIVERA': 'KUBITSCHEK DE OLIVEIRA'
    }
    for de, para in subst.items():
        txt = txt.replace(de, para)
    return txt

def carregar_dados_locais():
    file_turmas = "CadastroDeTurmas_26_10_2026.xlsx" if os.path.exists("CadastroDeTurmas_26_10_2026.xlsx") else None
    file_ta = "SRE - PARAISO  T.A 2026_2027 (5).xlsx" if os.path.exists("SRE - PARAISO  T.A 2026_2027 (5).xlsx") else None
    file_salas = "Levantamento de salas ociosas SRE-2026.xlsx" if os.path.exists("Levantamento de salas ociosas SRE-2026.xlsx") else None
    return file_turmas, file_ta, file_salas

def processar_reordenamento(df):
    df_regular = df[(df['TIPO DE TURMA'] == 'Escolarização') & (df['SITUAÇÃO'] == 'Ativa')].copy()
    
    mapeamento_transicao = {
        '1º ANO': '2º ANO', '2º ANO': '3º ANO', '3º ANO': '4º ANO',
        '4º ANO': '5º ANO', '5º ANO': '6º ANO', '6º ANO': '7º ANO',
        '7º ANO': '8º ANO', '8º ANO': '9º ANO', '9º ANO': '1ª SÉRIE',
        '1ª SÉRIE': '2ª SÉRIE', '2ª SÉRIE': '3ª SÉRIE', '2ª SÉRIE-IN': '3ª SÉRIE-IN',
        '3ª SÉRIE': 'CONCLUÍDO', '3ª SÉRIE-IN': 'CONCLUÍDO'
    }

    df_regular['ANO/SÉRIE 2027'] = df_regular['ANO/SÉRIE'].map(mapeamento_transicao)
    df_2027_base = df_regular[
        df_regular['ANO/SÉRIE 2027'].notna() & 
        (df_regular['ANO/SÉRIE 2027'] != 'CONCLUÍDO')
    ].copy()

    def identificar_local(row):
        turma = str(row['TURMA']).upper()
        multi = str(row['MULTISSERIADA']).upper()
        local_dif = str(row['LOCAL DE FUNCIONAMENTO DIFERENCIADO']).upper()
        
        if multi == 'SIM' or 'EXT.' in turma or 'EXTENSÃO' in turma or local_dif == 'SALA ANEXA':
            if 'EXT' in turma:
                partes = turma.split('EXT')
                nome_ext = 'EXT' + partes[-1]
                return f"Extensão ({nome_ext.strip()})"
            elif multi == 'SIM':
                return "Extensão / Multisseriada"
            else:
                return "Extensão (Sala Anexa)"
        return "Sede"

    df_2027_base['LOCAL_TIPO'] = df_2027_base.apply(identificar_local, axis=1)
    df_2027_base['CIDADE_NORM'] = df_2027_base['CIDADE'].apply(normalizar_texto)
    df_2027_base['ESCOLA_NORM'] = df_2027_base['ESCOLA'].apply(normalizar_texto)

    resultados = []
    agrupado = df_2027_base.groupby(['CIDADE_NORM', 'ESCOLA_NORM', 'LOCAL_TIPO', 'ANO/SÉRIE 2027'])

    for (cidade, escola, local_tipo, serie_2027), grupo in agrupado:
        total_estudantes = grupo['QUANTIDADE DE ESTUDANTES CURSANDO'].sum()
        if total_estudantes == 0:
            continue

        is_medio = 'SÉRIE' in serie_2027
        max_sala = 40 if is_medio else 35
        qtd_salas_2027 = 1 if total_estudantes <= max_sala else math.ceil(total_estudantes / max_sala)

        turnos_2026 = grupo['TURNO'].unique()
        turno_2027 = 'Matutino/Manhã' if 'Matutino/Manhã' in turnos_2026 or len(turnos_2026) > 1 else turnos_2026[0]
        obs = "Sede - Turma Regular" if local_tipo == "Sede" else "Turma de Extensão / Sala Anexa / Multisseriada vinculada"

        resultados.append({
            'CIDADE': cidade,
            'ESCOLA': escola,
            'UNIDADE / LOCAL': local_tipo,
            'SÉRIE 2027': serie_2027,
            'TURNO 2027': turno_2027,
            'TOTAL ESTUDANTES 2027': total_estudantes,
            'SALAS NECESSÁRIAS 2027': qtd_salas_2027,
            'MÉDIA ESTUDANTES/SALA': round(total_estudantes / qtd_salas_2027, 1),
            'TURMAS 2026 ORIGEM': len(grupo),
            'OBSERVAÇÃO': obs
        })

    return pd.DataFrame(resultados)

def processar_transferencia_automatica(file):
    df_ta = pd.read_excel(file, sheet_name='Município -> Estado', header=1)
    df_clean = df_ta.dropna(subset=['MUNICIPIO', 'QUANT. DE ALUNOS']).copy()
    
    df_clean['CIDADE'] = df_clean['MUNICIPIO'].apply(normalizar_texto)
    df_clean['ESCOLA_ORIGEM'] = df_clean['U.E. ORIGEM'].apply(normalizar_texto)
    df_clean['ESCOLA_DESTINO'] = df_clean['U.E. DESTINO'].apply(normalizar_texto)
    df_clean['SERIE_ORIGEM'] = df_clean['ANO/\nSÉRIE/PERÍODO DE ORIGEM']
    df_clean['SERIE_DESTINO'] = df_clean['ANO/\nSÉRIE/PERÍODODE DESTINO']
    df_clean['TOTAL_ALUNOS'] = pd.to_numeric(df_clean['TOTAL DE ALUNOS'], errors='coerce').fillna(0).astype(int)
    df_clean['ALUNOS_PCD'] = pd.to_numeric(df_clean['ALUNOS ESPECIAIS'], errors='coerce').fillna(0).astype(int)
    
    return df_clean

def processar_salas_ociosas(file):
    df_raw = pd.read_excel(file, header=1)
    df_salas = df_raw.iloc[2:].copy()
    df_salas.columns = [
        'Nº', 'INEP', 'ESCOLA_RAW', 'TOTAL_SALAS', 
        'TURMAS_2026_MAT', 'TURMAS_2026_VESP', 'TURMAS_2026_NOT', 'TURMAS_2026_INT',
        'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 'SALAS_LIVRES_2027_NOT', 'SALAS_LIVRES_2027_INT',
        'OBSERVACAO'
    ]
    df_salas = df_salas.dropna(subset=['ESCOLA_RAW']).copy()
    df_salas = df_salas[df_salas['ESCOLA_RAW'] != 'Escola'].copy()

    df_salas['ESCOLA'] = df_salas['ESCOLA_RAW'].apply(normalizar_texto)

    cols_num = ['TOTAL_SALAS', 'TURMAS_2026_MAT', 'TURMAS_2026_VESP', 'TURMAS_2026_NOT', 'TURMAS_2026_INT',
                'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 'SALAS_LIVRES_2027_NOT', 'SALAS_LIVRES_2027_INT']
    
    for c in cols_num:
        df_salas[c] = pd.to_numeric(df_salas[c], errors='coerce').fillna(0).astype(int)
        
    return df_salas

# ==========================================
# 2. INTERFACE STREAMLIT E DADOS
# ==========================================
st.set_page_config(page_title="Sistema Integrado 2027 - SRE Paraíso", page_icon="🏛️", layout="wide")

st.markdown("""
    <style>
    .kpi-container { display: flex; justify-content: space-between; gap: 15px; margin-bottom: 20px; }
    .kpi-card { background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); border-radius: 12px; padding: 18px 22px; flex: 1; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #e9ecef; }
    .kpi-card-blue { border-left: 6px solid #1E88E5; }
    .kpi-card-green { border-left: 6px solid #2E7D32; }
    .kpi-card-orange { border-left: 6px solid #EF6C00; }
    .kpi-card-purple { border-left: 6px solid #8E24AA; }
    .kpi-title { font-size: 0.85rem; font-weight: 700; text-transform: uppercase; color: #6c757d; margin-bottom: 6px; }
    .kpi-value { font-size: 2rem; font-weight: 800; color: #2c3e50; line-height: 1.1; }
    .kpi-sub { font-size: 0.78rem; color: #95a5a6; margin-top: 5px; }
    </style>
""", unsafe_allow_html=True)

st.title("🏛️ Painel Integrado de Reordenamento Escolar (2026 ➔ 2027)")
st.markdown("##### **Superintendência Regional de Educação de Paraíso do Tocantins - SEDUC/TO**")
st.markdown("---")

auto_turmas, auto_ta, auto_salas = carregar_dados_locais()

st.sidebar.image("https://img.icons8.com/color/96/school.png", width=70)
st.sidebar.title("Central de Dados")

file_turmas = st.sidebar.file_uploader("1. Cadastro de Turmas (.xlsx)", type=["xlsx"], key="turmas") or auto_turmas
file_ta = st.sidebar.file_uploader("2. Transferência Automática (.xlsx)", type=["xlsx"], key="ta") or auto_ta
file_salas = st.sidebar.file_uploader("3. Levantamento de Salas (.xlsx)", type=["xlsx"], key="salas") or auto_salas

if file_turmas:
    df_raw = pd.read_excel(file_turmas)
    df_projetado = processar_reordenamento(df_raw)

    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filtros Globais")
    
    cidades = ["Todas"] + sorted(list(df_projetado['CIDADE'].unique()))
    cidade_sel = st.sidebar.selectbox("Município:", cidades)

    df_proj_f = df_projetado.copy()
    if cidade_sel != "Todas":
        df_proj_f = df_proj_f[df_proj_f['CIDADE'] == cidade_sel]

    escolas = ["Todas"] + sorted(list(df_proj_f['ESCOLA'].unique()))
    escola_sel = st.sidebar.selectbox("Unidade Escolar:", escolas)

    if escola_sel != "Todas":
        df_proj_f = df_proj_f[df_proj_f['ESCOLA'] == escola_sel]

    aba1, aba2, aba3, aba4, aba5 = st.tabs([
        "📋 Projeção Detalhada", 
        "📊 Dashboard Reordenamento", 
        "🔄 Transferência Automática (T.A.)",
        "🏢 Capacidade Física & Salas",
        "🎯 Matriz de Viabilidade (Salas vs Demandas)"
    ])

    with aba1:
        st.subheader("📋 Planilha Otimizada de Reordenamento 2027")
        st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card kpi-card-blue">
                <div class="kpi-title">Estudantes Projetados</div>
                <div class="kpi-value">{df_proj_f["TOTAL ESTUDANTES 2027"].sum():,}</div>
                <div class="kpi-sub">Fluxo regular 2026 -> 2027</div>
            </div>
            <div class="kpi-card kpi-card-green">
                <div class="kpi-title">Salas Necessárias</div>
                <div class="kpi-value">{df_proj_f["SALAS NECESSÁRIAS 2027"].sum():,}</div>
                <div class="kpi-sub">Turmas otimizadas</div>
            </div>
            <div class="kpi-card kpi-card-orange">
                <div class="kpi-title">Unidades Escolas</div>
                <div class="kpi-value">{df_proj_f["ESCOLA"].nunique()}</div>
                <div class="kpi-sub">Escolas no filtro</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(df_proj_f, use_container_width=True)

        def converter_para_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Reordenamento_2027')
            return output.getvalue()

        st.download_button(
            label="📥 Baixar Planilha Otimizada (Excel .xlsx)",
            data=converter_para_excel(df_proj_f),
            file_name="Projecao_Reordenamento_2027_Paraiso.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with aba2:
        st.subheader("📊 Análise Gráfica do Reordenamento")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            df_cid = df_proj_f.groupby('CIDADE')['TOTAL ESTUDANTES 2027'].sum().reset_index().sort_values('TOTAL ESTUDANTES 2027', ascending=True)
            fig_cid = px.bar(df_cid, x='TOTAL ESTUDANTES 2027', y='CIDADE', orientation='h', title="Estudantes por Município", text_auto=True, color='TOTAL ESTUDANTES 2027', color_continuous_scale='Blues')
            fig_cid.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_cid, use_container_width=True)

        with col_g2:
            df_ser = df_proj_f.groupby('SÉRIE 2027')['TOTAL ESTUDANTES 2027'].sum().reset_index()
            fig_ser = px.bar(df_ser, x='SÉRIE 2027', y='TOTAL ESTUDANTES 2027', title="Estudantes por Série / Ano", text_auto=True, color='TOTAL ESTUDANTES 2027', color_continuous_scale='Teal')
            fig_ser.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_ser, use_container_width=True)

    with aba3:
        st.subheader("🔄 Ingressantes da Rede Municipal (Transferência Automática)")

        if file_ta:
            df_ta = processar_transferencia_automatica(file_ta)
            if cidade_sel != "Todas":
                df_ta = df_ta[df_ta['CIDADE'] == cidade_sel]
            if escola_sel != "Todas":
                df_ta = df_ta[df_ta['ESCOLA_DESTINO'] == escola_sel]

            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card kpi-card-orange">
                    <div class="kpi-title">Ingressantes T.A.</div>
                    <div class="kpi-value">{int(df_ta['TOTAL_ALUNOS'].sum()):,}</div>
                    <div class="kpi-sub">Município -> Estado</div>
                </div>
                <div class="kpi-card kpi-card-purple">
                    <div class="kpi-title">Estudantes PCDs</div>
                    <div class="kpi-value">{int(df_ta['ALUNOS_PCD'].sum())}</div>
                    <div class="kpi-sub">Educação Especial</div>
                </div>
                <div class="kpi-card kpi-card-blue">
                    <div class="kpi-title">Escolas Doadoras</div>
                    <div class="kpi-value">{df_ta['ESCOLA_ORIGEM'].nunique()}</div>
                    <div class="kpi-sub">Rede Municipal</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.dataframe(df_ta[['CIDADE', 'ESCOLA_ORIGEM', 'ESCOLA_DESTINO', 'SERIE_DESTINO', 'TOTAL_ALUNOS', 'ALUNOS_PCD']], use_container_width=True)

    with aba4:
        st.subheader("🏢 Estrutura Física e Ociosidade de Salas")

        if file_salas:
            df_s = processar_salas_ociosas(file_salas)
            if escola_sel != "Todas":
                df_s = df_s[df_s['ESCOLA'] == escola_sel]

            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card kpi-card-blue">
                    <div class="kpi-title">Salas Existentes</div>
                    <div class="kpi-value">{int(df_s['TOTAL_SALAS'].sum())}</div>
                </div>
                <div class="kpi-card kpi-card-green">
                    <div class="kpi-title">Livres Matutino</div>
                    <div class="kpi-value">{int(df_s['SALAS_LIVRES_2027_MAT'].sum())}</div>
                </div>
                <div class="kpi-card kpi-card-orange">
                    <div class="kpi-title">Livres Vespertino</div>
                    <div class="kpi-value">{int(df_s['SALAS_LIVRES_2027_VESP'].sum())}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.dataframe(df_s[['ESCOLA', 'TOTAL_SALAS', 'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 'SALAS_LIVRES_2027_NOT', 'OBSERVACAO']], use_container_width=True)

    with aba5:
        st.subheader("🎯 Matriz de Viabilidade (Salas Ociosas vs Demanda T.A.)")

        if file_ta and file_salas:
            df_ta = processar_transferencia_automatica(file_ta)
            df_s = processar_salas_ociosas(file_salas)

            if cidade_sel != "Todas":
                df_ta = df_ta[df_ta['CIDADE'] == cidade_sel]
                df_s = df_s[df_s['ESCOLA'].isin(df_proj_f['ESCOLA'])]

            if escola_sel != "Todas":
                df_ta = df_ta[df_ta['ESCOLA_DESTINO'] == escola_sel]
                df_s = df_s[df_s['ESCOLA'] == escola_sel]

            df_ta_resumo = df_ta.groupby('ESCOLA_DESTINO')['TOTAL_ALUNOS'].sum().reset_index()
            df_ta_resumo['TURMAS_NOVAS_ESTIMADAS'] = df_ta_resumo['TOTAL_ALUNOS'].apply(lambda x: math.ceil(x / 35))

            df_matriz = pd.merge(df_s, df_ta_resumo, left_on='ESCOLA', right_on='ESCOLA_DESTINO', how='left')

            df_matriz['TOTAL_ALUNOS'] = df_matriz['TOTAL_ALUNOS'].fillna(0).astype(int)
            df_matriz['TURMAS_NOVAS_ESTIMADAS'] = df_matriz['TURMAS_NOVAS_ESTIMADAS'].fillna(0).astype(int)

            df_matriz['SALAS_DISPONIVEIS_TOTAL'] = df_matriz['SALAS_LIVRES_2027_MAT'] + df_matriz['SALAS_LIVRES_2027_VESP']
            df_matriz['STATUS_CAPACIDADE'] = df_matriz.apply(
                lambda row: "✅ Capacidade Atendida" if row['SALAS_DISPONIVEIS_TOTAL'] >= row['TURMAS_NOVAS_ESTIMADAS'] else "⚠️ Atenção / Capacidade Limite", 
                axis=1
            )

            st.dataframe(
                df_matriz[[
                    'ESCOLA', 'TOTAL_ALUNOS', 'TURMAS_NOVAS_ESTIMADAS', 
                    'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 
                    'STATUS_CAPACIDADE'
                ]].rename(columns={
                    'TOTAL_ALUNOS': 'DEMANDA INGRESSANTES (T.A.)',
                    'TURMAS_NOVAS_ESTIMADAS': 'TURMAS NOVAS NECESSÁRIAS',
                    'SALAS_LIVRES_2027_MAT': 'SALAS LIVRES (MAT)',
                    'SALAS_LIVRES_2027_VESP': 'SALAS LIVRES (VESP)'
                }), use_container_width=True
            )
else:
    st.warning("⚠️ Carregue a planilha principal 'CadastroDeTurmas' no menu lateral.")