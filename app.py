import pandas as pd
import math
import io
import os
import streamlit as st
import plotly.express as px

# ==========================================
# 1. FUNÇÕES DE TRATAMENTO E NORMALIZAÇÃO
# ==========================================

def normalizar_texto(texto):
    """Limpa e padroniza nomes de escolas e municípios."""
    if pd.isna(texto):
        return ""
    txt = str(texto).upper().strip()
    txt = txt.replace("  ", " ")
    
    # Correção de anomalias conhecidas
    subst = {
        'PARAÍSO DO TOCANTINS DO TOCAN': 'PARAÍSO DO TOCANTINS',
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
        if de in txt:
            txt = txt.replace(de, para)
    return txt.strip()

def carregar_dados_locais():
    file_turmas = "CadastroDeTurmas_26_10_2026.xlsx" if os.path.exists("CadastroDeTurmas_26_10_2026.xlsx") else None
    file_ta = "SRE - PARAISO  T.A 2026_2027 (5).xlsx" if os.path.exists("SRE - PARAISO  T.A 2026_2027 (5).xlsx") else None
    file_salas = "Levantamento de salas ociosas SRE-2026.xlsx" if os.path.exists("Levantamento de salas ociosas SRE-2026.xlsx") else None
    return file_turmas, file_ta, file_salas

# ==========================================
# 2. PROCESSAMENTO DAS TRÊS BASES
# ==========================================

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
    excel_ta = pd.ExcelFile(file)
    dfs = []

    if 'Município -> Estado' in excel_ta.sheet_names:
        df_mun = pd.read_excel(file, sheet_name='Município -> Estado', header=1)
        df_mun = df_mun.dropna(subset=['MUNICIPIO', 'QUANT. DE ALUNOS']).copy()
        df_mun['TIPO_FLUXO'] = 'Município -> Estado'
        dfs.append(df_mun)

    if 'Estado -> Estado' in excel_ta.sheet_names:
        df_est = pd.read_excel(file, sheet_name='Estado -> Estado', header=1)
        df_est = df_est.dropna(subset=['MUNICIPIO', 'QUANT. DE ALUNOS']).copy()
        df_est['TIPO_FLUXO'] = 'Estado -> Estado'
        dfs.append(df_est)

    df_concat = pd.concat(dfs, ignore_index=True)
    df_concat['CIDADE'] = df_concat['MUNICIPIO'].apply(normalizar_texto)
    df_concat['ESCOLA_ORIGEM'] = df_concat['U.E. ORIGEM'].apply(normalizar_texto)
    df_concat['ESCOLA_DESTINO'] = df_concat['U.E. DESTINO'].apply(normalizar_texto)
    df_concat['SERIE_ORIGEM'] = df_concat['ANO/\nSÉRIE/PERÍODO DE ORIGEM']
    df_concat['SERIE_DESTINO'] = df_concat['ANO/\nSÉRIE/PERÍODODE DESTINO']
    df_concat['TOTAL_ALUNOS'] = pd.to_numeric(df_concat['TOTAL DE ALUNOS'], errors='coerce').fillna(0).astype(int)
    df_concat['ALUNOS_PCD'] = pd.to_numeric(df_concat['ALUNOS ESPECIAIS'], errors='coerce').fillna(0).astype(int)

    return df_concat

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
# 3. INTERFACE WEB STREAMLIT
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
    .kpi-card-red { border-left: 6px solid #D32F2F; }
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

    tipo_oferta = st.sidebar.radio("Tipo de Oferta Escolar:", ["Todas", "Parcial", "Tempo Integral"])

    aba1, aba2, aba3, aba4, aba5 = st.tabs([
        "📋 Projeção Detalhada", 
        "📊 Dashboard Reordenamento", 
        "🔄 Transferência Automática (T.A.)",
        "🏢 Capacidade Física & Salas",
        "🎯 Matriz de Viabilidade Escolar"
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

    with aba2:
        st.subheader("📊 Análise Gráfica do Reordenamento da Rede Estadual")
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            df_cid = df_proj_f.groupby('CIDADE')['TOTAL ESTUDANTES 2027'].sum().reset_index().sort_values('TOTAL ESTUDANTES 2027', ascending=True)
            fig_cid = px.bar(df_cid, x='TOTAL ESTUDANTES 2027', y='CIDADE', orientation='h', title="Estudantes Projetados por Município", text_auto=True, color='TOTAL ESTUDANTES 2027', color_continuous_scale='Blues')
            fig_cid.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_cid, use_container_width=True)

        with col_g2:
            df_ser = df_proj_f.groupby('SÉRIE 2027')['TOTAL ESTUDANTES 2027'].sum().reset_index()
            fig_ser = px.bar(df_ser, x='SÉRIE 2027', y='TOTAL ESTUDANTES 2027', title="Estudantes Projetados por Série / Ano (2027)", text_auto=True, color='TOTAL ESTUDANTES 2027', color_continuous_scale='Teal')
            fig_ser.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_ser, use_container_width=True)

    with aba3:
        st.subheader("🔄 Ingressantes da Rede Municipal para a Rede Estadual (T.A.)")
        if file_ta:
            df_ta = processar_transferencia_automatica(file_ta)
            
            # Filtro Inteligente Flexível por Nome
            if cidade_sel != "Todas":
                df_ta = df_ta[df_ta['CIDADE'].str.contains(cidade_sel[:8], case=False, na=False)]
            if escola_sel != "Todas":
                # Busca flexível para garantir escolas com nomes levemente divergentes
                termo_escola = escola_sel.replace("COLÉGIO ESTADUAL", "").replace("CENTRO DE ENSINO MÉDIO", "").strip()[:10]
                df_ta = df_ta[df_ta['ESCOLA_DESTINO'].str.contains(termo_escola, case=False, na=False)]

            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card kpi-card-orange">
                    <div class="kpi-title">Ingressantes T.A. (Novatos)</div>
                    <div class="kpi-value">{int(df_ta['TOTAL_ALUNOS'].sum()):,}</div>
                    <div class="kpi-sub">Alunos encaminhados</div>
                </div>
                <div class="kpi-card kpi-card-purple">
                    <div class="kpi-title">Estudantes PCDs / Ed. Especial</div>
                    <div class="kpi-value">{int(df_ta['ALUNOS_PCD'].sum())}</div>
                    <div class="kpi-sub">Atendimento especializado</div>
                </div>
                <div class="kpi-card kpi-card-blue">
                    <div class="kpi-title">Escolas Municipais Doadoras</div>
                    <div class="kpi-value">{df_ta['ESCOLA_ORIGEM'].nunique()}</div>
                    <div class="kpi-sub">Origem da demanda</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.dataframe(df_ta[['CIDADE', 'TIPO_FLUXO', 'ESCOLA_ORIGEM', 'ESCOLA_DESTINO', 'SERIE_DESTINO', 'TOTAL_ALUNOS', 'ALUNOS_PCD']], use_container_width=True)

    with aba4:
        st.subheader("🏢 Estrutura Física e Ociosidade de Salas")
        if file_salas:
            df_s = processar_salas_ociosas(file_salas)
            if escola_sel != "Todas":
                termo_escola = escola_sel.replace("COLÉGIO ESTADUAL", "").replace("CENTRO DE ENSINO MÉDIO", "").strip()[:10]
                df_s = df_s[df_s['ESCOLA'].str.contains(termo_escola, case=False, na=False)]

            st.dataframe(df_s[['ESCOLA', 'TOTAL_SALAS', 'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 'SALAS_LIVRES_2027_NOT', 'SALAS_LIVRES_2027_INT', 'OBSERVACAO']], use_container_width=True)

    with aba5:
        st.subheader("🎯 Matriz de Viabilidade Escolar (Cruzamento Reordenamento + T.A. vs Ociosidade)")
        if file_ta and file_salas:
            df_ta = processar_transferencia_automatica(file_ta)
            df_s = processar_salas_ociosas(file_salas)

            # Agrupar T.A. por escola de destino usando chave simplificada
            df_ta['CHAVE_ESCOLA'] = df_ta['ESCOLA_DESTINO'].apply(lambda x: x.replace("COLÉGIO ESTADUAL", "").replace("CENTRO DE ENSINO MÉDIO", "").replace("ESCOLA ESTADUAL", "").strip())
            df_ta_resumo = df_ta.groupby('CHAVE_ESCOLA')['TOTAL_ALUNOS'].sum().reset_index()
            df_ta_resumo.rename(columns={'TOTAL_ALUNOS': 'DEMANDA_TA_ALUNOS'}, inplace=True)
            df_ta_resumo['TURMAS_NOVAS_TA'] = df_ta_resumo['DEMANDA_TA_ALUNOS'].apply(lambda x: math.ceil(x / 35))

            # Agrupar Reordenamento
            df_proj_f['CHAVE_ESCOLA'] = df_proj_f['ESCOLA'].apply(lambda x: x.replace("COLÉGIO ESTADUAL", "").replace("CENTRO DE ENSINO MÉDIO", "").replace("ESCOLA ESTADUAL", "").strip())
            df_proj_resumo = df_proj_f.groupby('CHAVE_ESCOLA')[['TOTAL ESTUDANTES 2027', 'SALAS NECESSÁRIAS 2027']].sum().reset_index()
            df_proj_resumo.rename(columns={'TOTAL ESTUDANTES 2027': 'ALUNOS_REORD', 'SALAS NECESSÁRIAS 2027': 'SALAS_REORD'}, inplace=True)

            # Chave de junção nas salas
            df_s['CHAVE_ESCOLA'] = df_s['ESCOLA'].apply(lambda x: x.replace("COLÉGIO ESTADUAL", "").replace("CENTRO DE ENSINO MÉDIO", "").replace("ESCOLA ESTADUAL", "").strip())

            # Merge
            df_matriz = pd.merge(df_s, df_proj_resumo, on='CHAVE_ESCOLA', how='left')
            df_matriz = pd.merge(df_matriz, df_ta_resumo, on='CHAVE_ESCOLA', how='left')

            if escola_sel != "Todas":
                termo_escola = escola_sel.replace("COLÉGIO ESTADUAL", "").replace("CENTRO DE ENSINO MÉDIO", "").strip()[:10]
                df_matriz = df_matriz[df_matriz['ESCOLA'].str.contains(termo_escola, case=False, na=False)]

            df_matriz['ALUNOS_REORD'] = df_matriz['ALUNOS_REORD'].fillna(0).astype(int)
            df_matriz['SALAS_REORD'] = df_matriz['SALAS_REORD'].fillna(0).astype(int)
            df_matriz['DEMANDA_TA_ALUNOS'] = df_matriz['DEMANDA_TA_ALUNOS'].fillna(0).astype(int)
            df_matriz['TURMAS_NOVAS_TA'] = df_matriz['TURMAS_NOVAS_TA'].fillna(0).astype(int)

            df_matriz['TOTAL_ALUNOS_2027'] = df_matriz['ALUNOS_REORD'] + df_matriz['DEMANDA_TA_ALUNOS']
            df_matriz['CAPACIDADE_LIVRE_TOTAL'] = df_matriz['SALAS_LIVRES_2027_MAT'] + df_matriz['SALAS_LIVRES_2027_VESP'] + df_matriz['SALAS_LIVRES_2027_INT']

            df_matriz['STATUS_CAPACIDADE'] = df_matriz.apply(
                lambda r: "✅ Capacidade Atendida" if r['CAPACIDADE_LIVRE_TOTAL'] >= r['TURMAS_NOVAS_TA'] else "⚠️ Atenção / Capacidade Limite", 
                axis=1
            )

            qtd_atendida = len(df_matriz[df_matriz['STATUS_CAPACIDADE'].str.contains("✅")])
            qtd_alerta = len(df_matriz[df_matriz['STATUS_CAPACIDADE'].str.contains("⚠️")])

            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card kpi-card-green">
                    <div class="kpi-title">Escolas com Capacidade Ok</div>
                    <div class="kpi-value">{qtd_atendida}</div>
                    <div class="kpi-sub">Salas livres atende demanda novata</div>
                </div>
                <div class="kpi-card kpi-card-red">
                    <div class="kpi-title">Escolas em Limite / Alerta</div>
                    <div class="kpi-value">{qtd_alerta}</div>
                    <div class="kpi-sub">Requer readequação de espaço</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            df_exibir = df_matriz[[
                'ESCOLA', 'ALUNOS_REORD', 'DEMANDA_TA_ALUNOS', 'TOTAL_ALUNOS_2027',
                'TURMAS_NOVAS_TA', 'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 'SALAS_LIVRES_2027_INT',
                'STATUS_CAPACIDADE'
            ]].rename(columns={
                'ALUNOS_REORD': 'ALUNOS REORDENAMENTO',
                'DEMANDA_TA_ALUNOS': 'NOVATOS (T.A.)',
                'TOTAL_ALUNOS_2027': 'TOTAL ALUNOS 2027',
                'TURMAS_NOVAS_TA': 'TURMAS NOVAS T.A.',
                'SALAS_LIVRES_2027_MAT': 'LIVRES (MAT)',
                'SALAS_LIVRES_2027_VESP': 'LIVRES (VESP)',
                'SALAS_LIVRES_2027_INT': 'LIVRES (INTEGRAL)'
            })

            st.dataframe(df_exibir, use_container_width=True)
else:
    st.warning("⚠️ Carregue a planilha principal 'CadastroDeTurmas' no menu lateral para inicializar o dashboard.")