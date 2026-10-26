import pandas as pd
import math
import io
import streamlit as st
import plotly.express as px

# ==========================================
# 1. FUNÇÕES DE PROCESSAMENTO
# ==========================================

def processar_reordenamento(df):
    """Processa o reordenamento escolar com base na planilha CadastroDeTurmas."""
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
    resultados = []

    agrupado = df_2027_base.groupby(['CIDADE', 'ESCOLA', 'LOCAL_TIPO', 'ANO/SÉRIE 2027'])

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
    """Carrega dados da planilha de Transferência Automática (T.A.)."""
    df_ta = pd.read_excel(file, sheet_name='Município -> Estado', header=1)
    df_clean = df_ta.dropna(subset=['MUNICIPIO', 'QUANT. DE ALUNOS']).copy()
    
    df_clean.rename(columns={
        'MUNICIPIO': 'CIDADE',
        'U.E. ORIGEM': 'ESCOLA_ORIGEM',
        'U.E. DESTINO': 'ESCOLA_DESTINO',
        'ANO/\nSÉRIE/PERÍODO DE ORIGEM': 'SERIE_ORIGEM',
        'ANO/\nSÉRIE/PERÍODODE DESTINO': 'SERIE_DESTINO',
        'TOTAL DE ALUNOS': 'TOTAL_ALUNOS',
        'ALUNOS ESPECIAIS': 'ALUNOS_PCD'
    }, inplace=True)
    return df_clean


def processar_salas_ociosas(file):
    """Carrega dados do levantamento de capacidade física e salas livres."""
    df_raw = pd.read_excel(file, header=1)
    df_salas = df_raw.iloc[2:].copy()
    df_salas.columns = [
        'Nº', 'INEP', 'ESCOLA', 'TOTAL_SALAS', 
        'TURMAS_2026_MAT', 'TURMAS_2026_VESP', 'TURMAS_2026_NOT', 'TURMAS_2026_INT',
        'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 'SALAS_LIVRES_2027_NOT', 'SALAS_LIVRES_2027_INT',
        'OBSERVACAO'
    ]
    df_salas = df_salas.dropna(subset=['ESCOLA']).copy()
    df_salas = df_salas[df_salas['ESCOLA'] != 'Escola'].copy()

    cols_num = ['TOTAL_SALAS', 'TURMAS_2026_MAT', 'TURMAS_2026_VESP', 'TURMAS_2026_NOT', 'TURMAS_2026_INT',
                'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 'SALAS_LIVRES_2027_NOT', 'SALAS_LIVRES_2027_INT']
    
    for c in cols_num:
        df_salas[c] = pd.to_numeric(df_salas[c], errors='coerce').fillna(0)
        
    return df_salas

# ==========================================
# 2. CONFIGURAÇÃO DA PÁGINA E ESTILIZAÇÃO CSS
# ==========================================
st.set_page_config(
    page_title="Sistema Integrado 2027 - SRE Paraíso",
    page_icon="🏛️",
    layout="wide"
)

# Estilização CSS personalizada para os KPI Cards
st.markdown("""
    <style>
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 15px;
        margin-bottom: 20px;
    }
    .kpi-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 12px;
        padding: 18px 22px;
        flex: 1;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border: 1px solid #e9ecef;
        transition: transform 0.2s ease-in-out;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
    }
    .kpi-card-blue { border-left: 6px solid #1E88E5; }
    .kpi-card-green { border-left: 6px solid #2E7D32; }
    .kpi-card-orange { border-left: 6px solid #EF6C00; }
    .kpi-card-purple { border-left: 6px solid #8E24AA; }
    
    .kpi-title {
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #6c757d;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 2rem;
        font-weight: 800;
        color: #2c3e50;
        line-height: 1.1;
    }
    .kpi-sub {
        font-size: 0.78rem;
        color: #95a5a6;
        margin-top: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# Cabeçalho Principal
st.title("🏛️ Painel Integrado de Reordenamento Escolar (2026 ➔ 2027)")
st.markdown("##### **Superintendência Regional de Educação de Paraíso do Tocantins - SEDUC/TO**")
st.markdown("---")

# Sidebar / Carregamento dos Arquivos
st.sidebar.image("https://img.icons8.com/color/96/school.png", width=70)
st.sidebar.title("Central de Dados")
st.sidebar.markdown("Faça o upload das planilhas de monitoramento:")

file_turmas = st.sidebar.file_uploader("1. Cadastro de Turmas (.xlsx)", type=["xlsx"], key="turmas")
file_ta = st.sidebar.file_uploader("2. Transferência Automática (.xlsx)", type=["xlsx"], key="ta")
file_salas = st.sidebar.file_uploader("3. Levantamento de Salas (.xlsx)", type=["xlsx"], key="salas")

if file_turmas:
    df_raw = pd.read_excel(file_turmas)
    df_projetado = processar_reordenamento(df_raw)

    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filtros Dinâmicos")
    cidades = ["Todas"] + sorted(list(df_projetado['CIDADE'].unique()))
    cidade_sel = st.sidebar.selectbox("Município:", cidades)

    if cidade_sel != "Todas":
        df_projetado = df_projetado[df_projetado['CIDADE'] == cidade_sel]

    escolas = ["Todas"] + sorted(list(df_projetado['ESCOLA'].unique()))
    escola_sel = st.sidebar.selectbox("Unidade Escolar:", escolas)

    if escola_sel != "Todas":
        df_projetado = df_projetado[df_projetado['ESCOLA'] == escola_sel]

    # Estrutura Principal de Abas
    aba1, aba2, aba3, aba4, aba5 = st.tabs([
        "📋 Projeção Detalhada", 
        "📊 Dashboard Reordenamento", 
        "🔄 Transferência Automática (T.A.)",
        "🏢 Capacidade Física & Salas",
        "🎯 Matriz de Viabilidade (Salas vs T.A.)"
    ])

    # ------------------------------------------
    # ABA 1: PROJEÇÃO DETALHADA
    # ------------------------------------------
    with aba1:
        st.subheader("📋 Planilha Otimizada de Reordenamento 2027")
        
        # HTML dos Cards
        total_estudantes = df_projetado["TOTAL ESTUDANTES 2027"].sum()
        total_salas = df_projetado["SALAS NECESSÁRIAS 2027"].sum()
        total_escolas = df_projetado["ESCOLA"].nunique()

        st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card kpi-card-blue">
                <div class="kpi-title">Estudantes Projetados</div>
                <div class="kpi-value">{total_estudantes:,}</div>
                <div class="kpi-sub">Fluxo regular 2026 -> 2027</div>
            </div>
            <div class="kpi-card kpi-card-green">
                <div class="kpi-title">Salas Necessárias</div>
                <div class="kpi-value">{total_salas:,}</div>
                <div class="kpi-sub">Turmas otimizadas</div>
            </div>
            <div class="kpi-card kpi-card-orange">
                <div class="kpi-title">Unidades Escolas</div>
                <div class="kpi-value">{total_escolas}</div>
                <div class="kpi-sub">Escolas mapeadas</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(df_projetado, use_container_width=True)

        def converter_para_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Reordenamento_2027')
            return output.getvalue()

        st.download_button(
            label="📥 Baixar Planilha Otimizada de 2027 (Excel .xlsx)",
            data=converter_para_excel(df_projetado),
            file_name="Projecao_Reordenamento_2027_Paraiso.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ------------------------------------------
    # ABA 2: DASHBOARD REORDENAMENTO
    # ------------------------------------------
    with aba2:
        st.subheader("📊 Análise Gráfica do Reordenamento")

        col_g1, col_g2 = st.columns(2)

        with col_g1:
            df_cid = df_projetado.groupby('CIDADE')['TOTAL ESTUDANTES 2027'].sum().reset_index().sort_values('TOTAL ESTUDANTES 2027', ascending=True)
            fig_cid = px.bar(
                df_cid, x='TOTAL ESTUDANTES 2027', y='CIDADE', orientation='h',
                title="Estudantes Projetados por Município",
                text_auto=True, color='TOTAL ESTUDANTES 2027', color_continuous_scale='Blues'
            )
            fig_cid.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_cid, use_container_width=True)

        with col_g2:
            df_ser = df_projetado.groupby('SÉRIE 2027')['TOTAL ESTUDANTES 2027'].sum().reset_index()
            fig_ser = px.bar(
                df_ser, x='SÉRIE 2027', y='TOTAL ESTUDANTES 2027',
                title="Distribuição de Estudantes por Série / Ano (2027)",
                text_auto=True, color='TOTAL ESTUDANTES 2027', color_continuous_scale='Teal'
            )
            fig_ser.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig_ser, use_container_width=True)

        col_g3, col_g4 = st.columns(2)

        with col_g3:
            df_tur = df_projetado.groupby('TURNO 2027')['SALAS NECESSÁRIAS 2027'].sum().reset_index()
            fig_tur = px.pie(
                df_tur, names='TURNO 2027', values='SALAS NECESSÁRIAS 2027',
                title="Projeção de Alocação de Salas por Turno", hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig_tur.update_layout(height=380)
            st.plotly_chart(fig_tur, use_container_width=True)

        with col_g4:
            df_top = df_projetado.groupby('ESCOLA')[['TOTAL ESTUDANTES 2027', 'SALAS NECESSÁRIAS 2027']].sum().reset_index().sort_values('SALAS NECESSÁRIAS 2027', ascending=False).head(6)
            fig_top = px.bar(
                df_top, x='SALAS NECESSÁRIAS 2027', y='ESCOLA', orientation='h',
                title="Top Escolas com Maior Demanda de Salas Físicas",
                text_auto=True, color='SALAS NECESSÁRIAS 2027', color_continuous_scale='Reds'
            )
            fig_top.update_layout(showlegend=False, height=380)
            st.plotly_chart(fig_top, use_container_width=True)

    # ------------------------------------------
    # ABA 3: TRANSFERÊNCIA AUTOMÁTICA
    # ------------------------------------------
    with aba3:
        st.subheader("🔄 Ingressantes da Rede Municipal (Transferência Automática)")

        if file_ta:
            df_ta = processar_transferencia_automatica(file_ta)
            if cidade_sel != "Todas":
                df_ta = df_ta[df_ta['CIDADE'] == cidade_sel]

            m_ta1 = int(df_ta['TOTAL_ALUNOS'].sum())
            m_ta2 = int(df_ta['ALUNOS_PCD'].sum())
            m_ta3 = df_ta['ESCOLA_ORIGEM'].nunique()

            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card kpi-card-orange">
                    <div class="kpi-title">Ingressantes T.A. (Município -> Estado)</div>
                    <div class="kpi-value">{m_ta1:,}</div>
                    <div class="kpi-sub">Estudantes encaminhados</div>
                </div>
                <div class="kpi-card kpi-card-purple">
                    <div class="kpi-title">Estudantes PCDs / Ed. Especial</div>
                    <div class="kpi-value">{m_ta2}</div>
                    <div class="kpi-sub">Atendimento especializado</div>
                </div>
                <div class="kpi-card kpi-card-blue">
                    <div class="kpi-title">Escolas Municipais doadoras</div>
                    <div class="kpi-value">{m_ta3}</div>
                    <div class="kpi-sub">Origem da demanda</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col_ta1, col_ta2 = st.columns(2)

            with col_ta1:
                df_ta_cid = df_ta.groupby('CIDADE')['TOTAL_ALUNOS'].sum().reset_index().sort_values('TOTAL_ALUNOS', ascending=True)
                fig_ta_cid = px.bar(
                    df_ta_cid, x='TOTAL_ALUNOS', y='CIDADE', orientation='h',
                    title="Ingressantes T.A. por Município",
                    text_auto=True, color='TOTAL_ALUNOS', color_continuous_scale='Oranges'
                )
                fig_ta_cid.update_layout(showlegend=False, height=380)
                st.plotly_chart(fig_ta_cid, use_container_width=True)

            with col_ta2:
                df_ta_ser = df_ta.groupby('SERIE_DESTINO')['TOTAL_ALUNOS'].sum().reset_index()
                fig_ta_ser = px.pie(
                    df_ta_ser, names='SERIE_DESTINO', values='TOTAL_ALUNOS',
                    title="Etapa de Entrada na Rede Estadual", hole=0.4
                )
                fig_ta_ser.update_layout(height=380)
                st.plotly_chart(fig_ta_ser, use_container_width=True)

            st.markdown("##### 📜 Detalhamento do Fluxo de Origem e Destino")
            st.dataframe(df_ta[['CIDADE', 'ESCOLA_ORIGEM', 'ESCOLA_DESTINO', 'SERIE_DESTINO', 'TOTAL_ALUNOS', 'ALUNOS_PCD']], use_container_width=True)

        else:
            st.info("💡 Faça o upload do arquivo 'SRE - PARAISO T.A 2026_2027.xlsx' no menu lateral para habilitar este dashboard.")

    # ------------------------------------------
    # ABA 4: CAPACIDADE FÍSICA E SALAS LIVRES
    # ------------------------------------------
    with aba4:
        st.subheader("🏢 Estrutura Física e Ociosidade de Salas")

        if file_salas:
            df_s = processar_salas_ociosas(file_salas)

            k_s1 = int(df_s['TOTAL_SALAS'].sum())
            k_s2 = int(df_s['SALAS_LIVRES_2027_MAT'].sum())
            k_s3 = int(df_s['SALAS_LIVRES_2027_VESP'].sum())
            k_s4 = int(df_s['SALAS_LIVRES_2027_NOT'].sum())

            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card kpi-card-blue">
                    <div class="kpi-title">Salas Existentes</div>
                    <div class="kpi-value">{k_s1}</div>
                    <div class="kpi-sub">Capacidade física total</div>
                </div>
                <div class="kpi-card kpi-card-green">
                    <div class="kpi-title">Salas Livres (Matutino)</div>
                    <div class="kpi-value">{k_s2}</div>
                    <div class="kpi-sub">Disponíveis em 2027</div>
                </div>
                <div class="kpi-card kpi-card-orange">
                    <div class="kpi-title">Salas Livres (Vespertino)</div>
                    <div class="kpi-value">{k_s3}</div>
                    <div class="kpi-sub">Disponíveis em 2027</div>
                </div>
                <div class="kpi-card kpi-card-purple">
                    <div class="kpi-title">Salas Livres (Noturno)</div>
                    <div class="kpi-value">{k_s4}</div>
                    <div class="kpi-sub">Disponíveis em 2027</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col_s1, col_s2 = st.columns(2)

            with col_s1:
                df_s_topm = df_s.sort_values('SALAS_LIVRES_2027_MAT', ascending=False).head(7)
                fig_s1 = px.bar(
                    df_s_topm, x='SALAS_LIVRES_2027_MAT', y='ESCOLA', orientation='h',
                    title="Escolas com Maior Reserva no Matutino",
                    text_auto=True, color='SALAS_LIVRES_2027_MAT', color_continuous_scale='Greens'
                )
                fig_s1.update_layout(showlegend=False, height=380)
                st.plotly_chart(fig_s1, use_container_width=True)

            with col_s2:
                df_s_topv = df_s.sort_values('SALAS_LIVRES_2027_VESP', ascending=False).head(7)
                fig_s2 = px.bar(
                    df_s_topv, x='SALAS_LIVRES_2027_VESP', y='ESCOLA', orientation='h',
                    title="Escolas com Maior Reserva no Vespertino",
                    text_auto=True, color='SALAS_LIVRES_2027_VESP', color_continuous_scale='Purples'
                )
                fig_s2.update_layout(showlegend=False, height=380)
                st.plotly_chart(fig_s2, use_container_width=True)

            st.markdown("##### 🏢 Quadro Geral por Unidade Escolar")
            st.dataframe(
                df_s[[
                    'ESCOLA', 'TOTAL_SALAS', 
                    'SALAS_LIVRES_2027_MAT', 'SALAS_LIVRES_2027_VESP', 
                    'SALAS_LIVRES_2027_NOT', 'SALAS_LIVRES_2027_INT', 'OBSERVACAO'
                ]], use_container_width=True
            )

        else:
            st.info("💡 Faça o upload do arquivo 'Levantamento de salas ociosas SRE-2026.xlsx' no menu lateral para habilitar este dashboard.")

    # ------------------------------------------
    # ABA 5: MATRIZ DE VIABILIDADE
    # ------------------------------------------
    with aba5:
        st.subheader("🎯 Matriz de Viabilidade Escolar (Cruzamento T.A. vs Salas Ociosas)")

        if file_ta and file_salas:
            df_ta = processar_transferencia_automatica(file_ta)
            df_s = processar_salas_ociosas(file_salas)

            # Agrupar demanda por escola de destino
            df_ta_resumo = df_ta.groupby('ESCOLA_DESTINO')['TOTAL_ALUNOS'].sum().reset_index()
            df_ta_resumo['TURMAS_NOVAS_ESTIMADAS'] = df_ta_resumo['TOTAL_ALUNOS'].apply(lambda x: math.ceil(x / 35))

            # Unir com o levantamento de salas livres
            df_matriz = pd.merge(
                df_s, df_ta_resumo, 
                left_on='ESCOLA', right_on='ESCOLA_DESTINO', 
                how='inner'
            )

            df_matriz['SALAS_DISPONIVEIS_TOTAL'] = df_matriz['SALAS_LIVRES_2027_MAT'] + df_matriz['SALAS_LIVRES_2027_VESP']
            df_matriz['STATUS_CAPACIDADE'] = df_matriz.apply(
                lambda row: "✅ Capacidade Atendida" if row['SALAS_DISPONIVEIS_TOTAL'] >= row['TURMAS_NOVAS_ESTIMADAS'] else "⚠️ Atenção / Capacidade Limite", 
                axis=1
            )

            st.markdown("##### 🔄 Balanço da Capacidade para Receber Estudantes Novatos em 2027")
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
            st.info("💡 Envie as planilhas 'Transferência Automática' e 'Levantamento de Salas' no menu lateral para visualizar o cruzamento de viabilidade.")

else:
    st.warning("⚠️ Carregue a planilha principal 'CadastroDeTurmas' no menu lateral para inicializar o dashboard.")