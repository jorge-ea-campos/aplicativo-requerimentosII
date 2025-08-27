import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

# --- Constantes ---
# Nomes de colunas para evitar "magic strings" e facilitar a manutenção
COL_NUSP = "nusp"
COL_PROBLEMA = "problema"
COL_PARECER = "parecer"
COL_NOME = "Nome completo"
COL_DISCIPLINA = "disciplina"
COL_ANO = "Ano"
COL_SEMESTRE = "Semestre"
COL_LINK = "link_requerimento"
COL_PLANO = "plano_estudo"
COL_PLANO_PRESENCA = "plano_presenca"
# Constantes para as colunas de exportação (J e K)
COL_PARECER_SG = "Parecer Serviço de Graduação"
COL_OBSERVACAO_SG = "Observação SG"


# Colunas obrigatórias nos arquivos de entrada
REQUIRED_COLS_CONSOLIDADO = [COL_NUSP, COL_DISCIPLINA, COL_ANO, COL_SEMESTRE, COL_PROBLEMA, COL_PARECER]
REQUIRED_COLS_REQUERIMENTOS = [COL_NUSP, COL_NOME, COL_PROBLEMA, COL_LINK, COL_PLANO, COL_PLANO_PRESENCA]

# --- Configuração da Página e Estado da Sessão ---
st.set_page_config(
    page_title="Sistema de Conferência de Requerimentos",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicializa o estado da sessão para armazenar as decisões
if 'decisions' not in st.session_state:
    st.session_state.decisions = {}

# --- CSS Customizado ---
st.markdown("""
<style>
    /* Estilos do Cabeçalho de Identidade Visual */
    .header-container {
        margin-bottom: 2rem;
    }
    .logo-and-title {
        display: flex;
        align-items: center;
        margin-bottom: 1rem;
    }
    .header-logo {
        height: 70px;
        margin-right: 20px;
    }
    .header-title-text {
        display: flex;
        flex-direction: column;
    }
    .university-name {
        font-size: 1.7rem;
        font-weight: bold;
        color: #333;
    }
    .department-name {
        font-size: 1.3rem;
        color: #0072b5;
    }
    .color-bar-yellow { height: 8px; background-color: #FDB913; }
    .color-bar-lightblue { height: 4px; background-color: #89cff0; }
    .color-bar-darkblue { height: 12px; background-color: #003366; }

    /* Estilo do cabeçalho principal do App */
    .main-header {
        font-size: 2.5rem; color: #1f77b4; text-align: center;
        padding: 1rem 0; border-bottom: 3px solid #1f77b4; margin-bottom: 2rem;
        margin-top: 0;
    }
    .metric-card {
        background-color: #f0f2f6; padding: 1.5rem; border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)


# --- Funções de Carregamento e Preparação de Dados ---

def load_data(uploaded_file):
    """Tenta ler um arquivo como Excel e, se falhar, tenta como CSV."""
    try:
        return pd.read_excel(uploaded_file)
    except Exception:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file)
        except Exception:
            st.error(f"Falha ao ler o arquivo '{uploaded_file.name}'. Verifique o formato.")
            return None

def find_and_rename_columns(df, target_col_name, possible_names, other_renames=None):
    """Encontra e renomeia colunas para um padrão definido, evitando duplicatas."""
    rename_dict = {}
    
    all_rules = {}
    if other_renames:
        for original, new in other_renames.items():
            all_rules.setdefault(new, []).append(original.lower().strip())
    
    all_rules.setdefault(target_col_name, []).extend([p.lower().strip() for p in possible_names])

    processed_original_cols = set()
    assigned_target_names = set()

    for col in df.columns:
        if col in processed_original_cols:
            continue
        
        normalized_col = col.lower().strip()
        
        found_match = False
        for target, originals in all_rules.items():
            if normalized_col in originals:
                if target not in assigned_target_names:
                    rename_dict[col] = target
                    processed_original_cols.add(col)
                    assigned_target_names.add(target)
                    found_match = True
                    break
        
        if not found_match and target_col_name == COL_NUSP:
            if any(keyword in normalized_col for keyword in ['nusp', 'numero usp', 'número usp', 'n° usp']):
                 if COL_NUSP not in assigned_target_names:
                    rename_dict[col] = COL_NUSP
                    processed_original_cols.add(col)
                    assigned_target_names.add(COL_NUSP)

    df.rename(columns=rename_dict, inplace=True)
    
    if target_col_name not in df.columns:
         raise ValueError(f"Coluna principal '{target_col_name}' não encontrada ou mapeada. Colunas disponíveis: {', '.join(df.columns.tolist())}")

    return df

def validate_dataframes(df_consolidado, df_requerimentos):
    """Verifica se os DataFrames contêm as colunas necessárias."""
    missing_consolidado = [col for col in REQUIRED_COLS_CONSOLIDADO if col not in df_consolidado.columns]
    missing_requerimentos = [col for col in REQUIRED_COLS_REQUERIMENTOS if col not in df_requerimentos.columns]
    errors = []
    if missing_consolidado: errors.append(f"Arquivo consolidado: colunas faltando - {', '.join(missing_consolidado)}")
    if missing_requerimentos: errors.append(f"Arquivo requerimentos: colunas faltando - {', '.join(missing_requerimentos)}")
    if errors: raise ValueError("\n".join(errors))

def clean_nusp_column(df, file_name):
    """Converte a coluna NUSP para numérico e remove registros inválidos."""
    if COL_NUSP not in df.columns: return df
    df[COL_NUSP] = pd.to_numeric(df[COL_NUSP], errors='coerce')
    invalid_count = df[COL_NUSP].isna().sum()
    if invalid_count > 0:
        st.warning(f"⚠️ Removidos {invalid_count} registros com NUSP inválido do arquivo {file_name}")
        df.dropna(subset=[COL_NUSP], inplace=True)
        df[COL_NUSP] = df[COL_NUSP].astype(int)
    return df

# --- Funções de Análise e Métricas ---

def calculate_metrics(df_merged_with_history):
    """Calcula métricas adicionais a partir do DataFrame com histórico."""
    metrics = {}
    if df_merged_with_history.empty: return metrics
    pareceres = df_merged_with_history['parecer_historico'].str.lower()
    aprovados = pareceres.str.contains('aprovado', na=False) & ~pareceres.str.contains('indeferido|negado', na=False)
    negados = pareceres.str.contains('indeferido|negado', na=False)
    total_com_parecer = aprovados.sum() + negados.sum()
    metrics['taxa_aprovacao'] = (aprovados.sum() / total_com_parecer * 100) if total_com_parecer > 0 else 0
    metrics['top_disciplinas'] = df_merged_with_history['disciplina_historico'].value_counts().head(5)
    if 'Ano_historico' in df_merged_with_history.columns and 'Semestre_historico' in df_merged_with_history.columns:
        df_merged_with_history['periodo'] = df_merged_with_history['Ano_historico'].astype(str) + '/' + df_merged_with_history['Semestre_historico'].astype(str)
        metrics['distribuicao_temporal'] = df_merged_with_history['periodo'].value_counts().sort_index()
    return metrics

# --- Funções de Formatação e Exibição (UI) ---

def format_parecer(parecer):
    """Formata o parecer para exibição com ícones."""
    if pd.isna(parecer): return "📝 Pendente"
    p_str = str(parecer).lower()
    if "negado" in p_str or "indeferido" in p_str: return f"❌ {parecer}"
    if "aprovado" in p_str: return f"✅ {parecer}"
    return f"📝 {parecer}"

def format_problem_type(problem):
    """Formata o tipo de problema para exibição."""
    if pd.isna(problem): return "⚪ Não especificado"
    p_str = str(problem).upper()
    if p_str == "QR": return "🔴 Quebra de Requisito"
    if p_str == "CH": return "🟡 Conflito de Horário"
    return f"⚪ {problem}"

@st.cache_data
def to_excel(df):
    """Converte um DataFrame para um arquivo Excel em memória."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
        worksheet = writer.sheets['Relatorio']
        header_format = writer.book.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BD', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for i, col in enumerate(df.columns):
            width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, min(width, 50))
    return output.getvalue()

def display_metrics(df_req, df_merged_with_history, metrics):
    """Exibe os cartões de métricas principais."""
    st.markdown("### 📊 Métricas Principais")
    cols = st.columns(5)
    with cols[0]: st.metric("Total de Requerimentos", len(df_req))
    alunos_unicos_hist = df_merged_with_history[COL_NUSP].nunique()
    total_alunos_req = df_req[COL_NUSP].nunique()
    percentual_hist = (alunos_unicos_hist / total_alunos_req * 100) if total_alunos_req > 0 else 0
    with cols[1]: st.metric("Alunos com Histórico", alunos_unicos_hist, f"{percentual_hist:.1f}%")
    with cols[2]: st.metric("Quebras de Requisito (Hist.)", (df_merged_with_history["problema_historico"].str.upper() == "QR").sum())
    with cols[3]: st.metric("Conflitos de Horário (Hist.)", (df_merged_with_history["problema_historico"].str.upper() == "CH").sum())
    with cols[4]: st.metric("Taxa de Aprovação (Hist.)", f"{metrics.get('taxa_aprovacao', 0):.1f}%")

def display_charts(metrics):
    """Exibe os gráficos de análise."""
    st.markdown("### 📈 Análise Gráfica dos Alunos com Histórico")
    col1, col2 = st.columns(2)
    if 'top_disciplinas' in metrics and not metrics['top_disciplinas'].empty:
        with col1:
            st.markdown("##### 📚 Top 5 Disciplinas")
            top_d = metrics['top_disciplinas']
            fig = px.bar(top_d, x=top_d.values, y=top_d.index, orientation='h', text=top_d.values)
            fig.update_layout(yaxis_title="Disciplina", xaxis_title="Nº de Pedidos", yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)
    if 'distribuicao_temporal' in metrics and not metrics['distribuicao_temporal'].empty:
        with col2:
            st.markdown("##### 🗓️ Pedidos por Período")
            dist_t = metrics['distribuicao_temporal']
            fig2 = px.line(dist_t, x=dist_t.index, y=dist_t.values, markers=True)
            fig2.update_layout(xaxis_title="Período", yaxis_title="Nº de Pedidos")
            st.plotly_chart(fig2, use_container_width=True)

def display_student_details(df_requerimentos, df_merged):
    """Exibe a seção interativa de detalhes por aluno."""
    st.markdown("### 📋 Análise de Requerimentos por Aluno")
    st.info("Clique no nome para ver o histórico e dar o parecer nos pedidos atuais.")
    
    # CORREÇÃO: Garante que cada aluno (por NUSP) seja processado apenas uma vez.
    alunos_unicos = df_requerimentos[[COL_NUSP, COL_NOME]].drop_duplicates(subset=[COL_NUSP]).sort_values(COL_NOME)

    for _, aluno in alunos_unicos.iterrows():
        nusp_aluno = aluno[COL_NUSP]
        with st.expander(f"👤 {aluno[COL_NOME]} (NUSP: {nusp_aluno})"):
            current_requests = df_requerimentos[df_requerimentos[COL_NUSP] == nusp_aluno]
            
            st.markdown("##### 📌 Requerimento(s) no Semestre Atual para Análise")
            if current_requests.empty:
                st.write("Nenhum requerimento encontrado para este aluno no arquivo atual.")
            else:
                for index, request in current_requests.iterrows():
                    link = request.get(COL_LINK, "")
                    plano_estudo_link = request.get(COL_PLANO, "")
                    plano_presenca_link = request.get(COL_PLANO_PRESENCA, "")
                    
                    # A chave agora usa apenas o índice, que é garantido ser único.
                    decision_key = f"req_{index}"
                    st.session_state.decisions.setdefault(decision_key, {'status': 'Pendente', 'justificativa': ''})
                    
                    # O problema é exibido, mas não usado na chave para evitar colisões com NaN
                    problema_display = request.get('problema_atual', 'Não especificado')
                    st.markdown(f"**Problema/Pedido:** `{problema_display}`")

                    if pd.notna(link) and str(link).strip():
                        st.markdown(f"**🔗 Link para o Requerimento:** [Acessar Link]({link})")
                    else:
                        st.markdown("**🔗 Link para o Requerimento:** Não informado")

                    if pd.notna(plano_estudo_link) and str(plano_estudo_link).strip():
                        st.markdown(f"**📄 Link para o Plano de Estudo:** [Acessar Documento]({plano_estudo_link})")
                    else:
                        st.markdown("**📄 Plano de Estudo:** Não informado")
                    
                    if pd.notna(plano_presenca_link) and str(plano_presenca_link).strip():
                        st.markdown(f"**📋 Link para o Plano de Presença:** [Acessar Documento]({plano_presenca_link})")
                    else:
                        st.markdown("**📋 Plano de Presença:** Não informado")
                    
                    parecer_options = ('Pendente', 'Deferido SG', 'Indeferido SG', 'Para análise COC.')
                    current_status = st.session_state.decisions[decision_key]['status']
                    if current_status not in parecer_options:
                        current_status = 'Pendente'

                    status = st.radio("Parecer:", parecer_options,
                                      key=f"status_{decision_key}",
                                      index=parecer_options.index(current_status),
                                      horizontal=True)
                    st.session_state.decisions[decision_key]['status'] = status

                    if status != 'Pendente':
                        label = "Justificativa (Opcional):"
                        if status == 'Deferido SG':
                            label = "Justificativa para o deferimento:"
                        elif status == 'Indeferido SG':
                            label = "Justificativa para o indeferimento:"
                        elif status == 'Para análise COC.':
                            label = "Observações para o COC:"

                        justificativa_input = st.text_area(
                            label,
                            value=st.session_state.decisions[decision_key]['justificativa'],
                            key=f"just_input_{decision_key}"
                        )

                        if st.button("Salvar Justificativa", key=f"save_btn_{decision_key}"):
                            st.session_state.decisions[decision_key]['justificativa'] = justificativa_input
                            st.toast("Justificativa salva!", icon="✔️")
                    else:
                        st.session_state.decisions[decision_key]['justificativa'] = ''
                    
                    st.divider()

            historico_aluno = df_merged[df_merged[COL_NUSP] == nusp_aluno].copy()
            st.markdown("##### 📜 Histórico Completo de Pedidos")
            
            if not historico_aluno.empty and not historico_aluno['disciplina_historico'].isnull().all():
                historico_aluno['problema_formatado'] = historico_aluno['problema_historico'].apply(format_problem_type)
                historico_aluno['parecer_formatado'] = historico_aluno['parecer_historico'].apply(format_parecer)
                cols_hist = ['disciplina_historico', 'Ano_historico', 'Semestre_historico', 'problema_formatado', 'parecer_formatado']
                df_hist_display = historico_aluno[cols_hist].rename(columns=lambda c: c.replace('_historico', '').replace('_formatado',''))
                st.dataframe(df_hist_display, hide_index=True, use_container_width=True)
            else:
                st.info("Este aluno não possui histórico de pedidos anteriores.")


# --- Funções de Exportação ---
def prepare_export_data(df_req, decisions):
    """Aplica as decisões do parecer ao DataFrame de requerimentos para exportação."""
    df_export = df_req.copy()

    if COL_PARECER_SG not in df_export.columns:
        df_export[COL_PARECER_SG] = ""
    if COL_OBSERVACAO_SG not in df_export.columns:
        df_export[COL_OBSERVACAO_SG] = ""

    # A chave de decisão agora é baseada apenas no índice
    df_export['decision_key'] = "req_" + df_export.index.astype(str)
    df_export['parecer_temp'] = df_export['decision_key'].map(lambda k: decisions.get(k, {}).get('status', 'Pendente'))
    df_export['justificativa_temp'] = df_export['decision_key'].map(lambda k: decisions.get(k, {}).get('justificativa', ''))

    mask = df_export['parecer_temp'] != 'Pendente'
    df_export.loc[mask, COL_PARECER_SG] = df_export.loc[mask, 'parecer_temp']
    df_export.loc[mask, COL_OBSERVACAO_SG] = df_export.loc[mask, 'justificativa_temp']

    df_export = df_export.drop(columns=['decision_key', 'parecer_temp', 'justificativa_temp'])
    
    return df_export

# --- Função Principal da Aplicação ---
def run_app():
    # --- Cabeçalho de Identidade Visual ---
    st.markdown("""
        <div class="header-container">
            <div class="logo-and-title">
                <img src="http://googleusercontent.com/file_content/2" class="header-logo">
                <div class="header-title-text">
                    <span class="university-name">Universidade de São Paulo</span>
                    <span class="department-name">Serviço de Graduação - FZEA</span>
                </div>
            </div>
            <div class="color-bar-container">
                <div class="color-bar-yellow"></div>
                <div class="color-bar-lightblue"></div>
                <div class="color-bar-darkblue"></div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<h1 class="main-header">📋 Sistema de Conferência de Requerimentos</h1>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("📁 Upload de Arquivos")
        file_consolidado = st.file_uploader("**Histórico de Pedidos (consolidado)**", type=["xlsx", "csv"])
        file_requerimentos = st.file_uploader("**Pedidos do Semestre Atual (requerimentos)**", type=["xlsx", "csv"])
        st.info("💡 Os arquivos devem ter uma coluna com o número USP.")
        with st.expander("⚙️ Configurações Avançadas"):
            show_debug = st.checkbox("Mostrar informações de debug", value=False)

    if not (file_consolidado and file_requerimentos):
        st.markdown("### 🚀 Bem-vindo! Para começar, faça o upload dos arquivos.")
        with st.expander("📋 Estrutura esperada dos arquivos"):
            st.markdown(f"**Consolidado:** `{', '.join(REQUIRED_COLS_CONSOLIDADO)}`")
            st.markdown(f"**Requerimentos:** `{', '.join(REQUIRED_COLS_REQUERIMENTOS)}`")
            st.markdown("> A coluna `link_requerimento` deve conter o link (geralmente da coluna G da planilha original).")
            st.markdown("> As colunas `plano_estudo` e `plano_presenca` devem conter os links para os respectivos documentos.")
        return

    try:
        with st.spinner("Processando arquivos..."):
            df_consolidado = load_data(file_consolidado)
            df_requerimentos = load_data(file_requerimentos)
            if df_consolidado is None or df_requerimentos is None: st.stop()

            if show_debug:
                with st.expander("🔍 Debug - Colunas originais"):
                    st.write("**Consolidado:**", df_consolidado.columns.tolist())
                    st.write("**Requerimentos:**", df_requerimentos.columns.tolist())
            
            possible_nusp = ["nusp", "numero usp", "número usp", "n° usp", "n usp"]
            df_consolidado = find_and_rename_columns(df_consolidado, COL_NUSP, possible_nusp, {COL_PROBLEMA: COL_PROBLEMA})
            df_requerimentos = find_and_rename_columns(df_requerimentos, COL_NUSP, possible_nusp, {
                COL_PROBLEMA: COL_PROBLEMA,
                "link para o requerimento": COL_LINK,
                "links pedidos requerimento": COL_LINK,
                "plano de estudo": COL_PLANO,
                "link plano de estudos": COL_PLANO,
                "plano de presença": COL_PLANO_PRESENCA,
                "link plano de presença": COL_PLANO_PRESENCA,
                "observação sg": COL_OBSERVACAO_SG
            })
            validate_dataframes(df_consolidado, df_requerimentos)
            
            df_consolidado = clean_nusp_column(df_consolidado, "consolidado")
            df_requerimentos = clean_nusp_column(df_requerimentos, "requerimentos")
            
            cols_hist = {c: f"{c}_historico" for c in [COL_DISCIPLINA, COL_ANO, COL_SEMESTRE, COL_PROBLEMA, COL_PARECER]}
            df_consolidado.rename(columns=cols_hist, inplace=True)
            df_requerimentos.rename(columns={COL_PROBLEMA: 'problema_atual'}, inplace=True)

            df_merged = df_requerimentos.merge(df_consolidado, on=COL_NUSP, how="left")
            
            df_merged_with_history = df_merged.dropna(subset=['disciplina_historico'])
            metrics = calculate_metrics(df_merged_with_history)

        display_metrics(df_requerimentos, df_merged_with_history, metrics)
        st.divider()
        if not df_merged_with_history.empty:
            display_charts(metrics)
        st.divider()
        display_student_details(df_requerimentos, df_merged)
        st.divider()

        st.markdown("### 📥 Exportar Relatórios")
        df_com_pareceres = prepare_export_data(df_requerimentos, st.session_state.decisions)
        
        df_nao_indeferidos = df_com_pareceres[df_com_pareceres[COL_PARECER_SG] != 'Indeferido SG'].copy()

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### Relatório Completo com Pareceres")
            st.download_button(label="📥 Baixar como Excel", data=to_excel(df_com_pareceres),
                                file_name=f"relatorio_completo_pareceres_{datetime.now().strftime('%Y%m%d')}.xlsx")
        with col2:
            st.markdown("##### Relatório de Pedidos Não Indeferidos")
            st.download_button(label="📥 Baixar como Excel", data=to_excel(df_nao_indeferidos),
                                file_name=f"relatorio_nao_indeferidos_{datetime.now().strftime('%Y%m%d')}.xlsx")

    except ValueError as e: st.error(f"❌ **Erro de Validação:**\n\n{e}")
    except Exception as e:
        st.error(f"❌ **Ocorreu um erro inesperado:**\n\n{e}")
        if show_debug: st.exception(e)

# --- Ponto de Entrada e Autenticação ---
# CORREÇÃO: Lógica de autenticação refatorada para evitar erro de form duplicado.

# Inicializa o estado da sessão se não existir
if "password_correct" not in st.session_state:
    st.session_state["password_correct"] = False

# Se o usuário não estiver logado, mostra o formulário de login
if not st.session_state["password_correct"]:
    st.title("🔒 Acesso Restrito")
    try:
        correct_password = st.secrets["passwords"]["senha_mestra"]
    except (AttributeError, KeyError):
        st.error("Aplicação não configurada. Contate o administrador.")
        st.info("Dev: Configure a senha em `secrets.toml`:\n\n```toml\n[passwords]\nsenha_mestra = \"sua_senha\"\n```")
    else:
        with st.form("login_form"):
            password = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar")
            if submitted:
                if password == correct_password:
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("Senha incorreta.")
# Se o usuário estiver logado, executa o aplicativo principal
else:
    run_app()

import plotly.express as px
from io import BytesIO
from datetime import datetime
from PyPDF2 import PdfReader

# --- Constantes ---
# Nomes de colunas para evitar "magic strings" e facilitar a manutenção
COL_NUSP = "nusp"
COL_PROBLEMA = "problema"
COL_PARECER = "parecer"
COL_NOME = "Nome completo"
COL_DISCIPLINA = "disciplina"
COL_ANO = "Ano"
COL_SEMESTRE = "Semestre"
COL_LINK = "link_requerimento"
COL_PLANO = "plano_estudo"
COL_PLANO_PRESENCA = "plano_presenca"
# Constantes para as colunas de exportação (J e K)
COL_PARECER_SG = "Parecer Serviço de Graduação"
COL_OBSERVACAO_SG = "Observação SG"


# Colunas obrigatórias nos arquivos de entrada
REQUIRED_COLS_CONSOLIDADO = [COL_NUSP, COL_DISCIPLINA, COL_ANO, COL_SEMESTRE, COL_PROBLEMA, COL_PARECER]
REQUIRED_COLS_REQUERIMENTOS = [COL_NUSP, COL_NOME, COL_PROBLEMA, COL_LINK, COL_PLANO, COL_PLANO_PRESENCA]

# --- Configuração da Página e Estado da Sessão ---
st.set_page_config(
    page_title="Sistema de Conferência de Requerimentos",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicializa o estado da sessão para armazenar as decisões
if 'decisions' not in st.session_state:
    st.session_state.decisions = {}

# --- CSS Customizado ---
st.markdown("""
<style>
    /* Estilos do Cabeçalho de Identidade Visual */
    .header-container {
        margin-bottom: 2rem;
    }
    .logo-and-title {
        display: flex;
        align-items: center;
        margin-bottom: 1rem;
    }
    .header-logo {
        height: 70px;
        margin-right: 20px;
    }
    .header-title-text {
        display: flex;
        flex-direction: column;
    }
    .university-name {
        font-size: 1.7rem;
        font-weight: bold;
        color: #333;
    }
    .department-name {
        font-size: 1.3rem;
        color: #0072b5;
    }
    .color-bar-yellow { height: 8px; background-color: #FDB913; }
    .color-bar-lightblue { height: 4px; background-color: #89cff0; }
    .color-bar-darkblue { height: 12px; background-color: #003366; }

    /* Estilo do cabeçalho principal do App */
    .main-header {
        font-size: 2.5rem; color: #1f77b4; text-align: center;
        padding: 1rem 0; border-bottom: 3px solid #1f77b4; margin-bottom: 2rem;
        margin-top: 0;
    }
    .metric-card {
        background-color: #f0f2f6; padding: 1.5rem; border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)


# --- Funções de Carregamento e Preparação de Dados ---

def load_data(uploaded_file):
    """Tenta ler um arquivo como Excel e, se falhar, tenta como CSV."""
    try:
        return pd.read_excel(uploaded_file)
    except Exception:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file)
        except Exception:
            st.error(f"Falha ao ler o arquivo '{uploaded_file.name}'. Verifique o formato.")
            return None

def find_and_rename_columns(df, target_col_name, possible_names, other_renames=None):
    """Encontra e renomeia colunas para um padrão definido, evitando duplicatas."""
    rename_dict = {}
    
    all_rules = {}
    if other_renames:
        for original, new in other_renames.items():
            all_rules.setdefault(new, []).append(original.lower().strip())
    
    all_rules.setdefault(target_col_name, []).extend([p.lower().strip() for p in possible_names])

    processed_original_cols = set()
    assigned_target_names = set()

    for col in df.columns:
        if col in processed_original_cols:
            continue
        
        normalized_col = col.lower().strip()
        
        found_match = False
        for target, originals in all_rules.items():
            if normalized_col in originals:
                if target not in assigned_target_names:
                    rename_dict[col] = target
                    processed_original_cols.add(col)
                    assigned_target_names.add(target)
                    found_match = True
                    break
        
        if not found_match and target_col_name == COL_NUSP:
            if any(keyword in normalized_col for keyword in ['nusp', 'numero usp', 'número usp', 'n° usp']):
                 if COL_NUSP not in assigned_target_names:
                    rename_dict[col] = COL_NUSP
                    processed_original_cols.add(col)
                    assigned_target_names.add(COL_NUSP)

    df.rename(columns=rename_dict, inplace=True)
    
    if target_col_name not in df.columns:
         raise ValueError(f"Coluna principal '{target_col_name}' não encontrada ou mapeada. Colunas disponíveis: {', '.join(df.columns.tolist())}")

    return df

def validate_dataframes(df_consolidado, df_requerimentos):
    """Verifica se os DataFrames contêm as colunas necessárias."""
    missing_consolidado = [col for col in REQUIRED_COLS_CONSOLIDADO if col not in df_consolidado.columns]
    missing_requerimentos = [col for col in REQUIRED_COLS_REQUERIMENTOS if col not in df_requerimentos.columns]
    errors = []
    if missing_consolidado: errors.append(f"Arquivo consolidado: colunas faltando - {', '.join(missing_consolidado)}")
    if missing_requerimentos: errors.append(f"Arquivo requerimentos: colunas faltando - {', '.join(missing_requerimentos)}")
    if errors: raise ValueError("\n".join(errors))

def clean_nusp_column(df, file_name):
    """Converte a coluna NUSP para numérico e remove registros inválidos."""
    if COL_NUSP not in df.columns: return df
    df[COL_NUSP] = pd.to_numeric(df[COL_NUSP], errors='coerce')
    invalid_count = df[COL_NUSP].isna().sum()
    if invalid_count > 0:
        st.warning(f"⚠️ Removidos {invalid_count} registros com NUSP inválido do arquivo {file_name}")
        df.dropna(subset=[COL_NUSP], inplace=True)
        df[COL_NUSP] = df[COL_NUSP].astype(int)
    return df

# --- Funções de Análise e Métricas ---

def calculate_metrics(df_merged_with_history):
    """Calcula métricas adicionais a partir do DataFrame com histórico."""
    metrics = {}
    if df_merged_with_history.empty: return metrics
    pareceres = df_merged_with_history['parecer_historico'].str.lower()
    aprovados = pareceres.str.contains('aprovado', na=False) & ~pareceres.str.contains('indeferido|negado', na=False)
    negados = pareceres.str.contains('indeferido|negado', na=False)
    total_com_parecer = aprovados.sum() + negados.sum()
    metrics['taxa_aprovacao'] = (aprovados.sum() / total_com_parecer * 100) if total_com_parecer > 0 else 0
    metrics['top_disciplinas'] = df_merged_with_history['disciplina_historico'].value_counts().head(5)
    if 'Ano_historico' in df_merged_with_history.columns and 'Semestre_historico' in df_merged_with_history.columns:
        df_merged_with_history['periodo'] = df_merged_with_history['Ano_historico'].astype(str) + '/' + df_merged_with_history['Semestre_historico'].astype(str)
        metrics['distribuicao_temporal'] = df_merged_with_history['periodo'].value_counts().sort_index()
    return metrics

# --- Funções de Formatação e Exibição (UI) ---

def format_parecer(parecer):
    """Formata o parecer para exibição com ícones."""
    if pd.isna(parecer): return "📝 Pendente"
    p_str = str(parecer).lower()
    if "negado" in p_str or "indeferido" in p_str: return f"❌ {parecer}"
    if "aprovado" in p_str: return f"✅ {parecer}"
    return f"📝 {parecer}"

def format_problem_type(problem):
    """Formata o tipo de problema para exibição."""
    if pd.isna(problem): return "⚪ Não especificado"
    p_str = str(problem).upper()
    if p_str == "QR": return "🔴 Quebra de Requisito"
    if p_str == "CH": return "🟡 Conflito de Horário"
    return f"⚪ {problem}"

@st.cache_data
def to_excel(df):
    """Converte um DataFrame para um arquivo Excel em memória."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
        worksheet = writer.sheets['Relatorio']
        header_format = writer.book.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BD', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for i, col in enumerate(df.columns):
            width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, min(width, 50))
    return output.getvalue()

def display_metrics(df_req, df_merged_with_history, metrics):
    """Exibe os cartões de métricas principais."""
    st.markdown("### 📊 Métricas Principais")
    cols = st.columns(5)
    with cols[0]: st.metric("Total de Requerimentos", len(df_req))
    alunos_unicos_hist = df_merged_with_history[COL_NUSP].nunique()
    total_alunos_req = df_req[COL_NUSP].nunique()
    percentual_hist = (alunos_unicos_hist / total_alunos_req * 100) if total_alunos_req > 0 else 0
    with cols[1]: st.metric("Alunos com Histórico", alunos_unicos_hist, f"{percentual_hist:.1f}%")
    with cols[2]: st.metric("Quebras de Requisito (Hist.)", (df_merged_with_history["problema_historico"].str.upper() == "QR").sum())
    with cols[3]: st.metric("Conflitos de Horário (Hist.)", (df_merged_with_history["problema_historico"].str.upper() == "CH").sum())
    with cols[4]: st.metric("Taxa de Aprovação (Hist.)", f"{metrics.get('taxa_aprovacao', 0):.1f}%")

def display_charts(metrics):
    """Exibe os gráficos de análise."""
    st.markdown("### 📈 Análise Gráfica dos Alunos com Histórico")
    col1, col2 = st.columns(2)
    if 'top_disciplinas' in metrics and not metrics['top_disciplinas'].empty:
        with col1:
            st.markdown("##### 📚 Top 5 Disciplinas")
            top_d = metrics['top_disciplinas']
            fig = px.bar(top_d, x=top_d.values, y=top_d.index, orientation='h', text=top_d.values)
            fig.update_layout(yaxis_title="Disciplina", xaxis_title="Nº de Pedidos", yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)
    if 'distribuicao_temporal' in metrics and not metrics['distribuicao_temporal'].empty:
        with col2:
            st.markdown("##### 🗓️ Pedidos por Período")
            dist_t = metrics['distribuicao_temporal']
            fig2 = px.line(dist_t, x=dist_t.index, y=dist_t.values, markers=True)
            fig2.update_layout(xaxis_title="Período", yaxis_title="Nº de Pedidos")
            st.plotly_chart(fig2, use_container_width=True)

def display_student_details(df_requerimentos, df_merged, uploaded_hes):
    """Exibe a seção interativa de detalhes por aluno."""
    st.markdown("### 📋 Análise de Requerimentos por Aluno")
    st.info("Clique no nome para ver o histórico e dar o parecer nos pedidos atuais.")
    
    alunos_unicos = df_requerimentos[[COL_NUSP, COL_NOME]].drop_duplicates(subset=[COL_NUSP]).sort_values(COL_NOME)

    for _, aluno in alunos_unicos.iterrows():
        nusp_aluno = aluno[COL_NUSP]
        with st.expander(f"👤 {aluno[COL_NOME]} (NUSP: {nusp_aluno})"):
            current_requests = df_requerimentos[df_requerimentos[COL_NUSP] == nusp_aluno]
            
            # --- NOVO: Exibição do Histórico Escolar em PDF ---
            he_file = uploaded_hes.get(nusp_aluno)
            if he_file:
                with st.expander("🔍 Visualizar Histórico Escolar (Texto Extraído)"):
                    try:
                        pdf_reader = PdfReader(he_file)
                        text = ""
                        for page in pdf_reader.pages:
                            text += page.extract_text() + "\n--- Próxima Página ---\n"
                        st.text_area("Conteúdo do HE", text, height=300, key=f"he_text_{nusp_aluno}")
                    except Exception as e:
                        st.error(f"Não foi possível ler o arquivo PDF do histórico. Erro: {e}")
            
            st.markdown("##### 📌 Requerimento(s) no Semestre Atual para Análise")
            if current_requests.empty:
                st.write("Nenhum requerimento encontrado para este aluno no arquivo atual.")
            else:
                for index, request in current_requests.iterrows():
                    link = request.get(COL_LINK, "")
                    plano_estudo_link = request.get(COL_PLANO, "")
                    plano_presenca_link = request.get(COL_PLANO_PRESENCA, "")
                    
                    decision_key = f"req_{index}"
                    st.session_state.decisions.setdefault(decision_key, {'status': 'Pendente', 'justificativa': ''})
                    
                    problema_display = request.get('problema_atual', 'Não especificado')
                    st.markdown(f"**Problema/Pedido:** `{problema_display}`")

                    if pd.notna(link) and str(link).strip():
                        st.markdown(f"**🔗 Link para o Requerimento:** [Acessar Link]({link})")
                    else:
                        st.markdown("**🔗 Link para o Requerimento:** Não informado")

                    if pd.notna(plano_estudo_link) and str(plano_estudo_link).strip():
                        st.markdown(f"**📄 Link para o Plano de Estudo:** [Acessar Documento]({plano_estudo_link})")
                    else:
                        st.markdown("**📄 Plano de Estudo:** Não informado")
                    
                    if pd.notna(plano_presenca_link) and str(plano_presenca_link).strip():
                        st.markdown(f"**📋 Link para o Plano de Presença:** [Acessar Documento]({plano_presenca_link})")
                    else:
                        st.markdown("**📋 Plano de Presença:** Não informado")
                    
                    parecer_options = ('Pendente', 'Deferido SG', 'Indeferido SG', 'Para análise COC.')
                    current_status = st.session_state.decisions[decision_key]['status']
                    if current_status not in parecer_options:
                        current_status = 'Pendente'

                    status = st.radio("Parecer:", parecer_options,
                                      key=f"status_{decision_key}",
                                      index=parecer_options.index(current_status),
                                      horizontal=True)
                    st.session_state.decisions[decision_key]['status'] = status

                    if status != 'Pendente':
                        label = "Justificativa (Opcional):"
                        if status == 'Deferido SG':
                            label = "Justificativa para o deferimento:"
                        elif status == 'Indeferido SG':
                            label = "Justificativa para o indeferimento:"
                        elif status == 'Para análise COC.':
                            label = "Observações para o COC:"

                        justificativa_input = st.text_area(
                            label,
                            value=st.session_state.decisions[decision_key]['justificativa'],
                            key=f"just_input_{decision_key}"
                        )

                        if st.button("Salvar Justificativa", key=f"save_btn_{decision_key}"):
                            st.session_state.decisions[decision_key]['justificativa'] = justificativa_input
                            st.toast("Justificativa salva!", icon="✔️")
                    else:
                        st.session_state.decisions[decision_key]['justificativa'] = ''
                    
                    st.divider()

            historico_aluno = df_merged[df_merged[COL_NUSP] == nusp_aluno].copy()
            st.markdown("##### 📜 Histórico de Pedidos de Requerimento")
            
            if not historico_aluno.empty and not historico_aluno['disciplina_historico'].isnull().all():
                historico_aluno['problema_formatado'] = historico_aluno['problema_historico'].apply(format_problem_type)
                historico_aluno['parecer_formatado'] = historico_aluno['parecer_historico'].apply(format_parecer)
                cols_hist = ['disciplina_historico', 'Ano_historico', 'Semestre_historico', 'problema_formatado', 'parecer_formatado']
                df_hist_display = historico_aluno[cols_hist].rename(columns=lambda c: c.replace('_historico', '').replace('_formatado',''))
                st.dataframe(df_hist_display, hide_index=True, use_container_width=True)
            else:
                st.info("Este aluno não possui histórico de pedidos de requerimento anteriores.")


# --- Funções de Exportação ---
def prepare_export_data(df_req, decisions):
    """Aplica as decisões do parecer ao DataFrame de requerimentos para exportação."""
    df_export = df_req.copy()

    if COL_PARECER_SG not in df_export.columns:
        df_export[COL_PARECER_SG] = ""
    if COL_OBSERVACAO_SG not in df_export.columns:
        df_export[COL_OBSERVACAO_SG] = ""

    df_export['decision_key'] = "req_" + df_export.index.astype(str)
    df_export['parecer_temp'] = df_export['decision_key'].map(lambda k: decisions.get(k, {}).get('status', 'Pendente'))
    df_export['justificativa_temp'] = df_export['decision_key'].map(lambda k: decisions.get(k, {}).get('justificativa', ''))

    mask = df_export['parecer_temp'] != 'Pendente'
    df_export.loc[mask, COL_PARECER_SG] = df_export.loc[mask, 'parecer_temp']
    df_export.loc[mask, COL_OBSERVACAO_SG] = df_export.loc[mask, 'justificativa_temp']

    df_export = df_export.drop(columns=['decision_key', 'parecer_temp', 'justificativa_temp'])
    
    return df_export

# --- Função Principal da Aplicação ---
def run_app():
    # --- Cabeçalho de Identidade Visual ---
    st.markdown("""
        <div class="header-container">
            <div class="logo-and-title">
                <img src="http://googleusercontent.com/file_content/2" class="header-logo">
                <div class="header-title-text">
                    <span class="university-name">Universidade de São Paulo</span>
                    <span class="department-name">Serviço de Graduação - FZEA</span>
                </div>
            </div>
            <div class="color-bar-container">
                <div class="color-bar-yellow"></div>
                <div class="color-bar-lightblue"></div>
                <div class="color-bar-darkblue"></div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<h1 class="main-header">📋 Sistema de Conferência de Requerimentos</h1>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("📁 Upload de Arquivos")
        file_consolidado = st.file_uploader("**1. Histórico de Pedidos (consolidado)**", type=["xlsx", "csv"])
        file_requerimentos = st.file_uploader("**2. Pedidos do Semestre Atual**", type=["xlsx", "csv"])
        files_he = st.file_uploader(
            "**3. Históricos Escolares (PDF)**", 
            type="pdf", 
            accept_multiple_files=True,
            help="Nomeie cada arquivo com o NUSP do aluno (ex: 12345678.pdf)"
        )
        st.info("💡 Os arquivos devem ter uma coluna com o número USP.")
        with st.expander("⚙️ Configurações Avançadas"):
            show_debug = st.checkbox("Mostrar informações de debug", value=False)

    if not (file_consolidado and file_requerimentos):
        st.markdown("### 🚀 Bem-vindo! Para começar, faça o upload dos arquivos 1 e 2.")
        with st.expander("📋 Estrutura esperada dos arquivos"):
            st.markdown(f"**Consolidado:** `{', '.join(REQUIRED_COLS_CONSOLIDADO)}`")
            st.markdown(f"**Requerimentos:** `{', '.join(REQUIRED_COLS_REQUERIMENTOS)}`")
            st.markdown("> As colunas `plano_estudo` e `plano_presenca` devem conter os links para os documentos.")
        return

    try:
        with st.spinner("Processando arquivos..."):
            df_consolidado = load_data(file_consolidado)
            df_requerimentos = load_data(file_requerimentos)
            if df_consolidado is None or df_requerimentos is None: st.stop()

            # --- NOVO: Processamento dos PDFs de Histórico Escolar ---
            uploaded_hes = {}
            if files_he:
                for file in files_he:
                    try:
                        nusp = int(file.name.split('.')[0])
                        uploaded_hes[nusp] = file
                    except (ValueError, IndexError):
                        st.warning(f"Não foi possível extrair o NUSP do nome do arquivo: {file.name}. O arquivo deve ser nomeado como 'NUMEROUSP.pdf'.")

            if show_debug:
                with st.expander("🔍 Debug - Colunas originais"):
                    st.write("**Consolidado:**", df_consolidado.columns.tolist())
                    st.write("**Requerimentos:**", df_requerimentos.columns.tolist())
            
            possible_nusp = ["nusp", "numero usp", "número usp", "n° usp", "n usp"]
            df_consolidado = find_and_rename_columns(df_consolidado, COL_NUSP, possible_nusp, {COL_PROBLEMA: COL_PROBLEMA})
            df_requerimentos = find_and_rename_columns(df_requerimentos, COL_NUSP, possible_nusp, {
                COL_PROBLEMA: COL_PROBLEMA,
                "link para o requerimento": COL_LINK,
                "links pedidos requerimento": COL_LINK,
                "plano de estudo": COL_PLANO,
                "link plano de estudos": COL_PLANO,
                "plano de presença": COL_PLANO_PRESENCA,
                "link plano de presença": COL_PLANO_PRESENCA,
                "observação sg": COL_OBSERVACAO_SG
            })
            validate_dataframes(df_consolidado, df_requerimentos)
            
            df_consolidado = clean_nusp_column(df_consolidado, "consolidado")
            df_requerimentos = clean_nusp_column(df_requerimentos, "requerimentos")
            
            cols_hist = {c: f"{c}_historico" for c in [COL_DISCIPLINA, COL_ANO, COL_SEMESTRE, COL_PROBLEMA, COL_PARECER]}
            df_consolidado.rename(columns=cols_hist, inplace=True)
            df_requerimentos.rename(columns={COL_PROBLEMA: 'problema_atual'}, inplace=True)

            df_merged = df_requerimentos.merge(df_consolidado, on=COL_NUSP, how="left")
            
            df_merged_with_history = df_merged.dropna(subset=['disciplina_historico'])
            metrics = calculate_metrics(df_merged_with_history)

        display_metrics(df_requerimentos, df_merged_with_history, metrics)
        st.divider()
        if not df_merged_with_history.empty:
            display_charts(metrics)
        st.divider()
        display_student_details(df_requerimentos, df_merged, uploaded_hes)
        st.divider()

        st.markdown("### 📥 Exportar Relatórios")
        df_com_pareceres = prepare_export_data(df_requerimentos, st.session_state.decisions)
        
        df_nao_indeferidos = df_com_pareceres[df_com_pareceres[COL_PARECER_SG] != 'Indeferido SG'].copy()

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### Relatório Completo com Pareceres")
            st.download_button(label="📥 Baixar como Excel", data=to_excel(df_com_pareceres),
                                file_name=f"relatorio_completo_pareceres_{datetime.now().strftime('%Y%m%d')}.xlsx")
        with col2:
            st.markdown("##### Relatório de Pedidos Não Indeferidos")
            st.download_button(label="📥 Baixar como Excel", data=to_excel(df_nao_indeferidos),
                                file_name=f"relatorio_nao_indeferidos_{datetime.now().strftime('%Y%m%d')}.xlsx")

    except ValueError as e: st.error(f"❌ **Erro de Validação:**\n\n{e}")
    except Exception as e:
        st.error(f"❌ **Ocorreu um erro inesperado:**\n\n{e}")
        if show_debug: st.exception(e)

# --- Ponto de Entrada e Autenticação ---
if "password_correct" not in st.session_state:
    st.session_state["password_correct"] = False

if not st.session_state["password_correct"]:
    st.title("� Acesso Restrito")
    try:
        correct_password = st.secrets["passwords"]["senha_mestra"]
    except (AttributeError, KeyError):
        st.error("Aplicação não configurada. Contate o administrador.")
        st.info("Dev: Configure a senha em `secrets.toml`:\n\n```toml\n[passwords]\nsenha_mestra = \"sua_senha\"\n```")
    else:
        with st.form("login_form"):
            password = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar")
            if submitted:
                if password == correct_password:
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("Senha incorreta.")
else:
    run_app()

