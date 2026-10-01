import streamlit as st
import pandas as pd
import sqlite3
from datetime import date, datetime, timedelta
import io
import os
import re
import time
from streamlit.runtime import runtime
import subprocess
import threading

# ==========================================
# AUTO-FECHAMENTO (opcional - mesmo padrão do original)
# ==========================================
def monitorar_abas_ativas():
    time.sleep(8)
    while True:
        time.sleep(2)
        try:
            instancia_rt = runtime.get_instance()
            if instancia_rt:
                abas = instancia_rt._session_mgr.list_active_sessions()
                if len(abas) == 0:
                    pid_atual = os.getpid()
                    pid_pai = os.getppid()
                    subprocess.run(f"taskkill /F /PID {pid_pai}", shell=True, capture_output=True)
                    subprocess.run(f"taskkill /F /PID {pid_atual}", shell=True, capture_output=True)
                    os._exit(0)
        except Exception:
            pass

if "monitor_auto_close" not in st.session_state:
    threading.Thread(target=monitorar_abas_ativas, daemon=True).start()
    st.session_state["monitor_auto_close"] = True

DB_PATH = 'estoque_mercado.db'

# ==========================================
# 1. BACKUP E INICIALIZAÇÃO DO BANCO
# ==========================================
def realizar_backup_seguranca():
    if os.path.exists(DB_PATH):
        try:
            with sqlite3.connect(DB_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='produtos'")
                if cursor.fetchone():
                    df_produtos = pd.read_sql_query("SELECT * FROM produtos", conn)
                    df_movimentacoes = pd.read_sql_query("SELECT * FROM movimentacoes", conn)
                    
                    if not df_produtos.empty or not df_movimentacoes.empty:
                        if not os.path.exists('backups_automaticos'):
                            os.makedirs('backups_automaticos')
                        
                        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                        caminho_backup = f"backups_automaticos/Backup_Estoque_{timestamp}.xlsx"
                        
                        with pd.ExcelWriter(caminho_backup, engine='openpyxl') as writer:
                            df_produtos.to_excel(writer, index=False, sheet_name='Produtos')
                            df_movimentacoes.to_excel(writer, index=False, sheet_name='Movimentacoes')
        except Exception:
            pass

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        
        c.execute('''CREATE TABLE IF NOT EXISTS produtos 
                     (id INTEGER PRIMARY KEY,
                      codigo TEXT UNIQUE,
                      nome TEXT NOT NULL,
                      categoria TEXT,
                      unidade TEXT,
                      estoque_atual REAL DEFAULT 0,
                      estoque_minimo REAL DEFAULT 0,
                      preco_custo REAL DEFAULT 0,
                      preco_venda REAL DEFAULT 0,
                      fornecedor TEXT,
                      data_cadastro TEXT)''')
        
        c.execute('''CREATE TABLE IF NOT EXISTS movimentacoes 
                     (id INTEGER PRIMARY KEY,
                      produto_id INTEGER,
                      produto_nome TEXT,
                      tipo TEXT,                    -- ENTRADA | SAIDA | AJUSTE
                      quantidade REAL,
                      valor_unitario REAL,
                      valor_total REAL,
                      motivo TEXT,
                      usuario TEXT,
                      data_movimento TEXT,
                      observacao TEXT,
                      FOREIGN KEY (produto_id) REFERENCES produtos (id))''')
        
        conn.commit()

if 'backup_inicial_rodado' not in st.session_state:
    realizar_backup_seguranca()
    st.session_state['backup_inicial_rodado'] = True

init_db()

# ==========================================
# 2. FUNÇÕES AUXILIARES
# ==========================================
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def gerar_codigo_produto():
    return f"P{datetime.now().strftime('%y%m%d%H%M%S')}"

# ==========================================
# 3. INTERFACE
# ==========================================
st.set_page_config(
    page_title="Estoque Pro - Gestão de Mercado",
    page_icon="📦",
    layout="wide"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    
    html, body, [data-testid="stSidebar"] {
        font-family: 'Inter', sans-serif;
    }
    
    .header-banner {
        background: linear-gradient(135deg, #1e3a5f 0%, #0f172a 50%, #020617 100%);
        padding: 36px 40px;
        border-radius: 16px;
        color: white;
        margin-bottom: 28px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 8px 30px rgba(15, 23, 42, 0.25);
    }
    .header-banner::after {
        content: '';
        position: absolute;
        right: -8%; top: -40%;
        width: 380px; height: 380px;
        background: rgba(59, 130, 246, 0.12);
        border-radius: 50%;
        pointer-events: none;
    }
    .header-banner h1 { color: #FFFFFF !important; font-weight: 700; margin: 0; font-size: 2.1rem; }
    .header-banner p { color: #93C5FD !important; margin: 6px 0 0 0; opacity: 0.95; font-size: 1.05rem; }
    
    .ag-card {
        background: white;
        padding: 20px 22px;
        border-radius: 14px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04);
        border: 1px solid #f1f5f9;
        transition: transform 0.25s ease;
    }
    .ag-card:hover { transform: translateY(-3px); }
    .ag-title { font-size: 11px; font-weight: 700; letter-spacing: 0.8px; color: #64748B; margin-bottom: 6px; text-transform: uppercase; }
    .ag-value { font-size: 24px; font-weight: 700; margin: 0; }
    
    .ficha-produto { 
        background: #ffffff; 
        padding: 22px; 
        border-radius: 12px; 
        border: 1px solid #e2e8f0; 
        box-shadow: 0 4px 18px rgba(0,0,0,0.04);
        margin-bottom: 18px; 
    }
    
    .alerta-exclusao { 
        background-color: #fff5f5; 
        padding: 14px; 
        border-left: 5px solid #e53e3e; 
        border-radius: 6px; 
        margin-top: 10px; 
    }
    
    .badge-baixo {
        background: #FEE2E2;
        color: #B91C1C;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    .badge-ok {
        background: #DCFCE7;
        color: #166534;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    </style>
    """, unsafe_allow_html=True)

# Sidebar
menu = st.sidebar.selectbox(
    "🗺️ Navegação",
    ["Dashboard", "Cadastrar Produto", "Consultar Produtos", 
     "Entrada de Estoque", "Saída de Estoque", "Ajuste de Inventário", "Relatórios"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("🔔 Alertas de Estoque")

with sqlite3.connect(DB_PATH) as conn:
    df_alertas = pd.read_sql_query(
        "SELECT nome, estoque_atual, estoque_minimo FROM produtos WHERE estoque_atual <= estoque_minimo AND estoque_minimo > 0",
        conn
    )

if not df_alertas.empty:
    st.sidebar.error(f"🚨 **{len(df_alertas)}** produto(s) abaixo do mínimo")
    with st.sidebar.expander("Ver produtos em risco"):
        for _, row in df_alertas.iterrows():
            st.write(f"• **{row['nome']}** → {row['estoque_atual']} (mín: {row['estoque_minimo']})")
else:
    st.sidebar.success("✅ Nenhum produto em estoque crítico")

# Banner
st.markdown("""
    <div class="header-banner">
        <h1>📦 Estoque Pro</h1>
        <p>Sistema de Controle de Estoque para Mercados e Pequenos Negócios</p>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# DASHBOARD
# ==========================================
if menu == "Dashboard":
    st.markdown("### 📊 Visão Geral do Estoque")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_prod = pd.read_sql_query("SELECT * FROM produtos", conn)
        df_mov = pd.read_sql_query("SELECT * FROM movimentacoes ORDER BY data_movimento DESC LIMIT 50", conn)
    
    if df_prod.empty:
        st.info("💡 Nenhum produto cadastrado ainda. Comece cadastrando seus primeiros itens.")
    else:
        total_produtos = len(df_prod)
        valor_estoque_custo = (df_prod['estoque_atual'] * df_prod['preco_custo']).sum()
        valor_estoque_venda = (df_prod['estoque_atual'] * df_prod['preco_venda']).sum()
        produtos_baixo = len(df_prod[(df_prod['estoque_atual'] <= df_prod['estoque_minimo']) & (df_prod['estoque_minimo'] > 0)])
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid #3B82F6;">
                    <div class="ag-title">Produtos Cadastrados</div>
                    <div class="ag-value" style="color: #1E40AF;">{total_produtos}</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid #0F766E; background-color: #F0FDFA;">
                    <div class="ag-title" style="color: #0F766E;">Valor em Custo</div>
                    <div class="ag-value" style="color: #0F766E;">{formatar_moeda(valor_estoque_custo)}</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid #D97706; background-color: #FFFBEB;">
                    <div class="ag-title" style="color: #B45309;">Valor Potencial de Venda</div>
                    <div class="ag-value" style="color: #B45309;">{formatar_moeda(valor_estoque_venda)}</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col4:
            border_c = "#DC2626" if produtos_baixo > 0 else "#94A3B8"
            bg_c = "#FEF2F2" if produtos_baixo > 0 else "#FFFFFF"
            text_c = "#DC2626" if produtos_baixo > 0 else "#64748B"
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid {border_c}; background-color: {bg_c};">
                    <div class="ag-title" style="color: {text_c};">Estoque Crítico</div>
                    <div class="ag-value" style="color: {text_c};">{produtos_baixo}</div>
                </div>
            """, unsafe_allow_html=True)
        
        st.write("")
        st.caption(f"💰 **Margem potencial estimada:** {formatar_moeda(valor_estoque_venda - valor_estoque_custo)}")
        
        st.write("---")
        
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("##### 📦 Distribuição por Categoria")
            if not df_prod.empty:
                cat_count = df_prod.groupby('categoria').size().reset_index(name='Qtd')
                cat_count = cat_count.set_index('categoria')
                st.bar_chart(cat_count, height=300, color="#3B82F6")
        
        with col_g2:
            st.markdown("##### 🔄 Últimas Movimentações")
            if not df_mov.empty:
                df_mov_show = df_mov[['data_movimento', 'produto_nome', 'tipo', 'quantidade']].head(12)
                df_mov_show.columns = ['Data', 'Produto', 'Tipo', 'Qtd']
                st.dataframe(df_mov_show, use_container_width=True, hide_index=True, height=300)
            else:
                st.info("Nenhuma movimentação registrada.")

# ==========================================
# CADASTRAR PRODUTO
# ==========================================
elif menu == "Cadastrar Produto":
    st.header("➕ Novo Produto no Estoque")
    
    with st.form("form_produto"):
        col1, col2 = st.columns(2)
        
        with col1:
            codigo = st.text_input("Código / SKU / Código de Barras", 
                                   value=gerar_codigo_produto(),
                                   help="Deixe o gerado ou digite o código de barras")
            nome = st.text_input("Nome do Produto *")
            categoria = st.selectbox("Categoria", 
                ["Alimentos", "Bebidas", "Limpeza", "Higiene", "Padaria", 
                 "Hortifruti", "Açougue", "Frios", "Outros"])
            unidade = st.selectbox("Unidade de Medida", 
                ["UN", "KG", "L", "CX", "PCT", "M"])
        
        with col2:
            estoque_atual = st.number_input("Estoque Inicial", min_value=0.0, step=1.0, value=0.0)
            estoque_minimo = st.number_input("Estoque Mínimo (alerta)", min_value=0.0, step=1.0, value=5.0)
            preco_custo = st.number_input("Preço de Custo (R$)", min_value=0.0, step=0.50, format="%.2f")
            preco_venda = st.number_input("Preço de Venda (R$)", min_value=0.0, step=0.50, format="%.2f")
            fornecedor = st.text_input("Fornecedor principal")
        
        submit = st.form_submit_button("Cadastrar Produto", type="primary")
        
        if submit:
            if not nome:
                st.error("O nome do produto é obrigatório.")
            else:
                try:
                    with sqlite3.connect(DB_PATH) as conn:
                        c = conn.cursor()
                        c.execute("""INSERT INTO produtos 
                                     (codigo, nome, categoria, unidade, estoque_atual, estoque_minimo, 
                                      preco_custo, preco_venda, fornecedor, data_cadastro)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                  (codigo, nome.strip(), categoria, unidade, estoque_atual, 
                                   estoque_minimo, preco_custo, preco_venda, fornecedor, 
                                   datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                        conn.commit()
                        
                        # Se tiver estoque inicial, registra movimentação
                        if estoque_atual > 0:
                            prod_id = c.lastrowid
                            c.execute("""INSERT INTO movimentacoes 
                                         (produto_id, produto_nome, tipo, quantidade, valor_unitario, 
                                          valor_total, motivo, usuario, data_movimento, observacao)
                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                      (prod_id, nome.strip(), "ENTRADA", estoque_atual, preco_custo,
                                       estoque_atual * preco_custo, "Estoque Inicial", "Sistema",
                                       datetime.now().strftime('%Y-%m-%d %H:%M:%S'), "Cadastro inicial"))
                            conn.commit()
                    
                    st.success(f"✅ Produto **{nome}** cadastrado com sucesso!")
                    st.balloons()
                except sqlite3.IntegrityError:
                    st.error("Já existe um produto com este código. Altere o código e tente novamente.")

# ==========================================
# CONSULTAR / EDITAR PRODUTOS
# ==========================================
elif menu == "Consultar Produtos":
    st.header("📋 Cadastro de Produtos")
    
    with sqlite3.connect(DB_PATH) as conn:
        df = pd.read_sql_query("""
            SELECT id, codigo AS [Código], nome AS [Produto], categoria AS [Categoria],
                   unidade AS [Un.], estoque_atual AS [Estoque], estoque_minimo AS [Mínimo],
                   preco_custo AS [Custo], preco_venda AS [Venda], fornecedor AS [Fornecedor]
            FROM produtos ORDER BY nome ASC
        """, conn)
    
    if df.empty:
        st.info("Nenhum produto cadastrado.")
    else:
        # Filtros rápidos
        col_f1, col_f2 = st.columns([2, 1])
        busca = col_f1.text_input("🔍 Buscar por nome ou código", placeholder="Digite para filtrar...")
        filtro_cat = col_f2.selectbox("Categoria", ["Todas"] + sorted(df['Categoria'].unique().tolist()))
        
        df_view = df.copy()
        if busca:
            df_view = df_view[
                df_view['Produto'].str.contains(busca, case=False, na=False) |
                df_view['Código'].str.contains(busca, case=False, na=False)
            ]
        if filtro_cat != "Todas":
            df_view = df_view[df_view['Categoria'] == filtro_cat]
        
        # Destacar estoque baixo
        def destacar_estoque(row):
            if row['Estoque'] <= row['Mínimo'] and row['Mínimo'] > 0:
                return ['background-color: #FEE2E2'] * len(row)
            return [''] * len(row)
        
        st.dataframe(
            df_view.style.apply(destacar_estoque, axis=1).format({
                'Custo': 'R$ {:.2f}',
                'Venda': 'R$ {:.2f}',
                'Estoque': '{:.1f}',
                'Mínimo': '{:.1f}'
            }),
            use_container_width=True,
            hide_index=True
        )
        
        st.write("---")
        st.subheader("✏️ Editar ou Excluir Produto")
        
        produto_sel = st.selectbox(
            "Selecione o produto",
            ["— Escolha —"] + df['Produto'].tolist()
        )
        
        if produto_sel != "— Escolha —":
            info = df[df['Produto'] == produto_sel].iloc[0]
            prod_id = int(info['id'])
            
            st.markdown(f"""
            <div class="ficha-produto">
                <h3 style='margin-top:0; color:#1e3a5f;'>📦 {info['Produto']}</h3>
                <p><strong>Código:</strong> {info['Código']} &nbsp;|&nbsp; 
                   <strong>Categoria:</strong> {info['Categoria']} &nbsp;|&nbsp; 
                   <strong>Unidade:</strong> {info['Un.']}</p>
                <p><strong>Estoque atual:</strong> {info['Estoque']:.1f} &nbsp;|&nbsp; 
                   <strong>Mínimo:</strong> {info['Mínimo']:.1f}</p>
                <p><strong>Custo:</strong> {formatar_moeda(info['Custo'])} &nbsp;|&nbsp; 
                   <strong>Venda:</strong> {formatar_moeda(info['Venda'])}</p>
                <p><strong>Fornecedor:</strong> {info['Fornecedor'] or '—'}</p>
            </div>
            """, unsafe_allow_html=True)
            
            with st.expander("Atualizar dados do produto"):
                with st.form("form_edicao"):
                    novo_nome = st.text_input("Nome", value=info['Produto'])
                    nova_cat = st.selectbox("Categoria", 
                        ["Alimentos", "Bebidas", "Limpeza", "Higiene", "Padaria", 
                         "Hortifruti", "Açougue", "Frios", "Outros"],
                        index=["Alimentos", "Bebidas", "Limpeza", "Higiene", "Padaria", 
                               "Hortifruti", "Açougue", "Frios", "Outros"].index(info['Categoria']) 
                              if info['Categoria'] in ["Alimentos", "Bebidas", "Limpeza", "Higiene", "Padaria", 
                                                       "Hortifruti", "Açougue", "Frios", "Outros"] else 8)
                    nova_un = st.selectbox("Unidade", ["UN", "KG", "L", "CX", "PCT", "M"],
                                           index=["UN", "KG", "L", "CX", "PCT", "M"].index(info['Un.']) 
                                           if info['Un.'] in ["UN", "KG", "L", "CX", "PCT", "M"] else 0)
                    novo_min = st.number_input("Estoque Mínimo", value=float(info['Mínimo']), step=1.0)
                    novo_custo = st.number_input("Preço de Custo", value=float(info['Custo']), step=0.50, format="%.2f")
                    novo_venda = st.number_input("Preço de Venda", value=float(info['Venda']), step=0.50, format="%.2f")
                    novo_forn = st.text_input("Fornecedor", value=info['Fornecedor'] or "")
                    
                    if st.form_submit_button("Salvar alterações"):
                        with sqlite3.connect(DB_PATH) as conn:
                            c = conn.cursor()
                            c.execute("""UPDATE produtos SET nome=?, categoria=?, unidade=?, 
                                         estoque_minimo=?, preco_custo=?, preco_venda=?, fornecedor=?
                                         WHERE id=?""",
                                      (novo_nome, nova_cat, nova_un, novo_min, novo_custo, 
                                       novo_venda, novo_forn, prod_id))
                            # Atualiza nome nas movimentações também
                            c.execute("UPDATE movimentacoes SET produto_nome=? WHERE produto_id=?",
                                      (novo_nome, prod_id))
                            conn.commit()
                        st.success("Produto atualizado!")
                        st.rerun()
            
            # Exclusão
            st.markdown("### ⚠️ Excluir produto")
            chave = f"conf_del_{prod_id}"
            if chave not in st.session_state:
                st.session_state[chave] = False
            
            if not st.session_state[chave]:
                if st.button("🗑️ Excluir este produto", type="primary"):
                    st.session_state[chave] = True
                    st.rerun()
            else:
                st.markdown(f'<div class="alerta-exclusao"><strong>Atenção:</strong> Esta ação remove o produto e todo o histórico de movimentações relacionadas de forma permanente.</div>', unsafe_allow_html=True)
                dig = st.text_input(f"Digite o nome exato para confirmar ({info['Produto']}):")
                
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("CONFIRMAR EXCLUSÃO", type="primary", disabled=(dig != info['Produto'])):
                        with sqlite3.connect(DB_PATH) as conn:
                            c = conn.cursor()
                            c.execute("DELETE FROM movimentacoes WHERE produto_id=?", (prod_id,))
                            c.execute("DELETE FROM produtos WHERE id=?", (prod_id,))
                            conn.commit()
                        st.session_state[chave] = False
                        st.rerun()
                with c2:
                    if st.button("Cancelar"):
                        st.session_state[chave] = False
                        st.rerun()

# ==========================================
# ENTRADA DE ESTOQUE
# ==========================================
elif menu == "Entrada de Estoque":
    st.header("📥 Entrada de Mercadoria (Compra / Recebimento)")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_prod = pd.read_sql_query("SELECT id, codigo, nome, estoque_atual, preco_custo, unidade FROM produtos ORDER BY nome", conn)
    
    if df_prod.empty:
        st.warning("Cadastre produtos antes de registrar entradas.")
    else:
        with st.form("form_entrada"):
            produto_nome = st.selectbox("Produto", df_prod['nome'].tolist())
            info = df_prod[df_prod['nome'] == produto_nome].iloc[0]
            
            st.caption(f"Estoque atual: **{info['estoque_atual']:.1f} {info['unidade']}** | Custo cadastrado: {formatar_moeda(info['preco_custo'])}")
            
            col1, col2, col3 = st.columns(3)
            qtd = col1.number_input(f"Quantidade ({info['unidade']})", min_value=0.01, step=1.0, value=1.0)
            valor_unit = col2.number_input("Valor unitário de custo (R$)", min_value=0.0, 
                                           value=float(info['preco_custo']), step=0.50, format="%.2f")
            motivo = col3.selectbox("Motivo", ["Compra", "Devolução de cliente", "Transferência", "Outros"])
            
            observacao = st.text_input("Observação (opcional)")
            
            if st.form_submit_button("Registrar Entrada", type="primary"):
                novo_estoque = info['estoque_atual'] + qtd
                valor_total = qtd * valor_unit
                
                with sqlite3.connect(DB_PATH) as conn:
                    c = conn.cursor()
                    # Atualiza estoque e opcionalmente o custo médio (aqui atualiza o preço de custo)
                    c.execute("UPDATE produtos SET estoque_atual=?, preco_custo=? WHERE id=?",
                              (novo_estoque, valor_unit, int(info['id'])))
                    
                    c.execute("""INSERT INTO movimentacoes 
                                 (produto_id, produto_nome, tipo, quantidade, valor_unitario, 
                                  valor_total, motivo, usuario, data_movimento, observacao)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                              (int(info['id']), produto_nome, "ENTRADA", qtd, valor_unit,
                               valor_total, motivo, "Operador", 
                               datetime.now().strftime('%Y-%m-%d %H:%M:%S'), observacao))
                    conn.commit()
                
                st.success(f"✅ Entrada de **{qtd} {info['unidade']}** registrada. Novo estoque: **{novo_estoque:.1f}**")
                st.rerun()

# ==========================================
# SAÍDA DE ESTOQUE
# ==========================================
elif menu == "Saída de Estoque":
    st.header("📤 Saída / Baixa de Estoque (Venda ou Consumo)")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_prod = pd.read_sql_query(
            "SELECT id, codigo, nome, estoque_atual, preco_venda, unidade FROM produtos WHERE estoque_atual > 0 ORDER BY nome", 
            conn
        )
    
    if df_prod.empty:
        st.warning("Não há produtos com estoque disponível.")
    else:
        with st.form("form_saida"):
            produto_nome = st.selectbox("Produto", df_prod['nome'].tolist())
            info = df_prod[df_prod['nome'] == produto_nome].iloc[0]
            
            st.caption(f"Estoque disponível: **{info['estoque_atual']:.1f} {info['unidade']}** | Preço de venda: {formatar_moeda(info['preco_venda'])}")
            
            col1, col2, col3 = st.columns(3)
            qtd = col1.number_input(f"Quantidade ({info['unidade']})", 
                                    min_value=0.01, 
                                    max_value=float(info['estoque_atual']),
                                    step=1.0, value=1.0)
            valor_unit = col2.number_input("Valor unitário (R$)", min_value=0.0, 
                                           value=float(info['preco_venda']), step=0.50, format="%.2f")
            motivo = col3.selectbox("Motivo", ["Venda", "Consumo interno", "Perda / Avaria", "Doação", "Outros"])
            
            observacao = st.text_input("Observação (opcional)")
            
            if st.form_submit_button("Registrar Saída", type="primary"):
                if qtd > info['estoque_atual']:
                    st.error("Quantidade maior que o estoque disponível!")
                else:
                    novo_estoque = info['estoque_atual'] - qtd
                    valor_total = qtd * valor_unit
                    
                    with sqlite3.connect(DB_PATH) as conn:
                        c = conn.cursor()
                        c.execute("UPDATE produtos SET estoque_atual=? WHERE id=?",
                                  (novo_estoque, int(info['id'])))
                        
                        c.execute("""INSERT INTO movimentacoes 
                                     (produto_id, produto_nome, tipo, quantidade, valor_unitario, 
                                      valor_total, motivo, usuario, data_movimento, observacao)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                  (int(info['id']), produto_nome, "SAIDA", qtd, valor_unit,
                                   valor_total, motivo, "Operador", 
                                   datetime.now().strftime('%Y-%m-%d %H:%M:%S'), observacao))
                        conn.commit()
                    
                    st.success(f"✅ Saída de **{qtd} {info['unidade']}** registrada. Novo estoque: **{novo_estoque:.1f}**")
                    st.rerun()

# ==========================================
# AJUSTE DE INVENTÁRIO
# ==========================================
elif menu == "Ajuste de Inventário":
    st.header("⚖️ Ajuste de Inventário (Contagem Física)")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_prod = pd.read_sql_query(
            "SELECT id, codigo, nome, estoque_atual, unidade FROM produtos ORDER BY nome", 
            conn
        )
    
    if df_prod.empty:
        st.info("Nenhum produto cadastrado.")
    else:
        with st.form("form_ajuste"):
            produto_nome = st.selectbox("Produto", df_prod['nome'].tolist())
            info = df_prod[df_prod['nome'] == produto_nome].iloc[0]
            
            st.info(f"Estoque atual no sistema: **{info['estoque_atual']:.1f} {info['unidade']}**")
            
            novo_qtd = st.number_input(
                f"Quantidade contada fisicamente ({info['unidade']})",
                min_value=0.0, step=1.0, value=float(info['estoque_atual'])
            )
            
            diferenca = novo_qtd - info['estoque_atual']
            if diferenca != 0:
                tipo_msg = "aumento" if diferenca > 0 else "redução"
                st.warning(f"Diferença detectada: **{diferenca:+.1f}** ({tipo_msg})")
            
            observacao = st.text_input("Motivo do ajuste / Observação", 
                                       placeholder="Ex: Contagem física mensal, correção de erro...")
            
            if st.form_submit_button("Aplicar Ajuste", type="primary"):
                with sqlite3.connect(DB_PATH) as conn:
                    c = conn.cursor()
                    c.execute("UPDATE produtos SET estoque_atual=? WHERE id=?",
                              (novo_qtd, int(info['id'])))
                    
                    c.execute("""INSERT INTO movimentacoes 
                                 (produto_id, produto_nome, tipo, quantidade, valor_unitario, 
                                  valor_total, motivo, usuario, data_movimento, observacao)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                              (int(info['id']), produto_nome, "AJUSTE", abs(diferenca), 0,
                               0, "Ajuste de inventário", "Operador", 
                               datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 
                               f"{observacao} | Diferença: {diferenca:+.1f}"))
                    conn.commit()
                
                st.success(f"✅ Estoque ajustado para **{novo_qtd:.1f} {info['unidade']}**")
                st.rerun()

# ==========================================
# RELATÓRIOS
# ==========================================
elif menu == "Relatórios":
    st.header("📊 Relatórios e Exportação")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_prod = pd.read_sql_query("SELECT * FROM produtos", conn)
        df_mov = pd.read_sql_query("SELECT * FROM movimentacoes ORDER BY data_movimento DESC", conn)
    
    tab1, tab2, tab3 = st.tabs(["📦 Posição de Estoque", "🔄 Movimentações", "⚠️ Estoque Crítico"])
    
    with tab1:
        st.subheader("Posição atual de estoque")
        if not df_prod.empty:
            df_pos = df_prod[['codigo', 'nome', 'categoria', 'unidade', 'estoque_atual', 
                              'estoque_minimo', 'preco_custo', 'preco_venda']].copy()
            df_pos['Valor Custo'] = df_pos['estoque_atual'] * df_pos['preco_custo']
            df_pos['Valor Venda'] = df_pos['estoque_atual'] * df_pos['preco_venda']
            df_pos.columns = ['Código', 'Produto', 'Categoria', 'Un.', 'Estoque', 'Mínimo',
                              'Custo Unit.', 'Venda Unit.', 'Valor em Custo', 'Valor em Venda']
            
            st.dataframe(df_pos.style.format({
                'Custo Unit.': 'R$ {:.2f}',
                'Venda Unit.': 'R$ {:.2f}',
                'Valor em Custo': 'R$ {:.2f}',
                'Valor em Venda': 'R$ {:.2f}',
                'Estoque': '{:.1f}',
                'Mínimo': '{:.1f}'
            }), use_container_width=True, hide_index=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_pos.to_excel(writer, index=False, sheet_name='Posicao_Estoque')
            st.download_button(
                "📥 Exportar Posição de Estoque (Excel)",
                data=buffer.getvalue(),
                file_name=f"Posicao_Estoque_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.info("Sem dados.")
    
    with tab2:
        st.subheader("Histórico de movimentações")
        
        col_f1, col_f2 = st.columns(2)
        tipo_filtro = col_f1.multiselect("Tipo", ["ENTRADA", "SAIDA", "AJUSTE"], default=["ENTRADA", "SAIDA", "AJUSTE"])
        
        if not df_mov.empty:
            df_f = df_mov[df_mov['tipo'].isin(tipo_filtro)].copy()
            df_show = df_f[['data_movimento', 'produto_nome', 'tipo', 'quantidade', 
                            'valor_unitario', 'valor_total', 'motivo', 'observacao']]
            df_show.columns = ['Data', 'Produto', 'Tipo', 'Qtd', 'Valor Unit.', 'Total', 'Motivo', 'Obs.']
            
            st.dataframe(df_show, use_container_width=True, hide_index=True)
            
            buffer2 = io.BytesIO()
            with pd.ExcelWriter(buffer2, engine='openpyxl') as writer:
                df_show.to_excel(writer, index=False, sheet_name='Movimentacoes')
            st.download_button(
                "📥 Exportar Movimentações (Excel)",
                data=buffer2.getvalue(),
                file_name=f"Movimentacoes_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Nenhuma movimentação registrada.")
    
    with tab3:
        st.subheader("Produtos abaixo do estoque mínimo")
        if not df_prod.empty:
            df_crit = df_prod[
                (df_prod['estoque_atual'] <= df_prod['estoque_minimo']) & 
                (df_prod['estoque_minimo'] > 0)
            ][['codigo', 'nome', 'categoria', 'estoque_atual', 'estoque_minimo', 'fornecedor']]
            
            if not df_crit.empty:
                df_crit.columns = ['Código', 'Produto', 'Categoria', 'Estoque Atual', 'Mínimo', 'Fornecedor']
                st.dataframe(df_crit, use_container_width=True, hide_index=True)
                
                buffer3 = io.BytesIO()
                with pd.ExcelWriter(buffer3, engine='openpyxl') as writer:
                    df_crit.to_excel(writer, index=False, sheet_name='Estoque_Critico')
                st.download_button(
                    "📥 Exportar Lista de Críticos (Excel)",
                    data=buffer3.getvalue(),
                    file_name=f"Estoque_Critico_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("Nenhum produto em situação crítica no momento.")
        else:
            st.info("Sem dados.")