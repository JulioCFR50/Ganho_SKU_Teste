from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import numpy as np
import os

app = Flask(__name__)
CORS(app)

# ─────────────────────────────────────────────────────────────
# DADOS EXEMPLO (fallback caso os arquivos Excel não sejam encontrados)
# ─────────────────────────────────────────────────────────────

def criar_dados_exemplo():
    np.random.seed(42)

    vendedores = ['S00117-JULIO NETO', '000535-ANTONIO PEDRO', 'S00200-MARIA COSTA']
    gerentes   = ['NO | NE', 'SPC | SPC 2', 'MG | CO | ES | RJ']
    ufs        = ['SP', 'RJ', 'MG', 'BA']

    produtos_list = [
        '1001-LC ATG COCO DO VALE 24X200ML',
        '1002-LC LEITE DE COCO CDV GARRAFA 500ML',
        '1003-CR COCO RALADO SACHE CDV 100G',
        '2001-CR COCO RALADO PACOTE 200G',
        '7205-OC OLEO DE COCO DO VALE 200ML',
        '9013-AC AGUA DE COCO CDV 12X1L',
    ]

    linha_map = {
        '1001': 'Leite de Coco', '1002': 'Leite de Coco',
        '1003': 'Coco Ralado',   '2001': 'Coco Ralado',
        '7205': 'Óleo de Coco',  '9013': 'Água de Coco',
    }

    n     = 300
    prods = np.random.choice(produtos_list, n)
    cods  = [p.split('-')[0] for p in prods]

    data = {
        'Vendedor':        np.random.choice(vendedores, n),
        'Gerente':         np.random.choice(gerentes, n),
        'UF':              np.random.choice(ufs, n),
        'Produtos':        prods,
        'Linha':           [linha_map.get(c, 'Outros') for c in cods],
        'Qtd Vend.':       np.random.uniform(100, 1000, n),
        'Preço UN.':       np.random.uniform(5, 50, n),
        'Receita Líquida': np.random.uniform(500, 5000, n),
    }

    return pd.DataFrame(data)


def criar_plano_exemplo():
    np.random.seed(42)

    codigos   = ['1001', '1002', '1003', '2001', '7205', '9013']
    descricoes = [
        'ATG COCO DO VALE 24X200ML',
        'LEITE DE COCO CDV GARRAFA 500ML',
        'COCO RALADO SACHE CDV 100G',
        'COCO RALADO PACOTE 200G',
        'OLEO DE COCO DO VALE 200ML',
        'AGUA DE COCO CDV 12X1L',
    ]
    meses = [202601, 202602, 202603, 202604, 202605, 202606,
             202607, 202608, 202609, 202610, 202611, 202612]

    dados_vol   = {'CODIGO': codigos, 'DESCRICAO': descricoes, 'Mercador': ['INTERNO'] * len(codigos)}
    dados_preco = {'CODIGO': codigos, 'DESCRICAO': descricoes, 'MERCADO':  ['Interno'] * len(codigos)}

    for mes in meses:
        dados_vol[str(mes)]   = np.random.uniform(50, 500, len(codigos))
        dados_preco[str(mes)] = np.random.uniform(5,  40,  len(codigos))

    return pd.DataFrame(dados_vol), pd.DataFrame(dados_preco)


# ─────────────────────────────────────────────────────────────
# CARREGAR DADOS
# ─────────────────────────────────────────────────────────────

def carregar_dados():

    BASE = os.path.dirname(os.path.abspath(__file__))

    try:
        # FIX 1: header=1 — a linha 0 do Dados.xlsx é um resumo;
        #         os cabeçalhos reais estão na linha 1 (índice 1).
        df_dados = pd.read_excel(
            os.path.join(BASE, 'Dados.xlsx'),
            header=1
        )

        # FIX 2: não existe Linha_produto.xlsx separado;
        #         o arquivo de volume/preço tem as duas abas necessárias.
        df_vol = pd.read_excel(
            os.path.join(BASE, 'XLS_VOLUME_PREÇO_PREVISTO.xlsx'),
            sheet_name='VOLUME PREVISTO'
        )

        df_preco = pd.read_excel(
            os.path.join(BASE, 'XLS_VOLUME_PREÇO_PREVISTO.xlsx'),
            sheet_name='PREÇO LÍQ. PREVISTO'
        )

        print('✅ Excel carregado com sucesso')

    except Exception as e:

        print(f'⚠️  Usando dados de exemplo: {e}')
        df_dados        = criar_dados_exemplo()
        df_vol, df_preco = criar_plano_exemplo()

    # ─────────────────────────────────────────
    # LIMPAR COLUNAS
    # ─────────────────────────────────────────

    df_dados.columns  = df_dados.columns.str.strip()
    df_vol.columns    = df_vol.columns.astype(str).str.strip()
    df_preco.columns  = df_preco.columns.astype(str).str.strip()

    # ─────────────────────────────────────────
    # FILTRAR MERCADO INTERNO
    # ─────────────────────────────────────────

    if 'Mercado' in df_dados.columns:
        df_dados = df_dados[df_dados['Mercado'] == 'Mercado Interno'].copy()

    # ─────────────────────────────────────────
    # EXTRAIR CÓDIGO DO PRODUTO
    # ─────────────────────────────────────────

    df_dados['cod_prod'] = (
        df_dados['Produtos']
        .astype(str)
        .str.extract(r'^(\d+)')[0]
        .fillna('0')
        .str.strip()
    )

    # ─────────────────────────────────────────
    # NUMÉRICOS
    # ─────────────────────────────────────────

    for col in ['Qtd Vend.', 'Preço UN.', 'Receita Líquida']:
        if col in df_dados.columns:
            df_dados[col] = pd.to_numeric(df_dados[col], errors='coerce').fillna(0)

    # ─────────────────────────────────────────
    # MAPEAMENTO cod_prod → Linha
    # FIX 3: derivado do próprio Dados.xlsx (não precisa de arquivo externo)
    # ─────────────────────────────────────────

    df_linha = (
        df_dados[['cod_prod', 'Linha']]
        .drop_duplicates(subset='cod_prod')
        .dropna()
        .rename(columns={'cod_prod': 'CODIGO'})
    )
    df_linha['CODIGO'] = df_linha['CODIGO'].astype(str).str.strip()

    # ─────────────────────────────────────────
    # PLANO: filtrar INTERNO
    # ─────────────────────────────────────────

    df_vol_int   = df_vol[df_vol['Mercador'] == 'INTERNO'].copy()
    df_preco_int = df_preco[df_preco['MERCADO'] == 'Interno'].copy()

    df_vol_int['CODIGO']   = df_vol_int['CODIGO'].astype(str).str.strip()
    df_preco_int['CODIGO'] = df_preco_int['CODIGO'].astype(str).str.strip()

    # Colunas de mês (inteiros no Excel como 202601, 202602 …)
    month_cols = [
        c for c in df_vol_int.columns
        if str(c).isdigit() and len(str(c)) == 6
    ]

    # Melt volume
    df_vol_m = df_vol_int.melt(
        id_vars=['CODIGO', 'DESCRICAO'],
        value_vars=month_cols,
        var_name='mes',
        value_name='vol_plano'
    )

    # Melt preço
    df_preco_m = df_preco_int.melt(
        id_vars=['CODIGO'],
        value_vars=month_cols,
        var_name='mes',
        value_name='preco_plano'
    )

    df_vol_m['mes']           = df_vol_m['mes'].astype(int)
    df_preco_m['mes']         = df_preco_m['mes'].astype(int)
    df_vol_m['vol_plano']     = pd.to_numeric(df_vol_m['vol_plano'],     errors='coerce').fillna(0)
    df_preco_m['preco_plano'] = pd.to_numeric(df_preco_m['preco_plano'], errors='coerce').fillna(0)

    plano = df_vol_m.merge(df_preco_m, on=['CODIGO', 'mes'], how='inner')
    plano['receita_plano'] = plano['vol_plano'] * plano['preco_plano']

    # FIX 4: join Linha usando mapeamento derivado do Dados.xlsx
    plano = plano.merge(df_linha, on='CODIGO', how='left')
    plano['Linha'] = plano['Linha'].fillna('Outros')

    return df_dados, plano


df_dados, plano = carregar_dados()


# ─────────────────────────────────────────────────────────────
# AUXILIAR
# ─────────────────────────────────────────────────────────────

def safe_div(a, b):
    return a / b if b != 0 else 0


def fmt_mes(m):
    """Converte 202601 → {'value': 202601, 'label': 'Jan/2026'}"""
    s     = str(int(m))
    nomes = ['Jan','Fev','Mar','Abr','Mai','Jun',
             'Jul','Ago','Set','Out','Nov','Dez']
    return {
        'value': int(m),
        'label': f"{nomes[int(s[4:])-1]}/{s[:4]}"
    }


# ─────────────────────────────────────────────────────────────
# ENDPOINT FILTROS
# ─────────────────────────────────────────────────────────────

@app.route('/filtros', methods=['GET'])
def filtros():

    meses = sorted(plano['mes'].unique().tolist())

    # FIX 5: retorna objetos {value, label} em vez de inteiros puros
    return jsonify({
        'vendedores': sorted(df_dados['Vendedor'].dropna().unique().tolist()),
        'gerentes':   sorted(df_dados['Gerente'].dropna().unique().tolist()),
        'ufs':        sorted(df_dados['UF'].dropna().unique().tolist()),
        'meses':      [fmt_mes(m) for m in meses],
    })


# ─────────────────────────────────────────────────────────────
# ENDPOINT DADOS
# ─────────────────────────────────────────────────────────────

@app.route('/dados', methods=['GET'])
def dados():

    vendedor = request.args.get('vendedor')
    gerente  = request.args.get('gerente')
    uf       = request.args.get('uf')
    mesano   = request.args.get('mesano')
    nivel    = request.args.get('nivel', 'linha')   # 'linha' ou 'sku'

    df_r = df_dados.copy()
    pl   = plano.copy()

    if vendedor:
        df_r = df_r[df_r['Vendedor'] == vendedor]

    if gerente:
        df_r = df_r[df_r['Gerente'] == gerente]

    if uf:
        df_r = df_r[df_r['UF'] == uf]

    if mesano:
        pl = pl[pl['mes'] == int(mesano)]

    # ─────────────────────────────────────────
    # FIX 6: suporte ao nível SKU (Por Produto)
    # ─────────────────────────────────────────

    if nivel == 'sku':

        # REAL agrupado por produto
        real = df_r.groupby(['cod_prod', 'Produtos', 'Linha']).agg(
            qtd_real=('Qtd Vend.',       'sum'),
            receita_real=('Receita Líquida', 'sum')
        ).reset_index()

        real['preco_real'] = real.apply(
            lambda x: safe_div(x['receita_real'], x['qtd_real']), axis=1
        )
        real = real.rename(columns={'cod_prod': 'CODIGO'})

        # PLANO agrupado por produto
        plan = pl.groupby(['CODIGO', 'DESCRICAO', 'Linha']).agg(
            qtd_plan=('vol_plano',     'sum'),
            receita_plan=('receita_plano', 'sum')
        ).reset_index()

        plan['preco_plan'] = plan.apply(
            lambda x: safe_div(x['receita_plan'], x['qtd_plan']), axis=1
        )
        plan['nome'] = plan['CODIGO'].astype(str) + ' - ' + plan['DESCRICAO'].astype(str)

        # Merge por CODIGO
        final = plan[['CODIGO','nome','Linha','qtd_plan','preco_plan']].merge(
            real[['CODIGO','qtd_real','preco_real']],
            on='CODIGO',
            how='outer'
        ).fillna(0)

        nome_col = 'nome'

    else:

        # REAL por Linha
        real = df_r.groupby('Linha').agg(
            qtd_real=('Qtd Vend.',       'sum'),
            receita_real=('Receita Líquida', 'sum')
        ).reset_index()

        real['preco_real'] = real.apply(
            lambda x: safe_div(x['receita_real'], x['qtd_real']), axis=1
        )

        # PLANO por Linha
        plan = pl.groupby('Linha').agg(
            qtd_plan=('vol_plano',     'sum'),
            receita_plan=('receita_plano', 'sum')
        ).reset_index()

        plan['preco_plan'] = plan.apply(
            lambda x: safe_div(x['receita_plan'], x['qtd_plan']), axis=1
        )

        final    = plan.merge(real, on='Linha', how='outer').fillna(0)
        nome_col = 'Linha'

    # ─────────────────────────────────────────
    # MONTAR LINHAS DE RESPOSTA
    # ─────────────────────────────────────────

    rows = []

    for _, r in final.iterrows():

        nome      = str(r.get(nome_col, '-'))
        qtd_plan  = float(r.get('qtd_plan',  0))
        qtd_real  = float(r.get('qtd_real',  0))
        preco_plan = float(r.get('preco_plan', 0))
        preco_real = float(r.get('preco_real', 0))

        var_qtd   = qtd_real  - qtd_plan
        var_preco = preco_real - preco_plan
        ganho_qtd   = var_qtd   * preco_plan
        ganho_preco = var_preco * qtd_real

        rows.append({
            'linha':       nome,
            'plano_qtd':   round(qtd_plan,   2),
            'plano_preco': round(preco_plan,  2),
            'real_qtd':    round(qtd_real,    2),
            'real_preco':  round(preco_real,  2),
            'var_qtd':     round(var_qtd,     2),
            'var_preco':   round(var_preco,   2),
            'ganho_qtd':   round(ganho_qtd,   2),
            'ganho_preco': round(ganho_preco, 2),
        })

    total_ganho_preco = sum(r['ganho_preco'] for r in rows)
    total_ganho_qtd   = sum(r['ganho_qtd']   for r in rows)

    return jsonify({
        'rows':              rows,
        'total_ganho_preco': round(total_ganho_preco, 2),
        'total_ganho_qtd':   round(total_ganho_qtd,   2),
    })


# ─────────────────────────────────────────────────────────────
# START
# ─────────────────────────────────────────────────────────────

if __name__ == '__main__':

    print('🚀 Servidor iniciado')
    print('🌐 http://localhost:5000')

    app.run(debug=True, host='0.0.0.0', port=5000)
