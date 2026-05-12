"""
app.py — Backend do Painel Comercial
Coloque na mesma pasta que: Dados.xlsx  e  XLS_VOLUME_PREÇO_PREVISTO.xlsx

Instalar dependências (uma única vez no terminal):
    pip install flask flask-cors pandas openpyxl

Rodar:
    python app.py
"""

from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import os, math

app = Flask(__name__)
CORS(app)

BASE = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────

def _safe(v):
    if v is None: return None
    try:
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
    except Exception: pass
    return v

MESES_PT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

def mesano_label(m):
    try:
        s = str(int(float(m)))
        return f"{MESES_PT[int(s[4:])-1]}/{s[:4]}"
    except Exception:
        return str(m)

# ─────────────────────────────────────────────────────────────
# CARGA DOS DADOS
# ─────────────────────────────────────────────────────────────

def carregar():
    # REALIZADO
    df = pd.read_excel(os.path.join(BASE, "Dados.xlsx"), header=1)
    df.columns = df.columns.str.strip()
    df["sku"]      = df["Produtos"].astype(str).str.extract(r"^(\d+)")[0].str.strip()
    df["produto"]  = df["Produtos"].astype(str).str.replace(r"^\d+-\s*", "", regex=True).str.strip()
    col_linha = next(c for c in df.columns if c.strip() == 'Linha')
    df["linha"]    = df[col_linha].astype(str).str.strip()
    df["uf"]       = df["UF"].astype(str).str.strip()
    df["vendedor"] = df["Vendedor"].astype(str).str.strip()
    df["gerente"]  = df["Gerente"].astype(str).str.strip()
    df["real_qtd"] = pd.to_numeric(df["Qtd Vend."],       errors="coerce").fillna(0)
    df["real_rl"]  = pd.to_numeric(df["Receita Líquida"], errors="coerce").fillna(0)
    df["real_preco"]= pd.to_numeric(df["Preço UN."],      errors="coerce").fillna(0)

    # PLANO
    xls    = pd.ExcelFile(os.path.join(BASE, "XLS_VOLUME_PREÇO_PREVISTO.xlsx"))
    df_vol = pd.read_excel(xls, "VOLUME PREVISTO")
    df_prc = pd.read_excel(xls, "PREÇO LÍQ. PREVISTO")

    meses_vol = [c for c in df_vol.columns if isinstance(c, (int,float)) and 200000 < c < 210000]
    meses_prc = [c for c in df_prc.columns if isinstance(c, (int,float)) and 200000 < c < 210000]

    # Volume wide→long
    id_vol = [c for c in df_vol.columns if c not in meses_vol]
    vol = df_vol.melt(id_vars=id_vol, value_vars=meses_vol, var_name="mesano", value_name="plano_qtd")
    vol.columns = [str(c).strip() for c in vol.columns]
    col_merc_vol = next((c for c in vol.columns if "MERC" in c.upper()), None)
    if col_merc_vol:
        vol = vol[vol[col_merc_vol].astype(str).str.upper().str.contains("INTERN", na=False)]

    # Preço wide→long
    id_prc = [c for c in df_prc.columns if c not in meses_prc]
    prc = df_prc.melt(id_vars=id_prc, value_vars=meses_prc, var_name="mesano", value_name="plano_preco")
    prc.columns = [str(c).strip() for c in prc.columns]
    if "MERCADO" in prc.columns:
        prc = prc[prc["MERCADO"].astype(str).str.upper().str.contains("INTERN", na=False)]

    vol["sku"]    = vol["CODIGO"].astype(str).str.strip()
    prc["sku"]    = prc["CODIGO"].astype(str).str.strip()
    vol["mesano"] = vol["mesano"].astype(str)
    prc["mesano"] = prc["mesano"].astype(str)
    vol["plano_qtd"]   = pd.to_numeric(vol["plano_qtd"],   errors="coerce").fillna(0)
    prc["plano_preco"] = pd.to_numeric(prc["plano_preco"], errors="coerce").fillna(0)

    plano = pd.merge(
        vol[["sku","mesano","plano_qtd"]],
        prc[["sku","mesano","plano_preco"]],
        on=["sku","mesano"], how="left"
    ).fillna(0)
    plano["plano_rl"] = plano["plano_qtd"] * plano["plano_preco"]

    mapa_linha   = df[["sku","linha"]].drop_duplicates().set_index("sku")["linha"].to_dict()
    mapa_produto = df[["sku","produto"]].drop_duplicates().set_index("sku")["produto"].to_dict()
    plano["linha"]   = plano["sku"].map(mapa_linha).fillna("Outros")
    plano["produto"] = plano["sku"].map(mapa_produto).fillna("")

    return df, plano


DF_REAL, DF_PLANO = carregar()

# ─────────────────────────────────────────────────────────────
# AGREGAÇÃO
# ─────────────────────────────────────────────────────────────

def agregar(dr, dp, nivel):
    rows = []

    if nivel == "linha":
        agg_r = dr.groupby("linha", as_index=False).agg(real_qtd=("real_qtd","sum"), real_rl=("real_rl","sum"))
        agg_r["real_preco"] = (agg_r["real_rl"] / agg_r["real_qtd"].replace(0, float("nan"))).fillna(0)

        agg_p = dp.groupby("linha", as_index=False).agg(plano_qtd=("plano_qtd","sum"), plano_rl=("plano_rl","sum"))
        agg_p["plano_preco"] = (agg_p["plano_rl"] / agg_p["plano_qtd"].replace(0, float("nan"))).fillna(0)

        merged = pd.merge(agg_p[["linha","plano_qtd","plano_preco"]], agg_r[["linha","real_qtd","real_preco"]], on="linha", how="outer").fillna(0)

        for _, row in merged.iterrows():
            pq=row["plano_qtd"]; pp=row["plano_preco"]
            rq=row["real_qtd"];  rp=row["real_preco"]
            vq=rq-pq; vp=rp-pp; gq=vq*pp; gp=vp*rq
            rows.append({"type":"linha","linha":str(row["linha"]),
                "plano_qtd":_safe(pq),"plano_preco":_safe(pp),
                "real_qtd":_safe(rq),"real_preco":_safe(rp),
                "var_qtd":_safe(vq),"var_preco":_safe(vp),
                "ganho_qtd":_safe(gq),"ganho_preco":_safe(gp)})

    else:  # sku
        agg_r = dr.groupby(["sku","produto","linha"], as_index=False).agg(real_qtd=("real_qtd","sum"), real_rl=("real_rl","sum"))
        agg_r["real_preco"] = (agg_r["real_rl"] / agg_r["real_qtd"].replace(0, float("nan"))).fillna(0)

        agg_p = dp.groupby(["sku","linha"], as_index=False).agg(plano_qtd=("plano_qtd","sum"), plano_rl=("plano_rl","sum"))
        agg_p["plano_preco"] = (agg_p["plano_rl"] / agg_p["plano_qtd"].replace(0, float("nan"))).fillna(0)

        merged = pd.merge(agg_p, agg_r, on=["sku","linha"], how="outer").fillna(0)
        if "produto" not in merged.columns:
            mapa = DF_REAL[["sku","produto"]].drop_duplicates().set_index("sku")["produto"].to_dict()
            merged["produto"] = merged["sku"].map(mapa).fillna("")

        for lin in sorted(merged["linha"].unique()):
            sub = merged[merged["linha"] == lin]
            pq=sub["plano_qtd"].sum(); prl=sub["plano_rl"].sum() if "plano_rl" in sub else 0
            pp=(prl/pq) if pq else 0
            rq=sub["real_qtd"].sum(); rrl=sub["real_rl"].sum() if "real_rl" in sub else 0
            rp=(rrl/rq) if rq else 0
            vq=rq-pq; vp=rp-pp; gq=vq*pp; gp=vp*rq
            rows.append({"type":"linha","linha":str(lin),
                "plano_qtd":_safe(pq),"plano_preco":_safe(pp),
                "real_qtd":_safe(rq),"real_preco":_safe(rp),
                "var_qtd":_safe(vq),"var_preco":_safe(vp),
                "ganho_qtd":_safe(gq),"ganho_preco":_safe(gp)})

            for _, row in sub.iterrows():
                pq2=row["plano_qtd"]; pp2=row["plano_preco"]
                rq2=row["real_qtd"];  rp2=row["real_preco"]
                vq2=rq2-pq2; vp2=rp2-pp2; gq2=vq2*pp2; gp2=vp2*rq2
                rows.append({"type":"sku","sku":str(row["sku"]),"produto":str(row.get("produto","")),"linha":str(lin),
                    "plano_qtd":_safe(pq2),"plano_preco":_safe(pp2),
                    "real_qtd":_safe(rq2),"real_preco":_safe(rp2),
                    "var_qtd":_safe(vq2),"var_preco":_safe(vp2),
                    "ganho_qtd":_safe(gq2),"ganho_preco":_safe(gp2)})

    return rows

# ─────────────────────────────────────────────────────────────
# ROTAS
# ─────────────────────────────────────────────────────────────

@app.route("/filtros")
def filtros():
    meses_raw = sorted(DF_PLANO["mesano"].dropna().unique().tolist())
    return jsonify({
        "vendedores": sorted(DF_REAL["vendedor"].dropna().unique().tolist()),
        "gerentes":   sorted(DF_REAL["gerente"].dropna().unique().tolist()),
        "ufs":        sorted(DF_REAL["uf"].dropna().unique().tolist()),
        "linhas":     sorted(DF_REAL["linha"].dropna().unique().tolist()),
        "skus":       sorted(DF_REAL["sku"].dropna().unique().tolist()),
        "meses":      [{"value": m, "label": mesano_label(m)} for m in meses_raw],
    })

@app.route("/dados")
def dados():
    nivel    = request.args.get("nivel",    "linha")
    vendedor = request.args.get("vendedor", "")
    gerente  = request.args.get("gerente",  "")
    uf       = request.args.get("uf",       "")
    linha    = request.args.get("linha",    "")
    sku      = request.args.get("sku",      "")
    mesano   = request.args.get("mesano",   "")

    dr = DF_REAL.copy()
    if vendedor: dr = dr[dr["vendedor"] == vendedor]
    if gerente:  dr = dr[dr["gerente"]  == gerente]
    if uf:       dr = dr[dr["uf"]       == uf]
    if linha:    dr = dr[dr["linha"]    == linha]
    if sku:      dr = dr[dr["sku"]      == sku]

    dp = DF_PLANO.copy()
    if mesano: dp = dp[dp["mesano"] == str(mesano)]
    if linha:  dp = dp[dp["linha"]  == linha]
    if sku:    dp = dp[dp["sku"]    == sku]

    rows = agregar(dr, dp, nivel)
    total_gp = sum(r["ganho_preco"] or 0 for r in rows if r["type"] == "linha")
    total_gq = sum(r["ganho_qtd"]   or 0 for r in rows if r["type"] == "linha")

    return jsonify({
        "rows":              rows,
        "total_ganho_preco": _safe(total_gp),
        "total_ganho_qtd":   _safe(total_gq),
    })

if __name__ == "__main__":
    print("\n✅  Servidor rodando em http://localhost:5000\n")
    app.run(debug=True, port=5000)