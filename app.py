from __future__ import annotations

import hashlib
import io
import os
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "assistencia.db"
FILES_DIR = APP_DIR / "anexos"
DEFAULT_XLSX = APP_DIR / "CONTROLE ASSISNTECIA TÉC.xlsx"
FILES_DIR.mkdir(exist_ok=True)

STATUS_OPTIONS = ["⚪ Aberto", "🔵 Agendado", "🟡 Em atendimento", "🔴 Atrasado", "✅ Concluído", "⛔ Cancelado"]
PRIORIDADE_OPTIONS = ["Baixa", "Média", "Alta", "Urgente"]
PERFIS = ["Administrador", "Comercial", "Técnico", "Consulta"]
SLA_DIAS = 4

st.set_page_config(page_title="Jirau | Assistência Técnica", page_icon="🛠️", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
:root { --jirau:#b00020; --jirau2:#7c0016; --dark:#171717; --soft:#f4f6f8; }
.stApp { background:var(--soft); }
[data-testid="stSidebar"] { background:linear-gradient(180deg,var(--jirau2),var(--jirau)); }
[data-testid="stSidebar"] * { color:white !important; }
.hero { background:linear-gradient(135deg,#790014,#c7002b); padding:24px 28px; border-radius:20px; color:white; margin-bottom:18px; box-shadow:0 10px 30px rgba(176,0,32,.2); }
.hero h1 { margin:0; font-size:31px; }.hero p{margin:6px 0 0;opacity:.92}
.kpi { background:white; border-radius:16px; padding:17px; box-shadow:0 5px 18px rgba(0,0,0,.06); border-top:4px solid var(--jirau); min-height:94px; }
.kpi-label { color:#6b7280; font-size:12px; text-transform:uppercase; letter-spacing:.04em; }
.kpi-value { color:#111827; font-size:28px; font-weight:850; margin-top:4px; }
.section { background:white; padding:16px; border-radius:16px; box-shadow:0 4px 16px rgba(0,0,0,.05); }
.badge { display:inline-block; padding:5px 10px; border-radius:999px; background:#f3f4f6; font-weight:700; font-size:12px; }
.stButton>button,.stDownloadButton>button {border-radius:10px;font-weight:750}
div[data-testid="stDataFrame"] {border-radius:14px;overflow:hidden}
</style>
""", unsafe_allow_html=True)


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def ensure_column(conn, table: str, column: str, definition: str) -> None:
    cols = {r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def init_db() -> None:
    with connect() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS assistencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT, numero INTEGER UNIQUE, entrada TEXT NOT NULL,
            cliente TEXT NOT NULL, endereco TEXT, produto TEXT, vendedor TEXT, tecnico TEXT,
            previsao_visita TEXT, data_visita TEXT, status TEXT NOT NULL, servico TEXT,
            observacoes TEXT, prioridade TEXT DEFAULT 'Média', criado_em TEXT NOT NULL, atualizado_em TEXT NOT NULL)""")
        for col, definition in [
            ("telefone", "TEXT"), ("email", "TEXT"), ("responsavel_cliente", "TEXT"),
            ("causa", "TEXT"), ("solucao", "TEXT"), ("assinatura_nome", "TEXT"),
            ("hora_inicio", "TEXT"), ("hora_fim", "TEXT")]:
            ensure_column(conn, "assistencias", col, definition)
        conn.execute("""CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT, assistencia_id INTEGER NOT NULL, data_hora TEXT NOT NULL,
            evento TEXT NOT NULL, observacao TEXT, usuario TEXT, FOREIGN KEY(assistencia_id) REFERENCES assistencias(id))""")
        ensure_column(conn, "historico", "usuario", "TEXT")
        conn.execute("""CREATE TABLE IF NOT EXISTS anexos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, assistencia_id INTEGER NOT NULL, nome TEXT NOT NULL,
            tipo TEXT, caminho TEXT NOT NULL, categoria TEXT, enviado_em TEXT NOT NULL,
            FOREIGN KEY(assistencia_id) REFERENCES assistencias(id))""")
        conn.execute("""CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, login TEXT UNIQUE NOT NULL,
            senha_hash TEXT NOT NULL, perfil TEXT NOT NULL, ativo INTEGER DEFAULT 1)""")
        if conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0:
            conn.execute("INSERT INTO usuarios(nome,login,senha_hash,perfil,ativo) VALUES(?,?,?,?,1)",
                         ("Administrador", "admin", hash_password("admin123"), "Administrador"))
        conn.commit()


def excel_date(value) -> Optional[str]:
    if pd.isna(value) or value in ("", None): return None
    try:
        if isinstance(value, (int, float)):
            return (pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")).date().isoformat()
        return pd.to_datetime(value, dayfirst=True).date().isoformat()
    except Exception: return None


def auto_status(previsao, visita, manual: str) -> str:
    if manual == "⛔ Cancelado": return manual
    if visita: return "✅ Concluído"
    if previsao and pd.Timestamp(previsao) < pd.Timestamp(date.today()): return "🔴 Atrasado"
    if previsao and manual == "⚪ Aberto": return "🔵 Agendado"
    return manual


def load_excel_to_db(file_obj) -> tuple[int, int]:
    df = pd.read_excel(file_obj, sheet_name="Registros", header=1)
    df.columns = [str(c).strip().upper() for c in df.columns]
    inserted = updated = 0
    now = datetime.now().isoformat(timespec="seconds")
    with connect() as conn:
        for _, row in df.iterrows():
            n = pd.to_numeric(row.get("Nº ASSISTÊNCIA"), errors="coerce")
            if pd.isna(n): continue
            numero = int(n)
            visita = excel_date(row.get("DATA DA VISTA")) or excel_date(row.get("DATA DA VISITA"))
            previsao = excel_date(row.get("PREVISÃO VISITA"))
            data = dict(numero=numero, entrada=excel_date(row.get("ENTRADA")) or date.today().isoformat(),
                        cliente=str(row.get("CLIENTE", "") or "").strip(), endereco=str(row.get("ENDEREÇO", "") or "").strip(),
                        produto=str(row.get("PRODUTO", "") or "").strip(), vendedor=str(row.get("VENDEDOR", "") or "").strip(),
                        tecnico=str(row.get("ASS. TÉCNICO", "") or "").strip(), previsao_visita=previsao,
                        data_visita=visita, status=auto_status(previsao, visita, "⚪ Aberto"),
                        servico=str(row.get("SERVIÇO", "") or "").strip(), observacoes=str(row.get("OBSERVAÇÕES", "") or "").strip(),
                        prioridade="Média")
            existing = conn.execute("SELECT id FROM assistencias WHERE numero=?", (numero,)).fetchone()
            if existing:
                conn.execute("""UPDATE assistencias SET entrada=:entrada,cliente=:cliente,endereco=:endereco,produto=:produto,
                    vendedor=:vendedor,tecnico=:tecnico,previsao_visita=:previsao_visita,data_visita=:data_visita,status=:status,
                    servico=:servico,observacoes=:observacoes,prioridade=:prioridade,atualizado_em=:atualizado_em WHERE numero=:numero""",
                    {**data, "atualizado_em": now}); updated += 1
            else:
                conn.execute("""INSERT INTO assistencias(numero,entrada,cliente,endereco,produto,vendedor,tecnico,previsao_visita,
                    data_visita,status,servico,observacoes,prioridade,criado_em,atualizado_em)
                    VALUES(:numero,:entrada,:cliente,:endereco,:produto,:vendedor,:tecnico,:previsao_visita,:data_visita,:status,
                    :servico,:observacoes,:prioridade,:criado_em,:atualizado_em)""", {**data,"criado_em":now,"atualizado_em":now}); inserted += 1
        conn.commit()
    return inserted, updated


def get_df() -> pd.DataFrame:
    with connect() as conn: df = pd.read_sql_query("SELECT * FROM assistencias ORDER BY entrada DESC,numero DESC", conn)
    if df.empty: return df
    for c in ["entrada","previsao_visita","data_visita"]: df[c] = pd.to_datetime(df[c], errors="coerce")
    hoje = pd.Timestamp(date.today())
    df["dias_atendimento"] = (df["data_visita"].fillna(hoje)-df["entrada"]).dt.days.clip(lower=0)
    df["status_exibicao"] = df.apply(lambda r: auto_status(r["previsao_visita"], r["data_visita"], r["status"]), axis=1)
    df["sla_ok"] = df["dias_atendimento"] <= SLA_DIAS
    return df


def next_number() -> int:
    with connect() as conn: m = conn.execute("SELECT MAX(numero) FROM assistencias").fetchone()[0]
    return int(m or 65000) + 1


def add_history(record_id: int, evento: str, obs: str = "") -> None:
    with connect() as conn:
        conn.execute("INSERT INTO historico(assistencia_id,data_hora,evento,observacao,usuario) VALUES(?,?,?,?,?)",
                     (record_id, datetime.now().isoformat(timespec="seconds"), evento, obs, st.session_state.get("user_name","Sistema")))
        conn.commit()


def save_record(data: dict, record_id: Optional[int] = None) -> int:
    now = datetime.now().isoformat(timespec="seconds")
    fields = ["numero","entrada","cliente","endereco","produto","vendedor","tecnico","previsao_visita","data_visita",
              "status","servico","observacoes","prioridade","telefone","email","responsavel_cliente","causa","solucao",
              "assinatura_nome","hora_inicio","hora_fim"]
    payload = {k:data.get(k) for k in fields}
    with connect() as conn:
        if record_id:
            sets = ",".join(f"{k}=:{k}" for k in fields)
            conn.execute(f"UPDATE assistencias SET {sets},atualizado_em=:atualizado_em WHERE id=:id", {**payload,"atualizado_em":now,"id":record_id})
            rid = record_id; event="Registro atualizado"
        else:
            cols = ",".join(fields+["criado_em","atualizado_em"]); vals = ",".join(":"+x for x in fields+["criado_em","atualizado_em"])
            cur=conn.execute(f"INSERT INTO assistencias({cols}) VALUES({vals})", {**payload,"criado_em":now,"atualizado_em":now})
            rid=cur.lastrowid; event="Assistência criada"
        conn.commit()
    add_history(rid,event,data.get("observacoes", "")); return rid


def create_os_pdf(rec: pd.Series) -> bytes:
    buf=io.BytesIO(); doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=15*mm,leftMargin=15*mm,topMargin=14*mm,bottomMargin=14*mm)
    styles=getSampleStyleSheet(); styles.add(ParagraphStyle(name="CenterTitle",parent=styles["Title"],alignment=TA_CENTER,textColor=colors.HexColor("#B00020")))
    story=[Paragraph("JIRAU ENGENHARIA",styles["CenterTitle"]),Paragraph("ORDEM DE SERVIÇO - ASSISTÊNCIA TÉCNICA",styles["Heading2"]),Spacer(1,6*mm)]
    def val(k):
        v=rec.get(k,"");
        if pd.isna(v): return "-"
        if isinstance(v,pd.Timestamp): return v.strftime("%d/%m/%Y")
        return str(v) if str(v).strip() else "-"
    data=[["Nº",val("numero"),"Entrada",val("entrada")],["Cliente",val("cliente"),"Responsável",val("responsavel_cliente")],
          ["Telefone",val("telefone"),"E-mail",val("email")],["Obra / Endereço",val("endereco"),"Prioridade",val("prioridade")],
          ["Produto",val("produto"),"Vendedor",val("vendedor")],["Técnico",val("tecnico"),"Status",val("status_exibicao")],
          ["Previsão",val("previsao_visita"),"Data da visita",val("data_visita")],["Início",val("hora_inicio"),"Fim",val("hora_fim")]]
    t=Table(data,colWidths=[28*mm,58*mm,32*mm,58*mm]); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.5,colors.grey),("BACKGROUND",(0,0),(0,-1),colors.HexColor("#f1f1f1")),("BACKGROUND",(2,0),(2,-1),colors.HexColor("#f1f1f1")),("VALIGN",(0,0),(-1,-1),"TOP"),("FONTNAME",(0,0),(-1,-1),"Helvetica"),("FONTSIZE",(0,0),(-1,-1),9),("PADDING",(0,0),(-1,-1),5)])); story += [t,Spacer(1,6*mm)]
    for title,key in [("Serviço solicitado","servico"),("Causa identificada","causa"),("Solução executada","solucao"),("Observações","observacoes")]:
        story += [Paragraph(title,styles["Heading3"]),Paragraph(val(key).replace("\n","<br/>"),styles["BodyText"]),Spacer(1,4*mm)]
    story += [Spacer(1,8*mm),Table([["________________________________","________________________________"],["Técnico responsável",f"Cliente: {val('assinatura_nome')}"]],colWidths=[85*mm,85*mm],style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("FONTSIZE",(0,0),(-1,-1),9)]))]
    doc.build(story); return buf.getvalue()


def export_excel(df: pd.DataFrame) -> bytes:
    out=io.BytesIO(); exp=df.copy()
    with pd.ExcelWriter(out,engine="openpyxl") as w: exp.to_excel(w,index=False,sheet_name="Assistências")
    return out.getvalue()


init_db()
if DEFAULT_XLSX.exists():
    with connect() as conn: count=conn.execute("SELECT COUNT(*) FROM assistencias").fetchone()[0]
    if count==0: load_excel_to_db(DEFAULT_XLSX)

# Login
if "logged" not in st.session_state: st.session_state.logged=False
if not st.session_state.logged:
    st.markdown('<div class="hero"><h1>Jirau Assistência Técnica</h1><p>Acesso seguro ao sistema de atendimento.</p></div>',unsafe_allow_html=True)
    c1,c2,c3=st.columns([1,1.2,1])
    with c2:
        with st.form("login"):
            login=st.text_input("Usuário",value="admin"); senha=st.text_input("Senha",type="password",value="admin123")
            ok=st.form_submit_button("Entrar",use_container_width=True)
        if ok:
            with connect() as conn: u=conn.execute("SELECT * FROM usuarios WHERE login=? AND senha_hash=? AND ativo=1",(login,hash_password(senha))).fetchone()
            if u:
                st.session_state.update(logged=True,user_name=u["nome"],perfil=u["perfil"]); st.rerun()
            else: st.error("Usuário ou senha inválidos.")
    st.stop()

st.sidebar.markdown("## 🛠️ JIRAU"); st.sidebar.caption(f"{st.session_state.user_name} · {st.session_state.perfil}")
page=st.sidebar.radio("Navegação",["Dashboard","Assistências","Nova assistência","Agenda","Importar / Exportar","Usuários","Configurações"])
if st.sidebar.button("Sair",use_container_width=True): st.session_state.clear(); st.rerun()
st.sidebar.caption("Versão 2.0 Premium")
st.markdown('<div class="hero"><h1>Assistência Técnica</h1><p>Chamados, visitas, SLA, documentos e histórico em um único sistema.</p></div>',unsafe_allow_html=True)
df=get_df()

if page=="Dashboard":
    if df.empty: st.info("Nenhum registro."); st.stop()
    total=len(df); concl=(df.status_exibicao=="✅ Concluído").sum(); atras=(df.status_exibicao=="🔴 Atrasado").sum(); abertas=df.status_exibicao.isin(["⚪ Aberto","🔵 Agendado","🟡 Em atendimento"]).sum()
    sla=(df.loc[df.data_visita.notna(),"sla_ok"].mean()*100) if df.data_visita.notna().any() else 0; medio=df.loc[df.data_visita.notna(),"dias_atendimento"].mean()
    cards=[("Total",total),("Em aberto",abertas),("Concluídas",concl),("Atrasadas",atras),("SLA até 4 dias",f"{sla:.0f}%"),("Tempo médio",f"{medio:.1f} d" if pd.notna(medio) else "-")]
    for col,(label,value) in zip(st.columns(6),cards): col.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div>',unsafe_allow_html=True)
    c1,c2=st.columns(2)
    sc=df.status_exibicao.value_counts().rename_axis("Status").reset_index(name="Quantidade"); c1.plotly_chart(px.bar(sc,x="Status",y="Quantidade",text_auto=True,title="Chamados por status"),use_container_width=True)
    monthly=df.assign(Mês=df.entrada.dt.to_period("M").astype(str)).groupby("Mês").size().reset_index(name="Chamados"); c2.plotly_chart(px.line(monthly,x="Mês",y="Chamados",markers=True,title="Evolução mensal"),use_container_width=True)
    c3,c4=st.columns(2)
    tech=df.replace("",pd.NA).dropna(subset=["tecnico"]).tecnico.value_counts().head(10).reset_index(); tech.columns=["Técnico","Quantidade"]; c3.plotly_chart(px.bar(tech,x="Quantidade",y="Técnico",orientation="h",text_auto=True,title="Ranking de técnicos"),use_container_width=True)
    prod=df.replace("",pd.NA).dropna(subset=["produto"]).produto.value_counts().head(10).reset_index(); prod.columns=["Produto","Quantidade"]; c4.plotly_chart(px.bar(prod,x="Produto",y="Quantidade",text_auto=True,title="Produtos mais atendidos"),use_container_width=True)
    st.subheader("Próximas visitas e atrasos"); agenda=df[df.status_exibicao.isin(["🔵 Agendado","🔴 Atrasado","🟡 Em atendimento"])].sort_values("previsao_visita")
    st.dataframe(agenda[["numero","cliente","endereco","tecnico","previsao_visita","status_exibicao","prioridade"]],use_container_width=True,hide_index=True)

elif page=="Assistências":
    f1,f2,f3,f4=st.columns([2,1,1,1]); busca=f1.text_input("Pesquisar",placeholder="Número, cliente, obra, produto..."); stat=f2.multiselect("Status",STATUS_OPTIONS); techs=f3.multiselect("Técnico",sorted(x for x in df.tecnico.dropna().unique() if x)); pri=f4.multiselect("Prioridade",PRIORIDADE_OPTIONS)
    filt=df.copy()
    if busca: filt=filt[filt.astype(str).apply(lambda c:c.str.contains(busca,case=False,na=False)).any(axis=1)]
    if stat: filt=filt[filt.status_exibicao.isin(stat)]
    if techs: filt=filt[filt.tecnico.isin(techs)]
    if pri: filt=filt[filt.prioridade.isin(pri)]
    st.dataframe(filt[["numero","entrada","cliente","endereco","produto","tecnico","previsao_visita","data_visita","status_exibicao","prioridade","dias_atendimento"]],use_container_width=True,hide_index=True)
    if not filt.empty:
        numero=st.selectbox("Abrir assistência",filt.numero.tolist(),format_func=lambda n:f"#{n} - {filt.loc[filt.numero==n,'cliente'].iloc[0]}"); rec=filt[filt.numero==numero].iloc[0]
        tabs=st.tabs(["Dados","Histórico","Anexos","Ordem de serviço"])
        with tabs[0]:
            with st.form("edit"):
                a,b,c=st.columns(3); entrada=a.date_input("Entrada",rec.entrada.date()); cliente=b.text_input("Cliente",rec.cliente); resp=c.text_input("Responsável no cliente",rec.responsavel_cliente or "")
                a,b,c=st.columns(3); telefone=a.text_input("Telefone",rec.telefone or ""); email=b.text_input("E-mail",rec.email or ""); endereco=c.text_input("Obra / Endereço",rec.endereco or "")
                a,b,c=st.columns(3); produto=a.text_input("Produto",rec.produto or ""); vendedor=b.text_input("Vendedor",rec.vendedor or ""); tecnico=c.text_input("Técnico",rec.tecnico or "")
                a,b,c,d=st.columns(4); previsao=a.date_input("Previsão",rec.previsao_visita.date() if pd.notna(rec.previsao_visita) else None); visita=b.date_input("Visita",rec.data_visita.date() if pd.notna(rec.data_visita) else None); prioridade=c.selectbox("Prioridade",PRIORIDADE_OPTIONS,index=PRIORIDADE_OPTIONS.index(rec.prioridade) if rec.prioridade in PRIORIDADE_OPTIONS else 1); status=d.selectbox("Status",STATUS_OPTIONS,index=STATUS_OPTIONS.index(rec.status) if rec.status in STATUS_OPTIONS else 0)
                a,b=st.columns(2); hi=a.text_input("Hora início",rec.hora_inicio or ""); hf=b.text_input("Hora fim",rec.hora_fim or "")
                servico=st.text_area("Serviço solicitado",rec.servico or ""); causa=st.text_area("Causa identificada",rec.causa or ""); solucao=st.text_area("Solução executada",rec.solucao or ""); obs=st.text_area("Observações",rec.observacoes or ""); assinatura=st.text_input("Nome de quem recebeu o atendimento",rec.assinatura_nome or "")
                save=st.form_submit_button("Salvar alterações",use_container_width=True)
            if save:
                save_record(dict(numero=int(rec.numero),entrada=entrada.isoformat(),cliente=cliente,endereco=endereco,produto=produto,vendedor=vendedor,tecnico=tecnico,previsao_visita=previsao.isoformat() if previsao else None,data_visita=visita.isoformat() if visita else None,status=auto_status(previsao,visita,status),servico=servico,observacoes=obs,prioridade=prioridade,telefone=telefone,email=email,responsavel_cliente=resp,causa=causa,solucao=solucao,assinatura_nome=assinatura,hora_inicio=hi,hora_fim=hf),int(rec.id)); st.success("Atualizado."); st.rerun()
        with tabs[1]:
            event=st.text_input("Novo evento",placeholder="Ex.: Cliente confirmou a visita"); note=st.text_area("Detalhes do evento")
            if st.button("Adicionar ao histórico") and event.strip(): add_history(int(rec.id),event,note); st.rerun()
            with connect() as conn: hist=pd.read_sql_query("SELECT data_hora,evento,observacao,usuario FROM historico WHERE assistencia_id=? ORDER BY id DESC",conn,params=(int(rec.id),))
            st.dataframe(hist,use_container_width=True,hide_index=True)
        with tabs[2]:
            categoria=st.selectbox("Categoria",["Antes","Depois","Laudo","Assinatura","Outro"]); files=st.file_uploader("Adicionar arquivos",accept_multiple_files=True,key=f"up{rec.id}")
            if st.button("Salvar anexos") and files:
                with connect() as conn:
                    for f in files:
                        safe=f"{rec.numero}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{Path(f.name).name}"; path=FILES_DIR/safe; path.write_bytes(f.getbuffer())
                        conn.execute("INSERT INTO anexos(assistencia_id,nome,tipo,caminho,categoria,enviado_em) VALUES(?,?,?,?,?,?)",(int(rec.id),f.name,f.type,str(path),categoria,datetime.now().isoformat(timespec="seconds")))
                    conn.commit(); add_history(int(rec.id),f"{len(files)} anexo(s) incluído(s)",categoria); st.rerun()
            with connect() as conn: ans=pd.read_sql_query("SELECT * FROM anexos WHERE assistencia_id=? ORDER BY id DESC",conn,params=(int(rec.id),))
            for _,a in ans.iterrows():
                p=Path(a.caminho)
                if p.exists(): st.download_button(f"{a.categoria}: {a.nome}",p.read_bytes(),file_name=a.nome,key=f"dl{a.id}")
        with tabs[3]:
            pdf=create_os_pdf(rec); st.download_button("Baixar Ordem de Serviço em PDF",pdf,file_name=f"OS_{rec.numero}.pdf",mime="application/pdf",use_container_width=True)

elif page=="Nova assistência":
    with st.form("new",clear_on_submit=True):
        a,b,c=st.columns(3); numero=a.number_input("Nº",value=next_number(),step=1); entrada=b.date_input("Entrada",date.today()); prioridade=c.selectbox("Prioridade",PRIORIDADE_OPTIONS,index=1)
        a,b,c=st.columns(3); cliente=a.text_input("Cliente *"); resp=b.text_input("Responsável"); telefone=c.text_input("Telefone")
        a,b=st.columns(2); email=a.text_input("E-mail"); endereco=b.text_input("Obra / Endereço")
        a,b,c=st.columns(3); produto=a.text_input("Produto"); vendedor=b.text_input("Vendedor"); tecnico=c.text_input("Técnico")
        a,b,c=st.columns(3); previsao=a.date_input("Previsão",value=None); visita=b.date_input("Visita",value=None); status=c.selectbox("Status",STATUS_OPTIONS)
        servico=st.text_area("Serviço solicitado"); obs=st.text_area("Observações"); submit=st.form_submit_button("Cadastrar",use_container_width=True)
    if submit:
        if not cliente.strip(): st.error("Informe o cliente.")
        else:
            try:
                save_record(dict(numero=int(numero),entrada=entrada.isoformat(),cliente=cliente,endereco=endereco,produto=produto,vendedor=vendedor,tecnico=tecnico,previsao_visita=previsao.isoformat() if previsao else None,data_visita=visita.isoformat() if visita else None,status=auto_status(previsao,visita,status),servico=servico,observacoes=obs,prioridade=prioridade,telefone=telefone,email=email,responsavel_cliente=resp,causa="",solucao="",assinatura_nome="",hora_inicio="",hora_fim="")); st.success("Cadastrada.")
            except sqlite3.IntegrityError: st.error("Número já cadastrado.")

elif page=="Agenda":
    st.subheader("Agenda de visitas")
    view=st.radio("Período",["Hoje","Próximos 7 dias","Próximos 30 dias","Todas"],horizontal=True)
    ag=df[df.previsao_visita.notna()].copy(); hoje=pd.Timestamp(date.today())
    if view=="Hoje": ag=ag[ag.previsao_visita.dt.date==date.today()]
    elif view=="Próximos 7 dias": ag=ag[(ag.previsao_visita>=hoje)&(ag.previsao_visita<=hoje+pd.Timedelta(days=7))]
    elif view=="Próximos 30 dias": ag=ag[(ag.previsao_visita>=hoje)&(ag.previsao_visita<=hoje+pd.Timedelta(days=30))]
    st.dataframe(ag.sort_values("previsao_visita")[["previsao_visita","numero","cliente","endereco","tecnico","status_exibicao","prioridade"]],use_container_width=True,hide_index=True)

elif page=="Importar / Exportar":
    up=st.file_uploader("Importar Excel",type=["xlsx","xls"])
    if up and st.button("Importar / atualizar"):
        try: i,u=load_excel_to_db(up); st.success(f"{i} novos e {u} atualizados."); st.rerun()
        except Exception as e: st.error(str(e))
    st.download_button("Baixar base Excel",export_excel(df),file_name=f"assistencias_{date.today()}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    if DB_PATH.exists(): st.download_button("Backup do banco",DB_PATH.read_bytes(),file_name=f"assistencia_backup_{date.today()}.db",use_container_width=True)

elif page=="Usuários":
    if st.session_state.perfil!="Administrador": st.warning("Acesso exclusivo do administrador.")
    else:
        with st.form("user"):
            a,b=st.columns(2); nome=a.text_input("Nome"); login=b.text_input("Login"); a,b=st.columns(2); senha=a.text_input("Senha",type="password"); perfil=b.selectbox("Perfil",PERFIS); add=st.form_submit_button("Criar usuário")
        if add:
            try:
                with connect() as conn: conn.execute("INSERT INTO usuarios(nome,login,senha_hash,perfil,ativo) VALUES(?,?,?,?,1)",(nome,login,hash_password(senha),perfil)); conn.commit()
                st.success("Usuário criado.")
            except sqlite3.IntegrityError: st.error("Login já existe.")
        with connect() as conn: users=pd.read_sql_query("SELECT nome,login,perfil,ativo FROM usuarios",conn)
        st.dataframe(users,use_container_width=True,hide_index=True)

elif page=="Configurações":
    st.subheader("Configurações")
    st.write(f"SLA padrão: **{SLA_DIAS} dias**")
    st.write(f"Banco de dados: `{DB_PATH.name}`")
    st.write(f"Pasta de anexos: `{FILES_DIR.name}`")
    st.info("Login inicial: admin | Senha inicial: admin123. Troque criando um novo administrador antes de publicar.")
