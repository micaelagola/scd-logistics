"""SCD Logistics — Fase 1 (backend + UI en un solo servicio, para Render).
Sirve la interfaz en / y la API en /productos, /importar, etc.
Al arrancar crea las tablas y, si el maestro está vacío, carga master_seed.csv.
Credenciales solo en DATABASE_URL (variable de entorno). Nada de secretos en el front.
"""
import os, io, csv, re, json
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import FileResponse, StreamingResponse
import openpyxl

# ---------------- Engine: regla SKU / volumen / importador / dedup -------------
WT = "WT-"
def normalize_sku(raw):
    if raw is None: return ""
    s = str(raw).strip()
    return s[len(WT):] if s.startswith(WT) else s

def is_set(raw_sku, desc=""):
    raw = str(raw_sku or "")
    sku = normalize_sku(raw)
    return ("\n" in raw) or (" + " in sku) or ("in one set" in str(desc).lower())

def _num(x):
    try:
        if x is None or str(x).strip() == "": return None
        return float(x)
    except (ValueError, TypeError): return None

def sanitize_volume(L, A, H, V0, tol=0.02):
    L, A, H, V0 = _num(L), _num(A), _num(H), _num(V0)
    if None in (L, A, H) or L <= 0 or A <= 0 or H <= 0:
        return {"volumen_calculado_m3": None, "estado_volumen": "sin dimensiones"}
    calc = round(L * A * H, 6)
    if V0 is None or V0 < 0: estado = "corregido"
    else:
        rel = abs(V0 - calc) / calc if calc else 1.0
        estado = "correcto" if rel <= tol else "corregido"
    return {"volumen_calculado_m3": calc, "estado_volumen": estado}

SYN = {
 "sku":["sku","article #","article","item no","item no.","item","codigo","code","model"],
 "descripcion":["description of goods","description","descripcion","goods","producto"],
 "cantidad":["qty (pcs)","qty","quantity","cantidad","pcs","total pcs"],
 "ctn":["ctn","ctns","cartons","no of pkgs","no of pkgs.","bultos"],
 "largo":["largo","length","l (cm)","l"],"ancho":["ancho","width","w (cm)","w"],
 "alto":["alto","height","h (cm)","h"],"cbm":["cbm","volume","volumen","m3","meas"],
 "peso_neto":["nt.wt","net weight","peso neto","nt. wt."],
 "peso_bruto":["gr.wt","gross weight","peso bruto","gr. wt."],
}
def _n(s): return re.sub(r"\s+"," ",str(s or "").strip().lower())
def detect_header(rows, scan=25):
    bi,bm,bs=None,None,0
    for i,row in enumerate(rows[:scan]):
        cells=[_n(c) for c in row]; m={}; sc=0
        for canon,syns in SYN.items():
            for j,c in enumerate(cells):
                if c in syns: m[canon]=j; sc+=1; break
        if ("sku" in m or "descripcion" in m) and sc>bs: bi,bm,bs=i,m,sc
    return bi,(bm or {})
def _m(v):
    try: x=float(v)
    except (ValueError,TypeError): return None
    if x<=0: return None
    return round(x/100.0,6) if x>3 else round(x,6)
def parse_rows(rows, archivo=""):
    hi,cm=detect_header(rows); res={"detectados":0,"no_interpretadas":0,"errores":[]}; items=[]
    if hi is None:
        res["errores"].append("No pude identificar la columna SKU."); return items,res
    for row in rows[hi+1:]:
        rs=row[cm["sku"]] if "sku" in cm and cm["sku"]<len(row) else None
        ds=row[cm["descripcion"]] if "descripcion" in cm and cm["descripcion"]<len(row) else ""
        if (rs is None or str(rs).strip()=="") and (ds is None or str(ds).strip()==""): continue
        if rs is None or str(rs).strip()=="": res["no_interpretadas"]+=1; continue
        g=lambda k: row[cm[k]] if k in cm and cm[k]<len(row) else None
        items.append({"sku":normalize_sku(rs),"descripcion":str(ds or "").strip(),
            "cantidad":g("cantidad"),"largo_m":_m(g("largo")),"ancho_m":_m(g("ancho")),
            "alto_m":_m(g("alto")),"cbm":g("cbm"),"tipo":"set" if is_set(rs,ds) else "normal",
            "archivo_origen":archivo}); res["detectados"]+=1
    return items,res
_ITEM=re.compile(r"^(?P<r>\d+\s*-\s*\d+)\s+(?P<sku>[A-Za-z0-9][A-Za-z0-9\-\+/]*)\s+(?P<a>[\d.,]+)\s+(?P<b>[\d.,]+)\s+(?P<c>[\d.,]+)\s+(?P<pcs>[\d.,]+)\s+(?P<u>[A-Za-z]+)")
def parse_pdf_text(text, archivo=""):
    items,res=[],{"detectados":0,"no_interpretadas":0,"errores":[]}
    for line in text.splitlines():
        m=_ITEM.match(line.strip())
        if not m: continue
        f=lambda x: (float(str(x).replace(",","")) if x else None)
        rs=m.group("sku")
        items.append({"sku":normalize_sku(rs),"descripcion":"","cantidad":f(m.group("pcs")),
            "largo_m":None,"ancho_m":None,"alto_m":None,"cbm":None,
            "tipo":"set" if is_set(rs,"") else "normal","archivo_origen":archivo})
        res["detectados"]+=1
    if not items: res["errores"].append("No pude identificar filas de productos en el PDF.")
    return items,res

# ---------------- DB -----------------------------------------------------------
import psycopg
from psycopg.rows import dict_row
def conn(): return psycopg.connect(os.environ["DATABASE_URL"], row_factory=dict_row)

SCHEMA = """
create table if not exists productos(
 id bigserial primary key, sku text unique not null, descripcion text,
 largo_m numeric, ancho_m numeric, alto_m numeric,
 volumen_calculado_m3 numeric, volumen_original_m3 numeric,
 estado_volumen text, tipo text default 'normal',
 peso_neto numeric, peso_bruto numeric, fuente text, archivo_origen text,
 fecha_actualizacion date default current_date, fecha_recalculo date, version int default 1);
create table if not exists importaciones(
 id bigserial primary key, nombre_archivo text, tipo text, proveedor text,
 fecha timestamptz default now(), resumen_json jsonb);
create table if not exists revision_pendiente(
 id bigserial primary key, sku text, importacion_id bigint,
 valores_actuales_json jsonb, valores_nuevos_json jsonb, motivo text,
 estado text default 'abierta', fecha timestamptz default now());
create table if not exists auditoria(
 id bigserial primary key, sku text, campo text, valor_anterior text,
 valor_nuevo text, motivo text, usuario text, fecha timestamptz default now());
"""

def init_and_seed():
    with conn() as c, c.cursor() as cur:
        cur.execute(SCHEMA)
        cur.execute("select count(*) n from productos"); 
        if cur.fetchone()["n"] > 0: return
        path = os.path.join(os.path.dirname(__file__), "master_seed.csv")
        if not os.path.exists(path): return
        with open(path, encoding="utf-8") as f:
            for r in csv.DictReader(f):
                nn=lambda x: (float(x) if x not in (None,"","None") else None)
                cur.execute("""insert into productos(sku,descripcion,largo_m,ancho_m,alto_m,
                   volumen_calculado_m3,volumen_original_m3,estado_volumen,tipo,fuente,
                   archivo_origen,fecha_recalculo) values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                   on conflict (sku) do nothing""",
                   [r["sku"],r["articulo"],nn(r["largo_m"]),nn(r["ancho_m"]),nn(r["alto_m"]),
                    nn(r["volumen_calculado_m3"]),nn(r["volumen_original_m3"]),r["estado_volumen"],
                    r["tipo"],r["fuente"],r["archivo_origen"],r["fecha_recalculo"]])

# ---------------- API ----------------------------------------------------------
app = FastAPI(title="SCD Logistics")

@app.on_event("startup")
def _startup():
    try: init_and_seed()
    except Exception as e: print("seed error:", e)

@app.get("/health")
def health(): return {"ok": True}

@app.get("/productos")
def productos(q: str | None = Query(default=None)):
    sql="select * from productos"; args=[]
    if q: sql+=" where sku ilike %s or descripcion ilike %s"; args=[f"%{q}%",f"%{q}%"]
    sql+=" order by sku"
    with conn() as c, c.cursor() as cur: cur.execute(sql,args); return cur.fetchall()

@app.get("/revisiones")
def revisiones():
    with conn() as c, c.cursor() as cur:
        cur.execute("select * from revision_pendiente where estado='abierta' order by fecha")
        return cur.fetchall()

def _rows(name, content):
    ext=name.lower().rsplit(".",1)[-1]
    if ext in ("xlsx","xlsm","xls"):
        wb=openpyxl.load_workbook(io.BytesIO(content),data_only=True); ws=wb[wb.sheetnames[0]]
        return parse_rows([list(r) for r in ws.iter_rows(values_only=True)], name)
    if ext=="csv":
        rows=list(csv.reader(io.StringIO(content.decode("utf-8","ignore"))))
        return parse_rows(rows, name)
    if ext=="pdf":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            text="\n".join((p.extract_text() or "") for p in pdf.pages)
            tab=[r for p in pdf.pages for t in (p.extract_tables() or []) for r in t]
        items,res=parse_rows(tab,name)
        if res["detectados"]==0: items,res=parse_pdf_text(text,name)
        return items,res
    raise HTTPException(400,"Formato no soportado (usá PDF, XLSX o CSV).")

@app.post("/importar")
async def importar(file: UploadFile = File(...), commit: bool = False):
    content=await file.read()
    if len(content) > 25*1024*1024: raise HTTPException(413,"El archivo supera 25 MB.")
    items,res=_rows(file.filename, content)
    for it in items:
        s=sanitize_volume(it.get("largo_m"),it.get("ancho_m"),it.get("alto_m"),it.get("cbm"))
        it.update(s)
    skus=[it["sku"] for it in items]
    master={}
    with conn() as c, c.cursor() as cur:
        if skus:
            cur.execute("select * from productos where sku = any(%s)",[skus])
            master={r["sku"]:r for r in cur.fetchall()}
    nuevos=[it for it in items if it["sku"] not in master]
    seen=set(); dups=[]
    for it in items:
        if it["sku"] in seen: dups.append(it["sku"])
        seen.add(it["sku"])
    modificados=[]
    for it in items:
        m=master.get(it["sku"])
        if m and any(_num(it.get(k))!=_num(m.get(k)) for k in ("largo_m","ancho_m","alto_m")):
            modificados.append(it["sku"])
    resumen={"detectados":res["detectados"],"nuevos":len(set(n["sku"] for n in nuevos)),
      "existentes":len([1 for it in items if it["sku"] in master]),
      "modificados":len(set(modificados)),"duplicados_en_archivo":len(dups),
      "no_interpretadas":res["no_interpretadas"],"errores":res["errores"]}
    if commit:
        with conn() as c, c.cursor() as cur:
            cur.execute("insert into importaciones(nombre_archivo,tipo,resumen_json) values(%s,%s,%s) returning id",
                        [file.filename,file.filename.rsplit('.',1)[-1],json.dumps(resumen)])
            imp=cur.fetchone()["id"]
            done=set()
            for it in nuevos:
                if it["sku"] in done: continue
                done.add(it["sku"])
                cur.execute("""insert into productos(sku,descripcion,largo_m,ancho_m,alto_m,
                   volumen_calculado_m3,estado_volumen,tipo,archivo_origen,fecha_recalculo)
                   values(%s,%s,%s,%s,%s,%s,%s,%s,%s,current_date) on conflict (sku) do nothing""",
                   [it["sku"],it["descripcion"],it.get("largo_m"),it.get("ancho_m"),it.get("alto_m"),
                    it.get("volumen_calculado_m3"),it.get("estado_volumen"),it["tipo"],it["archivo_origen"]])
            for sku in set(modificados):
                m=master[sku]; it=next(x for x in items if x["sku"]==sku)
                cur.execute("insert into revision_pendiente(sku,importacion_id,valores_actuales_json,valores_nuevos_json,motivo) values(%s,%s,%s,%s,%s)",
                    [sku,imp,json.dumps({k:float(m[k]) if m.get(k) is not None else None for k in ("largo_m","ancho_m","alto_m")}),
                     json.dumps({k:it.get(k) for k in ("largo_m","ancho_m","alto_m")}),"Encontré dimensiones diferentes"])
        resumen["importacion_id"]=imp
    return resumen

@app.get("/")
def home():
    return FileResponse(os.path.join(os.path.dirname(__file__), "index.html"))
