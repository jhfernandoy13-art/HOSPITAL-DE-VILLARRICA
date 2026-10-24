# -*- coding: utf-8 -*-
"""
Dashboard de fármacos e insumos: estado de stock, proyección de demanda y alertas.
Construye y cruza automáticamente la base de datos de fármacos desde los archivos crudos
de 'reportes.zip' (Stock centralizado, Licitaciones, ICP Cenabast, Consumos por bodega,
Programación y Vencimientos) sin necesidad de subir Reporte_Farmacos(NUEVA).xlsx.
"""
import base64
import calendar
import glob
import io
import lzma
import os
import re
import zipfile
from datetime import date, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Fármacos e insumos: estado y demanda", page_icon="💊", layout="wide")

ARSENAL_SEED_B64 = (
    "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"
    "SzixxGf7AgAAAAAEWVo="
)

DEFAULT_FILE = os.environ.get("FARMACOS_XLSX", "Reporte_Farmacos__NUEVA_.xlsx")

MES_NUM = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
           "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}
MES_ABR = {v: k for k, v in MES_NUM.items()}
MES_NOMBRE_MAP = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
    "NOVIEMBRE": 11, "DICIEMBRE": 12,
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "SEPT": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}

BODEGAS_FARM_STOCK = {
    "BODEGA FARMACOS 2 HV": "FARMACOS 2",
    "BODEGA FARMACOS HV": "FARMACOS",
    "BODEGA ACTIVA FARMACIA AT. ABIERTA HV": "STOCK AT ABIERTA",
    "BODEGA ACTIVA FARMACIA AT. CERRADA HV": "STOCK AT CERRADA",
    "BODEGA STOCKLIGTH AT. ABIERTA": "STOCKLIGTH ABIERTA",
    "BODEGA STOCKLIGTH AT. CERRADA": "STOCKTIGTH CERADA",
    "BODEGA BOTIQUIN FARMACIA URGENCIA HV": "BOTIQUIN URGENCIA",
    "BODEGA ACTIVA FARMACIA C.E.M.": "BODEGA CEM",
}

ESTADOS = ["QS", "CR", "OK", "SS"]
ESTADO_TXT = {"QS": "QS (quiebre de stock)", "CR": "CR (crítico)",
              "OK": "OK (normal)", "SS": "SS (sobre stock)"}
ESTADO_ICON = {"QS": "🔴 QS", "CR": "🟠 CR", "OK": "🟢 OK", "SS": "🔵 SS"}
ESTADO_COLOR = {"QS": "#B3261E", "CR": "#E8890C", "OK": "#2E7D5B", "SS": "#3B6EA8"}
ESTADO_TXT.update({"SM": "SM (con stock, sin salidas)", "IN": "IN (inactivo)"})
ESTADO_ICON.update({"SM": "⚪ SM", "IN": "⚫ IN"})
ESTADO_COLOR.update({"SM": "#8A8A8A", "IN": "#C9C9C9"})
ROT_ORDEN = ["SR", "BR", "MIR", "MEDR", "ALR", "GR", "MAXR", "OK"]
CEN_RIESGO = ["FALTANTE", "SUSPENDIDO", "ELIMINADO"]
Z80 = 1.2816
MIN_TRAIN = 3
HORIZONTE_FC = 18

st.markdown(
    """
<style>
.kpi{border-radius:8px;padding:.8rem 1rem;background:rgba(128,128,128,.10);margin-bottom:.4rem}
.kpi-l{font-size:.85rem;opacity:.75}
.kpi-v{font-size:2rem;font-weight:700;line-height:1.15}
.kpi-s{font-size:.8rem;opacity:.65}
</style>
""",
    unsafe_allow_html=True,
)


def n0(x):
    if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
        return "–"
    return f"{x:,.0f}".replace(",", ".")


def clp(x):
    s = n0(x)
    return s if s == "–" else "$ " + s


def kpi(label, value, sub="", color="#888888"):
    st.markdown(
        f'<div class="kpi" style="border-left:6px solid {color}">'
        f'<div class="kpi-l">{label}</div><div class="kpi-v">{value}</div>'
        f'<div class="kpi-s">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def parse_mes(label):
    m = re.match(r"^(?:PROG )?([A-Z]{3})-(\d{2})(?: CONS)?$", str(label).strip())
    if m and m.group(1) in MES_NUM:
        return (2000 + int(m.group(2)), MES_NUM[m.group(1)])
    return None


def mes_label(ym):
    return f"{MES_ABR[ym[1]]}-{str(ym[0])[2:]}"


def next_month(ym):
    return (ym[0] + (1 if ym[1] == 12 else 0), ym[1] % 12 + 1)


def mes_ts(ym):
    return pd.Timestamp(ym[0], ym[1], 1)


def parse_clp(v):
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = re.sub(r"[^\d,\-]", "", str(v)).replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


def rot_from_consumo(x):
    return np.select(
        [x <= 0, x < 10, x < 100, x < 1000, x < 5000, x < 10000, x < 190000],
        ["SR", "BR", "MIR", "MEDR", "ALR", "GR", "MAXR"],
        "OK",
    )


def _fecha_dmy(d, m, y):
    try:
        return date(int(y), int(m), int(d))
    except (ValueError, TypeError):
        return None


def leer_con_encabezado(file_bytes, clave="Código"):
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    hdr = None
    for i in range(min(25, len(raw))):
        if str(raw.iloc[i, 0]).strip().lower() == clave.lower():
            hdr = i
            break
    if hdr is None:
        raise ValueError(f"No encontré la fila de títulos que empieza con '{clave}'.")
    d = raw.iloc[hdr + 1:].copy()
    d.columns = [str(x).strip() for x in raw.iloc[hdr].tolist()]
    col_real = d.columns[0]
    d = d.dropna(subset=[col_real]).reset_index(drop=True)
    texto = " ".join(str(x) for x in raw.iloc[:hdr, 0].dropna().tolist())
    return texto, d


@st.cache_data(show_spinner=False)
def clasificar(nombre: str, file_bytes: bytes):
    nom_up = nombre.upper()
    if nom_up.endswith(".CSV"):
        try:
            inicio = file_bytes[:1500].decode("latin1", errors="ignore").upper()
            if "CENABAST" in inicio or "SEMAFORO;" in inicio or "INTERMEDIACION" in inicio:
                return "icp_cenabast"
        except Exception:
            pass
        return None

    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        if "ARSENAL_FARMACIA" in wb.sheetnames:
            return "farmacos"
        ws = wb[wb.sheetnames[0]]
        filas = []
        for fila in ws.iter_rows(min_row=1, max_row=5, values_only=True):
            filas.append([str(c or "").strip() for c in fila[:6]])
        a1 = (filas[0][0] if filas and filas[0] else "").lower()
        a2 = (filas[1][0] if len(filas) > 1 and filas[1] else "").lower()

        if a1.startswith("reporte de licitaciones"):
            return "licitaciones_raw"
        if a1.startswith("reporte de productos sin recepcionar"):
            return "sin_recepcionar"
        if a1.startswith("reporte de stock"):
            if "todas las bodegas" in a2:
                return "stock_centralizado_raw"
            return "vencimientos"
        if a1.startswith("stock"):
            return "stock_insumos"
        if a1.startswith("movimientos"):
            es_farm = any(k in a1 for k in ("farmacia", "stockligth", "stocklight", "botiquin", "c.e.m"))
            return "consumo_farm_raw" if es_farm else "salidas_insumos"
        if "programacion" in nom_up or any("cant. programada" in " ".join(f).lower() for f in filas):
            return "programacion_raw"
    except Exception:
        pass
    return None


def reunir_archivos(ups):
    def vacio():
        return {
            "farmacos": None,
            "stock_insumos": None,
            "vencimientos": None,
            "salidas_insumos": [],
            "stock_centralizado_raw": None,
            "licitaciones_raw": None,
            "icp_cenabast": [],
            "consumo_farm_raw": [],
            "programacion_raw": [],
        }

    def registrar(res, vistos, nombre, b):
        if nombre.upper().endswith(".ZIP"):
            try:
                with zipfile.ZipFile(io.BytesIO(b), "r") as z:
                    for zinfo in z.infolist():
                        if not zinfo.is_dir() and not os.path.basename(zinfo.filename).startswith("."):
                            registrar(res, vistos, zinfo.filename, z.read(zinfo.filename))
            except Exception:
                pass
            return
        h = hash(b)
        if h in vistos:
            return
        vistos.add(h)
        t = clasificar(nombre, b)
        if t in ("salidas_insumos", "icp_cenabast", "consumo_farm_raw", "programacion_raw"):
            res[t].append((nombre, b))
        elif t in ("farmacos", "stock_insumos", "vencimientos", "stock_centralizado_raw", "licitaciones_raw"):
            if res[t] is None:
                res[t] = (nombre, b)

    subidos, locales = vacio(), vacio()
    vistos_s, vistos_l = set(), set()
    for u in ups or []:
        registrar(subidos, vistos_s, u.name, u.getvalue())

    rutas = []
    if os.path.exists(DEFAULT_FILE):
        rutas.append(DEFAULT_FILE)
    rutas += sorted(glob.glob("*.xlsx") + glob.glob("*.csv") + glob.glob("*.zip"), key=os.path.getmtime, reverse=True)
    for r in rutas:
        try:
            with open(r, "rb") as fh:
                registrar(locales, vistos_l, os.path.basename(r), fh.read())
        except OSError:
            pass

    res = vacio()
    for t in ("farmacos", "stock_insumos", "vencimientos", "stock_centralizado_raw", "licitaciones_raw"):
        res[t] = subidos[t] or locales[t]
    for t in ("salidas_insumos", "icp_cenabast", "consumo_farm_raw", "programacion_raw"):
        res[t] = subidos[t] or locales[t]
    return res


def hay_archivos_crudos_farmacos(archivos: dict) -> bool:
    return any([
        archivos["stock_centralizado_raw"] is not None,
        archivos["licitaciones_raw"] is not None,
        bool(archivos["icp_cenabast"]),
        bool(archivos["consumo_farm_raw"]),
        bool(archivos["programacion_raw"]),
    ])


@st.cache_data(show_spinner=False)
def cargar_base_semilla():
    raw_csv = lzma.decompress(base64.b64decode(ARSENAL_SEED_B64))
    ars = pd.read_csv(io.BytesIO(raw_csv))
    ars["dato"] = ars["dato"].astype(str).str.strip()
    texto = {"dato", "DESCRIPCION", "SEGUIMIENTO"}
    for c in ars.columns:
        if c not in texto:
            ars[c] = pd.to_numeric(ars[c], errors="coerce")
    for col_st in list(BODEGAS_FARM_STOCK.values()) + ["TOTAL BODEGAS", "TOTAL FARMACIAS", "STOCK TOTAL", "STOCK SIN STOCKLIGHT"]:
        ars[col_st] = 0.0
    conv = pd.DataFrame(columns=["dato", "ID LICITACION", "SALDO", "VCTO"])
    venc = pd.DataFrame(columns=["Bodega", "Código", "Producto", "Lote", "Fecha Vencimiento", "Cantidad", "Precio Lote", "Total"])
    lic = pd.DataFrame(columns=["Código", "Licitación", "Saldo", "Proveedor", "Vencimiento"])
    precios = pd.Series(dtype=float)
    return {"ars": ars, "conv": conv, "venc": venc, "precios": precios, "fecha_txt": None, "licit": lic}


@st.cache_data(show_spinner="Leyendo el Excel de fármacos…")
def load_data(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    ars = xls.parse("ARSENAL_FARMACIA")
    ars.columns = [str(c).strip() for c in ars.columns]
    ars = ars[ars["dato"].notna()].copy()
    ars["dato"] = ars["dato"].astype(str).str.strip()
    ars = ars.drop_duplicates("dato").reset_index(drop=True)
    texto = {"dato", "DESCRIPCION", "ESTADO", "ROTACION", "ESTADO CENABAST",
             "FECHA ENTREGA", "ENTREGA EN DIA", "SEGUIMIENTO", "CTRL2"}
    for c in ars.columns:
        if c not in texto:
            ars[c] = pd.to_numeric(ars[c], errors="coerce")

    try:
        conv = xls.parse("CONVENIOS_VIGENTES")
        conv.columns = [str(c).strip() for c in conv.columns]
        conv["dato"] = conv["dato"].astype(str).str.strip()
        conv = conv.drop_duplicates("dato")
        conv["VCTO"] = pd.to_datetime(conv["VCTO"], errors="coerce")
        conv["SALDO"] = pd.to_numeric(conv["SALDO"], errors="coerce")
        conv = conv[["dato", "ID LICITACION", "SALDO", "VCTO"]]
    except Exception:
        conv = pd.DataFrame(columns=["dato", "ID LICITACION", "SALDO", "VCTO"])

    try:
        venc = xls.parse("VENCIMIENTOS", header=2)
        venc.columns = [str(c).strip() for c in venc.columns]
        venc["Fecha Vencimiento"] = pd.to_datetime(venc["Fecha Vencimiento"], errors="coerce")
        venc = venc.dropna(subset=["Fecha Vencimiento"]).copy()
        venc["Código"] = venc["Código"].astype(str).str.strip()
        venc["Lote"] = venc["Lote"].astype(str)
        venc["Cantidad"] = pd.to_numeric(venc["Cantidad"], errors="coerce").fillna(0)
        venc["Precio Lote"] = venc["Precio Lote"].map(parse_clp)
        venc["Total"] = venc["Total"].map(parse_clp)
    except Exception:
        venc = pd.DataFrame(columns=["Bodega", "Código", "Producto", "Lote",
                                     "Fecha Vencimiento", "Cantidad", "Precio Lote", "Total"])

    try:
        lic = xls.parse("LICITACIONES")
        lic.columns = [str(c).strip() for c in lic.columns]
        lic["Código"] = lic["Código"].astype(str).str.strip()
        lic["Saldo"] = pd.to_numeric(lic["Saldo"], errors="coerce")
        lic["Vencimiento"] = pd.to_datetime(lic["Vencimiento"], errors="coerce")
        lic = lic[["Código", "Licitación", "Saldo", "Proveedor", "Vencimiento"]]
    except Exception:
        lic = pd.DataFrame(columns=["Código", "Licitación", "Saldo", "Proveedor", "Vencimiento"])

    precios = pd.Series(dtype=float)
    try:
        sc = xls.parse("STOCK CENTRALIZADO", usecols=["Código", "Cantidad", "Precio"])
        sc["Código"] = sc["Código"].astype(str).str.strip()
        sc = sc[sc["Código"].isin(ars["dato"])].copy()
        sc["Cantidad"] = pd.to_numeric(sc["Cantidad"], errors="coerce")
        sc["Precio"] = pd.to_numeric(sc["Precio"], errors="coerce")
        pos = sc[(sc["Cantidad"] > 0) & (sc["Precio"] > 0)].copy()
        pos["pq"] = pos["Precio"] * sc["Cantidad"]
        g = pos.groupby("Código")[["pq", "Cantidad"]].sum()
        p1 = g["pq"] / g["Cantidad"]
        p2 = sc[sc["Precio"] > 0].groupby("Código")["Precio"].mean()
        precios = p1.combine_first(p2)
    except Exception:
        pass

    fecha_txt = None
    try:
        rep = xls.parse("REPORTE", header=None)
        for v in rep.to_numpy().ravel():
            m = re.search(r"ACTUALIZACI[OÓ]N\s*(\d{1,2})/(\d{1,2})", str(v), flags=re.I)
            if m:
                fecha_txt = (int(m.group(1)), int(m.group(2)))
                break
    except Exception:
        pass

    return {"ars": ars, "conv": conv, "venc": venc, "precios": precios, "fecha_txt": fecha_txt, "licit": lic}


def _extraer_ym_nombre(nombre, texto="", anio_defecto=None):
    s = f"{nombre} {texto}".upper()
    m = re.search(r"desde\s*(\d{1,2})/(\d{1,2})/(\d{4})", s, flags=re.I)
    if m:
        return (int(m.group(3)), int(m.group(2)))
    for mes_nom, m_num in sorted(MES_NOMBRE_MAP.items(), key=lambda x: -len(x[0])):
        pat = rf"(?:^|[^A-Z]){mes_nom}(?:[^A-Z]|$)"
        if re.search(pat, s):
            ym = re.search(r"20(2\d|3\d)", s)
            if ym:
                yr = 2000 + int(ym.group(1))
            else:
                yr = anio_defecto or date.today().year
            return (yr, m_num)
    return None


def cruzar_datos_crudos(data: dict, archivos: dict):
    ars = data["ars"].copy()
    conv = data["conv"].copy()
    lic = data["licit"].copy()
    venc = data["venc"].copy()
    precios = data["precios"].copy()
    fecha_txt = data["fecha_txt"]
    log_cruces = []

    if archivos["stock_centralizado_raw"] is not None:
        nom_s, b_s = archivos["stock_centralizado_raw"]
        try:
            texto_s, sc = leer_con_encabezado(b_s, clave="Código")
            sc["Código"] = sc["Código"].astype(str).str.strip()
            sc["Bodega_norm"] = sc["Bodega"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip().str.upper()
            sc["Cantidad"] = pd.to_numeric(sc["Cantidad"], errors="coerce").fillna(0)
            sc["Precio"] = pd.to_numeric(sc["Precio"], errors="coerce").fillna(0)

            piv_s = (
                sc[sc["Bodega_norm"].isin(BODEGAS_FARM_STOCK)]
                .groupby(["Código", "Bodega_norm"])["Cantidad"]
                .sum()
                .unstack(fill_value=0)
                .rename(columns=BODEGAS_FARM_STOCK)
            )
            for col_dest in BODEGAS_FARM_STOCK.values():
                if col_dest in piv_s.columns:
                    ars[col_dest] = ars["dato"].map(piv_s[col_dest]).fillna(0)
                else:
                    ars[col_dest] = 0.0

            ars["TOTAL BODEGAS"] = ars["FARMACOS"].fillna(0) + ars["FARMACOS 2"].fillna(0)
            cols_farm = ["STOCK AT ABIERTA", "STOCK AT CERRADA", "STOCKLIGTH ABIERTA",
                         "STOCKTIGTH CERADA", "BOTIQUIN URGENCIA", "BODEGA CEM"]
            ars["TOTAL FARMACIAS"] = ars[cols_farm].fillna(0).sum(axis=1)
            ars["STOCK TOTAL"] = ars["TOTAL BODEGAS"] + ars["TOTAL FARMACIAS"]
            ars["STOCK SIN STOCKLIGHT"] = (
                ars["STOCK TOTAL"] - ars["STOCKLIGTH ABIERTA"].fillna(0) - ars["STOCKTIGTH CERADA"].fillna(0)
            )

            sc_f = sc[sc["Código"].isin(ars["dato"])].copy()
            pos = sc_f[(sc_f["Cantidad"] > 0) & (sc_f["Precio"] > 0)].copy()
            pos["pq"] = pos["Precio"] * pos["Cantidad"]
            g = pos.groupby("Código")[["pq", "Cantidad"]].sum()
            p1 = g["pq"] / g["Cantidad"]
            p2 = sc_f[sc["Precio"] > 0].groupby("Código")["Precio"].mean() if len(sc_f) else pd.Series(dtype=float)
            precios = p1.combine_first(p2)

            m_f = re.search(r"Obtenido\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})", texto_s, flags=re.I)
            if m_f:
                fecha_txt = (int(m_f.group(1)), int(m_f.group(2)))
            log_cruces.append(f"Stock centralizado ({os.path.basename(nom_s)})")
        except Exception as e:
            st.sidebar.warning(f"No se pudo cruzar Stock Centralizado: {e}")

    if archivos["licitaciones_raw"] is not None:
        nom_l, b_l = archivos["licitaciones_raw"]
        try:
            _, df_l = leer_con_encabezado(b_l, clave="Licitación")
            df_l["Código"] = df_l["Código"].astype(str).str.strip()
            df_l["Saldo"] = pd.to_numeric(df_l["Saldo"], errors="coerce")
            vcto = pd.to_datetime(df_l["Vencimiento"], format="%d/%m/%Y", errors="coerce")
            if vcto.isna().any():
                vcto[vcto.isna()] = pd.to_datetime(df_l.loc[vcto.isna(), "Vencimiento"], errors="coerce")
            df_l["Vencimiento"] = vcto
            lic = df_l[["Código", "Licitación", "Saldo", "Proveedor", "Vencimiento"]].copy()

            prim_lic = lic.drop_duplicates("Código", keep="first").set_index("Código")
            conv = pd.DataFrame({
                "dato": ars["dato"],
                "ID LICITACION": ars["dato"].map(prim_lic["Licitación"]),
                "SALDO": pd.to_numeric(ars["dato"].map(prim_lic["Saldo"]), errors="coerce"),
                "VCTO": pd.to_datetime(ars["dato"].map(prim_lic["Vencimiento"]), errors="coerce"),
            })
            log_cruces.append(f"Licitaciones y Convenios ({os.path.basename(nom_l)})")
        except Exception as e:
            st.sidebar.warning(f"No se pudo cruzar Licitaciones: {e}")

    if archivos["icp_cenabast"]:
        def _orden_icp(item):
            ym = _extraer_ym_nombre(item[0])
            return ym if ym else (0, 0)

        nom_icp, b_icp = sorted(archivos["icp_cenabast"], key=_orden_icp)[-1]
        try:
            txt_icp = b_icp.decode("latin1", errors="ignore")
            lineas = txt_icp.splitlines()
            hdr_idx = next(i for i, l in enumerate(lineas[:15]) if "SEMAFORO" in l.upper() and ";" in l)
            df_icp = pd.read_csv(io.StringIO("\n".join(lineas[hdr_idx:])), sep=";")
            df_icp.columns = [str(c).strip().upper() for c in df_icp.columns]
            df_icp["CODIGO GENERICO"] = pd.to_numeric(df_icp["CODIGO GENERICO"], errors="coerce")
            df_icp = df_icp.dropna(subset=["CODIGO GENERICO"]).drop_duplicates("CODIGO GENERICO", keep="first")
            icp_map = df_icp.set_index("CODIGO GENERICO")

            zgen_s = pd.to_numeric(ars["ZGEN"], errors="coerce")
            if "CANTIDAD UNITARIA PROGRAMADA" in icp_map.columns:
                ars["CANT PROGRAMADA CENABAST"] = pd.to_numeric(
                    zgen_s.map(icp_map["CANTIDAD UNITARIA PROGRAMADA"]), errors="coerce"
                )
            if "Q FALTANTE" in icp_map.columns:
                ars["CANT FALTANTE CENABAST"] = pd.to_numeric(
                    zgen_s.map(icp_map["Q FALTANTE"]), errors="coerce"
                )
            if "CANTIDAD UNITARIA A DESPACHAR" in icp_map.columns:
                ars["CANT ENTREGA CENABAST"] = pd.to_numeric(
                    zgen_s.map(icp_map["CANTIDAD UNITARIA A DESPACHAR"]), errors="coerce"
                )
            if "ESTADO DEL MATERIAL" in icp_map.columns:
                ars["ESTADO CENABAST"] = zgen_s.map(icp_map["ESTADO DEL MATERIAL"].astype(str).str.strip())
            log_cruces.append(f"ICP Cenabast ({os.path.basename(nom_icp)})")
        except Exception as e:
            st.sidebar.warning(f"No se pudo cruzar ICP Cenabast: {e}")

    if archivos["consumo_farm_raw"]:
        por_mes = {}
        n_archivos_mes = {}
        for nom_c, b_c in archivos["consumo_farm_raw"]:
            try:
                texto_c, df_c = leer_con_encabezado(b_c, clave="Código")
                ym = _extraer_ym_nombre(nom_c, texto_c)
                if ym is None:
                    continue
                n_archivos_mes[ym] = n_archivos_mes.get(ym, 0) + 1
                df_c["Código"] = df_c["Código"].astype(str).str.strip()
                c_col = "cantidad" if "cantidad" in df_c.columns else "Cantidad"
                df_c["cant"] = pd.to_numeric(df_c[c_col], errors="coerce").fillna(0)
                if "Tipo Documento" in df_c.columns:
                    td = df_c["Tipo Documento"].astype(str).str.strip().str.upper()
                    dest = (
                        df_c["Bodega Destino"].astype(str).str.strip().str.upper()
                        if "Bodega Destino" in df_c.columns
                        else pd.Series("", index=df_c.index)
                    )
                    cond = (td != "SOLICITUD AJUSTE") & ~((td == "TRASPASO") & dest.str.startswith("VILLARRICA HOSP."))
                    df_c = df_c[cond]
                s_mes = df_c.groupby("Código")["cant"].sum()
                por_mes[ym] = por_mes.get(ym, pd.Series(0.0, index=ars["dato"])).add(s_mes, fill_value=0)
            except Exception:
                pass
        meses_act = []
        for ym, serie_m in sorted(por_mes.items()):
            col_c = f"{mes_label(ym)} CONS"
            if col_c in ars.columns and n_archivos_mes.get(ym, 0) < 6:
                ars[col_c] = np.maximum(ars[col_c].fillna(0), ars["dato"].map(serie_m).fillna(0))
            else:
                ars[col_c] = ars["dato"].map(serie_m).fillna(0)
            meses_act.append(mes_label(ym))
        if meses_act:
            log_cruces.append(f"Consumos de farmacia ({', '.join(meses_act)})")

    if archivos["programacion_raw"]:
        prog_mes = {}
        for nom_p, b_p in archivos["programacion_raw"]:
            try:
                ym = _extraer_ym_nombre(nom_p)
                if ym is None:
                    continue
                df_p = pd.read_excel(io.BytesIO(b_p))
                df_p.columns = [str(c).strip() for c in df_p.columns]
                c_cod = next((c for c in df_p.columns if "cód" in c.lower() or "cod" in c.lower()), df_p.columns[0])
                c_can = next((c for c in df_p.columns if "cant" in c.lower() or "prog" in c.lower()), df_p.columns[-1])
                df_p[c_cod] = df_p[c_cod].astype(str).str.strip()
                df_p[c_can] = pd.to_numeric(df_p[c_can], errors="coerce").fillna(0)
                s_p = df_p.groupby(c_cod)[c_can].sum()
                prog_mes[ym] = prog_mes.get(ym, pd.Series(0.0, index=ars["dato"])).add(s_p, fill_value=0)
            except Exception:
                pass
        for ym, serie_p in sorted(prog_mes.items()):
            col_p = f"PROG {mes_label(ym)}"
            ars[col_p] = ars["dato"].map(serie_p).fillna(0)
        if prog_mes:
            log_cruces.append(f"Programación ({', '.join(mes_label(k) for k in sorted(prog_mes))})")

    if venc.empty and archivos["vencimientos"] is not None:
        try:
            _, d_v = leer_con_encabezado(archivos["vencimientos"][1], clave="Establecimiento")
            d_v["Código"] = d_v["Código"].astype(str).str.strip()
            d_v = d_v[d_v["Código"].isin(set(ars["dato"]))].copy()
            fv = pd.to_datetime(d_v["Fecha Vencimiento"], format="%d/%m/%Y", errors="coerce")
            if fv.isna().any():
                fv[fv.isna()] = pd.to_datetime(d_v.loc[fv.isna(), "Fecha Vencimiento"], errors="coerce")
            d_v["Fecha Vencimiento"] = fv
            d_v = d_v.dropna(subset=["Fecha Vencimiento"]).copy()
            d_v["Cantidad"] = pd.to_numeric(d_v["Cantidad"], errors="coerce").fillna(0)
            d_v["Precio Lote"] = d_v["Precio Lote"].map(parse_clp)
            d_v["Total"] = d_v["Total"].map(parse_clp)
            venc = d_v[["Bodega", "Código", "Producto", "Lote", "Fecha Vencimiento", "Cantidad", "Precio Lote", "Total"]].reset_index(drop=True)
        except Exception:
            pass

    if log_cruces:
        cons_cols = sorted(
            [c for c in ars.columns if re.match(r"^[A-Z]{3}-\d{2} CONS$", c) and parse_mes(c)],
            key=lambda c: parse_mes(c),
        )
        ult3 = cons_cols[-3:] if len(cons_cols) >= 3 else cons_cols
        mat3 = ars[ult3].apply(pd.to_numeric, errors="coerce").fillna(0).to_numpy(float)
        con_nz = np.where(mat3 != 0, mat3, np.nan)
        with np.errstate(invalid="ignore"):
            prom = np.nan_to_num(np.nanmean(con_nz, axis=1), nan=0.0)
        ars["CONSUMO PROMEDIO"] = prom
        if mat3.shape[1] == 3:
            mayor_c = np.maximum(mat3[:, 0], mat3[:, 2])
        else:
            mayor_c = mat3.max(axis=1)
        ars["MAYOR CONSUMO"] = mayor_c
        ars["CONSUMO DIARIO"] = ars["CONSUMO PROMEDIO"] / 30.0
        ars["CONSUMO SEMANAL"] = ars["CONSUMO DIARIO"] * 7.0
        ars["CONSUMO QUINCENAL"] = ars["CONSUMO DIARIO"] * 15.0
        ars["STOCK CRITICO"] = ars["CONSUMO DIARIO"] * 20.0
        ars["STOCK IDEAL"] = ars["CONSUMO DIARIO"] * 60.0

        cd = ars["CONSUMO DIARIO"].replace(0, np.nan)
        ars["AUTONOMIA DIAS (BODEGAS)"] = (ars["TOTAL BODEGAS"] / cd).fillna(0)
        ars["AUTONOMIA DIAS (HOSPITAL)"] = (ars["STOCK TOTAL"] / cd).fillna(0)

        q_sin = ars["STOCK SIN STOCKLIGHT"].fillna(0).to_numpy(float)
        ad_prom = ars["CONSUMO PROMEDIO"].fillna(0).to_numpy(float)
        ar_prog = (
            pd.to_numeric(ars["PROG JUL-26"], errors="coerce").fillna(0).to_numpy(float)
            if "PROG JUL-26" in ars.columns
            else np.zeros(len(ars))
        )
        aut_m = np.where(
            ad_prom > 0,
            q_sin / np.where(ad_prom > 0, ad_prom, 1.0),
            np.where(ar_prog > 0, q_sin / np.where(ar_prog > 0, ar_prog, 1.0), np.where(q_sin == 0, 0.0, 2.0)),
        )
        ars["AUTONOMIA MENSUAL (HOSP - STOCKLIGHT)"] = aut_m
        ars["AUTONOMIA DIAS (HOSP - STOCKLIGHT)"] = aut_m * 30.0
        ars["ESTADO"] = np.select(
            [ars["AUTONOMIA DIAS (HOSP - STOCKLIGHT)"] <= 0,
             ars["AUTONOMIA DIAS (HOSP - STOCKLIGHT)"] < 30,
             ars["AUTONOMIA DIAS (HOSP - STOCKLIGHT)"] < 360],
            ["QS", "CR", "OK"],
            "SS",
        )
        ars["ROTACION"] = rot_from_consumo(ars["MAYOR CONSUMO"].to_numpy(float))

    return {
        "ars": ars,
        "conv": conv,
        "venc": venc,
        "precios": precios,
        "fecha_txt": fecha_txt,
        "licit": lic,
        "log_cruces": log_cruces,
    }


def exportar_excel_cruzado(data: dict) -> bytes:
    import openpyxl
    base_file = "Reporte Farmacos (NUEVA).xlsx"
    if os.path.exists(base_file):
        wb = openpyxl.load_workbook(base_file)
    else:
        wb = openpyxl.Workbook()
        for s in ['REPORTE', 'ARSENAL_FARMACIA', 'CONVENIOS_VIGENTES', 'NIVELES_STOCK', 'STOCK CENTRALIZADO', 'ICP', 'FALTAS 1708', 'VENCIMIENTOS', 'LICITACIONES', 'FALTAS', 'PEGAR_DATOS', 'CONSUMOS_X_BODEGA']:
            if s not in wb.sheetnames:
                wb.create_sheet(s)

    if "ARSENAL_FARMACIA" in wb.sheetnames:
        ws = wb["ARSENAL_FARMACIA"]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None
        ars_df = data["ars"]
        for c_idx, col_name in enumerate(ars_df.columns):
            ws.cell(row=1, column=c_idx + 1, value=col_name)
        for r_idx, row in ars_df.iterrows():
            for c_idx, val in enumerate(row):
                ws.cell(row=r_idx + 2, column=c_idx + 1, value=val if pd.notna(val) else None)

    if "CONVENIOS_VIGENTES" in wb.sheetnames:
        ws = wb["CONVENIOS_VIGENTES"]
        conv_df = data["conv"]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None
        for c_idx, col_name in enumerate(conv_df.columns):
            ws.cell(row=1, column=c_idx + 1, value=col_name)
        for r_idx, row in conv_df.iterrows():
            for c_idx, val in enumerate(row):
                ws.cell(row=r_idx + 2, column=c_idx + 1, value=val if pd.notna(val) else None)

    if "LICITACIONES" in wb.sheetnames:
        ws = wb["LICITACIONES"]
        lic_df = data["licit"]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None
        for c_idx, col_name in enumerate(lic_df.columns):
            ws.cell(row=1, column=c_idx + 1, value=col_name)
        for r_idx, row in lic_df.iterrows():
            for c_idx, val in enumerate(row):
                ws.cell(row=r_idx + 2, column=c_idx + 1, value=val if pd.notna(val) else None)

    if "VENCIMIENTOS" in wb.sheetnames:
        ws = wb["VENCIMIENTOS"]
        venc_df = data["venc"]
        for r in range(4, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None
        for r_idx, row in venc_df.iterrows():
            for c_idx, val in enumerate(row):
                ws.cell(row=r_idx + 4, column=c_idx + 1, value=val if pd.notna(val) else None)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def f_pm3(y, h):
    return np.full(h, y[-3:].mean())


def f_mean(y, h):
    return np.full(h, y.mean())


def f_ses(y, h, a=0.5):
    lvl = y[0]
    for v in y[1:]:
        lvl = a * v + (1 - a) * lvl
    return np.full(h, lvl)


def f_holt(y