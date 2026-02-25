# -*- coding: utf-8 -*-
"""
Created on Thu Feb 19 20:28:41 2026

@author: User
"""

# streamlit_app.py
# -*- coding: utf-8 -*-

import io
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import dropbox
import streamlit as st
import numpy as np
import json
from openai import OpenAI
import random
from faker import Faker

st.set_page_config(page_title="Manifiestos Astrid", layout="wide")

# -----------------------------
# LOGIN (antes de cargar la app)
# -----------------------------
def require_login():
    if "authed" not in st.session_state:
        st.session_state.authed = False

    if st.session_state.authed:
        return

    st.title("🔒 Login")

    pw = st.text_input("Contraseña", type="password")

    if st.button("Entrar"):
        if pw == st.secrets["auth"]["password"]:
            st.session_state.authed = True
            st.success("Acceso concedido ✅")
            st.rerun()
        else:
            st.error("Contraseña incorrecta ❌")

    st.stop()  # corta la app aquí si no está autenticado


require_login()





# Ruta EXACTA en Dropbox (siempre empieza con /)
DBX_FILE_PATH = "/Manifiestos/Manifiestos_astrid.xlsx"


# -----------------------------
# Dropbox client
# -----------------------------
def get_dbx() -> dropbox.Dropbox:
    cfg_dbx = st.secrets["dropbox"]
    return dropbox.Dropbox(
        app_key=cfg_dbx["app_key"],
        app_secret=cfg_dbx["app_secret"],
        oauth2_refresh_token=cfg_dbx["refresh_token"],
    )


# -----------------------------
# Helpers
# -----------------------------
def _clean_str_series(s: pd.Series) -> pd.Series:
    return (
        s.astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def _norm_casillero(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    digits = "".join(ch for ch in s if ch.isdigit())
    return f"CA{digits}" if digits else s


# -----------------------------
# UI: Uploaders
# -----------------------------
st.title("Manifiestos Astrid")

col1, col2, col3 = st.columns(3)
with col1:
    up_a = st.file_uploader("Sube Pistoleo de bodega (Envios pistoleo)", type=["xlsx", "xls"])
with col2:
    up_b = st.file_uploader("Sube Envíos Encargomio (Envios_Encargomio)", type=["xlsx", "xls"])
with col3:
    up_p = st.file_uploader("Sube Productos por casillero", type=["xlsx", "xls"])

run = st.button(
    "Procesar y actualizar histórico en Dropbox",
    type="primary",
    disabled=not (up_a and up_b and up_p)
)
if run:
    # -----------------------------
    # 1) Leer histórico de Dropbox
    # -----------------------------
    dbx = get_dbx()
    def load_historico_or_empty(dbx: dropbox.Dropbox, path: str) -> pd.DataFrame:
        # Columnas mínimas esperadas del histórico
        base_cols = [
            "FECHA GUIA", "guia", "COMPAÑÍA REMITENTE", "REMITENTE DIRECCION", "REMITENTE TELEFONO",
            "REMITENTE CIUDAD", "REMITENTE ESTADO", "NOMBRE DESTINO", "DESTINO DIRECCION",
            "DESTINO TELEFONO", "DESTINO CIUDAD", "CONTENIDO", "PESO LIBRAS", "PESO KILOS",
            "VALOR DECLARADO", "PIEZAS", "DESTINO ESTADO", "POSICION ARANCELARIA", "MANIFIESTO","INSTRUCCIONES",
            "CASILLERO"  # útil para lógica de manifiesto
        ]
        try:
            _, res = dbx.files_download(path)
            df = pd.read_excel(io.BytesIO(res.content), sheet_name=0)
            # asegurar columna guia si viene como GUIA u otra variante
            if "guia" not in df.columns and "GUIA" in df.columns:
                df = df.rename(columns={"GUIA": "guia"})
            return df
        except dropbox.exceptions.ApiError as e:
            # si no existe, arrancar vacío
            if getattr(e, "error", None) and e.error.is_path() and e.error.get_path().is_not_found():
                return pd.DataFrame(columns=base_cols)
            raise  # cualquier otro error sí lo mostramos
    df_historico = load_historico_or_empty(dbx, DBX_FILE_PATH)

    st.success(f"Histórico descargado: {df_historico.shape[0]} filas")

    # -----------------------------
    # 2) Leer archivos subidos (A y B)
    # -----------------------------
    df_a = pd.read_excel(up_a)
    df_b = pd.read_excel(up_b)

    st.info(f"A (Pistoleo): {df_a.shape} | B (Encargomio): {df_b.shape}")

    # -----------------------------
    # 3) Asegurar STR + limpieza .0 en llaves originales
    # -----------------------------
    df_a["Envio"] = _clean_str_series(df_a["Envio"])
    df_a = df_a.drop_duplicates(subset=["Envio"], keep="first").copy()
    
    df_b["NUMERO ENVIO"] = _clean_str_series(df_b["NUMERO ENVIO"])

    # -----------------------------
    # 4) Renombres remitente
    # -----------------------------
    rename_map = {
        "CLIENTE DESTINO": "NOMBRE DESTINO",
        "DIRECCIÓN DESTINO": "DESTINO DIRECCION",
        "TELÉFONO": "DESTINO TELEFONO",
        "CIUDAD DESTINO": "DESTINO CIUDAD",
        "DEPARTAMENTO DESTINO": "DESTINO ESTADO",
    }

    faltantes = [c for c in rename_map.keys() if c not in df_b.columns]
    if faltantes:
        st.error(f"Faltan columnas en B: {faltantes}")
        st.stop()

    df_b = df_b.rename(columns=rename_map)

    # -----------------------------


    

    df_b["COMPAÑÍA REMITENTE"] =  "Largo Easy Corp"
    df_b["REMITENTE DIRECCION"] = "11860 SW 144th Ct Ste 2"
    df_b["REMITENTE TELEFONO"] = "3053996614"
    df_b["REMITENTE CIUDAD"] = "Miami"
    df_b["REMITENTE ESTADO"] = "FL"

    
    


    # -----------------------------
    # 6) Renombrar llaves a guia
    # -----------------------------
    df_a = df_a.rename(columns={"Envio": "guia"})
    df_b = df_b.rename(columns={"NUMERO ENVIO": "guia"})

    df_a["guia"] = _clean_str_series(df_a["guia"])
    df_b["guia"] = _clean_str_series(df_b["guia"])
    if "CATEGORÍAS PRODUCTOS" in df_b.columns:
        df_b = df_b.rename(columns={"CATEGORÍAS PRODUCTOS": "CONTENIDO"})

    # -----------------------------
    # 7) Selección columnas B + PESO -> PESO LIBRAS
    # -----------------------------
    cols_b = [
        "guia",
        "CASILLERO",
        "COMPAÑÍA REMITENTE",
        "REMITENTE DIRECCION",
        "REMITENTE TELEFONO",
        "REMITENTE CIUDAD",
        "REMITENTE ESTADO",
        "NOMBRE DESTINO",
        "DESTINO DIRECCION",
        "DESTINO TELEFONO",
        "DESTINO CIUDAD",
        "DESTINO ESTADO",
        "PESO",
    ]
    falt_b = [c for c in cols_b if c not in df_b.columns]
    if falt_b:
        st.error(f"Faltan columnas requeridas en B para el cruce: {falt_b}")
        st.stop()

    df_b_sel = df_b[cols_b].rename(columns={"PESO": "PESO LIBRAS"})

    # -----------------------------
    # 8) Merge
    # -----------------------------
    df_final = df_a.merge(df_b_sel, how="left", on="guia")
    # -----------------------------
# 8.1) Columnas nuevas
# -----------------------------
    df_final["PIEZAS"] = 1
    df_final["VALOR DECLARADO"] = np.random.randint(91, 100, size=len(df_final))  # 91–99
    df_final["POSICION ARANCELARIA"] = "980720"
    df_final["POSICION ARANCELARIA"] = df_final["POSICION ARANCELARIA"].astype("string")

    # -----------------------------
    # 9) PESO KILOS + FECHA GUIA Miami
    # -----------------------------
    df_final["PESO LIBRAS"] = pd.to_numeric(df_final["PESO LIBRAS"], errors="coerce")
    df_final["PESO KILOS"] = df_final["PESO LIBRAS"] / 2.2

    hoy_miami = datetime.now(ZoneInfo("America/New_York")).date()
    df_final["FECHA GUIA"] = pd.to_datetime(hoy_miami)

    # -----------------------------
    # 10) Concat + dedup (histórico manda)
    # -----------------------------
    df_historico["guia"] = _clean_str_series(df_historico["guia"])
    df_final["guia"] = _clean_str_series(df_final["guia"])

    df_concat = pd.concat([df_historico, df_final], ignore_index=True)
    df_concat = df_concat.drop_duplicates(subset=["guia"], keep="first").reset_index(drop=True)

    # -----------------------------
    # 11) Crear MANIFIESTO solo a vacíos con regla 11591 vs otros
    # -----------------------------
    if "MANIFIESTO" not in df_concat.columns:
        df_concat["MANIFIESTO"] = pd.NA

    df_concat["CASILLERO_NORM"] = df_concat["CASILLERO"].apply(_norm_casillero)
    df_concat["MANIFIESTO_NUM"] = pd.to_numeric(df_concat["MANIFIESTO"], errors="coerce")

    mask_vacio = df_concat["MANIFIESTO_NUM"].isna()
    mask_11591 = df_concat["CASILLERO_NORM"] == "CA11591"

    max_11591 = df_concat.loc[~mask_vacio & mask_11591, "MANIFIESTO_NUM"].max()
    max_otros = df_concat.loc[~mask_vacio & ~mask_11591, "MANIFIESTO_NUM"].max()

    if pd.isna(max_11591):
        max_11591 = 900000
    if pd.isna(max_otros):
        max_otros = 100000

    nuevo_man_11591 = int(max_11591) + 1
    nuevo_man_otros = int(max_otros) + 1

    df_concat.loc[mask_vacio & mask_11591, "MANIFIESTO"] = nuevo_man_11591
    df_concat.loc[mask_vacio & ~mask_11591, "MANIFIESTO"] = nuevo_man_otros

    df_concat["MANIFIESTO"] = pd.to_numeric(df_concat["MANIFIESTO"], errors="coerce").astype("Int64")
    df_concat = df_concat.drop(columns=["CASILLERO_NORM", "MANIFIESTO_NUM"], errors="ignore")
    # -----------------------------
    # 11.5) Reordenar columnas para export
    # -----------------------------
    orden_cols = [
        "FECHA GUIA",
        "guia",
        "COMPAÑÍA REMITENTE",
        "REMITENTE DIRECCION",
        "REMITENTE TELEFONO",
        "REMITENTE CIUDAD",
        "REMITENTE ESTADO",
        "NOMBRE DESTINO",
        "DESTINO DIRECCION",
        "DESTINO TELEFONO",
        "DESTINO CIUDAD",
        "CONTENIDO",
        "PESO LIBRAS",
        "PESO KILOS",
        "VALOR DECLARADO",
        "PIEZAS",
        "DESTINO ESTADO",
        "POSICION ARANCELARIA",
        "MANIFIESTO",
        "INSTRUCCIONES",
        "CASILLERO"
    ]
    
    # solo deja las que existan (por si el histórico viejo trae extras)
    presentes = [c for c in orden_cols if c in df_concat.columns]
    extras = [c for c in df_concat.columns if c not in presentes]
    df_concat = df_concat[presentes + extras]
    # -----------------------------
    # X) COSTO (se guarda en Dropbox)
    # -----------------------------
    df_concat["PESO LIBRAS"] = pd.to_numeric(df_concat["PESO LIBRAS"], errors="coerce")
    w = df_concat["PESO LIBRAS"]
    
    df_concat["COSTO"] = np.where(
        w.isna() | (w <= 0),
        0,
        np.where(
            w <= 1,
            7.35,
            7.35 + (w - 1) * 2.6
        ) + np.where(w >= 20, 4.75, 0)
    ).round(2)
    
    
    
    # -----------------------------
    # X) Resumen COSTO por manifiesto (para hoja 2)
    # -----------------------------
    # Asegurar numéricos
    for col in ["PESO LIBRAS", "PESO KILOS", "PIEZAS", "COSTO"]:
        if col in df_concat.columns:
            df_concat[col] = pd.to_numeric(df_concat[col], errors="coerce")
    
    df_costos_manifiesto = (
        df_concat
        .groupby("MANIFIESTO", dropna=False, as_index=False)[["PESO LIBRAS", "PESO KILOS", "PIEZAS", "COSTO"]]
        .sum(min_count=1)
    )
    
    # (opcional) ordenar
    df_costos_manifiesto = df_costos_manifiesto.sort_values("MANIFIESTO", na_position="last")
    
    
    
    
    # =========================
    # PRODUCTOS -> PRODUCTOS_RAW (solo si CONTENIDO está vacío)
    # =========================
    
    # 0) Cargar archivo productos (si ya lo tienes, omite esta línea)
# =========================
# PRODUCTOS (df_p) -> resumen por Envio usando Peso
# y cruce SOLO para filas sin CONTENIDO
# =========================

# df_p ya leído (si no:)
    df_p = pd.read_excel(up_p)
    
    COL_ENVIO = "Envío"
    COL_PROD  = "Nombre producto"
    COL_PESO  = "Peso"
    
    df_prod = df_p[[COL_ENVIO, COL_PROD, COL_PESO]].copy()
    
    df_prod[COL_ENVIO] = _clean_str_series(df_prod[COL_ENVIO])
    df_prod[COL_PROD]  = df_prod[COL_PROD].astype("string").fillna("").str.strip()
    df_prod[COL_PESO]  = pd.to_numeric(df_prod[COL_PESO], errors="coerce").fillna(0)
    
    df_prod = df_prod[df_prod[COL_PROD] != ""].copy()
    
    # ✅ DEDUP real: colapsa repetidos (Envio, Producto) sumando Peso
    df_prod = (
        df_prod
        .groupby([COL_ENVIO, COL_PROD], as_index=False)[COL_PESO]
        .sum()
    )
    
    # Dominante (máximo peso total por envío)
    idx_dom = df_prod.groupby(COL_ENVIO)[COL_PESO].idxmax()
    df_dom = (
        df_prod.loc[idx_dom, [COL_ENVIO, COL_PROD, COL_PESO]]
        .rename(columns={COL_PROD: "PRODUCTO_DOMINANTE", COL_PESO: "PESO_DOMINANTE"})
    )
    
    # Lista (ordenada por peso desc, ya sin repetidos)
    df_list = (
        df_prod.sort_values([COL_ENVIO, COL_PESO], ascending=[True, False])
        .groupby(COL_ENVIO, as_index=False)[COL_PROD]
        .apply(lambda s: " | ".join(s.head(50).tolist()))
        .rename(columns={COL_PROD: "PRODUCTOS_LISTA"})
    )
    
    df_prod_agg = df_dom.merge(df_list, on=COL_ENVIO, how="left")
    
    # ----- cruzar solo los que NO tienen contenido -----
    df_concat_copy = df_concat.copy()
    
    df_concat_copy["CONTENIDO"] = df_concat_copy.get(
        "CONTENIDO", pd.Series([pd.NA] * len(df_concat_copy))
    ).astype("string")
    
    mask_sin_contenido = df_concat_copy["CONTENIDO"].isna() | (df_concat_copy["CONTENIDO"].str.strip() == "")
    
    df_to_ai = df_concat_copy.loc[mask_sin_contenido].copy()
    df_to_ai["guia"] = _clean_str_series(df_to_ai["guia"])
    
    df_to_ai = df_to_ai.merge(
        df_prod_agg,
        how="left",
        left_on="guia",
        right_on=COL_ENVIO
    )
    
    df_concat_copy.loc[mask_sin_contenido, "PRODUCTO_DOMINANTE"] = df_to_ai["PRODUCTO_DOMINANTE"].values
    df_concat_copy.loc[mask_sin_contenido, "PESO_DOMINANTE"] = df_to_ai["PESO_DOMINANTE"].values
    df_concat_copy.loc[mask_sin_contenido, "PRODUCTOS_LISTA"] = df_to_ai["PRODUCTOS_LISTA"].values
    
    def get_openai_client() -> OpenAI:
        return OpenAI(api_key=st.secrets["openai"]["api_key"])
    
    
    CATEGORIAS = [
        "Tenis", "Calzado", "Celular", "Computador", "Componente de computador",
        "Ropa", "Perfumes", "Cosmeticos", "Accesorios", "Reloj/Joyeria",
        "Hogar", "Electrodomestico", "Juguetes", "Herramientas",
        "Suplementos", "Medicamentos", "Alimentos",
        "Libros/Papeleria", "Miscelaneo"
    ]
    
    # ✅ ESTE es el JSON Schema REAL (raíz type=object)
    CLASIFICAR_ENVIO_SCHEMA = {
        "type": "object",
        "properties": {
            "categoria": {"type": "string", "enum": CATEGORIAS},
            "confianza": {"type": "integer", "minimum": 0, "maximum": 100},
            "contenido": {
                "type": "string",
                "description": "Texto corto para CONTENIDO (sin marcas). Ej: 'Calzado (tenis)', 'Celular', 'Ropa'."
            }
        },
        "required": ["categoria", "confianza", "contenido"],
        "additionalProperties": False
    }
    
    def gpt_clasificar_envio(client: OpenAI, producto_dominante: str, productos_lista: str) -> dict:
        prompt = f"""
    Eres un clasificador de envíos para manifiestos.
    Objetivo: llenar la columna CONTENIDO con una categoria GENERAL, sin marcas.
    
    Reglas:
    - Prioriza PRODUCTO_DOMINANTE (es el más pesado del envío).
    - Usa PRODUCTOS_LISTA solo como contexto.
    - NO menciones marcas (Nike, Apple, etc). Solo categoria general.
    - 'contenido' debe ser corto (1-4 palabras). Puedes aclarar entre paréntesis sin marcas.
    - Si es muy variado, usa Miscelaneo.
    
    PRODUCTO_DOMINANTE: {producto_dominante}
    PRODUCTOS_LISTA: {productos_lista}
    """.strip()
    
        resp = client.responses.create(
            model="gpt-4o-mini",
            input=prompt,
            text={
                "format": {
                    "type": "json_schema",
                    "name": "clasificar_envio",   # <=64 chars
                    "schema": CLASIFICAR_ENVIO_SCHEMA,
                    "strict": True
                }
            }
        )
    
        # ✅ resp.output_text ya es JSON válido
        return json.loads(resp.output_text)   

    client = get_openai_client()
    
    df_concat_copy["CONTENIDO"] = df_concat_copy.get(
        "CONTENIDO", pd.Series([pd.NA] * len(df_concat_copy))
    ).astype("string")
    
    mask_sin_contenido = df_concat_copy["CONTENIDO"].isna() | (df_concat_copy["CONTENIDO"].str.strip() == "")
    
    mask_listo_para_gpt = (
        mask_sin_contenido
        & df_concat_copy["PRODUCTO_DOMINANTE"].notna()
        & (df_concat_copy["PRODUCTO_DOMINANTE"].astype(str).str.strip() != "")
    )
    
    cache = {}
    
    cats, confs, conts = [], [], []
    
    for dom, lista in zip(
        df_concat_copy.loc[mask_listo_para_gpt, "PRODUCTO_DOMINANTE"].astype(str).tolist(),
        df_concat_copy.loc[mask_listo_para_gpt, "PRODUCTOS_LISTA"].fillna("").astype(str).tolist()
    ):
        lista = lista[:3000]  # límite de texto
        key = (dom, lista)
    
        if key in cache:
            out = cache[key]
        else:
            out = gpt_clasificar_envio(client, dom, lista)
            cache[key] = out
    
        cats.append(out["categoria"])
        confs.append(out["confianza"])
        conts.append(out["contenido"])
    
    df_concat_copy.loc[mask_listo_para_gpt, "CATEGORIA_GPT"] = cats
    df_concat_copy.loc[mask_listo_para_gpt, "CONF_GPT"] = confs
    df_concat_copy.loc[mask_listo_para_gpt, "CONTENIDO"] = conts
    
    # Si sigue sin match (no había productos), se queda vacío

    # =========================
    # PASAR CONTENIDO de df_concat_copy -> df_concat (por guia)
    # =========================
    
    # 1) Normalizar llaves (por si acaso)
    df_concat_copy["guia"] = _clean_str_series(df_concat_copy["guia"])
    df_concat["guia"] = _clean_str_series(df_concat["guia"])
    
    # 2) Asegurar CONTENIDO como string
    df_concat["CONTENIDO"] = df_concat.get("CONTENIDO", pd.Series([pd.NA]*len(df_concat))).astype("string")
    df_concat_copy["CONTENIDO"] = df_concat_copy.get("CONTENIDO", pd.Series([pd.NA]*len(df_concat_copy))).astype("string")
    
    # 3) Tomar SOLO filas donde df_concat_copy tiene contenido generado (no vacío)
    mask_copy_con = df_concat_copy["CONTENIDO"].notna() & (df_concat_copy["CONTENIDO"].str.strip() != "")
    
    # Si hay duplicados de guia en copy, me quedo con el último no vacío (por seguridad)
    df_map = (
        df_concat_copy.loc[mask_copy_con, ["guia", "CONTENIDO"]]
        .drop_duplicates(subset=["guia"], keep="last")
    )
    
    map_contenido = df_map.set_index("guia")["CONTENIDO"].to_dict()
    
    # 4) POBLAR en df_concat SOLO donde estaba vacío
    mask_concat_vacio = df_concat["CONTENIDO"].isna() | (df_concat["CONTENIDO"].str.strip() == "")
    
    df_concat.loc[mask_concat_vacio, "CONTENIDO"] = df_concat.loc[mask_concat_vacio, "guia"].map(map_contenido)
        
 
    

    st.success(f"Manifiestos asignados. Nuevo 11591={nuevo_man_11591} | Otros={nuevo_man_otros}")

    # -----------------------------
    # 12) Subir histórico actualizado (overwrite) a Dropbox
    # -----------------------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_concat.to_excel(writer, sheet_name="HISTORICO", index=False)
        df_costos_manifiesto.to_excel(writer, sheet_name="costo por manifiesto", index=False)
    output.seek(0)
    excel_bytes = output.getvalue()
    
    dbx.files_upload(excel_bytes, DBX_FILE_PATH, mode=dropbox.files.WriteMode.overwrite)

    st.success("Histórico actualizado en Dropbox (reemplazado) ✅")

    # (opcional) mostrar muestra
    with st.expander("Ver muestra del histórico resultante"):
        st.dataframe(df_concat.head(100))
    st.session_state["df_concat"] = df_concat
    st.session_state["fecha_str"] = datetime.now(ZoneInfo("America/New_York")).strftime("%Y-%m-%d")
        
        
        
        
import zipfile

st.divider()
st.subheader("Descargas por manifiesto")

if "df_concat" not in st.session_state:
    st.info("Primero ejecuta: **Procesar y actualizar histórico en Dropbox** para habilitar descargas.")
else:
    df_concat = st.session_state["df_concat"]
    fecha_str = st.session_state.get("fecha_str") or datetime.now(ZoneInfo("America/New_York")).strftime("%Y-%m-%d")

    # -----------------------------
    # Vista SOLO para descargas (no afecta Dropbox)
    # - omite CASILLERO
    # - ordena columnas
    # - renombra guia -> GUIA
    # -----------------------------
    df_dl = df_concat.copy()
    df_dl = df_dl.drop(columns=["CASILLERO"], errors="ignore")

    if "guia" in df_dl.columns:
        df_dl = df_dl.rename(columns={"guia": "GUIA"})

    orden_descarga = [
        "FECHA GUIA",
        "GUIA",
        "COMPAÑÍA REMITENTE",
        "REMITENTE DIRECCION",
        "REMITENTE TELEFONO",
        "REMITENTE CIUDAD",
        "REMITENTE ESTADO",
        "NOMBRE DESTINO",
        "DESTINO DIRECCION",
        "DESTINO TELEFONO",
        "DESTINO CIUDAD",
        "CONTENIDO",
        "PESO LIBRAS",
        "PESO KILOS",
        "VALOR DECLARADO",
        "PIEZAS",
        "DESTINO ESTADO",
        "POSICION ARANCELARIA",
        "MANIFIESTO",
        "INSTRUCCIONES"
    ]

    presentes = [c for c in orden_descarga if c in df_dl.columns]
    extras = [c for c in df_dl.columns if c not in presentes]
    df_dl = df_dl[presentes + extras]

    # Lista de manifiestos disponibles (ordenados, sin nulos)
    if "MANIFIESTO" not in df_dl.columns:
        st.warning("No existe la columna MANIFIESTO en el histórico para descargas.")
    else:
        manifiestos = (
            df_dl["MANIFIESTO"]
            .dropna()
            .astype("int64")
            .sort_values()
            .unique()
            .tolist()
        )

        def dfs_to_excel_bytes(sheets: dict) -> bytes:
            """
            sheets: dict {nombre_hoja: dataframe}
            """
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                for name, df in sheets.items():
                    safe_name = str(name)[:31]  # límite de Excel
                    df.to_excel(writer, sheet_name=safe_name, index=False)
            out.seek(0)
            return out.getvalue()
        
        
        def resumen_costo_por_manifiesto(df: pd.DataFrame) -> pd.DataFrame:
            """
            Suma PESO LIBRAS, PESO KILOS, PIEZAS, COSTO por MANIFIESTO (solo para el df recibido).
            """
            df2 = df.copy()
        
            for col in ["PESO LIBRAS", "PESO KILOS", "PIEZAS", "COSTO"]:
                if col in df2.columns:
                    df2[col] = pd.to_numeric(df2[col], errors="coerce")
        
            if "MANIFIESTO" not in df2.columns:
                return pd.DataFrame(columns=["MANIFIESTO", "PESO LIBRAS", "PESO KILOS", "PIEZAS", "COSTO"])
        
            out = (
                df2
                .groupby("MANIFIESTO", as_index=False)[["PESO LIBRAS", "PESO KILOS", "PIEZAS", "COSTO"]]
                .sum(min_count=1)
            )
            return out
        
        
        def build_zip_all_manifiestos(df_all: pd.DataFrame) -> bytes:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for man in (
                    df_all["MANIFIESTO"].dropna().astype("int64").sort_values().unique().tolist()
                ):
                    df_m = df_all[df_all["MANIFIESTO"].astype("Int64") == man].copy()
        
                    df_costos_m = resumen_costo_por_manifiesto(df_m)
        
                    excel_bytes = dfs_to_excel_bytes({
                        f"MAN_{man}": df_m,
                        "costo por manifiesto": df_costos_m
                    })
        
                    filename = f"{fecha_str}-{man}.xlsx"
                    zf.writestr(filename, excel_bytes)
        
            zip_buf.seek(0)
            return zip_buf.getvalue()
        
        
        col_all, col_one = st.columns([1, 1])
        
        with col_all:
            st.markdown("**Descargar todos (ZIP)**")
            if st.button("Preparar ZIP con todos los manifiestos"):
                zip_bytes = build_zip_all_manifiestos(df_dl)
                st.download_button(
                    "Descargar ZIP",
                    data=zip_bytes,
                    file_name=f"{fecha_str}-manifiestos.zip",
                    mime="application/zip",
                )
        
        with col_one:
            st.markdown("**Descargar uno puntual**")
            if not manifiestos:
                st.info("No hay manifiestos disponibles para descargar.")
            else:
                man_sel = st.selectbox("Buscar/seleccionar manifiesto", options=manifiestos)
        
                df_sel = df_dl[df_dl["MANIFIESTO"].astype("Int64") == int(man_sel)].copy()
                df_costos_sel = resumen_costo_por_manifiesto(df_sel)
        
                excel_sel = dfs_to_excel_bytes({
                    f"MAN_{man_sel}": df_sel,
                    "costo por manifiesto": df_costos_sel
                })
        
                st.download_button(
                    f"Descargar {fecha_str}-{man_sel}.xlsx",
                    data=excel_sel,
                    file_name=f"{fecha_str}-{man_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )