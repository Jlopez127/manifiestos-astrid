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

col1, col2 = st.columns(2)
with col1:
    up_a = st.file_uploader("Sube Pistoleo de bodega (Envios pistoleo)", type=["xlsx", "xls"])
with col2:
    up_b = st.file_uploader("Sube Envíos Encargomio (Envios_Encargomio)", type=["xlsx", "xls"])

run = st.button("Procesar y actualizar histórico en Dropbox", type="primary", disabled=not (up_a and up_b))

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
    df_b["NUMERO ENVIO"] = _clean_str_series(df_b["NUMERO ENVIO"])

    # -----------------------------
    # 4) Renombres remitente
    # -----------------------------
    rename_map = {
        "CLIENTE": "COMPAÑÍA REMITENTE",
        "DIRECCIÓN DESTINO": "REMITENTE DIRECCION",
        "TELÉFONO": "REMITENTE TELEFONO",
        "CIUDAD DESTINO": "REMITENTE CIUDAD",
        "DEPARTAMENTO DESTINO": "REMITENTE ESTADO",
    }

    faltantes = [c for c in rename_map.keys() if c not in df_b.columns]
    if faltantes:
        st.error(f"Faltan columnas en B: {faltantes}")
        st.stop()

    df_b = df_b.rename(columns=rename_map)

    # -----------------------------
    # 5) Duplicados destino + destino estado
    # -----------------------------
    df_b["NOMBRE DESTINO"] = df_b["COMPAÑÍA REMITENTE"]
    df_b["DESTINO DIRECCION"] = df_b["REMITENTE DIRECCION"]
    df_b["DESTINO TELEFONO"] = df_b["REMITENTE TELEFONO"]
    df_b["DESTINO CIUDAD"] = df_b["REMITENTE CIUDAD"]
    df_b["DESTINO ESTADO"] = df_b["REMITENTE ESTADO"]

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

    st.success(f"Manifiestos asignados. Nuevo 11591={nuevo_man_11591} | Otros={nuevo_man_otros}")

    # -----------------------------
    # 12) Subir histórico actualizado (overwrite) a Dropbox
    # -----------------------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_concat.to_excel(writer, sheet_name="HISTORICO", index=False)
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

        def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "DATA") -> bytes:
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            out.seek(0)
            return out.getvalue()

        def build_zip_all_manifiestos(df_all: pd.DataFrame) -> bytes:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for man in (
                    df_all["MANIFIESTO"].dropna().astype("int64").sort_values().unique().tolist()
                ):
                    df_m = df_all[df_all["MANIFIESTO"].astype("Int64") == man].copy()
                    excel_bytes = df_to_excel_bytes(df_m, sheet_name=f"MAN_{man}")
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
                excel_sel = df_to_excel_bytes(df_sel, sheet_name=f"MAN_{man_sel}")

                st.download_button(
                    f"Descargar {fecha_str}-{man_sel}.xlsx",
                    data=excel_sel,
                    file_name=f"{fecha_str}-{man_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )