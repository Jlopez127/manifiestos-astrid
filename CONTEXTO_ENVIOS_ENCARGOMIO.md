# Contexto — Módulo "Envios Encargomio" (modelo final)

Documento de referencia del módulo **Envios Encargomio** dentro de `Manifiestos_astrid.py`
(app Streamlit). Refleja el estado actual tras la serie de cambios (pistoleo nuevo,
remitente/destino simulados, modo "agregar a existente" y doble descarga DIAN/Astrid).

## 1. Qué es y dónde vive
- Archivo único: `Manifiestos_astrid.py` (Streamlit). Selector `st.radio` con 3 procesos:
  **Manifiestos Luma**, **Celulares Fénix**, **Envios Encargomio**. Cada uno es un branch `if/elif`.
- Persistencia en **Dropbox**; clasificación de CONTENIDO con **OpenAI** (solo en línea).
- **Regla de oro:** un envío = una fila. Llave `guia` (número de envío). Dedup contra el
  histórico por `guia`, `keep="first"` → **el histórico manda** (un envío ya registrado no se
  reprocesa). NO hay multi-caja ni sufijos.

## 2. Entradas (3 archivos)
- **A — Pistoleo** (`leer_pistoleo_envio`): Excel con `Envio`, `Valor declarado`,
  `Posicion arancelaria`, `Descripción` (+ columnas basura `Unnamed` que se descartan).
  Devuelve SOLO 4 columnas normalizadas: `Envio`, `VALOR_A`, `POSICION_A`, `DESCRIPCION_A`.
  Tolera alias de Envio y archivo de una sola columna sin encabezado. NO deduplica.
- **B — Envíos Encargomio**: `NUMERO ENVIO`, `CASILLERO`, `PESO`, destinatario real
  (`CLIENTE DESTINO`, `DIRECCIÓN DESTINO`, `CIUDAD DESTINO`, `DEPARTAMENTO DESTINO`, `TELÉFONO`), `VALOR`.
- **P — Productos / Reporte Casillero** (opcional): `Envío`, `Nombre producto`, `Peso`. Solo para IA.
- **Llave de cruce:** `A.Envio == B.NUMERO ENVIO == P.Envío` (se limpian con `_clean_str_series`).

## 3. `procesar_envios_encargomio(df_a, df_b, hist, df_p=None, manifiesto_destino=None)`
Devuelve `(df_concat, manifiesto_usado, n_ia, guias_sin_peso, sin_categoria)`.
1. Limpia `Envio` → `guia`; **dedup de A por `guia`**.
2. Renombra B a los nombres del manifiesto; llave `guia`. Dedup B por `guia` (defensivo).
3. **Merge A×B por `guia`** (trae destinatario/peso/casillero).
4. Columnas calculadas por fila:
   - **VALOR DECLARADO** = `VALOR_A` si es número válido **y ≠ 0**; si es NaN o 0 → random 90–199.
   - **POSICION ARANCELARIA** = `POSICION_A` si matchea `^\d{4}\.\d` (código válido); cualquier
     otro caso (0, Comercial, Regular, vacío/NaN) → `"980720"`.
   - **PESO KILOS** = `PESO LIBRAS / 2.20462`. **PIEZAS** = 1. **FECHA GUIA** = hoy Miami.
   - **CONTENIDO** = `DESCRIPCION_A` (texto) o NA (vacío → lo llena la IA / garantías).
   - **Remitente SIMULADO por fila** (`_remitente_simulado_miami`): patrón Miami — nombre+apellido
     (Faker en_US), dirección `{4-5 díg} {cuadrante SW/NW/W/SE/E/NE} {ordinal} {vía St/Ct/Av/Terr}`,
     teléfono 10 díg que empieza en **305/786**, ciudad Miami, estado FL. (Reemplazó el remitente fijo.)
5. **Numeración / modo:**
   - `manifiesto_destino is None` → **CREAR NUEVO**: filas nuevas reciben `max(≥2000001)+1` (o 2000001).
   - `manifiesto_destino = N` → **AGREGAR**: las nuevas heredan `N`; el contador NO avanza.
6. **`FECHA_MANIFIESTO`** = hoy (Miami, fecha sola) sellada en las filas nuevas; las viejas conservan la suya.
7. Concat con histórico + **dedup por `guia`** (histórico manda).
8. **IA** (si hay P): `enriquecer_contenido_ia` reclasifica CONTENIDO vacío/genérico del lote nuevo.
9. **Garantía anti-"Otro"** (`_es_categoria_generica` + `_limpiar_cantidad`): nunca "Otro" ni cantidades;
   sin categoría → blanco + aviso.
10. **Simulación de destino** (ver §4) sobre las filas del manifiesto usado.

## 4. Simulación de destino — `simular_destinos_manifiesto(df, filas_a_evaluar_idx=None)`
- Regla: dentro de UN manifiesto, por cada **CASILLERO** la **primera** aparición conserva destino
  REAL; las siguientes reciben destino colombiano simulado (`_build_destinatario_fake`, Faker es_CO).
- La simulación va en columnas `*_SIM` (`NOMBRE DESTINO_SIM`, `DESTINO DIRECCION_SIM`,
  `DESTINO CIUDAD_SIM`, `DESTINO ESTADO_SIM`, `DESTINO TELEFONO_SIM`) + bandera `_DESTINO_SIMULADO` (bool).
  El destino REAL nunca se pierde.
- `filas_a_evaluar_idx` = filas NUEVAS (simulables). Las CONGELADAS (viejas) no se re-tocan, pero
  cuentan para "casillero ya visto" (solo si su aparición fue real). Casillero vacío/NaN nunca se simula.
- **Integración en procesar:** se simula el subconjunto `MANIFIESTO == manifiesto_usado` (en agregar
  incluye viejas + nuevas), evaluando SOLO las nuevas (`guias_nuevas`). Así, en "agregar", una fila
  nueva se simula si su casillero ya aparecía en ese manifiesto (viejo o nuevo), y las viejas quedan
  byte-idénticas.

## 5. Descarga — doble: DIAN y Astrid (branch UI, sección de descarga)
Se elige un manifiesto del histórico. Ambos parten de `df_m` (sus filas). Formato cajas =
`build_info_manifiesto_df` (`INFO_MANIFIESTO_COLS`) + totales (`info_manifiesto_excel_bytes`).
- **DIAN** (`{fecha}-{man}-DIAN-CAJAS.xlsx`): en las filas con `_DESTINO_SIMULADO == True`, el destino
  real se reemplaza por su `*_SIM`; las demás quedan reales. Formato cajas estándar (SIN casillero). Con totales.
- **Astrid** (`{fecha}-{man}-ASTRID-CAJAS.xlsx`): destino **REAL** siempre + columna **CASILLERO al final**.
  Las filas simuladas se **pintan de rojo** (`_paint_rows_red_by_guia`, FF4D4F) abriendo con openpyxl el
  Excel que ya trae totales → conserva totales + rojo.
- Consistencia garantizada: guía simulada → `DIAN.destino ≠ Astrid.destino`; guía real → `==`. Remitente
  idéntico en ambos. Totales (W.Lb/W.Kg/Value/Pieces) iguales (cambia el destino, no el peso/valor).

## 6. Modos en la UI (branch `elif modo == "Envios Encargomio"`)
- `st.radio`: **"Crear manifiesto nuevo"** vs **"Agregar a manifiesto existente"**.
- En "Agregar": `st.selectbox` con el **último** manifiesto (número más alto), como `"{FECHA} — {MANIFIESTO}"`;
  si el histórico está vacío, se deshabilita con aviso. Se pasa `manifiesto_destino` a `procesar`.
- Mensaje de éxito adaptado: "creó manifiesto nuevo N con X envíos" / "agregó X envíos al manifiesto N".

## 7. Formato cajas (INFO MANIFIESTO)
`INFO_MANIFIESTO_COLS`: `MASTER` (vacía), `FECHA GUIA`, `GUIA / consignment` (=guia), `COMPAÑÍA REMITENTE`,
`REMITENTE DIRECCION`, `REMITENTE CIUDAD`, `REMITENTE ESTADO`, `NOMBRE DESTINO`, `DESTINO DIRECCION`,
`DESTINO CIUDAD`, `CONTENIDO`, `PESO LIBRAS`, `PESO KILOS`, `VALOR DECLARADO`, `PIEZAS`, `DESTINO ESTADO`,
`POSICION ARANCELARIA`, `MANIFIESTO`. Sin teléfonos ni INSTRUCCIONES. Pie de TOTALES (W.Lb/W.Kg/Value/Pieces)
por manifiesto. (Astrid añade `CASILLERO` al final.)

## 8. Persistencia (Dropbox)
- Histórico Envios Encargomio: **`/Manifiestos/Envios_encargomio.xlsx`** (hoja `HISTORICO`), aparte de
  Luma (`/Manifiestos/Manifiestos_astrid.xlsx`) y Fénix (`/Manifiestos/Celulares_fenix.xlsx`).
- Columnas guardadas: las del manifiesto + `guia`, `MANIFIESTO`, `FECHA_MANIFIESTO`, `CASILLERO`,
  `_DESTINO_SIMULADO`, las `*_SIM`, y las crudas del pistoleo (`VALOR_A`, `POSICION_A`, `DESCRIPCION_A`).
- Credenciales: `get_dbx()` con `st.secrets["dropbox"]`. OpenAI: `st.secrets["openai"]["api_key"]`.

## 9. Reglas de negocio (resumen)
- kg = lb / **2.20462**. Numeración propia desde **2000001** (+1 por corrida; agregar no avanza).
- **VALOR** desde el pistoleo (válido y ≠0) o random 90–199. **POSICION** código válido o `980720`.
- **CONTENIDO** desde el pistoleo (DESCRIPCION_A) o IA; **nunca "Otro" ni cantidades**.
- **Remitente** simulado Miami por fila. **Destino** simulado por casillero (1ª real, resto _SIM).
- Histórico manda (dedup por `guia`); las simulaciones viejas nunca se re-simulan.

## 10. Funciones clave
`leer_pistoleo_envio`, `procesar_envios_encargomio`, `simular_destinos_manifiesto`,
`_remitente_simulado_miami` / `_ordinal_en`, `build_info_manifiesto_df`, `info_manifiesto_excel_bytes`,
`enriquecer_contenido_ia`, `_es_categoria_generica`, `_limpiar_cantidad`, `_paint_rows_red_by_guia`,
`_build_destinatario_fake`, `_norm_casillero`. Constantes: `INFO_MANIFIESTO_COLS`, `_DESTINO_SIM_COLS`.

## 11. Notas / pendientes
- La IA solo corre en línea (necesita la key en `st.secrets`).
- El teléfono del destinatario no va en el formato cajas (por eso no sale en las etiquetas 4×6 del
  repo `pdf_manifiestos`, que consume la columna `GUIA / consignment`).
- Validado con batería de 40 pruebas en memoria (6 escenarios, incl. idempotencia y agregar sobre
  casillero ya simulado): 40/40 PASS, sin Dropbox ni IA.
