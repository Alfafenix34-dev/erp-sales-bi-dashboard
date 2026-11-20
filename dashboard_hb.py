import io
import os
import re
import time
import glob
import calendar
import traceback
from datetime import datetime
import pandas as pd
from dotenv import load_dotenv
from flask import (
    Flask,
    jsonify,
    request,
    render_template,
    session,
    redirect,
    url_for,
    send_file,
    send_from_directory,
)
import database_hb as db
import reporte_ventas_hb as rpa

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

CARPETA_REPORTES = BASE_DIR
CARPETA_STATIC = os.path.join(BASE_DIR, "static")
os.makedirs(CARPETA_STATIC, exist_ok=True)

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=CARPETA_STATIC,
    static_url_path="/static",
)
app.secret_key = os.getenv("FLASK_SECRET_KEY") or os.urandom(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
)

db.init_db()

_CACHE_EXCEL = {}
_INTENTOS_LOGIN = {}

COLS_VENTAS = [
    "Asesor",
    "Total",
    "Descripción Artículo",
    "Código Artículo",
    "Marca",
    "Grupo",
    "Código Cliente",
    "Unidades",
]

COLS_DEV = [
    "Asesor",
    "Código Cliente",
    "Cliente",
    "Código Artículo",
    "Artículo",
    "Documento",
    "Motivo",
    "Unidades",
    "Total",
]

DIAS_SEMANA_ES = {
    0: "LUNES",
    1: "MARTES",
    2: "MIERCOLES",
    3: "JUEVES",
    4: "VIERNES",
    5: "SABADO",
    6: "DOMINGO",
}


def normalizar_codigo(serie):
    return (
        serie.fillna("")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )


def normalizar_texto(texto):
    return re.sub(r"\s+", " ", str(texto or "")).strip().upper()


def quitar_tildes(texto):
    reemplazos = str.maketrans("ÁÉÍÓÚÜ", "AEIOUU")
    return normalizar_texto(texto).translate(reemplazos)


def limpiar_telefono_wa(tel):
    digitos = re.sub(r"\D", "", str(tel or ""))
    if digitos.endswith("0") and len(digitos) == 11:
        digitos = digitos[:-1]
    return "57" + digitos if len(digitos) == 10 and digitos.startswith("3") else ""


def construir_llave_ac(df):
    return df["Asesor"].astype(str) + "||" + df["Código Cliente"].astype(str)


def calcular_metricas_dias_habiles(periodo_mes):
    try:
        anio, mes = map(int, periodo_mes.split("-"))
    except Exception:
        hoy_dt = datetime.now()
        anio, mes = hoy_dt.year, hoy_dt.month
        periodo_mes = f"{anio:04d}-{mes:02d}"

    ultimo_dia = calendar.monthrange(anio, mes)[1]
    hoy = datetime.now().date()
    es_mes_actual = (hoy.year == anio and hoy.month == mes)
    festivos = getattr(rpa, "FESTIVOS_CO", set())

    habiles_totales = 0
    habiles_transcurridos = 0
    fechas_habiles = []

    for dia_num in range(1, ultimo_dia + 1):
        fecha_actual = datetime(anio, mes, dia_num).date()
        fecha_str = fecha_actual.strftime("%Y-%m-%d")
        if fecha_actual.weekday() == 6 or fecha_str in festivos:
            continue
        habiles_totales += 1
        fechas_habiles.append(fecha_str)
        if not es_mes_actual or fecha_actual <= hoy:
            habiles_transcurridos += 1

    habiles_restantes = max(0, habiles_totales - habiles_transcurridos)
    return {
        "periodo_mes": periodo_mes,
        "habiles_totales": habiles_totales,
        "habiles_transcurridos": habiles_transcurridos,
        "habiles_restantes": habiles_restantes,
        "pct_tiempo": (habiles_transcurridos / habiles_totales * 100) if habiles_totales > 0 else 100.0,
        "fechas_habiles": fechas_habiles,
    }


def cargar_archivos_por_prefijo(prefijo, cols_esperadas):
    archivos = glob.glob(os.path.join(CARPETA_REPORTES, f"{prefijo}_*.xlsx"))
    archivos_diarios = [
        r for r in archivos if "_mensual_" not in os.path.basename(r).lower()
    ]

    fechas_archivos = {}
    for ruta in archivos_diarios:
        m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(ruta))
        if m:
            fecha_str = m.group(1)
            mtime = os.path.getmtime(ruta)
            if fecha_str not in fechas_archivos or mtime > fechas_archivos[fecha_str][1]:
                fechas_archivos[fecha_str] = (ruta, mtime)

    firma = tuple((info[0], info[1]) for info in sorted(fechas_archivos.values()))
    if _CACHE_EXCEL.get(prefijo, {}).get("firma") == firma:
        return _CACHE_EXCEL[prefijo]["df"].copy()

    frames = []
    for fecha_rep, (ruta, _) in sorted(fechas_archivos.items()):
        try:
            df = pd.read_excel(ruta, sheet_name=0)
            df.columns = [str(c).strip() for c in df.columns]
            for c in cols_esperadas:
                if c not in df.columns:
                    df[c] = 0 if c in ("Total", "Unidades") else ""
            df = df[cols_esperadas].copy()
            df["Fecha"] = fecha_rep
            df["Mes"] = fecha_rep[:7]
            df["Asesor"] = df["Asesor"].fillna("").astype(str).apply(normalizar_texto)
            df["Código Cliente"] = normalizar_codigo(df["Código Cliente"])
            df["Código Artículo"] = normalizar_codigo(df["Código Artículo"])
            if "Marca" in df.columns:
                df["Marca"] = df["Marca"].fillna("").astype(str).apply(normalizar_texto)
            if "Grupo" in df.columns:
                df["Grupo"] = df["Grupo"].fillna("").astype(str).apply(normalizar_texto)
            df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
            df["Unidades"] = pd.to_numeric(df["Unidades"], errors="coerce").fillna(0)
            df = df[df["Asesor"] != ""]
            df["LlaveAC"] = construir_llave_ac(df)
            frames.append(df)
        except Exception as e:
            print(f"[ERROR ETL] No se pudo leer {ruta}: {e}")

    resultado = (
        pd.concat(frames, ignore_index=True)
        if frames
        else pd.DataFrame(columns=cols_esperadas + ["Fecha", "Mes", "LlaveAC"])
    )
    _CACHE_EXCEL[prefijo] = {"firma": firma, "df": resultado}
    return resultado.copy()


def cargar_maestra_cruda():
    candidatos = [
        os.path.join(CARPETA_REPORTES, "MAESTRA_HB.XLS"),
        os.path.join(CARPETA_REPORTES, "MAESTRA_HB.xlsx"),
    ]
    ruta_maestra = next((p for p in candidatos if os.path.exists(p)), None)
    if not ruta_maestra:
        print("[ALERTA MAESTRA] No se encontró MAESTRA_HB.XLS en el directorio.")
        return pd.DataFrame()

    firma = (ruta_maestra, os.path.getmtime(ruta_maestra))
    if _CACHE_EXCEL.get("maestra", {}).get("firma") == firma:
        return _CACHE_EXCEL["maestra"]["df"].copy()

    try:
        df = pd.read_excel(ruta_maestra, sheet_name=0)
        df.columns = [str(c).strip() for c in df.columns]
        requeridas = {"Cod Cliente", "Cliente", "Asesor"}
        faltantes = requeridas - set(df.columns)
        if faltantes:
            raise ValueError(
                f"Columnas obligatorias faltantes en Maestra: {', '.join(sorted(faltantes))}"
            )

        col_dia = next(
            (c for c in df.columns if quitar_tildes(c) == "DIA"),
            None,
        )

        res = pd.DataFrame()
        res["Código Cliente"] = normalizar_codigo(df["Cod Cliente"])
        negocio = df["Cliente"].fillna("").astype(str).str.strip()
        contacto = (
            df["Nombre Común"].fillna("").astype(str).str.strip()
            if "Nombre Común" in df.columns
            else ""
        )
        res["Cliente"] = negocio.where(contacto == "", negocio + " (" + contacto + ")")
        res["Asesor"] = (
            df["Asesor"].fillna("SIN ASIGNAR").astype(str).apply(normalizar_texto)
        )
        res["Grupo Maestra"] = (
            df["Grupo"].fillna("").astype(str).apply(normalizar_texto)
            if "Grupo" in df.columns
            else ""
        )
        res["Día"] = (
            df[col_dia].fillna("").astype(str).apply(quitar_tildes)
            if col_dia
            else ""
        )
        res["Frecuencia"] = (
            df["Frecuencia"].fillna("").astype(str).apply(normalizar_texto)
            if "Frecuencia" in df.columns
            else ""
        )
        res["Teléfono"] = (
            normalizar_codigo(df["Telefono"]) if "Telefono" in df.columns else ""
        )
        res["Dirección"] = (
            df["Dirección"].fillna("").astype(str).str.strip()
            if "Dirección" in df.columns
            else ""
        )
        res["Barrio"] = (
            df["Barrio"].fillna("").astype(str).str.strip()
            if "Barrio" in df.columns
            else ""
        )
        resultado = res[res["Código Cliente"] != ""].copy()
        resultado["LlaveAC"] = construir_llave_ac(resultado)
        _CACHE_EXCEL["maestra"] = {"firma": firma, "df": resultado}
        return resultado.copy()
    except Exception as e:
        print(f"[ERROR CRÍTICO MAESTRA] Fallo al procesar {ruta_maestra}: {e}")
        return pd.DataFrame()


@app.before_request
def verificar_sesion_activa():
    rutas_libres = {"login", "logout", "servir_static", "static"}
    if request.endpoint in rutas_libres or not request.endpoint:
        return
    uid = session.get("user_id")
    if not uid:
        if request.path.startswith("/api/"):
            return jsonify({"error": "No autorizado"}), 401
        return redirect(url_for("login"))

    user_db = db.obtener_usuario_por_id(uid)
    if not user_db:
        session.clear()
        if request.path.startswith("/api/"):
            return jsonify({"error": "Sesión expirada"}), 401
        return redirect(url_for("login"))

    request.current_user = user_db
    if user_db["debe_cambiar_clave"] and request.endpoint != "cambiar_clave":
        return redirect(url_for("cambiar_clave"))


@app.route("/static/<path:filename>")
def servir_static(filename):
    return send_from_directory(CARPETA_STATIC, filename)


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    ip = request.remote_addr or "local"
    intentos, bloqueo_hasta = _INTENTOS_LOGIN.get(ip, [0, 0])

    if time.time() < bloqueo_hasta:
        error = "Demasiados intentos fallidos. Espera 2 minutos."
    elif request.method == "POST":
        user = db.autenticar_usuario(
            request.form.get("username", ""), request.form.get("password", "")
        )
        if user:
            _INTENTOS_LOGIN.pop(ip, None)
            session.clear()
            session["user_id"] = user["id"]
            if user["debe_cambiar_clave"]:
                return redirect(url_for("cambiar_clave"))
            return redirect(url_for("index"))
        intentos += 1
        bloqueo = time.time() + 120 if intentos >= 5 else 0
        _INTENTOS_LOGIN[ip] = [0 if bloqueo else intentos, bloqueo]
        error = "Usuario o contraseña incorrectos."

    logo_existe = os.path.exists(os.path.join(CARPETA_STATIC, "logo_hb.png"))
    return render_template("login.html", error=error, logo_existe=logo_existe)


@app.route("/cambiar_clave", methods=["GET", "POST"])
def cambiar_clave():
    error = ""
    if request.method == "POST":
        nueva = request.form.get("nueva", "").strip()
        confirmar = request.form.get("confirmar", "").strip()
        if len(nueva) < 8:
            error = "La contraseña debe tener al menos 8 caracteres."
        elif nueva != confirmar:
            error = "Las contraseñas no coinciden."
        else:
            db.actualizar_contrasena_propia(request.current_user["id"], nueva)
            return redirect(url_for("index"))
    return render_template(
        "cambiar_clave.html", user=request.current_user, error=error
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def aplicar_permisos_rol(df_v, df_d, df_m_raw, user):
    rol = user.get("rol", "asesor")
    if rol == "admin":
        return df_v, df_d, df_m_raw

    asesores_permitidos = set()
    if rol == "asesor":
        ase = normalizar_texto(user.get("asesor_abako", ""))
        if ase:
            asesores_permitidos.add(ase)
    elif rol == "supervisor":
        for a in user.get("asesores", []):
            asesores_permitidos.add(normalizar_texto(a))
        grupos_sup = {normalizar_texto(g) for g in user.get("grupos", [])}
        if grupos_sup and not df_m_raw.empty:
            asesores_permitidos.update(
                df_m_raw[df_m_raw["Grupo Maestra"].isin(grupos_sup)]["Asesor"].unique()
            )

    if not asesores_permitidos:
        return df_v.iloc[0:0], df_d.iloc[0:0], df_m_raw.iloc[0:0]

    return (
        df_v[df_v["Asesor"].isin(asesores_permitidos)].copy(),
        df_d[df_d["Asesor"].isin(asesores_permitidos)].copy(),
        df_m_raw[df_m_raw["Asesor"].isin(asesores_permitidos)].copy(),
    )


def construir_dataset_filtrado(args, user):
    df_v_all = cargar_archivos_por_prefijo("Ventas_Netas_HB", COLS_VENTAS)
    df_d_all = cargar_archivos_por_prefijo("Devoluciones_HB", COLS_DEV)
    df_m_all = cargar_maestra_cruda()
    df_v, df_d, df_m_raw = aplicar_permisos_rol(df_v_all, df_d_all, df_m_all, user)

    fechas_disp = sorted(
        set(df_v["Fecha"].unique()).union(set(df_d["Fecha"].unique())), reverse=True
    )
    meses_disp = sorted(
        set(df_v["Mes"].unique()).union(set(df_d["Mes"].unique())), reverse=True
    )
    if not fechas_disp:
        return {"vacio": True}

    periodo = args.get("periodo", "dia")
    valor_periodo = args.get("valor_periodo", "")
    asesor = normalizar_texto(args.get("asesor", ""))
    marca = normalizar_texto(args.get("marca", ""))
    grupo = normalizar_texto(args.get("grupo", ""))
    dia_ruta = quitar_tildes(args.get("dia_ruta", ""))
    cliente = args.get("cliente", "").strip()
    articulo = args.get("articulo", "").strip().lower()

    if periodo == "mes":
        target = valor_periodo if valor_periodo in meses_disp else meses_disp[0]
        v_f = df_v[df_v["Mes"] == target].copy()
        d_f = df_d[df_d["Mes"] == target].copy()
    else:
        target = valor_periodo if valor_periodo in fechas_disp else fechas_disp[0]
        v_f = df_v[df_v["Fecha"] == target].copy()
        d_f = df_d[df_d["Fecha"] == target].copy()

    grupos_opt = (
        sorted([g for g in df_m_raw["Grupo Maestra"].dropna().unique() if str(g).strip()])
        if not df_m_raw.empty
        else []
    )
    asesores_opt = sorted(
        set(v_f["Asesor"].unique()).union(set(d_f["Asesor"].unique()))
    )
    marcas_opt = sorted(
        [m for m in v_f["Marca"].dropna().unique() if str(m).strip()]
    )

    mapa_nombres = {}
    mapa_grupo_ac = {}
    mapa_grupo_asesor = {}
    mapa_tel_cliente = {}
    if not df_m_raw.empty:
        mapa_nombres.update(dict(zip(df_m_raw["Código Cliente"], df_m_raw["Cliente"])))
        mapa_tel_cliente.update(dict(zip(df_m_raw["Código Cliente"], df_m_raw["Teléfono"])))
        grp_ac_m = df_m_raw.groupby("LlaveAC", as_index=False).agg(
            {"Grupo Maestra": "first"}
        )
        mapa_grupo_ac = dict(zip(grp_ac_m["LlaveAC"], grp_ac_m["Grupo Maestra"]))

        grp_ase_m = df_m_raw[df_m_raw["Grupo Maestra"] != ""].groupby("Asesor", as_index=False).agg(
            {"Grupo Maestra": lambda s: " / ".join(sorted(set(x for x in s if x)))}
        )
        mapa_grupo_asesor = dict(zip(grp_ase_m["Asesor"], grp_ase_m["Grupo Maestra"]))

    if not df_d.empty:
        valid_d = df_d[df_d["Cliente"].astype(str).str.strip() != ""]
        mapa_nombres.update(
            dict(zip(valid_d["Código Cliente"], valid_d["Cliente"].astype(str)))
        )

    dia_fecha_actual = ""
    if periodo == "dia":
        try:
            dia_fecha_actual = DIAS_SEMANA_ES.get(
                datetime.strptime(target, "%Y-%m-%d").weekday(), ""
            )
        except Exception:
            dia_fecha_actual = ""

    dia_rutero_no_atendidos = ""
    if dia_ruta == "AUTO":
        dia_rutero_no_atendidos = dia_fecha_actual
    elif dia_ruta:
        dia_rutero_no_atendidos = dia_ruta

    if grupo and not df_m_raw.empty:
        asesores_del_grupo = set(
            df_m_raw[df_m_raw["Grupo Maestra"] == grupo]["Asesor"].unique()
        )
        v_f = v_f[v_f["Asesor"].isin(asesores_del_grupo)]
        d_f = d_f[d_f["Asesor"].isin(asesores_del_grupo)]

    if asesor:
        v_f = v_f[v_f["Asesor"] == asesor]
        d_f = d_f[d_f["Asesor"] == asesor]

    if marca:
        v_f = v_f[v_f["Marca"] == marca]

    if cliente:
        cods_match = {
            k for k, v in mapa_nombres.items() if cliente.lower() in v.lower()
        }
        v_f = v_f[
            v_f["Código Cliente"].str.contains(cliente, case=False, na=False)
            | v_f["Código Cliente"].isin(cods_match)
        ]
        d_f = d_f[
            d_f["Código Cliente"].str.contains(cliente, case=False, na=False)
            | d_f["Código Cliente"].isin(cods_match)
        ]
    if articulo:
        v_f = v_f[
            v_f["Descripción Artículo"]
            .astype(str)
            .str.lower()
            .str.contains(articulo, na=False)
            | v_f["Código Artículo"].str.lower().str.contains(articulo, na=False)
        ]
        d_f = d_f[
            d_f["Artículo"].astype(str).str.lower().str.contains(articulo, na=False)
            | d_f["Código Artículo"].str.lower().str.contains(articulo, na=False)
        ]

    llaves_atendidas_set = set(v_f["LlaveAC"].unique())
    clientes_atendidos_count = (
        int(v_f[["Asesor", "Código Cliente"]].drop_duplicates().shape[0])
        if not v_f.empty
        else 0
    )

    tiene_maestra = not df_m_raw.empty
    no_atendidos_df = pd.DataFrame()
    m_agrup = pd.DataFrame()
    total_no_atendidos = 0
    cobertura_pct = 0.0

    if tiene_maestra:
        m_f = df_m_raw.copy()
        if grupo:
            m_f = m_f[m_f["Grupo Maestra"] == grupo]
        if asesor:
            m_f = m_f[m_f["Asesor"] == asesor]
        if dia_rutero_no_atendidos:
            m_f = m_f[m_f["Día"] == dia_rutero_no_atendidos]
        if cliente:
            m_f = m_f[
                m_f["Código Cliente"].str.contains(cliente, case=False, na=False)
                | m_f["Cliente"].str.contains(cliente, case=False, na=False)
            ]

        m_agrup = m_f.groupby(["Asesor", "Código Cliente", "LlaveAC"], as_index=False).agg(
            {
                "Cliente": "first",
                "Día": lambda s: " / ".join(sorted(set(x for x in s if x))),
                "Frecuencia": "first",
                "Teléfono": "first",
                "Dirección": "first",
                "Barrio": "first",
                "Grupo Maestra": "first",
            }
        )
        no_atendidos_df = m_agrup[
            ~m_agrup["LlaveAC"].isin(llaves_atendidas_set)
        ].copy()
        if not no_atendidos_df.empty:
            no_atendidos_df["WA"] = no_atendidos_df["Teléfono"].apply(
                limpiar_telefono_wa
            )
        total_no_atendidos = int(len(no_atendidos_df))
        total_cartera = int(len(m_agrup))
        cobertura_pct = (
            ((total_cartera - total_no_atendidos) / total_cartera * 100)
            if total_cartera > 0
            else 0.0
        )

    # Analítica Gerencial: Comparativo MoM a igual número de días hábiles y Clientes en Riesgo
    periodo_mes_ref = str(target)[:7]
    info_habiles = calcular_metricas_dias_habiles(periodo_mes_ref)

    v_mes_actual = df_v[df_v["Mes"] == periodo_mes_ref].copy()
    d_mes_actual = df_d[df_d["Mes"] == periodo_mes_ref].copy()
    if grupo and not df_m_raw.empty:
        asesores_del_grupo = set(df_m_raw[df_m_raw["Grupo Maestra"] == grupo]["Asesor"].unique())
        v_mes_actual = v_mes_actual[v_mes_actual["Asesor"].isin(asesores_del_grupo)]
        d_mes_actual = d_mes_actual[d_mes_actual["Asesor"].isin(asesores_del_grupo)]
    if asesor:
        v_mes_actual = v_mes_actual[v_mes_actual["Asesor"] == asesor]
        d_mes_actual = d_mes_actual[d_mes_actual["Asesor"] == asesor]

    # Comparativo Mes Anterior (MoM) a corte equivalente de días hábiles
    anio_ref, mes_ref = map(int, periodo_mes_ref.split("-"))
    anio_ant = anio_ref - 1 if mes_ref == 1 else anio_ref
    mes_ant = 12 if mes_ref == 1 else mes_ref - 1
    periodo_mes_ant = f"{anio_ant:04d}-{mes_ant:02d}"
    hab_ant = calcular_metricas_dias_habiles(periodo_mes_ant)
    fechas_corte_ant = set(hab_ant["fechas_habiles"][: info_habiles["habiles_transcurridos"]])

    v_ant = df_v[df_v["Fecha"].isin(fechas_corte_ant)].copy()
    d_ant = df_d[df_d["Fecha"].isin(fechas_corte_ant)].copy()
    if grupo and not df_m_raw.empty:
        asesores_del_grupo = set(df_m_raw[df_m_raw["Grupo Maestra"] == grupo]["Asesor"].unique())
        v_ant = v_ant[v_ant["Asesor"].isin(asesores_del_grupo)]
        d_ant = d_ant[d_ant["Asesor"].isin(asesores_del_grupo)]
    if asesor:
        v_ant = v_ant[v_ant["Asesor"] == asesor]
        d_ant = d_ant[d_ant["Asesor"] == asesor]

    neta_mes_act = float(v_mes_actual["Total"].sum()) - float(d_mes_actual["Total"].sum())
    neta_mes_ant = float(v_ant["Total"].sum()) - float(d_ant["Total"].sum())
    variacion_mom_pct = (
        ((neta_mes_act - neta_mes_ant) / neta_mes_ant * 100) if neta_mes_ant > 0 else 0.0
    )

    # Clientes en Riesgo (Compraron en historial pero llevan >= 15 días sin facturar)
    clientes_riesgo_df = pd.DataFrame()
    if not df_v.empty:
        v_hist = df_v.copy()
        if grupo and not df_m_raw.empty:
            asesores_del_grupo = set(df_m_raw[df_m_raw["Grupo Maestra"] == grupo]["Asesor"].unique())
            v_hist = v_hist[v_hist["Asesor"].isin(asesores_del_grupo)]
        if asesor:
            v_hist = v_hist[v_hist["Asesor"] == asesor]

        fecha_max_dt = datetime.strptime(fechas_disp[0], "%Y-%m-%d").date()
        riesgo_grp = (
            v_hist.groupby(["Asesor", "Código Cliente"], as_index=False)
            .agg(
                UltimaCompra=("Fecha", "max"),
                Pedidos=("Fecha", "nunique"),
                VentaHistorica=("Total", "sum"),
            )
        )
        riesgo_grp["DiasSinCompra"] = riesgo_grp["UltimaCompra"].apply(
            lambda f: (fecha_max_dt - datetime.strptime(f, "%Y-%m-%d").date()).days
        )
        riesgo_grp = riesgo_grp[riesgo_grp["DiasSinCompra"] >= 15].copy()
        if not riesgo_grp.empty:
            riesgo_grp["Cliente"] = riesgo_grp["Código Cliente"].map(mapa_nombres).fillna("—")
            riesgo_grp["Grupo"] = riesgo_grp["Asesor"].map(mapa_grupo_asesor).fillna("SIN GRUPO")
            riesgo_grp["TicketHist"] = riesgo_grp["VentaHistorica"] / riesgo_grp["Pedidos"].clip(lower=1)
            riesgo_grp["WA"] = riesgo_grp["Código Cliente"].map(mapa_tel_cliente).apply(limpiar_telefono_wa)
            clientes_riesgo_df = riesgo_grp.sort_values("VentaHistorica", ascending=False)

    return {
        "vacio": False,
        "fechas_disp": fechas_disp,
        "meses_disp": meses_disp,
        "target": target,
        "dia_fecha_actual": dia_fecha_actual,
        "dia_ruta_efectivo": dia_rutero_no_atendidos,
        "v_f": v_f,
        "d_f": d_f,
        "v_mes_actual": v_mes_actual,
        "d_mes_actual": d_mes_actual,
        "m_agrup": m_agrup,
        "no_atendidos_df": no_atendidos_df,
        "clientes_riesgo_df": clientes_riesgo_df,
        "tiene_maestra": tiene_maestra,
        "asesores_opt": asesores_opt,
        "marcas_opt": marcas_opt,
        "grupos_opt": grupos_opt,
        "mapa_nombres": mapa_nombres,
        "mapa_grupo_ac": mapa_grupo_ac,
        "mapa_grupo_asesor": mapa_grupo_asesor,
        "clientes_atendidos": clientes_atendidos_count,
        "total_no_atendidos": total_no_atendidos,
        "cobertura_pct": cobertura_pct,
        "info_habiles": {k: v for k, v in info_habiles.items() if k != "fechas_habiles"},
        "mom": {
            "periodo_actual": periodo_mes_ref,
            "periodo_anterior": periodo_mes_ant,
            "neta_actual": neta_mes_act,
            "neta_anterior": neta_mes_ant,
            "variacion_pct": variacion_mom_pct,
            "tiene_base_anterior": neta_mes_ant > 0,
        },
    }


def calcular_tabla_cumplimiento(ds):
    periodo_mes_ref = ds["info_habiles"]["periodo_mes"]
    info_habiles = ds["info_habiles"]
    metas_mes = db.obtener_metas_mes(periodo_mes_ref)

    v_mes = ds["v_mes_actual"]
    d_mes = ds["d_mes_actual"]

    ventas_ase = v_mes.groupby("Asesor")["Total"].sum().to_dict() if not v_mes.empty else {}
    devs_ase = d_mes.groupby("Asesor")["Total"].sum().to_dict() if not d_mes.empty else {}

    todos_asesores = sorted(set(list(metas_mes.keys()) + list(ventas_ase.keys()) + list(devs_ase.keys())))
    filas = []

    for ase in todos_asesores:
        v_bruta = float(ventas_ase.get(ase, 0.0))
        v_dev = float(devs_ase.get(ase, 0.0))
        venta_neta = v_bruta - v_dev
        cuota = float(metas_mes.get(ase.upper(), 0.0))
        pct_cumpl = (venta_neta / cuota * 100) if cuota > 0 else 0.0
        prom_diario = (
            (venta_neta / info_habiles["habiles_transcurridos"])
            if info_habiles["habiles_transcurridos"] > 0
            else 0.0
        )
        proyeccion = prom_diario * info_habiles["habiles_totales"]
        pct_proy = (proyeccion / cuota * 100) if cuota > 0 else 0.0
        faltante = max(0.0, cuota - venta_neta)
        req_diario = (
            (faltante / info_habiles["habiles_restantes"])
            if info_habiles["habiles_restantes"] > 0
            else faltante
        )
        filas.append({
            "Asesor": ase,
            "Grupo": ds["mapa_grupo_asesor"].get(ase, "SIN GRUPO"),
            "Cuota": cuota,
            "VentaNeta": venta_neta,
            "PctCumplimiento": pct_cumpl,
            "PromDiario": prom_diario,
            "ProyeccionCierre": proyeccion,
            "PctProyectado": pct_proy,
            "Faltante": faltante,
            "RequeridoDiario": req_diario,
        })

    filas.sort(key=lambda x: (x["Cuota"] > 0, x["VentaNeta"]), reverse=True)
    return filas


@app.route("/api/datos")
def api_datos():
    ds = construir_dataset_filtrado(request.args, request.current_user)
    if ds.get("vacio"):
        return jsonify({"vacio": True})

    v_f, d_f = ds["v_f"], ds["d_f"]
    mapa_nombres = ds["mapa_nombres"]
    mapa_grupo_ac = ds["mapa_grupo_ac"]
    mapa_grupo_asesor = ds["mapa_grupo_asesor"]
    etiqueta_dia_venta = ds["dia_fecha_actual"] or ds["target"]

    venta_bruta = float(v_f["Total"].sum())
    devolucion_total = float(d_f["Total"].sum())
    venta_neta_real = venta_bruta - devolucion_total
    pct_devolucion = (
        (devolucion_total / venta_bruta * 100) if venta_bruta > 0 else 0.0
    )
    ticket_promedio = (
        (venta_bruta / ds["clientes_atendidos"])
        if ds["clientes_atendidos"] > 0
        else 0.0
    )

    top_asesores = []
    if not v_f.empty:
        grp_v = v_f.groupby("Asesor", as_index=False).agg(
            Venta=("Total", "sum"),
            Unidades=("Unidades", "sum"),
            Clientes=("Código Cliente", "nunique"),
            Lineas=("Código Artículo", "count"),
        )
        if not d_f.empty:
            grp_v = grp_v.merge(
                d_f.groupby("Asesor", as_index=False).agg(Devolucion=("Total", "sum")),
                on="Asesor",
                how="left",
            )
            grp_v["Devolucion"] = grp_v["Devolucion"].fillna(0.0)
        else:
            grp_v["Devolucion"] = 0.0
        grp_v["TicketProm"] = grp_v.apply(
            lambda r: r["Venta"] / r["Clientes"] if r["Clientes"] > 0 else 0, axis=1
        )
        grp_v["DropSize"] = grp_v.apply(
            lambda r: r["Lineas"] / r["Clientes"] if r["Clientes"] > 0 else 0, axis=1
        )
        grp_v["PctDev"] = grp_v.apply(
            lambda r: (r["Devolucion"] / r["Venta"] * 100) if r["Venta"] > 0 else 0,
            axis=1,
        )
        top_asesores = (
            grp_v.sort_values("Venta", ascending=False)
            .head(40)
            .to_dict(orient="records")
        )

    por_marca = (
        v_f.groupby("Marca", as_index=False)
        .agg(Total=("Total", "sum"), Unidades=("Unidades", "sum"))
        .sort_values("Total", ascending=False)
        .to_dict(orient="records")
        if not v_f.empty
        else []
    )

    motivos_dev = []
    if not d_f.empty:
        d_f["MotivoClean"] = (
            d_f["Motivo"]
            .fillna("SIN MOTIVO")
            .astype(str)
            .str.strip()
            .replace("", "SIN MOTIVO")
        )
        motivos_dev = (
            d_f.groupby("MotivoClean", as_index=False)
            .agg(
                Total=("Total", "sum"),
                Unidades=("Unidades", "sum"),
                Notas=("Documento", "nunique"),
            )
            .sort_values("Total", ascending=False)
            .to_dict(orient="records")
        )

    asesor_cliente = []
    if not v_f.empty:
        ac = (
            v_f.groupby(["Asesor", "Código Cliente", "LlaveAC"], as_index=False)
            .agg(
                Total=("Total", "sum"),
                Unidades=("Unidades", "sum"),
                Referencias=("Código Artículo", "nunique"),
                Marcas=("Marca", lambda s: ", ".join(sorted(set(x for x in s if x)))),
            )
            .sort_values("Total", ascending=False)
            .head(300)
        )
        ac["Cliente"] = ac["Código Cliente"].map(mapa_nombres).astype("object").fillna("—")
        ac["DiaRuta"] = etiqueta_dia_venta
        ac["Grupo"] = (
            ac["LlaveAC"]
            .map(mapa_grupo_ac)
            .fillna(ac["Asesor"].map(mapa_grupo_asesor))
            .astype("object")
            .fillna("SIN GRUPO")
        )
        asesor_cliente = ac.drop(columns=["LlaveAC"]).to_dict(orient="records")

    resumen_detallado = []
    if not v_f.empty:
        v_det = v_f.copy()
        v_det["DiaRuta"] = etiqueta_dia_venta
        v_det["GrupoRuta"] = (
            v_det["LlaveAC"]
            .map(mapa_grupo_ac)
            .fillna(v_det["Asesor"].map(mapa_grupo_asesor))
            .fillna("SIN GRUPO")
        )
        resumen_detallado = (
            v_det.groupby(["DiaRuta", "GrupoRuta", "Asesor"], as_index=False)
            .agg(
                ClientesAtendidos=("Código Cliente", "nunique"),
                Unidades=("Unidades", "sum"),
                Venta=("Total", "sum"),
            )
            .sort_values(["GrupoRuta", "Venta"], ascending=[True, False])
            .to_dict(orient="records")
        )

    dev_clientes, dev_detalle = [], []
    if not d_f.empty:
        cli_orig = d_f["Cliente"].fillna("").astype(str).str.strip()
        d_f["Cliente"] = cli_orig.where(
            cli_orig != "",
            d_f["Código Cliente"].map(mapa_nombres).fillna("—").astype(str),
        )
        dev_clientes = (
            d_f.groupby(["Asesor", "Código Cliente", "Cliente"], as_index=False)
            .agg(
                Total=("Total", "sum"),
                Unidades=("Unidades", "sum"),
                Referencias=("Código Artículo", "nunique"),
            )
            .sort_values("Total", ascending=False)
            .head(250)
            .to_dict(orient="records")
        )
        dev_detalle = (
            d_f[COLS_DEV]
            .sort_values("Total", ascending=False)
            .head(300)
            .to_dict(orient="records")
        )

    top_productos = (
        v_f.groupby(
            ["Código Artículo", "Descripción Artículo", "Marca", "Grupo"], as_index=False
        )
        .agg(
            Total=("Total", "sum"),
            Unidades=("Unidades", "sum"),
            Clientes=("Código Cliente", "nunique"),
        )
        .sort_values("Total", ascending=False)
        .head(200)
        .to_dict(orient="records")
        if not v_f.empty
        else []
    )

    cumplimiento_asesores = calcular_tabla_cumplimiento(ds)
    cuota_total_global = sum(x["Cuota"] for x in cumplimiento_asesores)
    neta_mes_global = sum(x["VentaNeta"] for x in cumplimiento_asesores)
    proy_mes_global = sum(x["ProyeccionCierre"] for x in cumplimiento_asesores)

    return jsonify(
        {
            "vacio": False,
            "fechas_disp": ds["fechas_disp"],
            "meses_disp": ds["meses_disp"],
            "target_actual": ds["target"],
            "dia_ruta_efectivo": ds["dia_ruta_efectivo"],
            "tiene_maestra": ds["tiene_maestra"],
            "opciones": {
                "asesores": ds["asesores_opt"],
                "marcas": ds["marcas_opt"],
                "grupos": ds["grupos_opt"],
            },
            "kpis": {
                "venta_bruta": venta_bruta,
                "devolucion_total": devolucion_total,
                "venta_neta_real": venta_neta_real,
                "pct_devolucion": pct_devolucion,
                "unidades_total": int(v_f["Unidades"].sum()),
                "unidades_dev": int(d_f["Unidades"].sum()),
                "clientes_atendidos": ds["clientes_atendidos"],
                "total_no_atendidos": ds["total_no_atendidos"],
                "cobertura_pct": ds["cobertura_pct"],
                "ticket_promedio": ticket_promedio,
            },
            "info_habiles": ds["info_habiles"],
            "mom": ds["mom"],
            "resumen_metas": {
                "cuota_total": cuota_total_global,
                "venta_neta_mes": neta_mes_global,
                "proyeccion_mes": proy_mes_global,
                "pct_cumplimiento_total": (neta_mes_global / cuota_total_global * 100) if cuota_total_global > 0 else 0.0,
                "pct_proyectado_total": (proy_mes_global / cuota_total_global * 100) if cuota_total_global > 0 else 0.0,
            },
            "cumplimiento_asesores": cumplimiento_asesores,
            "clientes_riesgo": ds["clientes_riesgo_df"].head(200).to_dict(orient="records")
            if not ds["clientes_riesgo_df"].empty
            else [],
            "top_asesores": top_asesores,
            "por_marca": por_marca,
            "motivos_dev": motivos_dev,
            "asesor_cliente": asesor_cliente,
            "resumen_detallado": resumen_detallado,
            "dev_clientes": dev_clientes,
            "dev_detalle": dev_detalle,
            "no_atendidos": ds["no_atendidos_df"].drop(columns=["LlaveAC"], errors="ignore").head(400).to_dict(orient="records")
            if not ds["no_atendidos_df"].empty
            else [],
            "top_productos": top_productos,
        }
    )


@app.route("/api/metas", methods=["GET", "POST"])
def api_metas():
    user = request.current_user
    if request.method == "GET":
        periodo = request.args.get("periodo", datetime.now().strftime("%Y-%m"))[:7]
        return jsonify({"periodo": periodo, "metas": db.obtener_metas_mes(periodo)})

    if user.get("rol") not in ("admin", "supervisor"):
        return jsonify({"ok": False, "msg": "No autorizado para editar cuotas."}), 403

    data = request.get_json(silent=True) or {}
    periodo = str(data.get("periodo", datetime.now().strftime("%Y-%m")))[:7]
    lista_metas = data.get("metas", [])
    db.guardar_metas_mes(periodo, lista_metas, usuario_editor=user["username"])
    return jsonify({"ok": True, "periodo": periodo})


@app.route("/api/exportar/<tipo>")
def api_exportar(tipo):
    if tipo not in ("no_atendidos", "devoluciones", "atendidos", "ejecutivo"):
        return "Tipo inválido", 400
    ds = construir_dataset_filtrado(request.args, request.current_user)
    if ds.get("vacio"):
        return "Sin datos", 400
    output = io.BytesIO()
    etiqueta_dia_venta = ds["dia_fecha_actual"] or ds["target"]

    if tipo == "ejecutivo":
        df_metas = pd.DataFrame(calcular_tabla_cumplimiento(ds))
        df_riesgo = ds["clientes_riesgo_df"].drop(columns=["WA"], errors="ignore") if not ds["clientes_riesgo_df"].empty else pd.DataFrame()
        nombre_arch = f"Informe_Ejecutivo_HB_{ds['info_habiles']['periodo_mes']}.xlsx"
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_metas.to_excel(writer, index=False, sheet_name="Metas_y_Proyeccion")
            df_riesgo.to_excel(writer, index=False, sheet_name="Clientes_en_Riesgo")
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name=nombre_arch,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if tipo == "no_atendidos":
        df_exp = ds["no_atendidos_df"].drop(columns=["WA", "LlaveAC"], errors="ignore")
        pref = "No_Atendidos"
    elif tipo == "devoluciones":
        df_exp = ds["d_f"][COLS_DEV].copy()
        pref = "Devoluciones"
    else:
        v_exp = ds["v_f"].copy()
        v_exp["Cliente"] = v_exp["Código Cliente"].map(ds["mapa_nombres"]).fillna("—")
        v_exp["Dia"] = etiqueta_dia_venta
        v_exp["GrupoRuta"] = (
            v_exp["LlaveAC"]
            .map(ds["mapa_grupo_ac"])
            .fillna(v_exp["Asesor"].map(ds["mapa_grupo_asesor"]))
            .fillna("SIN GRUPO")
        )
        df_exp = v_exp.groupby(
            ["Dia", "GrupoRuta", "Asesor", "Código Cliente", "Cliente"], as_index=False
        ).agg(
            Unidades=("Unidades", "sum"),
            Referencias=("Código Artículo", "nunique"),
            Marcas=("Marca", lambda s: ", ".join(sorted(set(x for x in s if x)))),
            Total=("Total", "sum"),
        ).sort_values(["GrupoRuta", "Asesor", "Total"], ascending=[True, True, False])
        pref = "Clientes_Atendidos"

    nombre_arch = f"{pref}_HB_{ds['target']}.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False, sheet_name="Reporte_HB")
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=nombre_arch,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/extraer", methods=["POST"])
def api_extraer():
    if request.current_user.get("rol") != "admin":
        return jsonify({"ok": False, "msg": "Acceso denegado."}), 403
    payload = request.json or {}
    tipo = payload.get("tipo", "")
    modo = payload.get("modo", "")
    if tipo not in ("ventas", "devoluciones") or modo not in ("ayer", "diario", "mensual"):
        return jsonify({"ok": False, "msg": "Parámetros inválidos."}), 400
    try:
        rpa.ejecutar_extraccion_unica(tipo=tipo, modo=modo, headless=False)
        _CACHE_EXCEL.clear()
        return jsonify({"ok": True, "msg": f"Extracción de {tipo} ({modo}) completada."})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/usuarios", methods=["GET", "POST"])
def api_usuarios():
    if request.current_user.get("rol") != "admin":
        return jsonify({"error": "No autorizado"}), 403
    if request.method == "POST":
        db.crear_o_actualizar_usuario(request.json or {})
        return jsonify({"ok": True})
    df_m = cargar_maestra_cruda()
    return jsonify(
        {
            "usuarios": db.listar_usuarios(),
            "grupos_maestra": sorted(
                [g for g in df_m["Grupo Maestra"].unique() if g]
            )
            if not df_m.empty
            else [],
            "asesores_maestra": sorted([a for a in df_m["Asesor"].unique() if a])
            if not df_m.empty
            else [],
        }
    )


@app.route("/api/usuarios/masivo_asesores", methods=["POST"])
def api_masivo_asesores():
    if request.current_user.get("rol") != "admin":
        return jsonify({"error": "No autorizado"}), 403
    df_m = cargar_maestra_cruda()
    if df_m.empty:
        return jsonify({"ok": False, "msg": "No se encontró MAESTRA_HB.XLS"}), 400
    creados = db.generar_usuarios_masivos_asesores(
        sorted([a for a in df_m["Asesor"].unique() if a])
    )
    return jsonify({"ok": True, "creados": creados})


@app.route("/api/usuarios/exportar_excel")
def api_exportar_usuarios():
    if request.current_user.get("rol") != "admin":
        return "No autorizado", 403
    filas = [
        {
            "Usuario": u["username"],
            "Nombre": u["nombre"],
            "Rol": u["rol"].upper(),
            "Estado": "Activo" if u["activo"] else "Inactivo",
            "Clave Temporal": u["clave_temporal"],
            "Asesor Abako": u["asesor_abako"],
            "Grupos": ", ".join(u["grupos"]),
        }
        for u in db.listar_usuarios()
    ]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(filas).to_excel(writer, index=False, sheet_name="Usuarios_HB")
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="Credenciales_Temporales_HB.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/usuarios/<int:uid>", methods=["DELETE"])
def api_eliminar_usuario(uid):
    if request.current_user.get("rol") != "admin":
        return jsonify({"error": "No autorizado"}), 403
    db.eliminar_usuario(uid)
    return jsonify({"ok": True})


@app.route("/")
def index():
    logo_existe = os.path.exists(os.path.join(CARPETA_STATIC, "logo_hb.png"))
    return render_template(
        "dashboard.html", user=request.current_user, logo_existe=logo_existe
    )


if __name__ == "__main__":
    from waitress import serve

    print("Dashboard HB v2.2 (Producción Waitress) activo en: http://0.0.0.0:5050")
    serve(app, host="0.0.0.0", port=5050, threads=6)