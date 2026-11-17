import os
import re
import calendar
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

URL_LOGIN = os.getenv("ABAKO_URL", "https://hberp.abakoerp.com:9443/PortalAbakoERP/login")
USUARIO = os.getenv("ABAKO_USER", "")
CONTRASENA = os.getenv("ABAKO_PASS", "")
CARPETA_DESCARGAS = BASE_DIR

FESTIVOS_CO = {
    # Festivos Colombia 2026 - 2027 (YYYY-MM-DD)
    "2026-01-01", "2026-01-12", "2026-03-23", "2026-04-02", "2026-04-03",
    "2026-05-01", "2026-05-18", "2026-06-08", "2026-06-15", "2026-06-29",
    "2026-07-20", "2026-08-07", "2026-08-17", "2026-10-12", "2026-11-02",
    "2026-11-16", "2026-12-08", "2026-12-25",
    "2027-01-01", "2027-01-11", "2027-03-22", "2027-03-25", "2027-03-26",
    "2027-05-01", "2027-05-10", "2027-05-31", "2027-06-07", "2027-06-28",
    "2027-07-05", "2027-07-20", "2027-08-07", "2027-08-16", "2027-10-18",
    "2027-11-01", "2027-11-15", "2027-12-08", "2027-12-25",
}

REPORTES_CONFIG = {
    "ventas": {
        "menu": "Ventas Netas BI",
        "prefijo_archivo": "Ventas_Netas_HB",
        "campos": ["Código Cliente", "Grupo", "Marca", "Código Artículo", "Descripción Artículo", "Total"],
    },
    "devoluciones": {
        "menu": "Detalle Notas Crédito",
        "prefijo_archivo": "Devoluciones_HB",
        "campos": ["Asesor", "Código Cliente", "Cliente", "Código Artículo", "Artículo", "Documento", "Motivo"],
    },
}


def validar_credenciales():
    if not USUARIO or not CONTRASENA:
        raise RuntimeError("Faltan ABAKO_USER o ABAKO_PASS en el archivo .env")


def crear_contexto(browser):
    return browser.new_context(
        accept_downloads=True,
        ignore_https_errors=True,
        locale="es-CO",
        viewport={"width": 1366, "height": 768},
    )


def ultimo_dia_habil(referencia=None):
    dia = (referencia or datetime.now()) - timedelta(days=1)
    while dia.weekday() == 6 or dia.strftime("%Y-%m-%d") in FESTIVOS_CO:
        dia -= timedelta(days=1)
    return dia


def obtener_rango_fechas(modo="ayer"):
    hoy = datetime.now()
    if modo == "ayer":
        habil = ultimo_dia_habil(hoy)
        return habil, habil
    elif modo == "diario":
        return hoy, hoy
    elif modo == "mensual":
        primer_dia = hoy.replace(day=1)
        ultimo_num = calendar.monthrange(hoy.year, hoy.month)[1]
        return primer_dia, hoy.replace(day=ultimo_num)
    raise ValueError(f"Modo no válido: {modo}")


def iniciar_sesion_abako(page):
    validar_credenciales()
    page.goto(URL_LOGIN, wait_until="domcontentloaded")
    page.get_by_placeholder("USUARIO").fill(USUARIO)
    page.get_by_placeholder("CONTRASEÑA").fill(CONTRASENA)
    page.get_by_role("button", name="INGRESAR").click()
    page.locator("a:visible").filter(has_text=re.compile(r"^\s*Portal Web\s*$")).first.wait_for(
        state="visible", timeout=20000
    )


def abrir_reporte_ventas(page, nombre_subreporte):
    page.get_by_text("Portal Web", exact=True).first.click()
    ventas_visible = page.locator("a:visible", has_text="Ventas").first
    ventas_visible.hover()
    ventas_visible.click()
    subreporte = page.locator("a:visible", has_text=nombre_subreporte).first
    subreporte.wait_for(state="visible", timeout=10000)
    subreporte.click(force=True)
    page.get_by_text("Filtrar").wait_for(state="visible", timeout=25000)


def digitar_fechas_filtro(page, dt_ini, dt_fin):
    page.get_by_text("Filtrar").click()
    sel = "mat-dialog-container input, .cdk-overlay-pane input, nb-dialog-container input"
    modal_inputs = page.locator(sel)
    modal_inputs.first.wait_for(state="visible", timeout=10000)

    for idx, dt in enumerate([dt_ini, dt_fin]):
        campo = modal_inputs.nth(idx)
        campo.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        campo.press_sequentially(dt.strftime("%m/%d/%Y"), delay=40)
        page.keyboard.press("Tab")


def configurar_campos_y_expandir(page, lista_campos):
    tab = page.locator("a.tab-link, span.tab-text, li.tab").filter(has_text="CAMPOS").first
    tab.click(force=True)

    encabezado_filas = page.get_by_text("Filas", exact=True).locator("visible=true").first
    encabezado_filas.wait_for(state="visible", timeout=10000)
    box_hdr = encabezado_filas.bounding_box()
    coords_destino = (box_hdr["x"] + 90, box_hdr["y"] + 65)
    zona_filas = encabezado_filas.locator("..").locator("..")

    for nombre_campo in lista_campos:
        if zona_filas.get_by_text(nombre_campo, exact=True).count() > 0:
            continue
        try:
            item = page.get_by_text(nombre_campo, exact=True).locator("visible=true").first
            item.scroll_into_view_if_needed(timeout=1500)
            box_orig = item.bounding_box()
            if box_orig:
                page.mouse.move(box_orig["x"] + box_orig["width"] / 2, box_orig["y"] + box_orig["height"] / 2)
                page.mouse.down()
                page.mouse.move(coords_destino[0], coords_destino[1], steps=12)
                page.mouse.up()
                page.wait_for_timeout(180)
        except Exception:
            pass

    x_filas, y_filas = coords_destino
    page.mouse.move(x_filas, y_filas)
    page.mouse.wheel(0, -600)
    for _ in range(6):
        etiquetas = page.get_by_text("Expandir", exact=True).locator("visible=true")
        for i in range(etiquetas.count()):
            cb = etiquetas.nth(i).locator("..").locator("input[type='checkbox']").first
            if cb.count() > 0 and not cb.is_checked():
                cb.check(force=True)
        page.mouse.move(x_filas, y_filas)
        page.mouse.wheel(0, 250)
        page.wait_for_timeout(120)


def ejecutar_extraccion_unica(tipo="ventas", modo="ayer", headless=False):
    cfg = REPORTES_CONFIG[tipo]
    dt_ini, dt_fin = obtener_rango_fechas(modo)
    print(f"Extrayendo {tipo.upper()} ({modo}): {dt_ini.strftime('%Y-%m-%d')}...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = crear_contexto(browser)
        page = context.new_page()

        iniciar_sesion_abako(page)
        abrir_reporte_ventas(page, cfg["menu"])
        digitar_fechas_filtro(page, dt_ini, dt_fin)
        configurar_campos_y_expandir(page, cfg["campos"])

        page.get_by_role("button", name="BUSCAR").click()
        page.get_by_text(re.compile(r"\d+\s+Ítems encontrados", re.IGNORECASE)).wait_for(
            state="visible", timeout=60000
        )
        page.wait_for_timeout(1200)

        with page.expect_download(timeout=35000) as dl_info:
            page.get_by_text("Exportar").click()
            page.get_by_text("Excel - Todos los Registros").click()

        fecha_nombre = dt_fin.strftime("%Y-%m-%d")
        ruta_final = os.path.join(CARPETA_DESCARGAS, f"{cfg['prefijo_archivo']}_{modo}_{fecha_nombre}.xlsx")
        dl_info.value.save_as(ruta_final)
        print(f"Guardado: {ruta_final}")
        browser.close()


def descargar_reporte_abako(modo="ayer", headless=False):
    ejecutar_extraccion_unica("ventas", modo=modo, headless=headless)


def descargar_devoluciones_abako(modo="ayer", headless=False):
    ejecutar_extraccion_unica("devoluciones", modo=modo, headless=headless)


if __name__ == "__main__":
    ejecutar_extraccion_unica("ventas", modo="ayer", headless=True)