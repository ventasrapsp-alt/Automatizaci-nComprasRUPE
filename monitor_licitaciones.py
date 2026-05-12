"""
╔══════════════════════════════════════════════════════╗
║   MONITOR DE LICITACIONES — ARCE Uruguay            ║
║   ANCAP · UTE · ANTEL · TODO EL PORTAL             ║
║   Fuente: comprasestatales.gub.uy                   ║
║   Ejecución: Automática cada lunes vía cron/Task    ║
╚══════════════════════════════════════════════════════╝

INSTALACIÓN:
    pip install requests beautifulsoup4 openpyxl

CONFIGURACIÓN:
    Editar la sección CONFIG antes de ejecutar.

AUTOMATIZACIÓN:
    Linux/Mac (cron):  0 8 * * 1 python3 /ruta/monitor_licitaciones.py
    Windows (Task):    Ejecutar cada lunes a las 08:00
"""

import requests
from bs4 import BeautifulSoup
import smtplib
import csv
import os
import json
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

import os

# ─────────────────────────────────────────────
#  ⚙️  CONFIG — EDITÁ ESTO ANTES DE EJECUTAR
# ─────────────────────────────────────────────
CONFIG = {
    "email_destino":      os.environ.get("EMAIL_DESTINO",      "ventas@rap-sp.com.ar"),
    "email_remitente":    os.environ.get("EMAIL_REMITENTE",    "ventas.rap.sp@gmail.com"),
    "gmail_app_password": os.environ.get("GMAIL_APP_PASSWORD", "ywqdzlbcgfmbsevi"),
    "csv_historial": "historial_licitaciones.csv",
    "cache_file": "cache_notificadas.json",
    "solo_alertar_con_hits": True,
}

# ─────────────────────────────────────────────
#  🎯  KEYWORDS A DETECTAR
#
#  Perfil de negocio:
#  → Suministro de equipos eléctricos AT/MT/BT
#  → Tendido de líneas y redes eléctricas
#  → EPP específico para trabajos con tensión
#  → NO interesa: limpieza, excavaciones, obra civil general
# ─────────────────────────────────────────────
KEYWORDS = [
    # Tensión — AT / MT / BT (todas las combinaciones)
    " AT ", "alta tensión", "alta tension",
    " MT ", "media tensión", "media tension",
    " BT ", "baja tensión", "baja tension",
    "MT/BT", "BT/BT", "MT/MT", "AT/MT", "AT/BT",
    "kV", "150kV", "500kV", "220kV", "66kV", "15kV",

    # Suministro — solo combinado con equipos del rubro
    "suministro de transformadores", "suministro de cables",
    "suministro de conductores", "suministro de celdas",
    "suministro de disyuntores", "suministro de equipos eléctricos",
    "suministro de materiales eléctricos", "suministro de aisladores",
    "suministro de herrajes", "suministro de conectores",
    "suministro de arneses", "suministro de EPP",
    "suministro de equipos de protección",
    "suministro e instalación",
    "autotrafo", "autotransformador",
    "celda", "celdas primarias", "seccionadora", "seccionadoras",
    "disyuntor", "disyuntores", "interruptor",
    "conductor", "conductores", "cable", "cables protegido",
    "bornera", "borneras", "pinza amperimétrica", "pinzas amperimétricas",
    "descargador", "pararrayos",
    "medidor", "transformador de medida",
    "banco de condensadores",

    # Tendido de líneas y redes
    "tendido", "líneas de transmisión", "lineas de transmision",
    "líneas de distribución", "lineas de distribucion",
    "red de distribución", "red de distribucion",
    "mástil", "mastil", "mástiles",
    "subestación", "subestaciones",
    "estación 150", "estación 500",

    # EPP altura y riesgo eléctrico
    "casco", "cascos", "casco dieléctrico", "casco dielectrico",
    "arnés", "arnes", "arneses",
    "soga dieléctrica", "soga dielectrica", "sogas dieléctricas",
    "cuerda dieléctrica", "cuerda dielectrica",
    "mosquetón", "mosqueton", "mosquetones",
    "línea de vida", "linea de vida", "líneas de vida",
    "absorbedor de impacto", "absorbedor de energía",
    "freno de soga", "freno de cuerda",
    "anclaje", "anclajes", "punto de anclaje",
    "eslinga", "eslingas",
    "posicionador", "posicionadores",
    "cinturón de seguridad", "cinturon de seguridad",
    "equipo de protección contra caídas", "protección contra caída",
    "proteccion contra caida",
    "detector de tensión", "detector de tension",
    "pértiga", "pertiga", "pértigas", "pertigas",
    "vara de maniobra", "varas de maniobra",
    "equipo de maniobra", "equipos de maniobra",
    "equipo de puesta a tierra", "puesta a tierra",
    "traje ignífugo", "ropa ignifuga", "ropa ignífuga",
    "calzado dieléctrico", "calzado dielectrico",
    "casco dieléctrico", "casco dielectrico",
    "pantalla facial", "careta facial",
    "alfombra dieléctrica", "alfombra dielectrica",
    "banqueta aislante", "escalera dieléctrica",

    # Conectores, herrajes y accesorios de línea
    "conector", "conectores",
    "herraje", "herrajes",
    "grapa", "grapas",
    "aislador", "aisladores",
    "vaina termorretráctil", "vainas termorretractil",
    "termorretráctil", "termorretractil",
    "empalme", "empalmes",
    "terminal de media tensión", "terminal de alta tensión",
    "terminal de cable", "terminales de cable",
    "terminal termorretráctil", "terminales termorretráctiles",
    "manguito", "manguitos",
    "preformado", "preformados",
    "amortiguador", "amortiguadores",
    "cruceta", "crucetas",
    "fitting", "fittings",

    # Trabajos con tensión
    "trabajo con tensión", "trabajo con tension",
    "trabajos con tensión", "trabajos con tension",
    "línea viva", "linea de vida", "lineas de vida", "líneas vivas",
    "mantenimiento con tensión", "mantenimiento en tensión",

    # Capacitación técnica eléctrica
    "capacitación técnica", "capacitacion tecnica",
    "curso trabajo con tensión", "curso trabajos con tension",
    "formación en tensión", "formacion en tension",
]

# Keywords que descartan una licitación aunque tenga hits (exclusiones)
KEYWORDS_EXCLUIR = [
    "limpieza", "excavacion", "excavación",
    "obra civil", "muro", "pintura",
    "venta de inmueble", "venta de monte",
    "formularios", "publicacion de avisos", "publicación de avisos",
    "arrendamiento de local",
]

# ─────────────────────────────────────────────────────────────────
#  🌐  FUENTES
#
#  Estrategia de doble cobertura:
#
#  A) Por organismo (ANCAP / UTE / ANTEL):
#     → Se descargan TODAS sus licitaciones vigentes.
#     → Útil para ver todo lo que publican, más allá de keywords.
#
#  B) Portal completo (TODO el Estado):
#     → RSS de "cambios de la última semana" sin filtro de organismo.
#     → Se filtra por keywords EPP / altura / capacitación.
#     → Captura cualquier organismo que publique algo relevante.
#     → Complementa A: si surge una oportunidad en OSE, ANP, ASSE, etc.,
#       la detectás igual.
# ─────────────────────────────────────────────────────────────────

# Incisos: 60=ANCAP · 61=UTE · 65=ANTEL
FUENTES_POR_ORGANISMO = [
    {
        "nombre": "ANCAP — Llamados Vigentes",
        "organismo": "ANCAP",
        "url_html": "https://www.comprasestatales.gub.uy/consultas/buscar/tipo-pub/VIG/inciso/60/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_rss":  "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/60/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_base": "https://www.comprasestatales.gub.uy",
        "filtrar_por_keywords": False,  # Mostrar TODAS sus licitaciones
    },
    {
        "nombre": "UTE — Llamados Vigentes",
        "organismo": "UTE",
        "url_html": "https://www.comprasestatales.gub.uy/consultas/buscar/tipo-pub/VIG/inciso/61/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_rss":  "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/61/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_base": "https://www.comprasestatales.gub.uy",
        "filtrar_por_keywords": False,  # Mostrar TODAS sus licitaciones
    },
    {
        "nombre": "ANTEL — Concursos de Precios Vigentes",
        "organismo": "ANTEL",
        "url_html": "https://www.comprasestatales.gub.uy/consultas/buscar/tipo-pub/VIG/inciso/65/tipo-doc/C/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_rss":  "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/65/tipo-doc/C/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_base": "https://www.comprasestatales.gub.uy",
        "filtrar_por_keywords": False,  # Mostrar TODAS sus licitaciones
    },
]

# RSS semanal sin filtro de organismo — todo el portal ARCE
# La URL se construye dinámicamente con las fechas de la semana actual
FUENTE_PORTAL_COMPLETO = {
    "nombre": "ARCE — Portal Completo (todos los organismos)",
    "organismo": "VARIOS",
    "url_base": "https://www.comprasestatales.gub.uy",
    "filtrar_por_keywords": True,  # SOLO mostrar los que tengan keywords
}

# Unión de todas las fuentes para iterar
FUENTES = FUENTES_POR_ORGANISMO + [FUENTE_PORTAL_COMPLETO]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; LicitacionesBot/1.0)"
}

# ─────────────────────────────────────────────
#  📦  FUNCIONES PRINCIPALES
# ─────────────────────────────────────────────

def cargar_cache():
    """Carga el listado de licitaciones ya notificadas para no repetir alertas."""
    if os.path.exists(CONFIG["cache_file"]):
        with open(CONFIG["cache_file"], "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def guardar_cache(cache: set):
    with open(CONFIG["cache_file"], "w", encoding="utf-8") as f:
        json.dump(list(cache), f, ensure_ascii=False)


def detectar_keywords(texto: str) -> list:
    """Retorna keywords encontradas. Retorna lista vacía si hay keywords de exclusión."""
    texto_lower = texto.lower()
    # Si el texto contiene algo que explícitamente NO nos interesa → descartar
    for excluir in KEYWORDS_EXCLUIR:
        if excluir.lower() in texto_lower:
            return []
    return [kw for kw in KEYWORDS if kw.lower() in texto_lower]


def scrape_rss(fuente: dict) -> list:
    """
    Intenta obtener licitaciones vía RSS feed de ARCE.
    Retorna lista de dicts con info de cada licitación.
    """
    resultados = []
    try:
        r = requests.get(fuente["url_rss"], headers=HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.content, "xml")
        items = soup.find_all("item")

        for item in items:
            titulo = item.find("title").get_text(strip=True) if item.find("title") else ""
            descripcion = item.find("description").get_text(strip=True) if item.find("description") else ""
            link = item.find("link").get_text(strip=True) if item.find("link") else ""
            pub_date = item.find("pubDate").get_text(strip=True) if item.find("pubDate") else ""

            texto_completo = f"{titulo} {descripcion}"
            keywords_encontradas = detectar_keywords(texto_completo)

            resultados.append({
                "organismo": fuente["organismo"],
                "titulo": titulo,
                "descripcion": descripcion[:300],
                "link": link,
                "fecha_publicacion": pub_date,
                "keywords": keywords_encontradas,
                "tiene_hits": len(keywords_encontradas) > 0,
                "fuente": "RSS",
            })

        print(f"  ✅ {fuente['organismo']} (RSS): {len(resultados)} licitaciones encontradas")

    except Exception as e:
        print(f"  ⚠️  RSS falló para {fuente['organismo']}: {e}")
        # Fallback a HTML
        resultados = scrape_html(fuente)

    return resultados


def scrape_html(fuente: dict) -> list:
    """
    Fallback: scraping HTML del portal ARCE.
    Extrae licitaciones de la tabla principal.
    """
    resultados = []
    try:
        r = requests.get(fuente["url_html"], headers=HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        # ARCE usa una tabla o lista con class "resultado"
        filas = soup.select("table tr") or soup.select(".resultado") or soup.select(".llamado")

        for fila in filas[1:]:  # Saltar header
            celdas = fila.find_all(["td", "th"])
            if len(celdas) < 2:
                continue

            texto_fila = fila.get_text(" ", strip=True)
            link_tag = fila.find("a", href=True)
            link = fuente["url_base"] + link_tag["href"] if link_tag else fuente["url_html"]

            keywords_encontradas = detectar_keywords(texto_fila)

            resultados.append({
                "organismo": fuente["organismo"],
                "titulo": texto_fila[:200],
                "descripcion": texto_fila[:300],
                "link": link,
                "fecha_publicacion": datetime.now().strftime("%Y-%m-%d"),
                "keywords": keywords_encontradas,
                "tiene_hits": len(keywords_encontradas) > 0,
                "fuente": "HTML",
            })

        print(f"  ✅ {fuente['organismo']} (HTML fallback): {len(resultados)} filas procesadas")

    except Exception as e:
        print(f"  ❌ Error crítico en {fuente['organismo']}: {e}")

    return resultados


def guardar_csv(licitaciones: list):
    """Agrega nuevas licitaciones al CSV histórico (no duplica)."""
    archivo = CONFIG["csv_historial"]
    existe = os.path.exists(archivo)

    with open(archivo, "a", newline="", encoding="utf-8") as f:
        campos = ["fecha_reporte", "organismo", "titulo", "descripcion",
                  "link", "fecha_publicacion", "keywords", "tiene_hits", "fuente"]
        writer = csv.DictWriter(f, fieldnames=campos)

        if not existe:
            writer.writeheader()

        hoy = datetime.now().strftime("%Y-%m-%d")
        for lic in licitaciones:
            writer.writerow({
                "fecha_reporte": hoy,
                "organismo": lic["organismo"],
                "titulo": lic["titulo"],
                "descripcion": lic["descripcion"],
                "link": lic["link"],
                "fecha_publicacion": lic["fecha_publicacion"],
                "keywords": ", ".join(lic["keywords"]),
                "tiene_hits": lic["tiene_hits"],
                "fuente": lic["fuente"],
            })

    print(f"  💾 CSV actualizado: {archivo}")


def construir_email_html(licitaciones_con_hits: list, total_revisadas: int) -> str:
    """Genera el cuerpo HTML del email de reporte."""
    hoy = datetime.now().strftime("%d/%m/%Y")
    n_hits = len(licitaciones_con_hits)

    filas_html = ""
    for lic in licitaciones_con_hits:
        kws = ", ".join([f"<strong>{k}</strong>" for k in lic["keywords"]])
        filas_html += f"""
        <tr>
            <td style="padding:10px;border-bottom:1px solid #eee;color:#333;">
                <span style="background:#1a3c5e;color:white;padding:2px 8px;border-radius:3px;font-size:11px;">
                    {lic['organismo']}
                </span>
            </td>
            <td style="padding:10px;border-bottom:1px solid #eee;">
                <a href="{lic['link']}" style="color:#1a3c5e;text-decoration:none;font-weight:bold;">
                    {lic['titulo'][:120]}
                </a>
                <br><small style="color:#888;">{lic['descripcion'][:150]}...</small>
            </td>
            <td style="padding:10px;border-bottom:1px solid #eee;font-size:12px;color:#555;">
                {kws}
            </td>
            <td style="padding:10px;border-bottom:1px solid #eee;font-size:12px;color:#888;white-space:nowrap;">
                {lic['fecha_publicacion']}
            </td>
        </tr>
        """

    return f"""
    <html><body style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;">
        <div style="background:#1a3c5e;padding:20px;border-radius:8px 8px 0 0;">
            <h2 style="color:white;margin:0;">📋 Reporte de Licitaciones</h2>
            <p style="color:#8ab4d4;margin:5px 0 0;">ANCAP & UTE Uruguay — {hoy}</p>
        </div>

        <div style="background:#f0f4f8;padding:15px;display:flex;gap:20px;">
            <div style="background:white;padding:15px 25px;border-radius:6px;text-align:center;flex:1;">
                <div style="font-size:28px;font-weight:bold;color:#1a3c5e;">{total_revisadas}</div>
                <div style="color:#888;font-size:12px;">Total revisadas</div>
            </div>
            <div style="background:white;padding:15px 25px;border-radius:6px;text-align:center;flex:1;">
                <div style="font-size:28px;font-weight:bold;color:#e85d04;">{n_hits}</div>
                <div style="color:#888;font-size:12px;">Con keywords relevantes</div>
            </div>
        </div>

        {"<div style='background:#fff3cd;padding:15px;border-left:4px solid #ffc107;margin:10px 0;'><strong>⚠️ Sin licitaciones relevantes esta semana.</strong> El historial CSV fue actualizado igual.</div>" if n_hits == 0 else ""}

        {"<table style='width:100%;border-collapse:collapse;background:white;'><thead><tr style='background:#f8f9fa;'><th style='padding:10px;text-align:left;font-size:12px;color:#666;'>ORGANISMO</th><th style='padding:10px;text-align:left;font-size:12px;color:#666;'>LICITACIÓN</th><th style='padding:10px;text-align:left;font-size:12px;color:#666;'>KEYWORDS</th><th style='padding:10px;text-align:left;font-size:12px;color:#666;'>FECHA</th></tr></thead><tbody>" + filas_html + "</tbody></table>" if n_hits > 0 else ""}

        <div style="padding:15px;background:#f8f9fa;border-radius:0 0 8px 8px;margin-top:10px;">
            <p style="color:#888;font-size:12px;margin:0;">
                📁 Ver historial completo en el CSV adjunto.<br>
                🔗 <a href="https://www.comprasestatales.gub.uy">comprasestatales.gub.uy</a>
            </p>
        </div>
    </body></html>
    """


def enviar_email(licitaciones_con_hits: list, total: int):
    """Envía el reporte por email con el CSV adjunto."""
    hoy = datetime.now().strftime("%d/%m/%Y")
    n = len(licitaciones_con_hits)

    asunto = f"📦 Licitaciones ANCAP/UTE — {hoy} — {n} relevante(s) encontrada(s)"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = CONFIG["email_remitente"]
    msg["To"] = CONFIG["email_destino"]

    # Adjuntar CSV
    if os.path.exists(CONFIG["csv_historial"]):
        with open(CONFIG["csv_historial"], "rb") as f:
            adjunto = MIMEBase("application", "octet-stream")
            adjunto.set_payload(f.read())
            encoders.encode_base64(adjunto)
            adjunto.add_header(
                "Content-Disposition",
                f"attachment; filename=historial_licitaciones_{datetime.now().strftime('%Y%m%d')}.csv"
            )
            msg.attach(adjunto)

    # Cuerpo HTML
    html_body = construir_email_html(licitaciones_con_hits, total)
    msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(CONFIG["email_remitente"], CONFIG["gmail_app_password"])
            server.send_message(msg)
        print(f"  📧 Email enviado a {CONFIG['email_destino']}")
    except Exception as e:
        print(f"  ❌ Error enviando email: {e}")
        print("     → Verificá tu Gmail App Password en: myaccount.google.com/apppasswords")


# ─────────────────────────────────────────────
#  🌍  SCRAPING PORTAL COMPLETO (sin filtro de organismo)
# ─────────────────────────────────────────────

def scrape_portal_completo() -> list:
    """
    Descarga el RSS de 'cambios de la última semana' de TODO el portal ARCE
    (sin filtro de organismo) y retorna solo los que contienen keywords.

    La URL se construye dinámicamente con el rango de fechas de los últimos 7 días,
    igual a como lo hace el portal en su link 'Cambios de la última semana'.
    """
    hoy = datetime.now()
    hace_7_dias = hoy - timedelta(days=7)

    desde = hace_7_dias.strftime("%Y-%m-%d+00%%3A00%%3A00")
    hasta = hoy.strftime("%Y-%m-%d+23%%3A59%%3A59")

    url_rss = (
        f"https://www.comprasestatales.gub.uy/consultas/rss/"
        f"tipo-pub/ALL/tipo-fecha/MOD/orden/ORD_MOD/tipo-orden/DESC/"
        f"rango-fecha/{desde}_{hasta}"
    )

    fuente = FUENTE_PORTAL_COMPLETO
    resultados = []

    print(f"\n🌍 Procesando: {fuente['nombre']}")
    print(f"   Período: {hace_7_dias.strftime('%d/%m/%Y')} → {hoy.strftime('%d/%m/%Y')}")

    try:
        r = requests.get(url_rss, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.content, "xml")
        items = soup.find_all("item")

        total_items = len(items)
        hits = 0

        for item in items:
            titulo      = item.find("title").get_text(strip=True)      if item.find("title")       else ""
            descripcion = item.find("description").get_text(strip=True) if item.find("description") else ""
            link        = item.find("link").get_text(strip=True)        if item.find("link")        else ""
            pub_date    = item.find("pubDate").get_text(strip=True)     if item.find("pubDate")     else ""

            texto_completo = f"{titulo} {descripcion}"
            keywords_encontradas = detectar_keywords(texto_completo)

            # Para el portal completo: solo guardamos los que tienen keywords
            if keywords_encontradas:
                hits += 1
                # Intentar extraer el organismo del título (ARCE suele incluirlo)
                organismo_detectado = "VARIOS"
                for org in ["ANCAP", "UTE", "ANTEL", "OSE", "ANP", "ASSE", "BPS", "BROU", "BSE", "INAU"]:
                    if org in titulo.upper() or org in descripcion.upper():
                        organismo_detectado = org
                        break

                resultados.append({
                    "organismo": organismo_detectado,
                    "titulo": titulo,
                    "descripcion": descripcion[:300],
                    "link": link,
                    "fecha_publicacion": pub_date,
                    "keywords": keywords_encontradas,
                    "tiene_hits": True,
                    "fuente": "RSS-PORTAL-COMPLETO",
                })

        print(f"  ✅ Portal completo: {total_items} licitaciones revisadas, {hits} con keywords relevantes")

    except Exception as e:
        print(f"  ❌ Error en portal completo: {e}")

    return resultados


# ─────────────────────────────────────────────
#  🚀  EJECUCIÓN PRINCIPAL
# ─────────────────────────────────────────────

def main():
    print("\n" + "="*60)
    print(f"  MONITOR LICITACIONES ARCE — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("="*60)
    print("  Cobertura: ANCAP · UTE · ANTEL · Portal completo")
    print("="*60)

    cache = cargar_cache()
    todas_las_licitaciones = []

    # ── BLOQUE A: organismos específicos (ANCAP, UTE, ANTEL) ──
    print("\n── BLOQUE A: Organismos específicos ──────────────────")
    for fuente in FUENTES_POR_ORGANISMO:
        print(f"\n🔍 Procesando: {fuente['nombre']}")
        resultados = scrape_rss(fuente)
        todas_las_licitaciones.extend(resultados)

    # ── BLOQUE B: portal completo filtrado por keywords ──
    print("\n── BLOQUE B: Portal completo (todos los organismos) ──")
    resultados_portal = scrape_portal_completo()

    # Evitar duplicar licitaciones que ya aparecieron en ANCAP/UTE/ANTEL
    links_ya_vistos = {l["link"] for l in todas_las_licitaciones}
    nuevos_del_portal = [r for r in resultados_portal if r["link"] not in links_ya_vistos]
    todas_las_licitaciones.extend(nuevos_del_portal)

    # ── Filtrar nuevas (no notificadas en ejecuciones anteriores) ──
    nuevas = []
    for lic in todas_las_licitaciones:
        id_unico = f"{lic['organismo']}_{lic['link']}"
        if id_unico not in cache:
            nuevas.append(lic)
            cache.add(id_unico)

    # Guardar historial completo en CSV
    if todas_las_licitaciones:
        guardar_csv(todas_las_licitaciones)

    # Para el email: licitaciones nuevas con keywords relevantes
    # ANCAP/UTE/ANTEL: todas las nuevas (tienen_hits puede ser False)
    # Portal completo: ya vienen pre-filtradas por keywords
    con_hits = [l for l in nuevas if l["tiene_hits"]]

    print(f"\n{'='*60}")
    print(f"  📊 RESUMEN FINAL")
    print(f"{'='*60}")
    print(f"  Total revisadas esta semana:   {len(todas_las_licitaciones)}")
    print(f"  Nuevas (no vistas antes):      {len(nuevas)}")
    print(f"  Con keywords EPP/Altura/Cap:   {len(con_hits)}")
    print(f"  Del portal completo (nuevas):  {len(nuevos_del_portal)}")

    # Decidir si enviar email
    debe_enviar = True
    if CONFIG["solo_alertar_con_hits"] and len(con_hits) == 0:
        debe_enviar = False
        print("\n  → Sin hits relevantes esta semana.")
        print("  → Cambiá solo_alertar_con_hits=False en CONFIG para recibir reporte siempre.")

    if debe_enviar:
        print("\n📧 Enviando email...")
        enviar_email(con_hits, len(todas_las_licitaciones))

    # Actualizar cache
    guardar_cache(cache)

    print("\n✅ Proceso completado.\n")


if __name__ == "__main__":
    main()
