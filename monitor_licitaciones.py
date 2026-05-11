"""
╔══════════════════════════════════════════════════════╗
║   MONITOR DE LICITACIONES — ANCAP & UTE URUGUAY     ║
║   Fuente: ARCE (comprasestatales.gub.uy)            ║
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

# ─────────────────────────────────────────────
#  ⚙️  CONFIG — EDITÁ ESTO ANTES DE EJECUTAR
# ─────────────────────────────────────────────
CONFIG = {
    # Email destino (donde llega el reporte)
    "email_destino": "tu@email.com",

    # Cuenta Gmail remitente (necesita App Password de Google)
    "email_remitente": "tubot@gmail.com",
    "gmail_app_password": "xxxx xxxx xxxx xxxx",  # Generá en: myaccount.google.com/apppasswords

    # Archivo CSV donde se acumula el historial
    "csv_historial": "historial_licitaciones.csv",

    # Archivo JSON para no re-alertar licitaciones ya notificadas
    "cache_file": "cache_notificadas.json",

    # Solo alertar si hay al menos una keyword encontrada
    "solo_alertar_con_hits": True,
}

# ─────────────────────────────────────────────
#  🎯  KEYWORDS A DETECTAR
# ─────────────────────────────────────────────
KEYWORDS = [
    # EPP
    "EPP", "equipos de protección", "protección personal",
    "casco", "arnés", "guantes", "calzado de seguridad",
    "ropa de trabajo", "indumentaria de seguridad",
    "gafas de protección", "protección auditiva",
    # Trabajo en altura
    "trabajo en altura", "trabajos en altura",
    "rescate en altura", "líneas de vida",
    "andamios", "plataformas elevadoras",
    # Capacitación / Cursos
    "capacitación", "capacitacion", "curso",
    "formación", "formacion", "entrenamiento",
    "seguridad laboral", "seguridad e higiene",
    "prevención de riesgos", "prevencion de riesgos",
]

# ─────────────────────────────────────────────
#  🌐  FUENTES — URLs DIRECTAS POR ORGANISMO
# ─────────────────────────────────────────────
# ARCE asigna inciso 60 = ANCAP, inciso 61 = UTE
FUENTES = [
    {
        "nombre": "ANCAP — Llamados Vigentes",
        "organismo": "ANCAP",
        "url_html": "https://www.comprasestatales.gub.uy/consultas/buscar/tipo-pub/VIG/inciso/60/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_rss":  "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/60/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_base": "https://www.comprasestatales.gub.uy",
    },
    {
        "nombre": "UTE — Llamados Vigentes",
        "organismo": "UTE",
        "url_html": "https://www.comprasestatales.gub.uy/consultas/buscar/tipo-pub/VIG/inciso/61/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_rss":  "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/61/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC",
        "url_base": "https://www.comprasestatales.gub.uy",
    },
]

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
    """Retorna lista de keywords encontradas en el texto (case-insensitive)."""
    texto_lower = texto.lower()
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
#  🚀  EJECUCIÓN PRINCIPAL
# ─────────────────────────────────────────────

def main():
    print("\n" + "="*55)
    print(f"  MONITOR LICITACIONES — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("="*55)

    cache = cargar_cache()
    todas_las_licitaciones = []

    for fuente in FUENTES:
        print(f"\n🔍 Procesando: {fuente['nombre']}")
        resultados = scrape_rss(fuente)
        todas_las_licitaciones.extend(resultados)

    # Filtrar nuevas (no notificadas antes)
    nuevas = []
    for lic in todas_las_licitaciones:
        id_unico = f"{lic['organismo']}_{lic['link']}"
        if id_unico not in cache:
            nuevas.append(lic)
            cache.add(id_unico)

    # Guardar todo en CSV (historial completo)
    if todas_las_licitaciones:
        guardar_csv(todas_las_licitaciones)

    # Filtrar solo las que tienen keywords relevantes (entre las nuevas)
    con_hits = [l for l in nuevas if l["tiene_hits"]]

    print(f"\n📊 Resumen:")
    print(f"   Total revisadas:      {len(todas_las_licitaciones)}")
    print(f"   Nuevas (no vistas):   {len(nuevas)}")
    print(f"   Con keywords EPP/Alt: {len(con_hits)}")

    # Decidir si enviar email
    debe_enviar = True
    if CONFIG["solo_alertar_con_hits"] and len(con_hits) == 0:
        debe_enviar = False
        print("   → Sin hits relevantes, email no enviado (config: solo_alertar_con_hits=True)")
        print("   → Cambiá a False en CONFIG para recibir reporte igual.")

    if debe_enviar:
        print("\n📧 Enviando email...")
        enviar_email(con_hits, len(todas_las_licitaciones))

    # Actualizar cache
    guardar_cache(cache)

    print("\n✅ Proceso completado.\n")


if __name__ == "__main__":
    main()
