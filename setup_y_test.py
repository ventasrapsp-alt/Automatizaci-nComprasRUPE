"""
setup_y_test.py — Verificador de instalación y test sin email

Ejecutá esto PRIMERO para verificar que todo funciona
antes de configurar el envío de emails.

Uso:
    python setup_y_test.py
"""

import subprocess
import sys

def instalar_dependencias():
    paquetes = ["requests", "beautifulsoup4", "lxml", "openpyxl"]
    print("📦 Instalando dependencias...")
    for pkg in paquetes:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
    print("   ✅ Dependencias OK\n")

def test_conexion():
    import requests

    urls_test = [
        ("ANCAP en ARCE", "https://www.comprasestatales.gub.uy/consultas/buscar/tipo-pub/VIG/inciso/60/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC"),
        ("UTE en ARCE",   "https://www.comprasestatales.gub.uy/consultas/buscar/tipo-pub/VIG/inciso/61/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC"),
        ("RSS ANCAP",     "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/60/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC"),
        ("RSS UTE",       "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/61/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC"),
    ]

    headers = {"User-Agent": "Mozilla/5.0 (compatible; LicitacionesBot/1.0)"}

    print("🌐 Testeando conectividad a ARCE...")
    for nombre, url in urls_test:
        try:
            r = requests.get(url, headers=headers, timeout=10)
            estado = "✅" if r.status_code == 200 else f"⚠️  HTTP {r.status_code}"
            print(f"   {estado} {nombre}: {len(r.content)} bytes recibidos")
        except Exception as e:
            print(f"   ❌ {nombre}: {e}")

def test_scraping():
    import requests
    from bs4 import BeautifulSoup

    print("\n🔍 Testeando extracción de datos...")

    url_rss = "https://www.comprasestatales.gub.uy/consultas/rss/tipo-pub/VIG/inciso/60/tipo-doc/R/tipo-fecha/ROF/filtro-cat/CAT/orden/ORD_ROF/tipo-orden/DESC"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        r = requests.get(url_rss, headers=headers, timeout=10)
        soup = BeautifulSoup(r.content, "xml")
        items = soup.find_all("item")

        print(f"   ✅ ANCAP RSS: {len(items)} licitaciones encontradas")
        if items:
            primer_titulo = items[0].find("title").get_text(strip=True) if items[0].find("title") else "N/A"
            print(f"   📄 Primera licitación: {primer_titulo[:80]}...")

    except Exception as e:
        print(f"   ⚠️  RSS: {e} — El script usará HTML como fallback")

    print("\n✅ Test completado. Si todo está OK, configurá el email en monitor_licitaciones.py y ejecutalo.")
    print("\n📋 PRÓXIMOS PASOS:")
    print("   1. Abrir monitor_licitaciones.py")
    print("   2. Editar la sección CONFIG con tu email y Gmail App Password")
    print("   3. Ejecutar: python monitor_licitaciones.py")
    print("   4. Configurar ejecución automática (ver README.md)\n")

if __name__ == "__main__":
    instalar_dependencias()
    test_conexion()
    test_scraping()
